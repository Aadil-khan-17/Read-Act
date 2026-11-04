import os
import re
import sqlite3
import uuid
import json
import csv
from io import BytesIO
from pathlib import Path
from typing import Literal

import fitz
import httpx
import pytesseract
from docx import Document
from openpyxl import load_workbook
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from PIL import Image
from pydantic import BaseModel

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
UPLOAD_DIR = DATA_DIR / "uploads"
OUTPUT_DIR = DATA_DIR / "outputs"
DB_PATH = DATA_DIR / "redactions.db"
DATA_DIR.mkdir(exist_ok=True)
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

app = FastAPI(title="Redactly", version="0.1.0")

for candidate in (
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
):
    if Path(candidate).exists():
        pytesseract.pytesseract.tesseract_cmd = candidate
        break


class Detection(BaseModel):
    id: str
    page: int
    category: str
    text: str
    bbox: list[float]
    confidence: float
    source: Literal["deterministic", "ai", "layout"]
    risk: Literal["high", "medium", "low"]
    decision: Literal["proposed", "approved", "rejected"] = "proposed"


class AnalyzeResponse(BaseModel):
    document_id: str
    filename: str
    pages: int
    text: str
    detections: list[Detection]


class DecisionRequest(BaseModel):
    document_id: str
    detections: list[Detection]


class Policy(BaseModel):
    pii: bool = True
    financial: bool = True
    credentials: bool = True
    demographic: bool = False


class PolicyCreate(BaseModel):
    name: str
    description: str = ""
    prompt: str
    enabled: bool = True


class PolicyResponse(PolicyCreate):
    id: str
    created_at: str | None = None


class InstructionRequest(BaseModel):
    document_id: str
    prompt: str
    policy_id: str | None = None
    keep_existing: bool = True


PATTERNS = {
    "email": (r"(?<![\w.-])[\w.+-]+@[\w-]+(?:\.[\w-]+)+(?![\w.-])", "medium", 0.99),
    "phone": (
        r"(?<!\d)(?:\+?\s*(?:1[\s.-]?)?(?:\(\d{3}\)|\d{3})[\s.-]?\d{3}[\s.-]?\d{4}(?:\s*(?:x|ext\.?)\s*\d{1,6})?|(?:\+91[\s.-]?)?[6-9]\d{9})(?!\d)",
        "medium",
        0.95,
    ),
    "aadhaar": (r"(?<!\d)\d{4}[\s-]?\d{4}[\s-]?\d{4}(?!\d)", "high", 0.98),
    "pan": (r"(?<![A-Z0-9])[A-Z]{5}\d{4}[A-Z](?![A-Z0-9])", "high", 0.98),
    "credit_card": (r"(?<!\d)(?:\d[ -]*){13,19}(?!\d)", "high", 0.92),
    "ip_address": (r"(?<![\d.])(?:\d{1,3}\.){3}\d{1,3}(?![\d.])", "low", 0.95),
    "passport": (r"(?<![A-Z0-9])[A-Z][0-9]{7}(?![A-Z0-9])", "high", 0.90),
}

SENSITIVE_LABELS = re.compile(
    r"\b(password|secret|api[_ -]?key|token|access[_ -]?key)\b",
    re.IGNORECASE,
)


def init_db() -> None:
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """CREATE TABLE IF NOT EXISTS documents (
                id TEXT PRIMARY KEY, filename TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )"""
        )
        conn.execute(
            """CREATE TABLE IF NOT EXISTS decisions (
                id TEXT PRIMARY KEY, document_id TEXT, category TEXT, text_hash TEXT,
                decision TEXT, source TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )"""
        )
        conn.execute(
            """CREATE TABLE IF NOT EXISTS policies (
                id TEXT PRIMARY KEY, name TEXT UNIQUE, description TEXT, prompt TEXT,
                enabled INTEGER DEFAULT 1, created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )"""
        )
        conn.execute(
            """CREATE TABLE IF NOT EXISTS prompt_history (
                id TEXT PRIMARY KEY, document_id TEXT, policy_id TEXT, prompt TEXT,
                response_json TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )"""
        )


init_db()


def get_pdf_text_and_words(path: Path, page_count: int | None = None):
    doc = fitz.open(path)
    pages = []
    full_text = []
    for page_index in range(doc.page_count if page_count is None else page_count):
        page = doc.load_page(page_index)
        text = page.get_text("text").strip()
        words = page.get_text("words")
        source = "native"

        if len(text) < 20:
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            image = Image.open(BytesIO(pix.tobytes("png")))
            text = pytesseract.image_to_string(image)
            data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DICT)
            words = []
            scale = 0.5
            for i, word in enumerate(data["text"]):
                if word.strip():
                    words.append((
                        data["left"][i] * scale,
                        data["top"][i] * scale,
                        (data["left"][i] + data["width"][i]) * scale,
                        (data["top"][i] + data["height"][i]) * scale,
                        word,
                        data["conf"][i] / 100,
                        0,
                        0,
                        0,
                    ))
            source = "ocr"
        full_text.append(f"--- PAGE {page_index + 1} ---\n{text}")
        pages.append({"page": page_index + 1, "text": text, "words": words, "source": source})
    doc.close()
    return pages, "\n".join(full_text)


def image_to_pdf(path: Path) -> Path:
    image = Image.open(path)
    if image.mode not in ("RGB", "L"):
        image = image.convert("RGB")
    converted = path.with_suffix(".pdf")
    image.save(converted, "PDF", resolution=150.0)
    return converted


def extract_tabular_or_office(path: Path, suffix: str):
    pages = []
    full_text = []

    if suffix == ".csv":
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            sample = handle.read(4096)
            handle.seek(0)
            dialect = csv.Sniffer().sniff(sample) if sample else csv.excel
            rows = list(csv.reader(handle, dialect))
        text = "\n".join("\t".join(row) for row in rows)
        pages.append({"page": 1, "text": text, "words": [], "source": "csv"})
        full_text.append("--- SHEET ROWS ---\n" + text)

    elif suffix in {".xlsx", ".xlsm"}:
        workbook = load_workbook(path, data_only=False)
        for sheet_index, sheet in enumerate(workbook.worksheets, start=1):
            lines = []
            for row in sheet.iter_rows():
                values = ["" if cell.value is None else str(cell.value) for cell in row]
                lines.append("\t".join(values))
            text = "\n".join(lines)
            pages.append({"page": sheet_index, "text": text, "words": [], "source": f"xlsx:{sheet.title}"})
            full_text.append(f"--- SHEET: {sheet.title} ---\n{text}")
        workbook.close()

    elif suffix == ".docx":
        document = Document(path)
        lines = []
        for paragraph in document.paragraphs:
            if paragraph.text.strip():
                lines.append(paragraph.text)
        for table in document.tables:
            for row in table.rows:
                lines.append("\t".join(cell.text for cell in row.cells))
        for section in document.sections:
            lines.extend(paragraph.text for paragraph in section.header.paragraphs if paragraph.text.strip())
            lines.extend(paragraph.text for paragraph in section.footer.paragraphs if paragraph.text.strip())
        text = "\n".join(lines)
        pages.append({"page": 1, "text": text, "words": [], "source": "docx"})
        full_text.append("--- DOCUMENT ---\n" + text)

    return pages, "\n".join(full_text)


def deterministic_detections(pages: list[dict]) -> list[Detection]:
    detections = []
    for page in pages:
        text = page["text"]
        for category, (pattern, risk, base_confidence) in PATTERNS.items():
            for match in re.finditer(pattern, text):
                value = match.group(0)
                bbox = find_word_bbox(page["words"], value)
                detections.append(Detection(
                    id=str(uuid.uuid4()),
                    page=page["page"],
                    category=category,
                    text=value,
                    bbox=bbox or [0, 0, 0, 0],
                    confidence=base_confidence,
                    source="deterministic",
                    risk=risk,
                ))

        for match in SENSITIVE_LABELS.finditer(text):
            bbox = find_word_bbox(page["words"], match.group(0))
            detections.append(Detection(
                id=str(uuid.uuid4()),
                page=page["page"],
                category="credential_label",
                text=match.group(0),
                bbox=bbox or [0, 0, 0, 0],
                confidence=0.86,
                source="layout",
                risk="high",
            ))
    return deduplicate(detections)


def find_word_bbox(words, value: str) -> list[float] | None:
    clean = re.sub(r"\s+", "", value)
    lines: dict[int, list] = {}
    for word in words:
        x0, y0, x1, y1 = word[0], word[1], word[2], word[3]
        token = re.sub(r"\s+", "", str(word[4]))
        line_key = round(y0 / 4) * 4
        lines.setdefault(line_key, []).append((x0, y0, x1, y1, token))

    for line in lines.values():
        line.sort(key=lambda item: item[0])
        for start in range(len(line)):
            accumulated = ""
            xs: list[float] = []
            ys: list[float] = []
            x2s: list[float] = []
            y2s: list[float] = []
            for item in line[start:]:
                x0, y0, x1, y1, token = item
                accumulated += token
                xs.append(x0)
                ys.append(y0)
                x2s.append(x1)
                y2s.append(y1)
                if clean in accumulated:
                    return [min(xs), min(ys), max(x2s), max(y2s)]
                if len(accumulated) > len(clean) + 20:
                    break
    return None


def deduplicate(items: list[Detection]) -> list[Detection]:
    unique = {}
    for item in items:
        key = (item.page, item.category, item.text, tuple(item.bbox))
        if key not in unique or item.confidence > unique[key].confidence:
            unique[key] = item
    return list(unique.values())


def labeled_instruction_findings(pages: list[dict], instruction: str) -> list[Detection]:
    instruction_lower = instruction.lower()
    wants_personal_info = any(
        term in instruction_lower
        for term in ("personal", "pii", "name", "phone", "email", "contact", "individual")
    )
    if not wants_personal_info:
        return []

    findings: list[Detection] = []
    patterns = {
        "person_name": (r"(?im)^\s*name\s*:\s*([^\r\n]+)", "high", 0.94),
        "phone": (
            r"(?im)^\s*(?:phone|telephone|mobile|contact(?:\s+number)?)\s*:\s*([^\r\n]+)",
            "medium",
            0.94,
        ),
    }
    for page in pages:
        for category, (pattern, risk, confidence) in patterns.items():
            for match in re.finditer(pattern, page["text"]):
                value = match.group(1).strip()
                if not value:
                    continue
                findings.append(Detection(
                    id=str(uuid.uuid4()),
                    page=page["page"],
                    category=category,
                    text=value,
                    bbox=find_word_bbox(page["words"], value) or [0, 0, 0, 0],
                    confidence=confidence,
                    source="layout",
                    risk=risk,
                ))
    return findings


def literal_instruction_findings(pages: list[dict], instruction: str) -> list[Detection]:
    literals: list[str] = []

    for match in re.finditer(r"""["'`]([^"'`\r\n]{1,200})["'`]""", instruction):
        literals.append(match.group(1).strip())

    for match in re.finditer(
        r"""(?i)\b(?:word|term|phrase)\s+(?:called\s+|named\s+)?["']?([^\s"'`,]{2,200})["']?""",
        instruction,
    ):
        literals.append(match.group(1).strip(" ._-"))

    findings: list[Detection] = []
    for literal in dict.fromkeys(literals):
        if len(literal) < 2:
            continue
        for page in pages:
            start = 0
            while True:
                index = page["text"].lower().find(literal.lower(), start)
                if index < 0:
                    break
                exact = page["text"][index:index + len(literal)]
                findings.append(Detection(
                    id=str(uuid.uuid4()),
                    page=page["page"],
                    category="custom_sensitive",
                    text=exact,
                    bbox=find_word_bbox(page["words"], exact) or [0, 0, 0, 0],
                    confidence=1.0,
                    source="layout",
                    risk="high",
                ))
                start = index + len(literal)
    return findings


DOCUMENT_PAGES: dict[str, list[dict]] = {}


def locate_ai_finding(
    pages: list[dict],
    value: str,
    category: str,
    confidence: float,
    risk: str,
    expected_page: int | None = None,
):
    if not isinstance(value, str) or not value.strip():
        return []

    findings = []
    needle = value.strip()
    for page in pages:
        if expected_page is not None and page["page"] != expected_page:
            continue
        text = page["text"]
        start = 0
        while True:
            index = text.lower().find(needle.lower(), start)
            if index < 0:
                break
            exact = text[index:index + len(needle)]
            findings.append(Detection(
                id=str(uuid.uuid4()),
                page=page["page"],
                category=category[:80] or "custom_instruction",
                text=exact,
                bbox=find_word_bbox(page["words"], exact) or [0, 0, 0, 0],
                confidence=max(0.0, min(float(confidence), 1.0)),
                source="ai",
                risk=risk if risk in {"high", "medium", "low"} else "high",
            ))
            start = index + len(needle)
    return findings


def call_openai_compatible(text: str, instruction: str):
    base_url = os.getenv("LLM_API_BASE", "http://127.0.0.1:11434/v1")
    api_key = os.getenv("LLM_API_KEY")
    model = os.getenv("LLM_MODEL", "qwen2.5:3b")

    if not base_url:
        raise HTTPException(
            503,
            "LLM is not configured. Set LLM_API_BASE, optionally LLM_API_KEY and LLM_MODEL.",
        )

    system_prompt = """You are a conservative document redaction analyst for a regulated pharmaceutical company.
Use only the supplied document text. Return valid JSON only.
Every item must contain an exact contiguous quote copied verbatim from the document.
When the instruction requests personal information, names, phone numbers, emails, addresses, or identifiers, include the exact value after each label (for example, a Name label's value, not the label itself).
Do not infer unprovided values, invent text, summarize, or redact broad sections.
Classify each item as one of: patient_name, employee_name, clinician_name, sponsor_name,
study_identifier, protocol_identifier, site_name, address, email, phone, date_of_birth,
medical_record, lab_result, diagnosis, dose, adverse_event, batch_number, document_number,
financial, credential, custom_sensitive, or other.
For each item return: {"text":"exact quote","page":1,"category":"...","confidence":0.0,"risk":"low|medium|high","reason":"brief evidence"}.
Return an object with exactly: {"items":[...]}."""

    payload = {
        "model": model,
        "temperature": 0,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": f"REDACTION INSTRUCTION:\n{instruction}\n\nDOCUMENT TEXT:\n{text[:120000]}"},
        ],
        "response_format": {"type": "json_object"},
    }
    headers = {"Content-Type": "application/json"}
    if api_key:
        headers["Authorization"] = "Bearer " + api_key

    url = base_url.rstrip("/") + "/chat/completions"
    try:
        with httpx.Client(timeout=60) as client:
            response = client.post(url, headers=headers, json=payload)
            response.raise_for_status()
            content = response.json()["choices"][0]["message"]["content"]
    except httpx.HTTPStatusError as exc:
        raise HTTPException(502, f"LLM provider returned HTTP {exc.response.status_code}") from exc
    except Exception as exc:
        raise HTTPException(502, f"LLM request failed: {exc}") from exc

    content = content.strip()
    if content.startswith("```"):
        content = re.sub(r"^```(?:json)?\s*", "", content, flags=re.IGNORECASE)
        content = re.sub(r"\s*```$", "", content)
    try:
        parsed = json.loads(content)
    except json.JSONDecodeError as exc:
        raise HTTPException(502, "LLM returned invalid JSON") from exc
    return parsed.get("items", []) if isinstance(parsed, dict) else []


@app.post("/api/analyze", response_model=AnalyzeResponse)
async def analyze(file: UploadFile = File(...)):
    allowed = {".pdf", ".png", ".jpg", ".jpeg", ".tif", ".tiff", ".csv", ".xlsx", ".xlsm", ".docx"}
    suffix = Path(file.filename or "upload").suffix.lower()
    if suffix not in allowed:
        raise HTTPException(400, "Supported formats: PDF, PNG, JPG, TIFF, CSV, XLSX, XLSM, DOCX")

    document_id = str(uuid.uuid4())
    source_path = UPLOAD_DIR / f"{document_id}{suffix}"
    source_path.write_bytes(await file.read())

    if suffix in {".png", ".jpg", ".jpeg", ".tif", ".tiff"}:
        source_path = image_to_pdf(source_path)

    if suffix in {".csv", ".xlsx", ".xlsm", ".docx"}:
        pages, text = extract_tabular_or_office(source_path, suffix)
    else:
        pages, text = get_pdf_text_and_words(source_path)
    detections = deterministic_detections(pages)
    DOCUMENT_PAGES[document_id] = pages
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            "INSERT INTO documents (id, filename) VALUES (?, ?)",
            (document_id, file.filename),
        )

    return AnalyzeResponse(
        document_id=document_id,
        filename=file.filename or "upload",
        pages=len(pages),
        text=text,
        detections=detections,
    )


@app.get("/api/policies")
async def list_policies():
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            "SELECT id, name, description, prompt, enabled, created_at FROM policies ORDER BY created_at DESC"
        ).fetchall()
    return {"policies": [
        PolicyResponse(
            id=row["id"],
            name=row["name"],
            description=row["description"],
            prompt=row["prompt"],
            enabled=bool(row["enabled"]),
            created_at=row["created_at"],
        ).model_dump()
        for row in rows
    ]}


@app.post("/api/policies", response_model=PolicyResponse)
async def create_policy(policy: PolicyCreate):
    name = policy.name.strip()
    if not name:
        raise HTTPException(400, "Policy name is required")
    policy_id = str(uuid.uuid4())
    try:
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute(
                "INSERT INTO policies (id, name, description, prompt, enabled) VALUES (?, ?, ?, ?, ?)",
                (policy_id, name, policy.description.strip(), policy.prompt.strip(), int(policy.enabled)),
            )
    except sqlite3.IntegrityError as exc:
        raise HTTPException(409, "A policy with this name already exists") from exc
    return PolicyResponse(
        id=policy_id,
        name=name,
        description=policy.description.strip(),
        prompt=policy.prompt.strip(),
        enabled=policy.enabled,
    )


@app.delete("/api/policies/{policy_id}")
async def delete_policy(policy_id: str):
    with sqlite3.connect(DB_PATH) as conn:
        result = conn.execute("DELETE FROM policies WHERE id = ?", (policy_id,))
        if result.rowcount == 0:
            raise HTTPException(404, "Policy not found")
    return {"deleted": True}


@app.post("/api/instruct")
async def instruct(request: InstructionRequest):
    pages = DOCUMENT_PAGES.get(request.document_id)
    if not pages:
        raise HTTPException(404, "Document session not found; analyze it again")

    instruction = request.prompt.strip()
    policy = None
    if request.policy_id:
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            policy = conn.execute(
                "SELECT id, name, description, prompt FROM policies WHERE id = ? AND enabled = 1",
                (request.policy_id,),
            ).fetchone()
        if policy:
            instruction = (
                f"Mode: {policy['name']}\n"
                f"Description: {policy['description']}\n"
                f"Instructions: {policy['prompt']}\n"
                f"User iteration: {request.prompt}"
            )
    if not instruction:
        raise HTTPException(400, "Instruction cannot be empty")

    text = "\n".join(
        f"--- PAGE {page['page']} ---\n{page['text']}"
        for page in pages
    )
    raw_items = call_openai_compatible(text, instruction)
    ai_findings: list[Detection] = []
    for item in raw_items:
        if not isinstance(item, dict):
            continue
        try:
            expected_page = int(item.get("page")) if item.get("page") is not None else None
        except (TypeError, ValueError):
            expected_page = None
        ai_findings.extend(locate_ai_finding(
            pages,
            item.get("text", ""),
            item.get("category", "custom_sensitive"),
            item.get("confidence", 0.7),
            item.get("risk", "high"),
            expected_page,
        ))

    ai_findings.extend(labeled_instruction_findings(pages, instruction))
    ai_findings.extend(literal_instruction_findings(pages, instruction))
    ai_findings = deduplicate(ai_findings)
    history_id = str(uuid.uuid4())
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            "INSERT INTO prompt_history (id, document_id, policy_id, prompt, response_json) VALUES (?, ?, ?, ?, ?)",
            (
                history_id,
                request.document_id,
                request.policy_id,
                instruction,
                json.dumps([finding.model_dump() for finding in ai_findings]),
            ),
        )

    return {
        "document_id": request.document_id,
        "instruction_id": history_id,
        "llm_configured": True,
        "policy": policy["name"] if policy else None,
        "detections": ai_findings,
    }


@app.get("/api/session/{document_id}")
async def get_session(document_id: str):
    if document_id not in DOCUMENT_PAGES:
        raise HTTPException(404, "Document session not found; analyze it again")
    pages = DOCUMENT_PAGES[document_id]
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        history = conn.execute(
            """SELECT id, policy_id, prompt, response_json, created_at
               FROM prompt_history WHERE document_id = ? ORDER BY created_at DESC""",
            (document_id,),
        ).fetchall()
        decisions = conn.execute(
            """SELECT category, text_hash, decision, source, created_at
               FROM decisions WHERE document_id = ? ORDER BY created_at DESC""",
            (document_id,),
        ).fetchall()
    return {
        "document_id": document_id,
        "pages": len(pages),
        "history": [dict(row) for row in history],
        "decisions": [dict(row) for row in decisions],
    }


@app.post("/api/verify/{document_id}")
async def verify_export(document_id: str, request: DecisionRequest):
    output_record = next(
        (path for path in OUTPUT_DIR.iterdir() if path.name.startswith(document_id)),
        None,
    )
    if not output_record:
        raise HTTPException(404, "Redacted export not found")

    suffix = output_record.suffix.lower()
    if suffix == ".pdf":
        return verify_pdf_export(document_id, output_record, request.detections)
    if suffix == ".csv":
        text = output_record.read_text(encoding="utf-8-sig")
        return build_verification(document_id, text, request.detections)
    if suffix in {".xlsx", ".xlsm"}:
        workbook = load_workbook(output_record, data_only=False)
        text = "\n".join(
            str(cell.value) for sheet in workbook.worksheets
            for row in sheet.iter_rows() for cell in row if cell.value is not None
        )
        workbook.close()
        return build_verification(document_id, text, request.detections)
    if suffix == ".docx":
        document = Document(output_record)
        text_parts = (
            [p.text for p in document.paragraphs]
            + [cell.text for table in document.tables for row in table.rows for cell in row.cells]
        )
        for section in document.sections:
            text_parts.extend(p.text for p in section.header.paragraphs)
            text_parts.extend(p.text for p in section.footer.paragraphs)
        text = "\n".join(text_parts)
        return build_verification(document_id, text, request.detections)
    raise HTTPException(404, "Unsupported redacted format")


def verify_pdf_export(document_id: str, output_path: Path, detections: list[Detection]):
    approved = [finding.text for finding in detections if finding.decision == "approved"]
    remaining: list[dict] = []
    with fitz.open(output_path) as doc:
        for finding in detections:
            if finding.decision != "approved":
                continue
            if 1 <= finding.page <= doc.page_count:
                page_text = doc.load_page(finding.page - 1).get_text("text")
                if finding.text.lower() in page_text.lower():
                    remaining.append({"text": finding.text, "category": finding.category, "page": finding.page})
    return {"document_id": document_id, "approved_count": len(approved), "remaining_count": len(remaining), "verified": not remaining, "remaining": remaining}


def build_verification(document_id: str, text: str, detections: list[Detection]):
    approved = [finding.text for finding in detections if finding.decision == "approved"]
    remaining = [
        {"text": finding.text, "category": finding.category, "page": finding.page}
        for finding in detections
        if finding.decision == "approved" and finding.text.lower() in text.lower()
    ]
    return {"document_id": document_id, "approved_count": len(approved), "remaining_count": len(remaining), "verified": not remaining, "remaining": remaining}


@app.post("/api/export")
async def export(request: DecisionRequest):
    doc_records = [p for p in UPLOAD_DIR.iterdir() if p.name.startswith(request.document_id)]
    doc_record = next((p for p in doc_records if p.suffix.lower() == ".pdf"), doc_records[0] if doc_records else None)
    if not doc_record:
        raise HTTPException(404, "Document not found")

    approved_values = [finding.text for finding in request.detections if finding.decision == "approved"]
    suffix = doc_record.suffix.lower()

    if suffix == ".csv":
        output_path = OUTPUT_DIR / f"{request.document_id}-redacted.csv"
        with doc_record.open("r", encoding="utf-8-sig", newline="") as handle:
            sample = handle.read(4096)
            handle.seek(0)
            dialect = csv.Sniffer().sniff(sample) if sample else csv.excel
            rows = list(csv.reader(handle, dialect))
        redacted_rows = [[replace_values(value, approved_values) for value in row] for row in rows]
        with output_path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle, dialect)
            writer.writerows(redacted_rows)
        return native_export_response(request, doc_record, output_path, "text/csv", ".csv")

    if suffix in {".xlsx", ".xlsm"}:
        output_path = OUTPUT_DIR / f"{request.document_id}-redacted{suffix}"
        workbook = load_workbook(doc_record, data_only=False)
        for sheet in workbook.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and not cell.value.startswith("="):
                        cell.value = replace_values(cell.value, approved_values)
        workbook.save(output_path)
        workbook.close()
        media_type = "application/vnd.ms-excel.sheet.macroEnabled.12" if suffix == ".xlsm" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        return native_export_response(request, doc_record, output_path, media_type, suffix)

    if suffix == ".docx":
        output_path = OUTPUT_DIR / f"{request.document_id}-redacted.docx"
        document = Document(doc_record)
        redact_docx_paragraphs(document.paragraphs, approved_values)
        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    redact_docx_paragraphs(cell.paragraphs, approved_values)
        for section in document.sections:
            redact_docx_paragraphs(section.header.paragraphs, approved_values)
            redact_docx_paragraphs(section.footer.paragraphs, approved_values)
        document.save(output_path)
        return native_export_response(
            request,
            doc_record,
            output_path,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".docx",
        )

    if suffix == ".pdf":
        output_path = OUTPUT_DIR / f"{request.document_id}-redacted.pdf"
        return pdf_export_response(request, doc_record, output_path)

    if suffix in {".png", ".jpg", ".jpeg", ".tif", ".tiff"}:
        raise HTTPException(400, "Raster uploads are exported through the generated PDF; re-analyze after conversion")
    raise HTTPException(400, f"Unsupported export format: {suffix}")


def replace_values(value: str, approved_values: list[str]) -> str:
    for approved in sorted(set(approved_values), key=len, reverse=True):
        value = value.replace(approved, "[REDACTED]")
    return value


def redact_docx_paragraphs(paragraphs, approved_values: list[str]) -> None:
    for paragraph in paragraphs:
        original = paragraph.text
        redacted = replace_values(original, approved_values)
        if redacted == original:
            continue
        if paragraph.runs:
            first_run = paragraph.runs[0]
            first_run.text = redacted
            for run in paragraph.runs[1:]:
                run.text = ""
        else:
            paragraph.add_run(redacted)


def record_export_decisions(request: DecisionRequest) -> None:
    with sqlite3.connect(DB_PATH) as conn:
        for finding in request.detections:
            if finding.decision != "proposed":
                conn.execute(
                    "INSERT INTO decisions (id, document_id, category, text_hash, decision, source) VALUES (?, ?, ?, ?, ?, ?)",
                    (
                        str(uuid.uuid4()),
                        request.document_id,
                        finding.category,
                        str(abs(hash(finding.text))),
                        finding.decision,
                        finding.source,
                    ),
                )


def native_export_response(request: DecisionRequest, doc_record: Path, output_path: Path, media_type: str, suffix: str):
    record_export_decisions(request)
    return FileResponse(output_path, filename=f"{request.document_id}-redacted{suffix}", media_type=media_type)


def pdf_export_response(request: DecisionRequest, doc_record: Path, output_path: Path):
    doc = fitz.open(doc_record)
    for finding in request.detections:
        if finding.decision != "approved":
            continue
        page = doc.load_page(finding.page - 1)
        rect = fitz.Rect(finding.bbox)
        if rect.is_empty or rect.width <= 0 or rect.height <= 0:
            continue
        page.add_redact_annot(rect, fill=(0, 0, 0))
    for page in doc:
        page.apply_redactions()
    doc.save(output_path)
    doc.close()
    record_export_decisions(request)
    return FileResponse(output_path, filename=f"{request.document_id}-redacted.pdf", media_type="application/pdf")


app.mount("/", StaticFiles(directory=ROOT / "static", html=True), name="static")
