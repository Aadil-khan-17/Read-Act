# Read-Act

A hackathon prototype for hybrid deterministic and AI-assisted document redaction.

## Architecture

- **Ingestion**: PDF, TIFF, PNG, JPEG, CSV, XLSX, XLSM, and DOCX uploads.
- **OCR adapter**: Tesseract for local scanned-document text extraction.
- **Evidence graph**: Text spans with page and bounding-box evidence.
- **Deterministic detector**: High-confidence rules for email, phone, Aadhaar, PAN, passport, credit card, IP, and names matched from an allow/deny policy.
- **AI adjudicator**: Optional OpenAI-compatible endpoint for ambiguous context; every AI decision is restricted by evidence spans.
- **Safety gate**: High-risk findings require human review; low-risk deterministic findings may auto-apply.
- **Audit trail**: SQLite records document, findings, decisions, and redaction policy.

The UI supports upload, preview of extracted evidence, policy selection, applying redactions, and downloading a redacted copy.

## Run locally

```powershell
python -m pip install -r requirements.txt
python -m uvicorn app.main:app --reload
```

Open `http://127.0.0.1:8000`.

## Configure an LLM

The backend uses an OpenAI-compatible chat-completions endpoint. Configure it with environment variables:

```powershell
$env:LLM_API_BASE = "https://your-provider.example/v1"
$env:LLM_API_KEY = "your-key"
$env:LLM_MODEL = "your-model-name"
```

For an Ollama endpoint, use a base URL such as `http://localhost:11434/v1`.

The LLM receives document text and a user instruction. It is required to return exact quoted spans and page numbers.
Those spans are located against OCR/native word coordinates before becoming review findings. Without an exact text match,
the backend discards the finding rather than applying an ungrounded redaction.

## Safety and verification

- Every approved finding is matched to document evidence before export.
- PDF redactions are applied as PyMuPDF redaction annotations and permanently removed during saving.
- CSV values are replaced while retaining the detected CSV dialect.
- XLSX/XLSM cell values are replaced while preserving workbook structure and styles; formulas are not modified.
- DOCX text is replaced in body paragraphs, tables, headers, and footers.
- The exported PDF is reopened and checked for remaining occurrences of approved values.
- Approved/rejected decisions and prompt history are retained in SQLite.
- LLM output is isolated behind exact-string grounding; the model cannot directly create PDF coordinates.
- The original uploaded document is not overwritten.

## Current format scope

The prototype supports PDF, common raster images, CSV, XLSX, XLSM, and DOCX. Native Office exports preserve the source
structure and replace approved exact text values. Enterprise deployment should add encrypted storage, authentication,
retention controls, richer format fidelity, and versioned document objects.
