const fileInput = document.querySelector("#file");
const dropzone = document.querySelector("#dropzone");
const analyzeButton = document.querySelector("#analyze");
const exportButton = document.querySelector("#export");
const status = document.querySelector("#status");
let result = null;

loadPolicies();

fileInput.addEventListener("change", () => {
  analyzeButton.disabled = !fileInput.files.length;
});

for (const eventName of ["dragenter", "dragover"]) {
  dropzone.addEventListener(eventName, event => {
    event.preventDefault();
    dropzone.classList.add("dragging");
  });
}
for (const eventName of ["dragleave", "drop"]) {
  dropzone.addEventListener(eventName, event => {
    event.preventDefault();
    dropzone.classList.remove("dragging");
  });
}
dropzone.addEventListener("drop", event => {
  fileInput.files = event.dataTransfer.files;
  analyzeButton.disabled = !fileInput.files.length;
});

analyzeButton.addEventListener("click", async () => {
  if (!fileInput.files.length) return;
  status.textContent = "Analyzing...";
  analyzeButton.disabled = true;
  const body = new FormData();
  body.append("file", fileInput.files[0]);
  try {
    const response = await fetch("/api/analyze", { method: "POST", body });
    if (!response.ok) throw new Error(await response.text());
    result = await response.json();
    document.querySelector("#text").textContent = result.text;
    document.querySelector("#pageCount").textContent = `${result.pages} page${result.pages === 1 ? "" : "s"}`;
    renderFindings();
    document.querySelector("#instruct").disabled = false;
    status.textContent = "Review needed";
  } catch (error) {
    status.textContent = "Analysis failed";
    alert(error.message);
  } finally {
    analyzeButton.disabled = false;
  }
});

function renderFindings() {
  const container = document.querySelector("#findings");
  container.innerHTML = "";
  document.querySelector("#findingCount").textContent = result.detections.length;
  result.detections.forEach((finding, index) => {
    const card = document.createElement("div");
    card.className = `finding ${finding.decision === "approved" ? "approved" : finding.decision === "rejected" ? "rejected" : ""}`;
    card.innerHTML = `
      <strong>${finding.category}</strong>
      <p>${escapeHtml(finding.text)}</p>
      <div class="meta">Page ${finding.page} · ${(finding.confidence * 100).toFixed(0)}% · ${finding.risk} risk · ${finding.source}</div>
      <button class="${finding.decision === "approved" ? "active" : ""}" data-action="approve" data-index="${index}">Approve</button>
      <button class="reject ${finding.decision === "rejected" ? "active" : ""}" data-action="reject" data-index="${index}">Reject</button>
    `;
    container.appendChild(card);
  });
  exportButton.disabled = result.detections.length === 0;
}

document.querySelector("#findings").addEventListener("click", event => {
  const button = event.target.closest("button");
  if (!button || !result) return;
  const finding = result.detections[Number(button.dataset.index)];
  finding.decision = button.dataset.action === "approve" ? "approved" : "rejected";
  const card = button.parentElement;
  card.classList.toggle("approved", finding.decision === "approved");
  card.classList.toggle("rejected", finding.decision === "rejected");
  card.querySelector('[data-action="approve"]').classList.toggle("active", finding.decision === "approved");
  card.querySelector('[data-action="reject"]').classList.toggle("active", finding.decision === "rejected");
  status.textContent = finding.decision === "approved"
    ? `${finding.category} approved for redaction`
    : `${finding.category} rejected`;
});

exportButton.addEventListener("click", async () => {
  status.textContent = "Creating redacted document...";
  const exportBody = JSON.stringify({ document_id: result.document_id, detections: result.detections });
  const response = await fetch("/api/export", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: exportBody,
  });
  if (!response.ok) {
    status.textContent = "Export failed";
    alert(await response.text());
    return;
  }
  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  const originalName = result.filename || "document";
  const dotIndex = originalName.lastIndexOf(".");
  const baseName = dotIndex > 0 ? originalName.slice(0, dotIndex) : originalName;
  const extension = dotIndex > 0 ? originalName.slice(dotIndex).toLowerCase() : ".pdf";
  anchor.download = `${baseName}-redacted${extension}`;
  anchor.click();
  URL.revokeObjectURL(url);
  const verification = await fetch(`/api/verify/${result.document_id}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: exportBody,
  });
  if (verification.ok) {
    const report = await verification.json();
    status.textContent = report.verified
      ? `Verified: ${report.approved_count} approved redaction${report.approved_count === 1 ? "" : "s"} removed`
      : `Warning: ${report.remaining_count} approved value${report.remaining_count === 1 ? "" : "s"} may remain`;
  } else {
    status.textContent = "Export complete; verification unavailable";
  }
});

document.querySelector("#instruct").addEventListener("click", async () => {
  if (!result) return;
  const prompt = document.querySelector("#instruction").value.trim();
  if (!prompt) {
    alert("Enter an instruction describing what should be redacted.");
    return;
  }
  const button = document.querySelector("#instruct");
  button.disabled = true;
  status.textContent = "AI is interpreting...";
  try {
    const response = await fetch("/api/instruct", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        document_id: result.document_id,
        prompt,
        policy_id: document.querySelector("#policySelect").value || null,
        keep_existing: true,
      }),
    });
    if (!response.ok) throw new Error(await response.text());
    const data = await response.json();
    result.detections = mergeFindings(result.detections, data.detections);
    renderFindings();
    status.textContent = data.detections.length
      ? `${data.detections.length} new AI finding${data.detections.length === 1 ? "" : "s"}`
      : "No new exact evidence found";
  } catch (error) {
    status.textContent = "Instruction failed";
    alert(error.message);
  } finally {
    button.disabled = false;
  }
});

document.querySelector("#savePolicy").addEventListener("click", () => {
  document.querySelector("#saveRow").classList.add("visible");
});
document.querySelector("#cancelSave").addEventListener("click", () => {
  document.querySelector("#saveRow").classList.remove("visible");
});
document.querySelector("#confirmSave").addEventListener("click", async () => {
  const name = document.querySelector("#policyName").value.trim();
  const prompt = document.querySelector("#instruction").value.trim();
  if (!name || !prompt) {
    alert("A mode name and instruction are required.");
    return;
  }
  const response = await fetch("/api/policies", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      name,
      prompt,
      description: document.querySelector("#policyDescription").value.trim(),
    }),
  });
  if (!response.ok) {
    alert(await response.text());
    return;
  }
  document.querySelector("#saveRow").classList.remove("visible");
  document.querySelector("#policyName").value = "";
  document.querySelector("#policyDescription").value = "";
  await loadPolicies();
  status.textContent = "Redaction mode saved";
});

async function loadPolicies() {
  const response = await fetch("/api/policies");
  if (!response.ok) return;
  const data = await response.json();
  const select = document.querySelector("#policySelect");
  select.innerHTML = '<option value="">Custom instruction</option>';
  data.policies.forEach(policy => {
    const option = document.createElement("option");
    option.value = policy.id;
    option.textContent = policy.name;
    option.title = policy.description || policy.prompt;
    select.appendChild(option);
  });
}

document.querySelector("#policySelect").addEventListener("change", async () => {
  const policyId = document.querySelector("#policySelect").value;
  if (!policyId) return;
  const response = await fetch("/api/policies");
  if (!response.ok) return;
  const data = await response.json();
  const policy = data.policies.find(item => item.id === policyId);
  if (policy) document.querySelector("#instruction").value = policy.prompt;
});

function mergeFindings(existing, incoming) {
  const merged = [...existing];
  incoming.forEach(finding => {
    const key = `${finding.page}|${finding.category}|${finding.text}|${finding.bbox.join(",")}`;
    const duplicate = merged.find(item =>
      `${item.page}|${item.category}|${item.text}|${item.bbox.join(",")}` === key
    );
    if (!duplicate) merged.push(finding);
  });
  return merged;
}

function escapeHtml(value) {
  const node = document.createElement("div");
  node.textContent = value;
  return node.innerHTML;
}
