(() => {
  const state = { token: null, columns: [], rows: [], detections: [] };

  const el = (id) => document.getElementById(id);
  const fileInput = el("file-input");
  const uploadBtn = el("upload-btn");
  const uploadError = el("upload-error");
  const detectionsCard = el("detections-card");
  const detectionsList = el("detections-list");
  const gridCard = el("grid-card");
  const tableWrap = el("table-wrap");
  const statusMsg = el("status-msg");

  function escapeHtml(value) {
    return String(value ?? "").replace(/[&<>"']/g, (c) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
    }[c]));
  }

  function showError(message) {
    uploadError.textContent = message;
    uploadError.hidden = !message;
  }

  function setStatus(message) {
    statusMsg.textContent = message;
    if (message) setTimeout(() => { if (statusMsg.textContent === message) statusMsg.textContent = ""; }, 4000);
  }

  async function postJson(url, body) {
    const resp = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    if (!resp.ok) {
      let message = "Request failed.";
      try { message = (await resp.json()).error || message; } catch (_) { /* ignore */ }
      throw new Error(message);
    }
    return resp;
  }

  async function downloadFromEndpoint(url, body) {
    const resp = await postJson(url, body);
    const blob = await resp.blob();
    const cd = resp.headers.get("Content-Disposition") || "";
    const match = /filename="?([^"]+)"?/.exec(cd);
    const filename = match ? match[1] : "download.xlsx";
    const blobUrl = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = blobUrl;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(blobUrl);
  }

  function renderDetections() {
    const lowConfidence = state.detections.filter((d) => d.confidence === "low");
    if (lowConfidence.length === 0) {
      detectionsCard.hidden = true;
      return;
    }
    detectionsCard.hidden = false;
    detectionsList.innerHTML = lowConfidence
      .map((d) => `
        <div class="detection-row" data-sheet="${escapeHtml(d.sheet)}">
          <span class="low">${escapeHtml(d.sheet)}</span>
          <span>header row guessed at</span>
          <input type="number" min="1" value="${d.header_row}" class="override-row">
          <button class="apply-override">Use this row</button>
        </div>`)
      .join("");
  }

  function renderTable() {
    const cols = state.columns;
    let headHtml = '<th class="source-col">Source Tab</th>';
    headHtml += cols
      .map((c, i) => `
        <th data-col-idx="${i}">
          <span class="col-name">${escapeHtml(c)}</span>
          <button class="link delete-col" data-col-idx="${i}" title="Delete column">&times;</button>
        </th>`)
      .join("");

    let bodyHtml = state.rows
      .map((row, rIdx) => {
        const sourceCell = `<td class="source-col">${escapeHtml(row.__source_tab)}</td>`;
        const cells = cols
          .map((c, cIdx) => `<td contenteditable="true" data-row-idx="${rIdx}" data-col-idx="${cIdx}">${escapeHtml(row[c])}</td>`)
          .join("");
        return `<tr>${sourceCell}${cells}</tr>`;
      })
      .join("");

    tableWrap.innerHTML = `<table><thead><tr>${headHtml}</tr></thead><tbody>${bodyHtml}</tbody></table>`;
  }

  function loadMasterPayload(payload) {
    state.columns = payload.columns;
    state.rows = payload.rows;
    if (payload.detections) state.detections = payload.detections;
    renderDetections();
    renderTable();
    gridCard.hidden = false;
  }

  uploadBtn.addEventListener("click", async () => {
    showError("");
    const file = fileInput.files[0];
    if (!file) { showError("Choose a .xlsx file first."); return; }

    const formData = new FormData();
    formData.append("workbook", file);
    uploadBtn.disabled = true;
    try {
      const resp = await fetch("/upload", { method: "POST", body: formData });
      const payload = await resp.json();
      if (!resp.ok) throw new Error(payload.error || "Upload failed.");
      state.token = payload.token;
      loadMasterPayload(payload);
    } catch (err) {
      showError(err.message);
    } finally {
      uploadBtn.disabled = false;
    }
  });

  detectionsList.addEventListener("click", async (e) => {
    if (!e.target.classList.contains("apply-override")) return;
    const row = e.target.closest(".detection-row");
    const sheet = row.dataset.sheet;
    const headerRow = parseInt(row.querySelector(".override-row").value, 10);
    try {
      const resp = await postJson("/header-override", { token: state.token, sheet_name: sheet, header_row: headerRow });
      const payload = await resp.json();
      loadMasterPayload(payload);
      setStatus(`Re-parsed '${sheet}' using row ${headerRow} as the header.`);
    } catch (err) {
      showError(err.message);
    }
  });

  tableWrap.addEventListener("input", (e) => {
    const target = e.target;
    if (target.tagName !== "TD" || target.dataset.rowIdx === undefined) return;
    const rIdx = parseInt(target.dataset.rowIdx, 10);
    const cIdx = parseInt(target.dataset.colIdx, 10);
    const col = state.columns[cIdx];
    state.rows[rIdx][col] = target.textContent;
  });

  tableWrap.addEventListener("click", (e) => {
    if (!e.target.classList.contains("delete-col")) return;
    const cIdx = parseInt(e.target.dataset.colIdx, 10);
    const colName = state.columns[cIdx];
    if (!confirm(`Delete column "${colName}"? On push, this removes it from every original tab that has it.`)) return;
    state.columns.splice(cIdx, 1);
    renderTable();
  });

  el("add-column-btn").addEventListener("click", () => {
    const name = prompt("New column name:");
    if (!name) return;
    const trimmed = name.trim();
    if (!trimmed) return;
    if (state.columns.some((c) => c.toLowerCase() === trimmed.toLowerCase())) {
      alert("A column with that name already exists.");
      return;
    }
    state.columns.push(trimmed);
    state.rows.forEach((row) => { row[trimmed] = ""; });
    renderTable();
  });

  el("download-btn").addEventListener("click", async () => {
    try {
      await downloadFromEndpoint("/download-master", { columns: state.columns, rows: state.rows });
      setStatus("Downloaded consolidated_master.xlsx");
    } catch (err) {
      showError(err.message);
    }
  });

  el("push-btn").addEventListener("click", async () => {
    if (!confirm("Push changes back to the original tabs? Deleted columns will be removed from every tab that had them. This downloads a new file — your original upload is untouched.")) return;
    try {
      await downloadFromEndpoint("/push", { token: state.token, columns: state.columns, rows: state.rows });
      setStatus("Downloaded synced_result.xlsx with highlighted changes.");
    } catch (err) {
      showError(err.message);
    }
  });
})();
