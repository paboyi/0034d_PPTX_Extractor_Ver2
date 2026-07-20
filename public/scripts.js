const frame = document.getElementById("frame");
const drop = document.getElementById("drop");
const fileInput = document.getElementById("file");
const statusEl = document.getElementById("status");
const results = document.getElementById("results");
const output = document.getElementById("output");
const specEl = document.getElementById("spec");
const copyBtn = document.getElementById("copyBtn");
const downloadBtn = document.getElementById("downloadBtn");
const truncNote = document.getElementById("truncNote");

// Backend URL
const BACKEND_API_BASE = "https://your-service-xxxxx.a.run.app"; 

const PREVIEW_CAP = 60000; // chars shown in preview; full data still copies/downloads
let fullJson = "";
let baseName = "presentation";

//  open picker 
drop.addEventListener("click", () => fileInput.click());
drop.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); fileInput.click(); }
});
fileInput.addEventListener("change", () => {
    if (fileInput.files.length) handleFile(fileInput.files[0]);
});

//  drag & drop 
["dragenter", "dragover"].forEach((ev) =>
    frame.addEventListener(ev, (e) => { e.preventDefault(); frame.classList.add("dragging"); })
);
["dragleave", "drop"].forEach((ev) =>
    frame.addEventListener(ev, (e) => {
    e.preventDefault();
    if (ev === "dragleave" && frame.contains(e.relatedTarget)) return;
    frame.classList.remove("dragging");
    })
);
frame.addEventListener("drop", (e) => {
    const f = e.dataTransfer.files[0];
    if (f) handleFile(f);
});

//  the pptx-only gate (client side) 
function isPptx(file) {
    const nameOk = file.name.toLowerCase().endsWith(".pptx");
    const typeOk =
    file.type === "" ||
    file.type === "application/octet-stream" ||
    file.type === "application/vnd.openxmlformats-officedocument.presentationml.presentation";
    return nameOk && typeOk;
}

function setStatus(msg, kind) {
    statusEl.className = "status" + (kind ? " " + kind : "");
    statusEl.innerHTML = msg;
}

async function handleFile(file) {
    results.classList.remove("show");

    if (!isPptx(file)) {
    setStatus("✕ " + escapeHtml(file.name) + " isn't a .pptx. Choose a PowerPoint file.", "err");
    return;
    }

    baseName = file.name.replace(/\.pptx$/i, "");
    setStatus('Reading <strong>' + escapeHtml(file.name) + '</strong><span class="blink">…</span>', "busy");

    const form = new FormData();
    form.append("file", file);

    try {
    // const res = await fetch("/api/extract", { method: "POST", body: form }); //local dev .env
    const res = await fetch(BACKEND_API_BASE + "/api/extract", { method: "POST", body: form }); //API .env
    const data = await res.json();

    if (!res.ok) {
        setStatus("✕ " + escapeHtml(data.error || "Extraction failed."), "err");
        return;
    }

    render(data, file.name);
    const slides = data.head?.slides ?? 0;
    setStatus("✓ Extracted " + slides + " slide" + (slides === 1 ? "" : "s") + " from " + escapeHtml(file.name), "ok");
    } catch (err) {
    setStatus("✕ Could not reach the server. Is it running?", "err");
    }
}

function render(data, filename) {
    // spec strip
    const w = data.head?.size?.width ?? "—";
    const h = data.head?.size?.height ?? "—";
    let textN = 0, imgN = 0;
    for (const key in (data.slides || {})) {
    for (const el of data.slides[key].elements || []) {
        if (el.type === "text") textN++;
        else if (el.type === "image") imgN++;
    }
    }
    specEl.innerHTML = "";
    addCell("Slides", data.head?.slides ?? "—");
    addCell("Canvas", w + " <small>× " + h + " px</small>");
    addCell("Text blocks", textN);
    addCell("Images", imgN);

    fullJson = JSON.stringify(data, null, 2);
    if (fullJson.length > PREVIEW_CAP) {
    output.textContent = fullJson.slice(0, PREVIEW_CAP) + "\n… ";
    truncNote.hidden = false;
    truncNote.textContent =
        "Preview truncated at " + PREVIEW_CAP.toLocaleString() +
        " characters (images are inline). Use Download for the full file.";
    } else {
    output.textContent = fullJson;
    truncNote.hidden = true;
    }
    results.classList.add("show");
}

function addCell(k, v) {
    const cell = document.createElement("div");
    cell.className = "cell";
    cell.innerHTML = '<div class="k">' + k + '</div><div class="v">' + v + "</div>";
    specEl.appendChild(cell);
}

copyBtn.addEventListener("click", async () => {
    try {
    await navigator.clipboard.writeText(fullJson);
    flash(copyBtn, "Copied");
    } catch {
    flash(copyBtn, "Copy failed");
    }
});

downloadBtn.addEventListener("click", () => {
    const blob = new Blob([fullJson], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = baseName + ".json";
    a.click();
    URL.revokeObjectURL(url);
});

function flash(btn, text) {
    const old = btn.textContent;
    btn.textContent = text;
    setTimeout(() => (btn.textContent = old), 1400);
}

function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])
    );
}