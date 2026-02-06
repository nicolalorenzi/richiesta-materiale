const EMAIL_TO = "acquisti@oleodinamicaseguini.it";

/* =========================
   Modulo – righe fisse (mail)
   ========================= */
const DOC_INFO = {
  mod: "Mod. 07-05.01",
  rev: "Rev. 1 del 01/09/2003",
  agg: "Agg. 1 del 01/09/2003",
};

/* =========================
   Utils
   ========================= */
const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (m) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;",
  }[m]));
}

const pad = (n) => String(n).padStart(2, "0");

function formatNowText() {
  const d = new Date();
  return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function formatDateForFile() {
  const d = new Date();
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; // YYYY-MM-DD
}

function sanitizeFilePart(s) {
  return (s || "")
    .trim()
    .replace(/[\\/:*?"<>|]/g, "-")
    .replace(/\s+/g, " ")
    .slice(0, 60);
}

/* =========================
   DOM refs
   ========================= */
const els = {
  supplier: $("#supplier"),
  operator: $("#operator"),
  items: $("#items"),
  output: $("#output"),
  addRow: $("#addRow"),
  generate: $("#generate"),
  emailExcel: $("#emailExcel"),
  copy: $("#copy"),
  share: $("#share"),
  print: $("#print"),
  reset: $("#reset"),

  // scanner modal
  scanModal: $("#scanModal"),
  scanVideo: $("#scanVideo"),
  closeScan: $("#closeScan"),
  scanPill: $("#scanPill"),
  lastCodePill: $("#lastCodePill"),
};

function mustExist(el, name) {
  if (!el) throw new Error(`Elemento mancante: ${name}. Controlla gli id nell'HTML.`);
  return el;
}

// Fail fast se mancano pezzi in HTML
[
  ["supplier", els.supplier],
  ["operator", els.operator],
  ["items", els.items],
  ["output", els.output],
  ["addRow", els.addRow],
  ["generate", els.generate],
  ["emailExcel", els.emailExcel],
  ["copy", els.copy],
  ["share", els.share],
  ["print", els.print],
  ["reset", els.reset],
  ["scanModal", els.scanModal],
  ["scanVideo", els.scanVideo],
  ["closeScan", els.closeScan],
  ["scanPill", els.scanPill],
  ["lastCodePill", els.lastCodePill],
].forEach(([n, el]) => mustExist(el, n));

/* =========================
   Template riga articolo
   ========================= */
function itemRowTemplate(code = "", qty = "") {
  return `
    <div class="item" data-item>
      <div class="item-grid">
        <div class="code-block">
          <label>Codice</label>
          <div class="code-actions">
            <input class="code" data-code placeholder="Codice articolo" value="${escapeHtml(code)}" autocomplete="off" />
            <button class="scan-btn" data-action="scan" type="button" title="Scansiona">📷</button>
          </div>
        </div>

        <div class="qty-del">
          <div class="qty-block">
            <label>Q.tà</label>
            <input class="qty" data-qty type="number" min="0" max="9999" step="1" inputmode="numeric"
                   placeholder="0" value="${escapeHtml(String(qty))}" />
          </div>
          <button class="del-btn" data-action="del" type="button" title="Elimina riga">✕</button>
        </div>
      </div>
    </div>
  `;
}

function addItemRow({ code = "", qty = "" } = {}) {
  els.items.insertAdjacentHTML("beforeend", itemRowTemplate(code, qty));
}

function getItems() {
  return $$("[data-item]", els.items)
    .map((row) => {
      const code = ($("[data-code]", row)?.value || "").trim();
      const qtyRaw = ($("[data-qty]", row)?.value || "").trim();
      const qtyNum = qtyRaw === "" ? "" : Number(qtyRaw);
      return { code, qty: qtyNum };
    })
    .filter((x) => x.code || (x.qty !== "" && !Number.isNaN(x.qty)));
}

function getValidatedData() {
  const supplier = els.supplier.value.trim();
  const operator = els.operator.value.trim();
  const items = getItems();

  if (!supplier) { alert("Inserisci il fornitore."); return null; }
  if (!operator) { alert("Inserisci il nome operatore."); return null; }
  if (items.length === 0) { alert("Inserisci almeno una riga (codice + quantità)."); return null; }

  return { supplier, operator, items };
}

/* =========================
   Auto aggiunta riga (ultima riga completa)
   + apertura scanner sulla nuova
   ========================= */
function maybeAddNextRowAndOpenScanner(fromRowEl) {
  const codeInput = $("[data-code]", fromRowEl);
  const qtyInput = $("[data-qty]", fromRowEl);
  if (!codeInput || !qtyInput) return;

  const code = codeInput.value.trim();
  const qty = qtyInput.value.trim();

  if (!code) return;
  if (qty === "" || Number.isNaN(Number(qty))) return;

  const isLast = fromRowEl === els.items.lastElementChild;
  if (!isLast) return;

  addItemRow();

  const newRow = els.items.lastElementChild;
  const newCode = $("[data-code]", newRow);
  const newScan = $('[data-action="scan"]', newRow);

  newCode?.focus();
  if (newScan) setTimeout(() => newScan.click(), 150);
}

/* =========================
   Output testo
   ========================= */
function generateText() {
  const data = getValidatedData();
  if (!data) return;

  const { supplier, operator, items } = data;

  const lines = [];
  lines.push("RICHIESTA INTERNA MATERIALE");
  lines.push(`Data: ${formatNowText()}`);
  lines.push(`Fornitore: ${supplier}`);
  lines.push(`Operatore: ${operator}`);
  lines.push("");

  items.forEach((it, idx) => {
    const q = (it.qty === "" || Number.isNaN(it.qty)) ? "" : ` — Q.tà ${it.qty}`;
    lines.push(`${idx + 1}) ${it.code}${q}`);
  });

  lines.push("");
  lines.push(`Totale righe: ${items.length}`);

  els.output.value = lines.join("\n");
}

/* =========================
   Excel
   ========================= */
function buildExcelAndDownload() {
  if (!window.XLSX) {
    alert("Libreria Excel non caricata. Controlla connessione o blocchi script.");
    return null;
  }

  const data = getValidatedData();
  if (!data) return null;

  const { supplier, operator, items } = data;
  const nowText = formatNowText();
  const dateForFile = formatDateForFile();

  const aoa = [
    ["ORDINE FORNITORE"],
    ["Fornitore", supplier],
    ["Operatore", operator],
    ["Data/Ora", nowText],
    [],
    ["Codice", "Q.tà"],
    ...items.map((i) => [i.code, (i.qty === "" || Number.isNaN(i.qty)) ? "" : i.qty]),
  ];

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{ wch: 30 }, { wch: 10 }];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Ordine");

  const safeSupplier = sanitizeFilePart(supplier);
  const filename = `ORDINE ${safeSupplier} - ${dateForFile}.xlsx`;

  XLSX.writeFile(wb, filename);
  return { filename, nowText, ...data };
}

/* =========================
   Email (mailto) + Mod/Rev/Agg
   ========================= */
function openEmailDraft({ supplier, operator, nowText, items, filename }) {
  const subject = `ORDINE ${supplier} - ${formatDateForFile()}`;

  const lines = [];
  lines.push("RICHIESTA INTERNA MATERIALE");
  lines.push("");
  lines.push(DOC_INFO.mod);
  lines.push(DOC_INFO.rev);
  lines.push(DOC_INFO.agg);
  lines.push("");
  lines.push("Buongiorno,");
  lines.push("");
  lines.push("in allegato invio ordine fornitore.");
  lines.push("");
  lines.push(`Fornitore: ${supplier}`);
  lines.push(`Operatore: ${operator}`);
  lines.push(`Data/Ora: ${nowText}`);
  lines.push(`File: ${filename}`);
  lines.push("");
  lines.push("Riepilogo righe:");
  items.forEach((it, idx) => {
    const q = (it.qty === "" || Number.isNaN(it.qty)) ? "" : it.qty;
    lines.push(`${idx + 1}) ${it.code} — Q.tà ${q}`);
  });
  lines.push("");
  lines.push("Grazie.");

  const body = lines.join("\n");
  const mailto = `mailto:${encodeURIComponent(EMAIL_TO)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
  window.location.href = mailto;
}

/* =========================
   Copy / Share / Print / Reset
   ========================= */
async function copyText() {
  if (!els.output.value.trim()) return alert("Genera prima la richiesta.");
  await navigator.clipboard.writeText(els.output.value);
  alert("Copiato ✅");
}

async function shareText() {
  if (!els.output.value.trim()) return alert("Genera prima la richiesta.");
  if (navigator.share) {
    await navigator.share({ text: els.output.value, title: "Richiesta interna materiale" });
  } else {
    alert("Condivisione non supportata qui. Usa “Copia testo”.");
  }
}

function resetAll() {
  if (!confirm("Svuotare tutto?")) return;
  els.supplier.value = "";
  els.operator.value = "";
  els.items.innerHTML = "";
  els.output.value = "";
  addItemRow(); // 1 riga
}

/* =========================
   Scanner Barcode (camera)
   ========================= */
let activeCodeInput = null;
let scanStream = null;
let scanTimer = null;
let barcodeDetector = null;
let lastDetected = { value: null, ts: 0 };

const cropCanvas = document.createElement("canvas");
const cropCtx = cropCanvas.getContext("2d", { willReadFrequently: true });

function setScanPill(text, ok = true) {
  els.scanPill.textContent = text;
  els.scanPill.classList.toggle("bad", !ok);
}

function cleanCode(raw) {
  return (raw || "").trim().replace(/^\*+|\*+$/g, "").replace(/\s+/g, "");
}

async function tryLockLandscape() {
  try {
    if (screen.orientation && screen.orientation.lock) {
      await screen.orientation.lock("landscape");
    }
  } catch {}
}

async function unlockOrientation() {
  try {
    if (screen.orientation && screen.orientation.unlock) screen.orientation.unlock();
  } catch {}
}

async function openScannerFor(codeInput) {
  activeCodeInput = codeInput;

  if (!navigator.mediaDevices?.getUserMedia) {
    alert("Fotocamera non disponibile su questo browser.");
    return;
  }
  if (!("BarcodeDetector" in window)) {
    alert("Scanner non supportato qui. Usa Chrome/Android oppure inserisci a mano.");
    return;
  }

  barcodeDetector = new window.BarcodeDetector({
    formats: ["code_39", "code_128", "ean_13", "ean_8", "itf", "upc_a", "upc_e", "qr_code"],
  });

  els.scanModal.classList.add("open");
  els.scanModal.setAttribute("aria-hidden", "false");
  setScanPill("Richiesta permesso fotocamera…", true);
  els.lastCodePill.textContent = "Ultimo: —";
  lastDetected = { value: null, ts: 0 };

  await tryLockLandscape();

  try {
    scanStream = await navigator.mediaDevices.getUserMedia({
      video: {
        facingMode: { ideal: "environment" },
        width: { ideal: 1920 },
        height: { ideal: 1080 },
      },
      audio: false,
    });

    els.scanVideo.srcObject = scanStream;
    await els.scanVideo.play();

    // Zoom se disponibile
    const track = scanStream.getVideoTracks()[0];
    const caps = track.getCapabilities?.() || {};
    if (caps.zoom) {
      const targetZoom = Math.min(caps.zoom.max, Math.max(caps.zoom.min, 2));
      await track.applyConstraints({ advanced: [{ zoom: targetZoom }] }).catch(() => {});
    }

    setScanPill("Tieni il barcode ORIZZONTALE nel riquadro", true);
    startDetectLoop();
  } catch (e) {
    console.error(e);
    setScanPill("Permesso negato o errore fotocamera", false);
  }
}

function startDetectLoop() {
  stopDetectLoop();

  scanTimer = setInterval(async () => {
    const v = els.scanVideo;
    if (!v || v.readyState < 2) return;

    try {
      const vw = v.videoWidth, vh = v.videoHeight;
      if (!vw || !vh) return;

      // crop centrale "orizzontale"
      const cropW = Math.floor(vw * 0.78);
      const cropH = Math.floor(vh * 0.22);
      const sx = Math.floor((vw - cropW) / 2);
      const sy = Math.floor((vh - cropH) / 2);

      cropCanvas.width = cropW;
      cropCanvas.height = cropH;
      cropCtx.drawImage(v, sx, sy, cropW, cropH, 0, 0, cropW, cropH);

      const barcodes = await barcodeDetector.detect(cropCanvas);
      if (!barcodes || barcodes.length === 0) return;

      const raw = cleanCode(barcodes[0].rawValue);
      if (!raw) return;

      const now = Date.now();
      if (lastDetected.value === raw && (now - lastDetected.ts) < 1500) return;
      lastDetected = { value: raw, ts: now };

      els.lastCodePill.textContent = `Ultimo: ${raw}`;

      if (activeCodeInput) {
        activeCodeInput.value = raw;

        // focus sulla quantità della stessa riga
        const item = activeCodeInput.closest("[data-item]") || activeCodeInput.closest(".item");
        const qtyInput = item?.querySelector("[data-qty]") || item?.querySelector(".qty");
        if (qtyInput) qtyInput.focus();
      }

      await closeScanner();
    } catch {}
  }, 180);
}

function stopDetectLoop() {
  if (scanTimer) {
    clearInterval(scanTimer);
    scanTimer = null;
  }
}

async function closeScanner() {
  stopDetectLoop();

  if (scanStream) {
    scanStream.getTracks().forEach((t) => t.stop());
    scanStream = null;
  }
  els.scanVideo.srcObject = null;

  els.scanModal.classList.remove("open");
  els.scanModal.setAttribute("aria-hidden", "true");

  await unlockOrientation();
  activeCodeInput = null;
}

/* =========================
   Event delegation righe
   ========================= */
els.items.addEventListener("click", (e) => {
  const btn = e.target.closest("button[data-action]");
  if (!btn) return;

  const row = btn.closest("[data-item]");
  if (!row) return;

  const action = btn.dataset.action;

  if (action === "del") {
    row.remove();
    if (!els.items.children.length) addItemRow(); // sempre almeno 1 riga
    return;
  }

  if (action === "scan") {
    const codeInput = $("[data-code]", row);
    if (codeInput) openScannerFor(codeInput);
  }
});

// Enter su quantità -> auto-add + scanner nuova riga
els.items.addEventListener("keydown", (e) => {
  if (!e.target.matches("input[data-qty]")) return;
  if (e.key !== "Enter") return;

  e.preventDefault();
  const row = e.target.closest("[data-item]");
  if (row) maybeAddNextRowAndOpenScanner(row);
});

// Blur quantità -> stesso comportamento
els.items.addEventListener(
  "blur",
  (e) => {
    if (!e.target.matches("input[data-qty]")) return;
    const row = e.target.closest("[data-item]");
    if (row) setTimeout(() => maybeAddNextRowAndOpenScanner(row), 0);
  },
  true
);

/* =========================
   Eventi UI
   ========================= */
els.closeScan.addEventListener("click", closeScanner);
els.scanModal.addEventListener("click", (e) => {
  if (e.target === els.scanModal) closeScanner();
});

els.addRow.addEventListener("click", () => addItemRow());
els.generate.addEventListener("click", generateText);

els.emailExcel.addEventListener("click", () => {
  const info = buildExcelAndDownload();
  if (!info) return;
  openEmailDraft(info);
});

els.copy.addEventListener("click", copyText);
els.share.addEventListener("click", shareText);
els.print.addEventListener("click", () => {
  if (!els.output.value.trim()) generateText();
  window.print();
});
els.reset.addEventListener("click", resetAll);

/* init: 1 sola riga */
addItemRow();
