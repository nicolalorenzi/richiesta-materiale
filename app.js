const EMAIL_TO = "acquisti@oleodinamicaseguini.it";

/** =========================
 *  Utilities
 *  ========================= */
const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];

const escapeHtml = (s) =>
  String(s).replace(/[&<>"']/g, (m) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m]));

const pad2 = (n) => String(n).padStart(2, "0");

function formatNowText() {
  const d = new Date();
  return `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)}/${d.getFullYear()} ${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
}

function formatDateForFile() {
  const d = new Date();
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

function sanitizeFilePart(s) {
  return (s || "")
    .trim()
    .replace(/[\\/:*?"<>|]/g, "-")
    .replace(/\s+/g, " ")
    .slice(0, 60);
}

function alertAndReturn(msg) {
  alert(msg);
  return null;
}

/** =========================
 *  DOM refs
 *  ========================= */
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
  // scanner
  scanModal: $("#scanModal"),
  scanVideo: $("#scanVideo"),
  closeScan: $("#closeScan"),
  scanPill: $("#scanPill"),
  lastCodePill: $("#lastCodePill"),
};

/** =========================
 *  Items rendering + data access
 *  ========================= */
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

function getFormDataValidated() {
  const supplier = els.supplier.value.trim();
  const operator = els.operator.value.trim();
  const items = getItems();

  if (!supplier) return alertAndReturn("Inserisci il fornitore.");
  if (!operator) return alertAndReturn("Inserisci il nome operatore.");
  if (items.length === 0) return alertAndReturn("Inserisci almeno una riga (codice + quantità).");

  return { supplier, operator, items };
}

/** =========================
 *  Output generation
 *  ========================= */
function generateText() {
  const data = getFormDataValidated();
  if (!data) return;

  const { supplier, operator, items } = data;

  const lines = [
    "RICHIESTA INTERNA MATERIALE",
    `Data: ${formatNowText()}`,
    `Fornitore: ${supplier}`,
    `Operatore: ${operator}`,
    "",
    ...items.map((it, idx) => {
      const q = it.qty === "" || Number.isNaN(it.qty) ? "" : ` — Q.tà ${it.qty}`;
      return `${idx + 1}) ${it.code}${q}`;
    }),
    "",
    `Totale righe: ${items.length}`,
  ];

  els.output.value = lines.join("\n");
}

/** =========================
 *  Excel
 *  ========================= */
function buildExcelAndDownload() {
  if (!window.XLSX) return alertAndReturn("Libreria Excel non caricata. Controlla connessione o blocchi script.");

  const data = getFormDataValidated();
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
    ...items.map((i) => [i.code, i.qty === "" || Number.isNaN(i.qty) ? "" : i.qty]),
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

function openEmailDraft({ supplier, operator, nowText, items, filename }) {
  const subject = `ORDINE ${supplier} - ${formatDateForFile()}`;

  const lines = [
    "Buongiorno,",
    "",
    "in allegato invio ordine fornitore.",
    "",
    `Fornitore: ${supplier}`,
    `Operatore: ${operator}`,
    `Data/Ora: ${nowText}`,
    `File: ${filename}`,
    "",
    "Riepilogo righe:",
    ...items.map((it, idx) => `${idx + 1}) ${it.code} — Q.tà ${it.qty === "" ? "" : it.qty}`),
    "",
    "Grazie.",
  ];

  const body = lines.join("\n");
  const mailto = `mailto:${encodeURIComponent(EMAIL_TO)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
  window.location.href = mailto;
}

/** =========================
 *  Clipboard / Share
 *  ========================= */
async function copyText() {
  if (!els.output.value.trim()) return alert("Genera prima la richiesta.");
  await navigator.clipboard.writeText(els.output.value);
  alert("Copiato ✅");
}

async function shareText() {
  if (!els.output.value.trim()) return alert("Genera prima la richiesta.");
  if (!navigator.share) return alert("Condivisione non supportata qui. Usa “Copia testo”.");
  await navigator.share({ text: els.output.value, title: "Richiesta interna materiale" });
}

/** =========================
 *  Reset
 *  ========================= */
function resetAll() {
  if (!confirm("Svuotare tutto?")) return;
  els.supplier.value = "";
  els.operator.value = "";
  els.items.innerHTML = "";
  els.output.value = "";
  addItemRow(); // 1 sola riga
}

/** =========================
 *  Auto-add row when last row has code+qty (delegated)
 *  ========================= */
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

/** =========================
 *  Scanner (incapsulato)
 *  ========================= */
const Scanner = (() => {
  let activeCodeInput = null;
  let scanStream = null;
  let scanTimer = null;
  let barcodeDetector = null;
  let lastDetected = { value: null, ts: 0 };

  const cropCanvas = document.createElement("canvas");
  const cropCtx = cropCanvas.getContext("2d", { willReadFrequently: true });

  const setPill = (text, ok = true) => {
    els.scanPill.textContent = text;
    els.scanPill.classList.toggle("bad", !ok);
  };

  const cleanCode = (raw) => (raw || "").trim().replace(/^\*+|\*+$/g, "").replace(/\s+/g, "");

  async function tryLockLandscape() {
    try {
      if (screen.orientation?.lock) await screen.orientation.lock("landscape");
    } catch {}
  }

  async function unlockOrientation() {
    try {
      if (screen.orientation?.unlock) screen.orientation.unlock();
    } catch {}
  }

  function stopLoop() {
    if (scanTimer) clearInterval(scanTimer);
    scanTimer = null;
  }

  async function close() {
    stopLoop();
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

  function startLoop() {
    stopLoop();
    scanTimer = setInterval(async () => {
      const v = els.scanVideo;
      if (!v || v.readyState < 2) return;

      try {
        const vw = v.videoWidth, vh = v.videoHeight;
        if (!vw || !vh) return;

        const cropW = Math.floor(vw * 0.78);
        const cropH = Math.floor(vh * 0.22);
        const sx = Math.floor((vw - cropW) / 2);
        const sy = Math.floor((vh - cropH) / 2);

        cropCanvas.width = cropW;
        cropCanvas.height = cropH;
        cropCtx.drawImage(v, sx, sy, cropW, cropH, 0, 0, cropW, cropH);

        const barcodes = await barcodeDetector.detect(cropCanvas);
        if (!barcodes?.length) return;

        const raw = cleanCode(barcodes[0].rawValue);
        if (!raw) return;

        const now = Date.now();
        if (lastDetected.value === raw && now - lastDetected.ts < 1500) return;
        lastDetected = { value: raw, ts: now };

        els.lastCodePill.textContent = `Ultimo: ${raw}`;

        if (activeCodeInput) {
          activeCodeInput.value = raw;
          const item = activeCodeInput.closest("[data-item]");
          const qtyInput = item ? $("[data-qty]", item) : null;
          qtyInput?.focus();
        }

        await close();
      } catch {}
    }, 180);
  }

  async function openFor(codeInput) {
    activeCodeInput = codeInput;

    if (!navigator.mediaDevices?.getUserMedia) return alert("Fotocamera non disponibile su questo browser.");
    if (!("BarcodeDetector" in window)) return alert("Scanner non supportato qui. Usa Chrome/Android oppure inserisci a mano.");

    barcodeDetector = new window.BarcodeDetector({
      formats: ["code_39", "code_128", "ean_13", "ean_8", "itf", "upc_a", "upc_e", "qr_code"],
    });

    els.scanModal.classList.add("open");
    els.scanModal.setAttribute("aria-hidden", "false");
    setPill("Richiesta permesso fotocamera…", true);
    els.lastCodePill.textContent = "Ultimo: —";
    lastDetected = { value: null, ts: 0 };

    await tryLockLandscape();

    try {
      scanStream = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: { ideal: "environment" }, width: { ideal: 1920 }, height: { ideal: 1080 } },
        audio: false,
      });

      els.scanVideo.srcObject = scanStream;
      await els.scanVideo.play();

      const track = scanStream.getVideoTracks()[0];
      const caps = track.getCapabilities?.() || {};
      if (caps.zoom) {
        const targetZoom = Math.min(caps.zoom.max, Math.max(caps.zoom.min, 2));
        await track.applyConstraints({ advanced: [{ zoom: targetZoom }] }).catch(() => {});
      }

      setPill("Tieni il barcode ORIZZONTALE nel riquadro", true);
      startLoop();
    } catch (e) {
      console.error(e);
      setPill("Permesso negato o errore fotocamera", false);
    }
  }

  return { openFor, close };
})();

/** =========================
 *  Events (delegation)
 *  ========================= */
els.items.addEventListener("click", (e) => {
  const btn = e.target.closest("button[data-action]");
  if (!btn) return;

  const row = btn.closest("[data-item]");
  if (!row) return;

  const action = btn.dataset.action;

  if (action === "del") {
    row.remove();
    if (!els.items.children.length) addItemRow();
    return;
  }

  if (action === "scan") {
    const codeInput = $("[data-code]", row);
    if (codeInput) Scanner.openFor(codeInput);
  }
});

els.items.addEventListener("keydown", (e) => {
  if (!e.target.matches("input[data-qty]")) return;
  if (e.key !== "Enter") return;

  e.preventDefault();
  const row = e.target.closest("[data-item]");
  if (row) maybeAddNextRowAndOpenScanner(row);
});

els.items.addEventListener(
  "blur",
  (e) => {
    if (!e.target.matches("input[data-qty]")) return;
    const row = e.target.closest("[data-item]");
    if (row) setTimeout(() => maybeAddNextRowAndOpenScanner(row), 0);
  },
  true
);

els.closeScan.addEventListener("click", () => Scanner.close());
els.scanModal.addEventListener("click", (e) => {
  if (e.target === els.scanModal) Scanner.close();
});

els.addRow.addEventListener("click", () => addItemRow());
els.generate.addEventListener("click", generateText);

els.emailExcel.addEventListener("click", () => {
  const info = buildExcelAndDownload();
  if (!info) return;
  openEmailDraft({
    supplier: info.supplier,
    operator: info.operator,
    nowText: info.nowText,
    items: info.items,
    filename: info.filename,
  });
});

els.copy.addEventListener("click", copyText);
els.share.addEventListener("click", shareText);
els.print.addEventListener("click", () => {
  if (!els.output.value.trim()) generateText();
  window.print();
});
els.reset.addEventListener("click", resetAll);

/** init */
addItemRow();
