const EMAIL_TO = "acquisti@oleodinamicaseguini.it";
const codeReader = new ZXing.BrowserMultiFormatReader();
let selectedDeviceId;
let activeInput = null;

const els = {
  items: document.getElementById('items'),
  addRow: document.getElementById('addRow'),
  generate: document.getElementById('generate'),
  output: document.getElementById('output'),
  copy: document.getElementById('copy'),
  reset: document.getElementById('reset'),
  scanModal: document.getElementById('scanModal'),
  closeScan: document.getElementById('closeScan'),
  scanVideo: document.getElementById('scanVideo'),
  scanPill: document.getElementById('scanPill'),
  emailExcel: document.getElementById('emailExcel')
};

// --- Inizializzazione ---
function addItemRow(code = "", qty = "") {
  const div = document.createElement('div');
  div.className = 'item-grid';
  div.innerHTML = `
    <div>
      <label>Codice Articolo</label>
      <div style="display:flex; gap:5px;">
        <input class="code-input" value="${code}" placeholder="Scansiona o scrivi...">
        <button type="button" class="scan-btn" style="background:#ddd;">📷</button>
      </div>
    </div>
    <div>
      <label>Qtà</label>
      <input type="number" class="qty-input" value="${qty}" placeholder="0">
    </div>
    <button type="button" class="danger delete-btn">X</button>
  `;
  els.items.appendChild(div);
}

// --- Logica Scanner (Ottimizzata iPhone) ---
async function openScanner(inputEl) {
  activeInput = inputEl;
  els.scanModal.classList.add('open');
  els.scanPill.textContent = "Accesso fotocamera...";

  try {
    const videoDevices = await codeReader.listVideoInputDevices();
    // Seleziona la fotocamera posteriore
    selectedDeviceId = videoDevices.length > 1 ? videoDevices[1].deviceId : videoDevices[0].deviceId;
    
    els.scanPill.textContent = "Inquadra il codice...";
    
    codeReader.decodeFromVideoDevice(selectedDeviceId, 'scanVideo', (result, err) => {
      if (result) {
        activeInput.value = result.text;
        vibratePhone();
        closeScanner();
      }
    });
  } catch (err) {
    console.error(err);
    alert("Errore fotocamera: " + err);
    closeScanner();
  }
}

function closeScanner() {
  codeReader.reset();
  els.scanModal.classList.remove('open');
  activeInput = null;
}

function vibratePhone() {
  if (navigator.vibrate) navigator.vibrate(200);
}

// --- Event Listeners ---
els.addRow.addEventListener('click', () => addItemRow());

els.items.addEventListener('click', (e) => {
  if (e.target.classList.contains('delete-btn')) {
    e.target.closest('.item-grid').remove();
  }
  if (e.target.classList.contains('scan-btn')) {
    const input = e.target.parentElement.querySelector('.code-input');
    openScanner(input);
  }
});

els.closeScan.addEventListener('click', closeScanner);

els.reset.addEventListener('click', () => {
  if (confirm("Vuoi svuotare tutto?")) {
    els.items.innerHTML = "";
    els.output.value = "";
    addItemRow();
  }
});

els.generate.addEventListener('click', () => {
  const rows = [...document.querySelectorAll('.item-grid')];
  let text = `Richiesta da: ${document.getElementById('operator').value || 'N/D'}\n`;
  text += `Fornitore: ${document.getElementById('supplier').value || 'N/D'}\n\n`;
  
  rows.forEach(row => {
    const code = row.querySelector('.code-input').value;
    const qty = row.querySelector('.qty-input').value;
    if (code) text += `- ${code} (Qtà: ${qty})\n`;
  });
  
  els.output.value = text;
});

els.copy.addEventListener('click', () => {
  els.output.select();
  document.execCommand('copy');
  alert("Copiato!");
});

// Avvia con una riga vuota
addItemRow();
