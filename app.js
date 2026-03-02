const EMAIL_TO = "acquisti@oleodinamicaseguini.it";
const LOGO_URL = "./logo.jpg";
const codeReader = new ZXing.BrowserMultiFormatReader();
let activeInput = null;

// Riferimenti DOM
const itemsContainer = document.getElementById('items');
const scanModal = document.getElementById('scanModal');
const scanVideo = document.getElementById('scanVideo');

// Funzione per aggiungere riga
function addItemRow(code = "", qty = "") {
    const div = document.createElement('div');
    div.className = 'item-grid';
    div.innerHTML = `
        <div>
            <div style="display:flex; gap:4px;">
                <input class="code-input" value="${code}" placeholder="Codice">
                <button type="button" class="scan-trigger" style="background:#eee;">📷</button>
            </div>
        </div>
        <div><input type="number" class="qty-input" value="${qty}" placeholder="Q.tà"></div>
        <button class="danger del-row">X</button>
    `;
    itemsContainer.appendChild(div);
}

// Logica Scanner
async function startScan(input) {
    activeInput = input;
    scanModal.classList.add('open');
    try {
        const devices = await codeReader.listVideoInputDevices();
        const backCamera = devices.find(d => d.label.toLowerCase().includes('back')) || devices[0];
        
        codeReader.decodeFromVideoDevice(backCamera.deviceId, 'scanVideo', (result) => {
            if (result) {
                activeInput.value = result.text;
                if(navigator.vibrate) navigator.vibrate(100);
                closeScanner();
                generateText(); // Aggiorna l'output
            }
        });
    } catch (err) {
        alert("Errore fotocamera: " + err);
        closeScanner();
    }
}

function closeScanner() {
    codeReader.reset();
    scanModal.classList.remove('open');
}

// Genera il testo per l'anteprima
function generateText() {
    const supplier = document.getElementById('supplier').value;
    const operator = document.getElementById('operator').value;
    let txt = `Richiesta Materiale\nFornitore: ${supplier}\nOperatore: ${operator}\n\n`;
    
    document.querySelectorAll('.item-grid').forEach(row => {
        const c = row.querySelector('.code-input').value;
        const q = row.querySelector('.qty-input').value;
        if(c) txt += `- ${c} (Qtà: ${q})\n`;
    });
    document.getElementById('output').value = txt;
}

// Funzione Excel (Recuperata dal tuo vecchio codice)
async function sendExcel() {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Richiesta');
    
    // Aggiungi dati
    worksheet.addRow(["RICHIESTA MATERIALE"]);
    worksheet.addRow(["Fornitore:", document.getElementById('supplier').value]);
    worksheet.addRow(["Operatore:", document.getElementById('operator').value]);
    worksheet.addRow([]);
    worksheet.addRow(["Codice Articolo", "Quantità"]);

    document.querySelectorAll('.item-grid').forEach(row => {
        const c = row.querySelector('.code-input').value;
        const q = row.querySelector('.qty-input').value;
        if(c) worksheet.addRow([c, q]);
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    saveAs(blob, "Richiesta_Materiale.xlsx");

    // Apri Email
    const subject = `Richiesta Materiale - ${document.getElementById('supplier').value}`;
    const body = encodeURIComponent(document.getElementById('output').value + "\n\n(L'allegato Excel è stato scaricato, caricalo manualmente)");
    window.location.href = `mailto:${EMAIL_TO}?subject=${subject}&body=${body}`;
}

// Eventi
document.getElementById('addRow').onclick = () => addItemRow();
document.getElementById('closeScan').onclick = closeScanner;
document.getElementById('emailExcel').onclick = sendExcel;
document.getElementById('copy').onclick = () => {
    document.getElementById('output').select();
    document.execCommand('copy');
    alert("Copiato!");
};
document.getElementById('reset').onclick = () => { if(confirm("Svuoto?")) location.reload(); };

// Gestione click su righe (per cancella e scanner)
itemsContainer.onclick = (e) => {
    if(e.target.classList.contains('del-row')) e.target.closest('.item-grid').remove();
    if(e.target.classList.contains('scan-trigger')) {
        const inp = e.target.closest('.item-grid').querySelector('.code-input');
        startScan(inp);
    }
};

// Avvio
addItemRow();
