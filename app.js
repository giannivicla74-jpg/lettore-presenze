// Lettore Presenze AI - versione PWA (senza server): legge e scrive i file Excel nella cartella autorizzata.
// Elenco operai: viene letto dalla tabella presenze del file Excel attivo (loadWorkersRoster).
// Nel codice non ci sono nomi di persone: i dati restano solo nei file Excel sul PC.
let OPERAI_ROSTER = [];

// Stato dell'applicazione
const state = {
  currentPdf: null,
  pageImages: [], // [frontDataUrl, backDataUrl]
  currentPageIndex: 0,
  zoomLevel: 1.0,
  rotationDegrees: 0,
  backRotation: 180, // Rotazione per raddrizzare il retro: usata anche per l'immagine inviata all'AI
  reviewingBatchItem: null, // rapportino della scansione multipla aperto con "Controlla"
  panX: 0,
  panY: 0,
  isDragging: false,
  dragStartX: 0,
  dragStartY: 0,
  tableData: [],
  selectedExcelPath: '',
  availableExcelFiles: []
};

// Inizializzazione PDF.js
if (window.pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'libs/pdf.worker.min.js';
}

// Elementi DOM
const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const docCanvas = document.getElementById('docCanvas');
const canvasWrapper = document.getElementById('canvasWrapper');
const viewport = document.getElementById('viewport');
const pageTabs = document.getElementById('pageTabs');
const tabFront = document.getElementById('tabFront');
const tabBack = document.getElementById('tabBack');
const btnZoomIn = document.getElementById('btnZoomIn');
const btnZoomOut = document.getElementById('btnZoomOut');
const btnZoomReset = document.getElementById('btnZoomReset');
const btnRotate180 = document.getElementById('btnRotate180');
const btnRotate90 = document.getElementById('btnRotate90');
const btnAnalyzeAI = document.getElementById('btnAnalyzeAI');
const btnWriteExcel = document.getElementById('btnWriteExcel');
const btnAddRow = document.getElementById('btnAddRow');
const btnClearTable = document.getElementById('btnClearTable');
const btnExportJson = document.getElementById('btnExportJson');
const tableBody = document.getElementById('tableBody');
const detectionBanner = document.getElementById('detectionBanner');
const bannerTitle = document.getElementById('bannerTitle');
const bannerDesc = document.getElementById('bannerDesc');
const statWorkers = document.getElementById('statWorkers');
const statHours = document.getElementById('statHours');

// Tasto espandi/riduci tabella
const btnToggleWide = document.getElementById('btnToggleWide');
const appContainer = document.querySelector('.app-container');
if (btnToggleWide && appContainer) {
  btnToggleWide.addEventListener('click', () => {
    appContainer.classList.toggle('wide-mode');
    const isWide = appContainer.classList.contains('wide-mode');
    btnToggleWide.textContent = isWide ? '⚖️ Vista 50/50' : '⛶ Allarga Tabella';
    setTimeout(() => { if (typeof fitToViewport === 'function') fitToViewport(); }, 320);
  });
}

// Metadati form
const metaDate = document.getElementById('metaDate');
const metaDay = document.getElementById('metaDay');
const metaCommessa = document.getElementById('metaCommessa');
const metaCantiere = document.getElementById('metaCantiere');
const metaCaposquadra = document.getElementById('metaCaposquadra');
const metaRespCommessa = document.getElementById('metaRespCommessa');

// Settings Modal
const settingsModal = document.getElementById('settingsModal');
const btnOpenSettings = document.getElementById('btnOpenSettings');
const btnCloseSettings = document.getElementById('btnCloseSettings');
const btnCancelSettings = document.getElementById('btnCancelSettings');
const btnSaveSettings = document.getElementById('btnSaveSettings');
const geminiApiKeyInput = document.getElementById('geminiApiKey');
const btnTestApiKey = document.getElementById('btnTestApiKey');
const apiKeyTestStatus = document.getElementById('apiKeyTestStatus');

// Carica API Key da localStorage
const savedApiKey = localStorage.getItem('gemini_api_key') || '';
if (geminiApiKeyInput) geminiApiKeyInput.value = savedApiKey;
const btnToggleApiKey = document.getElementById('btnToggleApiKey');
if (btnToggleApiKey && geminiApiKeyInput) {
  btnToggleApiKey.addEventListener('click', () => geminiApiKeyInput.classList.toggle('api-key-hidden'));
}

// Elementi Gestione File Excel Mensile
const excelFileSelect = document.getElementById('excelFileSelect');
const currentDestFile = document.getElementById('currentDestFile');
const btnNewMonth = document.getElementById('btnNewMonth');
const newMonthModal = document.getElementById('newMonthModal');
const btnCloseNewMonth = document.getElementById('btnCloseNewMonth');
const btnCancelNewMonth = document.getElementById('btnCancelNewMonth');
const btnConfirmNewMonth = document.getElementById('btnConfirmNewMonth');
const newMonthFileName = document.getElementById('newMonthFileName');

// Elenco dei file Excel della cartella autorizzata. state.selectedExcelPath contiene il NOME del file
// (es. "Presenze_10-Ottobre_2026.xlsm"): la cartella è sempre quella scelta dall'utente
async function loadAvailableExcelFiles(autoSelectName = null) {
  if (!excelFileSelect || !cartellaPronta) return;
  try {
    const files = await CartellaExcel.listExcelFiles();
    state.availableExcelFiles = files.map(name => ({ name, fullPath: name }));
    state.masterFileName = CartellaExcel.masterName(files);

    excelFileSelect.innerHTML = '';
    if (files.length === 0) {
      excelFileSelect.innerHTML = '<option value="">Nessun file Excel nella cartella</option>';
      state.selectedExcelPath = '';
      return;
    }

    const saved = autoSelectName || localStorage.getItem('selected_excel_name') || state.masterFileName;
    const selected = files.includes(saved) ? saved : files[0];
    files.forEach(name => {
      const opt = document.createElement('option');
      opt.value = name;
      opt.textContent = name;
      if (name === selected) opt.selected = true;
      excelFileSelect.appendChild(opt);
    });

    state.selectedExcelPath = selected;
    localStorage.setItem('selected_excel_name', selected);
    if (currentDestFile) currentDestFile.textContent = selected;
    loadWorkersRoster();
  } catch (err) {
    console.error('Errore lettura cartella Excel:', err);
    showToast(`Non riesco a leggere la cartella Excel: ${err.message}`, 'error');
  }
}

if (excelFileSelect) {
  excelFileSelect.addEventListener('change', (e) => {
    state.selectedExcelPath = e.target.value;
    localStorage.setItem('selected_excel_name', state.selectedExcelPath);
    if (currentDestFile) currentDestFile.textContent = state.selectedExcelPath;
    showToast(`File presenze attivo: ${state.selectedExcelPath}`, 'info');
    loadWorkersRoster();
  });
}

if (btnNewMonth && newMonthModal) {
  btnNewMonth.addEventListener('click', () => {
    newMonthModal.classList.add('active');
    if (newMonthFileName) {
      newMonthFileName.focus();
      newMonthFileName.select();
    }
  });
}

if (btnCloseNewMonth && newMonthModal) {
  btnCloseNewMonth.addEventListener('click', () => {
    newMonthModal.classList.remove('active');
  });
}

if (btnCancelNewMonth && newMonthModal) {
  btnCancelNewMonth.addEventListener('click', () => {
    newMonthModal.classList.remove('active');
  });
}

// Chiudi cliccando sullo sfondo oscurato
if (newMonthModal) {
  newMonthModal.addEventListener('click', (e) => {
    if (e.target === newMonthModal) {
      newMonthModal.classList.remove('active');
    }
  });
}

document.querySelectorAll('.btn-quick-month').forEach(btn => {
  btn.addEventListener('click', (e) => {
    const val = e.currentTarget.dataset.month;
    if (newMonthFileName && val) {
      newMonthFileName.value = val;
      newMonthFileName.focus();
    }
  });
});

if (newMonthFileName) {
  newMonthFileName.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      if (btnConfirmNewMonth) btnConfirmNewMonth.click();
    } else if (e.key === 'Escape') {
      newMonthModal.classList.remove('active');
    }
  });
}

if (btnConfirmNewMonth) {
  btnConfirmNewMonth.addEventListener('click', async () => {
    const rawName = (newMonthFileName ? newMonthFileName.value : '').trim();
    if (!rawName) {
      showToast('Inserisci un nome per il file Excel', 'error');
      return;
    }

    btnConfirmNewMonth.disabled = true;
    btnConfirmNewMonth.textContent = 'Caricamento in corso...';

    try {
      if (!cartellaPronta) throw new Error('collega prima la cartella dei file Excel');
      const fileName = /\.xlsm$/i.test(rawName) ? rawName : `${rawName}.xlsm`;
      if (await CartellaExcel.exists(fileName)) {
        // File già esistente: viene solo selezionato, SENZA toccarlo (prepararlo cancellerebbe i dati del mese)
        showToast(`${fileName} esisteva già: selezionato come mese attivo (dati non modificati)`, 'info');
      } else {
        const files = await CartellaExcel.listExcelFiles();
        const master = CartellaExcel.masterName(files);
        if (!master) throw new Error('nella cartella non trovo il file Master_Operai da copiare');
        const { month, year } = ExcelXlsm.parseMonthAndYear(fileName);
        const bytes = await ExcelXlsm.prepareNewMonth(await CartellaExcel.readBytes(master), month, year);
        await CartellaExcel.writeBytes(fileName, bytes, true);
        showToast(`Nuovo mese creato: ${fileName} (copia di ${master})`, 'success');
      }
      newMonthModal.classList.remove('active');
      await loadAvailableExcelFiles(fileName);
    } catch (err) {
      showToast(`Errore creazione file: ${err.message}`, 'error');
    } finally {
      btnConfirmNewMonth.disabled = false;
      btnConfirmNewMonth.textContent = '➕ Crea e Imposta Come Attivo';
    }
  });
}

// ============================================================
// COLLEGAMENTO ALLA CARTELLA DEI FILE EXCEL
// Barra in alto finché la cartella non è collegata: alla prima apertura si sceglie la cartella,
// alle successive il browser può chiedere solo una conferma (un clic)
// ============================================================
const folderBanner = document.getElementById('folderBanner');
const folderBannerText = document.getElementById('folderBannerText');
const btnFolderAction = document.getElementById('btnFolderAction');
const btnChangeFolder = document.getElementById('btnChangeFolder');
let cartellaPronta = false;
let folderAction = 'scegli';

function showFolderBanner(stato) {
  if (!folderBanner) return;
  if (stato === 'pronta') {
    folderBanner.hidden = true;
    return;
  }
  folderBanner.hidden = false;
  if (!CartellaExcel.supported()) {
    folderBannerText.textContent = 'Questo browser non può accedere ai file del PC: apri l\'app con Microsoft Edge o Google Chrome.';
    btnFolderAction.hidden = true;
  } else if (stato === 'da-confermare') {
    folderAction = 'conferma';
    folderBannerText.textContent = `Conferma l'accesso alla cartella "${CartellaExcel.name()}" per leggere e scrivere i file Excel.`;
    btnFolderAction.textContent = '✅ Consenti accesso';
  } else {
    folderAction = 'scegli';
    folderBannerText.textContent = 'Collega la cartella dei file Excel ("Cartella di lavoro con attivazione macro di Excel"): serve una sola volta.';
    btnFolderAction.textContent = '📁 Collega cartella Excel';
  }
}

async function onFolderReady() {
  cartellaPronta = true;
  showFolderBanner('pronta');
  await loadAvailableExcelFiles();
}

if (btnFolderAction) {
  btnFolderAction.addEventListener('click', async () => {
    try {
      if (folderAction === 'conferma' && await CartellaExcel.confirmAccess()) {
        await onFolderReady();
      } else {
        const name = await CartellaExcel.choose();
        await onFolderReady();
        showToast(`Cartella collegata: ${name}`, 'success');
      }
    } catch (err) {
      if (err.name !== 'AbortError') showToast(`Cartella non collegata: ${err.message}`, 'error');
    }
  });
}

if (btnChangeFolder) {
  btnChangeFolder.addEventListener('click', async () => {
    try {
      const name = await CartellaExcel.choose();
      await onFolderReady();
      showToast(`Cartella collegata: ${name}`, 'success');
    } catch (err) {
      if (err.name !== 'AbortError') showToast(`Cartella non cambiata: ${err.message}`, 'error');
    }
  });
}

// Le operazioni che leggono o scrivono Excel richiedono la cartella collegata
function requireFolder() {
  if (cartellaPronta) return true;
  showFolderBanner(CartellaExcel.name() ? 'da-confermare' : 'mancante');
  showToast('Collega prima la cartella dei file Excel (barra in alto)', 'error');
  return false;
}

CartellaExcel.restore().then(stato => (stato === 'pronta' ? onFolderReady() : showFolderBanner(stato)));

// Gestione Drag & Drop
dropzone.addEventListener('click', () => fileInput.click());
dropzone.addEventListener('dragover', (e) => { e.preventDefault(); dropzone.classList.add('drag-over'); });
dropzone.addEventListener('dragleave', () => dropzone.classList.remove('drag-over'));
dropzone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropzone.classList.remove('drag-over');
  if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
    handleFiles(e.dataTransfer.files);
  }
});
fileInput.addEventListener('change', (e) => {
  if (e.target.files && e.target.files.length > 0) {
    handleFiles(e.target.files);
  }
});

// Elaborazione file caricati
async function handleFiles(files) {
  state.pageImages = [];
  state.backRotation = 180;
  state.reviewingBatchItem = null; // un documento nuovo non è più il rapportino in controllo
  const file = files[0];
  if (!file) return;

  const isPdf = (file.name && file.name.toLowerCase().endsWith('.pdf')) || file.type === 'application/pdf';

  if (isPdf) {
    showToast('Caricamento e conversione PDF in corso...', 'info');
    try {
      const arrayBuffer = await file.arrayBuffer();
      const typedArray = new Uint8Array(arrayBuffer);

      const loadingTask = pdfjsLib.getDocument({
        data: typedArray
      });

      const pdf = await loadingTask.promise;
      state.currentPdf = pdf;

      // Renderizza pagina 1 e pagina 2 alla risoluzione ottimale (veloce e nitida)
      for (let pageNum = 1; pageNum <= Math.min(pdf.numPages, 2); pageNum++) {
        const page = await pdf.getPage(pageNum);
        const unscaledViewport = page.getViewport({ scale: 1.0 });
        // Calcola scala ottimale: max 1400px di larghezza per ridurre il payload del 90%
        const optimalScale = Math.min(1400 / unscaledViewport.width, 1.5);
        const viewport = page.getViewport({ scale: optimalScale });
        
        const tempCanvas = document.createElement('canvas');
        tempCanvas.width = viewport.width;
        tempCanvas.height = viewport.height;
        const ctx = tempCanvas.getContext('2d');
        await page.render({ canvasContext: ctx, viewport: viewport }).promise;
        // Compressione JPEG ottimizzata a 0.80 (da ~5MB a ~250KB per pagina)
        state.pageImages.push(tempCanvas.toDataURL('image/jpeg', 0.80));
      }

      pageTabs.style.display = 'flex';
      switchPage(0);
      btnAnalyzeAI.disabled = false;
      showToast(`PDF pronto (${state.pageImages.length} pagine). Clicca su 'Analizza con AI Vision'`, 'success');
      
      // Evidenzia visivamente il pulsante di analisi
      btnAnalyzeAI.style.animation = 'pulse 1.5s infinite';

    } catch (err) {
      console.error('Errore durante la lettura del PDF:', err);
      showToast(`Errore apertura PDF: ${err.message}`, 'error');
    }

  } else if (file.type.startsWith('image/') || /\.(jpe?g|png|webp|bmp)$/i.test(file.name)) {
    try {
      for (let f of files) {
        if (f.type.startsWith('image/') || /\.(jpe?g|png|webp|bmp)$/i.test(f.name)) {
          const rawUrl = await readFileAsDataURL(f);
          const optimizedUrl = await resizeImageIfNeeded(rawUrl, 1400);
          state.pageImages.push(optimizedUrl);
        }
      }
      if (state.pageImages.length > 1) {
        pageTabs.style.display = 'flex';
      }
      switchPage(0);
      btnAnalyzeAI.disabled = false;
      showToast(`${state.pageImages.length} immagine/i pronte per l'analisi`, 'success');
      btnAnalyzeAI.style.animation = 'pulse 1.5s infinite';
    } catch (err) {
      showToast(`Errore apertura immagine: ${err.message}`, 'error');
    }
  } else {
    showToast('Formato file non supportato. Trascina un PDF o un\'immagine.', 'error');
  }
}

function readFileAsDataURL(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function resizeImageIfNeeded(dataUrl, maxDim = 1400) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      let width = img.width;
      let height = img.height;
      if (width <= maxDim && height <= maxDim) {
        resolve(dataUrl);
        return;
      }
      if (width > height) {
        height = Math.round((height * maxDim) / width);
        width = maxDim;
      } else {
        width = Math.round((width * maxDim) / height);
        height = maxDim;
      }
      const canvas = document.createElement('canvas');
      canvas.width = width;
      canvas.height = height;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0, width, height);
      resolve(canvas.toDataURL('image/jpeg', 0.82));
    };
    img.onerror = () => resolve(dataUrl);
    img.src = dataUrl;
  });
}

// Cambio pagina (Fronte vs Retro)
function switchPage(index) {
  if (!state.pageImages[index]) return;
  state.currentPageIndex = index;
  
  if (index === 0) {
    tabFront.classList.add('active');
    tabBack.classList.remove('active');
    state.rotationDegrees = 0; // Il fronte solitamente è dritto
  } else {
    tabBack.classList.add('active');
    tabFront.classList.remove('active');
    state.rotationDegrees = state.backRotation; // Il retro nei rapportini scansionati è tipicamente capovolto!
  }

  state.panX = 0;
  state.panY = 0;
  renderCanvas();
}

tabFront.addEventListener('click', () => switchPage(0));
tabBack.addEventListener('click', () => switchPage(1));

// Renderizzazione immagine sul canvas con Adattamento, Pan e Zoom
function renderCanvas() {
  const dataUrl = state.pageImages[state.currentPageIndex];
  if (!dataUrl) return;

  const img = new Image();
  img.onload = () => {
    docCanvas.width = img.width;
    docCanvas.height = img.height;
    const ctx = docCanvas.getContext('2d');
    ctx.clearRect(0, 0, docCanvas.width, docCanvas.height);
    ctx.drawImage(img, 0, 0);

    fitToViewport();
  };
  img.src = dataUrl;
}

// Adatta perfettamente il documento all'area visibile (senza tagli o barre scomode)
function fitToViewport() {
  if (!docCanvas.width || !docCanvas.height || !viewport) return;
  const rect = viewport.getBoundingClientRect();
  const availW = rect.width - 32;
  const availH = rect.height - 32;
  if (availW <= 0 || availH <= 0) return;

  const isSideways = (state.rotationDegrees % 180 !== 0);
  const contentW = isSideways ? docCanvas.height : docCanvas.width;
  const contentH = isSideways ? docCanvas.width : docCanvas.height;

  const scaleW = availW / contentW;
  const scaleH = availH / contentH;
  const idealScale = Math.min(scaleW, scaleH);

  state.zoomLevel = Math.max(Math.round(idealScale * 100) / 100, 0.2);
  state.panX = 0;
  state.panY = 0;
  applyTransformations();
  updateZoomLabel();
}

function updateZoomLabel() {
  if (btnZoomReset) {
    btnZoomReset.textContent = `${Math.round(state.zoomLevel * 100)}%`;
  }
}

function applyTransformations() {
  canvasWrapper.style.transform = `translate(-50%, -50%) translate(${state.panX}px, ${state.panY}px) scale(${state.zoomLevel}) rotate(${state.rotationDegrees}deg)`;
}

// Gestione Drag & Pan con "Manina" (Grab / Grabbing)
if (viewport) {
  viewport.addEventListener('mousedown', (e) => {
    // Tasto sinistro (0) o rotella centrale (1)
    if (e.button !== 0 && e.button !== 1) return;
    state.isDragging = true;
    state.dragStartX = e.clientX - state.panX;
    state.dragStartY = e.clientY - state.panY;
    viewport.classList.add('is-dragging');
    e.preventDefault();
  });

  window.addEventListener('mousemove', (e) => {
    if (!state.isDragging) return;
    state.panX = e.clientX - state.dragStartX;
    state.panY = e.clientY - state.dragStartY;
    applyTransformations();
  });

  window.addEventListener('mouseup', () => {
    if (state.isDragging) {
      state.isDragging = false;
      viewport.classList.remove('is-dragging');
    }
  });

  // Zoom fluido con rotella del mouse centrato
  viewport.addEventListener('wheel', (e) => {
    e.preventDefault();
    const zoomFactor = e.deltaY < 0 ? 1.15 : 0.87;
    const newZoom = Math.min(Math.max(state.zoomLevel * zoomFactor, 0.15), 4.5);
    state.zoomLevel = Math.round(newZoom * 100) / 100;
    applyTransformations();
    updateZoomLabel();
  }, { passive: false });

  // Doppio clic per riadattare perfettamente allo schermo
  viewport.addEventListener('dblclick', () => {
    fitToViewport();
  });
}

// Auto-adattamento su ridimensionamento finestra
window.addEventListener('resize', () => {
  if (state.pageImages.length > 0 && state.panX === 0 && state.panY === 0) {
    fitToViewport();
  }
});

// Controlli Zoom e Rotazione da Pulsanti
btnZoomIn.addEventListener('click', () => {
  state.zoomLevel = Math.min(Math.round((state.zoomLevel + 0.15) * 100) / 100, 4.5);
  applyTransformations();
  updateZoomLabel();
});

btnZoomOut.addEventListener('click', () => {
  state.zoomLevel = Math.max(Math.round((state.zoomLevel - 0.15) * 100) / 100, 0.15);
  applyTransformations();
  updateZoomLabel();
});

btnZoomReset.addEventListener('click', () => {
  fitToViewport();
});

btnRotate180.addEventListener('click', () => {
  state.rotationDegrees = (state.rotationDegrees + 180) % 360;
  if (state.currentPageIndex === 1) state.backRotation = state.rotationDegrees;
  applyTransformations();
});

btnRotate90.addEventListener('click', () => {
  state.rotationDegrees = (state.rotationDegrees + 90) % 360;
  if (state.currentPageIndex === 1) state.backRotation = state.rotationDegrees;
  applyTransformations();
});

// Modale Impostazioni
btnOpenSettings.addEventListener('click', () => {
  if (apiKeyTestStatus) apiKeyTestStatus.style.display = 'none';
  settingsModal.classList.add('active');
});
btnCloseSettings.addEventListener('click', () => settingsModal.classList.remove('active'));
btnCancelSettings.addEventListener('click', () => settingsModal.classList.remove('active'));
btnSaveSettings.addEventListener('click', () => {
  const key = geminiApiKeyInput.value.trim();
  localStorage.setItem('gemini_api_key', key);
  cachedCandidateModels = null; // Resetta cache modelli
  settingsModal.classList.remove('active');
  showToast('API Key salvata correttamente', 'success');
});

// Test connessione API Key dalle Impostazioni
if (btnTestApiKey) {
  btnTestApiKey.addEventListener('click', async () => {
    const key = geminiApiKeyInput.value.trim();
    if (!key) {
      showApiKeyStatus('Inserisci prima una chiave API da testare.', 'error');
      return;
    }
    btnTestApiKey.disabled = true;
    btnTestApiKey.textContent = '⏳ Verifica in corso...';
    try {
      const models = await getAvailableGeminiModels(key);
      if (models && models.length > 0) {
        cachedCandidateModels = null;
        showApiKeyStatus(`✅ Connessione riuscita! Modelli attivi trovati: ${models.slice(0, 4).join(', ')}`, 'success');
      } else {
        showApiKeyStatus(`⚠️ Connessione stabilita, verranno usati i modelli standard (gemini-3.5-flash / gemini-flash-latest).`, 'warning');
      }
    } catch (err) {
      showApiKeyStatus(`❌ ${err.message}`, 'error');
    } finally {
      btnTestApiKey.disabled = false;
      btnTestApiKey.textContent = '🔍 Verifica Connessione';
    }
  });
}

function showApiKeyStatus(msg, type) {
  if (!apiKeyTestStatus) return;
  apiKeyTestStatus.style.display = 'block';
  apiKeyTestStatus.textContent = msg;
  if (type === 'success') {
    apiKeyTestStatus.style.background = 'rgba(16, 185, 129, 0.15)';
    apiKeyTestStatus.style.border = '1px solid rgba(16, 185, 129, 0.4)';
    apiKeyTestStatus.style.color = '#34d399';
  } else if (type === 'warning') {
    apiKeyTestStatus.style.background = 'rgba(245, 158, 11, 0.15)';
    apiKeyTestStatus.style.border = '1px solid rgba(245, 158, 11, 0.4)';
    apiKeyTestStatus.style.color = '#fbbf24';
  } else {
    apiKeyTestStatus.style.background = 'rgba(239, 68, 68, 0.15)';
    apiKeyTestStatus.style.border = '1px solid rgba(239, 68, 68, 0.4)';
    apiKeyTestStatus.style.color = '#f87171';
  }
}

// Cache modelli disponibili per evitare chiamate ripetute a ListModels
let cachedCandidateModels = null;

// Rileva automaticamente i modelli Gemini disponibili per questa API Key
async function getAvailableGeminiModels(apiKey) {
  // La chiave va nell'intestazione, non nell'indirizzo: gli indirizzi possono finire in cronologia e registri
  const res = await fetch('https://generativelanguage.googleapis.com/v1beta/models', { headers: { 'x-goog-api-key': apiKey } });
  if (!res.ok) {
    let errDetail = '';
    try {
      const errObj = await res.json();
      errDetail = errObj.error?.message || '';
    } catch (_) {
      errDetail = await res.text().catch(() => '');
    }
    if (res.status === 400 || res.status === 403) {
      throw new Error(`Chiave Google Gemini non valida o non abilitata (${res.status}): ${errDetail || 'Verifica la chiave su Google AI Studio'}`);
    }
    throw new Error(`Errore verifica API Google (${res.status}): ${errDetail}`);
  }

  const data = await res.json();
  const models = (data.models || [])
    .filter(m => Array.isArray(m.supportedGenerationMethods) && m.supportedGenerationMethods.includes('generateContent'))
    .map(m => m.name.replace(/^models\//, ''));
  return models;
}

// Risolve la lista ordinata dei migliori modelli da utilizzare
async function resolveCandidateModels(apiKey) {
  if (cachedCandidateModels && cachedCandidateModels.length > 0) {
    return cachedCandidateModels;
  }

  try {
    const list = await getAvailableGeminiModels(apiKey);
    if (list && list.length > 0) {
      const candidates = [];
      // Ordine di efficienza, velocità e stabilità
      const priority = [
        'gemini-3.5-flash',
        'gemini-flash-latest',
        'gemini-3.5-flash-lite',
        'gemini-flash-lite-latest',
        'gemini-2.5-flash' // ultimo: nei test legge peggio le caselle del retro
      ];
      for (const p of priority) {
        if (list.includes(p)) candidates.push(p);
      }
      // Esclude modelli specializzati (voce, immagini, live...) che non servono per leggere i fogli presenze
      const excluded = /tts|image|live|audio|transcribe|robotics|computer-use|omni|banana|embedding/;
      for (const m of list) {
        if (!candidates.includes(m) && m.startsWith('gemini-') && m.includes('flash') && !excluded.test(m)) {
          candidates.push(m);
        }
      }
      if (candidates.length > 0) {
        cachedCandidateModels = candidates;
        return candidates;
      }
    }
  } catch (err) {
    if (err.message && err.message.includes('non valida')) throw err;
    console.warn('Verifica modelli fallita, uso lista predefinita:', err.message);
  }

  // Fallback sicuro se offline o listModels non disponibile
  cachedCandidateModels = ['gemini-3.5-flash', 'gemini-flash-latest', 'gemini-2.5-flash'];
  return cachedCandidateModels;
}

// ANALISI AI VISION DEI RAPPORTINI MOD. 22.03
btnAnalyzeAI.addEventListener('click', async () => {
  if (state.pageImages.length === 0) {
    showToast('Nessun documento caricato da analizzare', 'error');
    return;
  }

  const apiKey = localStorage.getItem('gemini_api_key');
  if (!apiKey) {
    settingsModal.classList.add('active');
    showToast('Inserisci prima la tua Google Gemini API Key', 'error');
    return;
  }

  btnAnalyzeAI.disabled = true;
  btnAnalyzeAI.style.animation = 'none';
  let elapsed = 0;
  let activeModelText = 'AI';
  btnAnalyzeAI.innerHTML = `⏳ Analisi ${activeModelText} (0s)...`;
  const timerId = setInterval(() => {
    elapsed++;
    btnAnalyzeAI.innerHTML = `⏳ Analisi ${activeModelText} (${elapsed}s)...`;
  }, 1000);

  try {
    const analysisResult = await callGeminiVision(apiKey, state.pageImages, (modelName) => {
      activeModelText = modelName.replace('gemini-', '');
      btnAnalyzeAI.innerHTML = `⏳ Analisi ${activeModelText} (${elapsed}s)...`;
    }, state.backRotation);
    populateExtractedData(analysisResult);
    showToast(`Rapportino estratto con successo in ${elapsed}s con ${activeModelText}!`, 'success');
  } catch (err) {
    console.error('Errore analisi Gemini:', err);
    showToast(`Errore durante l'analisi: ${err.message}`, 'error');
  } finally {
    clearInterval(timerId);
    btnAnalyzeAI.disabled = false;
    btnAnalyzeAI.innerHTML = '🤖 Analizza con AI Vision';
  }
});

// ============================================================
// LETTURA IN DUE PASSAGGI INDIPENDENTI (FRONTE / RETRO) + UNIONE DETERMINISTICA NEL CODICE
// - Il FRONTE decide CHI è presente (solo X nella colonna P).
// - Il RETRO decide QUANTE ORE (blocchi con riquadro 1° TURNO / 2° TURNO).
// - Il codice unisce i due risultati: non può comparire un operaio senza X sul fronte.
// ============================================================

const FRONT_PROMPT = `Sei un sistema OCR di massima precisione. Ti viene fornita UNA SOLA immagine: il FRONTE (pagina 1) di un Rapportino di Cantiere Mod. 22.03 della ditta Vi.Cla Future.

COMPITO 1 - INTESTAZIONE
Leggi con cura l'intestazione:
- COMMESSA: il codice scritto a mano accanto alla scritta stampata "COMMESSA:". Spesso inizia per "CA" oppure "COMM." seguito da numeri e trattini o barre (es. "CA 4-3695/25", "CA 4-3696", "COMM. 0"). Trascrivi esattamente l'intero codice con lettere, trattini e barre (es. "CA 4-3695/25"). Non inventare e non confonderlo con orari.
- DATA: la data scritta a mano (es. "15/10/26", "16/10/26").
- GIORNO: solo il numero del giorno del mese (es. "15", "16").
- LOCALIZZAZIONE DEL SITO: la città/luogo scritto a mano sotto commessa (es. "CASERTA", "POZZUOLI").
- Caposquadra: il nome scritto a mano accanto a Caposquadra (es. "ROSSI", "BIANCHI M.").
- Responsabile di Commessa: il nome scritto a mano (es. "VERDI").

COMPITO 2 - OPERAI PRESENTI (colonna P)
- La tabella ha circa 30 righe numerate con il nome stampato di un operaio. Subito dopo il nome c'è una colonna stretta con intestazione "P" (Presenza).
- Esamina le righe UNA PER UNA, dall'alto verso il basso. Per ciascuna guarda SOLO la casella della colonna P.
- Una riga è "presente" SOLO se DENTRO la casella P c'è un segno tracciato a mano (X, crocetta, spunta, tratto). Se la casella è bianca la riga è ASSENTE.
- Di norma la maggior parte delle caselle è vuota e solo pochi operai sono presenti. Non indovinare: nel dubbio la riga è ASSENTE.
- Ignora completamente gli orari scritti nelle altre colonne (Dalle/Alle, orario partenza, ecc.): contano SOLO i segni dentro la colonna P.
- Per ogni riga presente trascrivi il nome stampato esattamente com'è scritto sul foglio. Se un operaio è stato aggiunto a mano in una riga vuota e ha la X, trascrivi il nome scritto a mano.
- Non usare elenchi esterni, non completare, non aggiungere nomi che non vedi con la X.

Restituisci ESCLUSIVAMENTE un oggetto JSON valido con questa struttura (i valori descrivono cosa inserire):
{
  "data": "data letta",
  "giorno": "numero del giorno del mese",
  "commessa": "codice commessa letto nell'intestazione",
  "cantiere": "localizzazione del sito letta",
  "caposquadra": "caposquadra letto",
  "responsabile": "responsabile di commessa letto",
  "totale_righe_con_X": "quante righe hanno la X nella colonna P",
  "righe_con_X_in_colonna_P": [
    { "numero_riga": "numero stampato a sinistra", "nome": "nome stampato o scritto in quella riga" }
  ]
}`;

const BACK_PROMPT = `Sei un sistema OCR di massima precisione. Ti viene fornita UNA SOLA immagine: il RETRO (pagina 2) di un Rapportino di Cantiere Mod. 22.03. La pagina potrebbe essere capovolta di 180 gradi: raddrizzala mentalmente prima di leggere.

Il retro contiene fino a DUE blocchi identici (uno in alto, uno in basso). Ogni blocco è composto da:
 (a) una tabella "TURNO DI LAVORO" con colonne Commessa | Entrata | Uscita | Sub-ore | A detrarre PAUSA | Parziali | Totale, e 2 righe (riga 1 = primo turno, riga 2 = secondo turno);
 (b) sotto la tabella, eventuali NOMI di operai scritti a mano (gli operai a cui si applica quel blocco) ed eventualmente il nome di una località;
 (c) a destra, il riquadro ore con DUE colonne "1° TURNO" e "2° TURNO" e le righe: Ord. D, Straord. D, Ord. N, Straord. N, Viaggio, ROL, TOTALE ORE.

Per OGNI blocco che contiene almeno un valore scritto a mano, trascrivi:
- nomi_operai: SOLO i nomi di persona scritti a mano sotto/accanto al blocco (le località vanno in "localita")
- commessa_turno1 / commessa_turno2: il codice di commessa scritto a mano nella colonna Commessa della riga 1 / riga 2 (es. "CA 4-3695", "CA 4-3696"). Le due righe possono avere commesse DIVERSE che differiscono anche di una sola cifra: leggi ogni riga cifra per cifra, senza copiare una riga nell'altra. Se la casella è vuota o contiene solo orari/sigle scrivi "" (stringa vuota). Non copiare nulla da altre parti del foglio e non inserire località.
- entrata_turno1 / uscita_turno1 / entrata_turno2 / uscita_turno2: gli orari scritti nelle colonne Entrata e Uscita della riga 1 / riga 2 della tabella TURNO DI LAVORO ("" se la casella è vuota).
- turno1 / turno2: i numeri scritti nella colonna "1° TURNO" / "2° TURNO" del riquadro ore, riga per riga: or = Ord. D, sd = Straord. D, on = Ord. N, sn = Straord. N, v = Viaggio, rol = ROL.
- totale_ore_riquadro: il numero scritto nella casella TOTALE ORE del riquadro (0 se non leggibile).

REGOLE DI LETTURA (tassative):
- Casella vuota = 0. La virgola è decimale ("0,5" significa 0.5).
- Un valore cancellato (barrato, scarabocchiato, annerito) NON vale: la casella conta come vuota. Trascrivi solo i valori leggibili e non cancellati.
- Un numero va in un campo SOLO se è scritto nella casella di quella riga E di quella colonna. Non spostare valori da una colonna all'altra, non inventare, non arrotondare, non usare 8 come valore predefinito.
- Un blocco senza nessun valore scritto va omesso.
- Trascrivi ogni casella così com'è scritta, anche se la somma non coincide con TOTALE ORE: non correggere, non togliere e non aggiungere valori per far quadrare i conti (le differenze vengono segnalate a parte).

Restituisci ESCLUSIVAMENTE un oggetto JSON valido con questa struttura (gli zeri sono solo la forma, vanno sostituiti con ciò che leggi):
{
  "pagina_capovolta": false,
  "blocchi": [
    {
      "posizione": "alto oppure basso",
      "nomi_operai": [],
      "localita": [],
      "commessa_turno1": "",
      "commessa_turno2": "",
      "entrata_turno1": "",
      "uscita_turno1": "",
      "entrata_turno2": "",
      "uscita_turno2": "",
      "turno1": { "or": 0, "sd": 0, "on": 0, "sn": 0, "v": 0, "rol": 0 },
      "turno2": { "or": 0, "sd": 0, "on": 0, "sn": 0, "v": 0, "rol": 0 },
      "totale_ore_riquadro": 0
    }
  ]
}`;

// Numero tollerante (virgola decimale, stringhe)
function toNum(v) {
  if (v === null || v === undefined) return 0;
  const n = parseFloat(String(v).replace(',', '.').replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

// Distanza di Levenshtein: tollera gli errori di lettura della scrittura a mano
function editDistance(a, b) {
  const m = a.length, n = b.length;
  if (!m) return n;
  if (!n) return m;
  const dp = Array.from({ length: m + 1 }, (_, i) => [i]);
  for (let j = 1; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      dp[i][j] = Math.min(
        dp[i - 1][j] + 1,
        dp[i][j - 1] + 1,
        dp[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1)
      );
    }
  }
  return dp[m][n];
}

// Un nome scritto a mano sul retro corrisponde a questo operaio?
function nameMatchesWorker(written, worker) {
  const found = findWorkerFromScannedText(String(written || ''));
  if (found) return found.index === worker.index;
  const { surname } = surnameAndInitial(worker);
  const tokens = String(written || '').toUpperCase().split(/\s+/)
    .map(t => t.replace(/[^A-Z]/g, ''))
    .filter(t => t.length >= 3);
  const tol = surname.length > 6 ? 2 : 1;
  return tokens.some(t => editDistance(t, surname) <= tol);
}

// Numero principale di una commessa: la sequenza di cifre più lunga ("CA 4-3695/25" -> "3695")
function commessaCore(s) {
  const runs = String(s || '').match(/\d+/g) || [];
  return runs.reduce((best, r) => (r.length > best.length ? r : best), '');
}

// Commessa del retro che differisce da quella del fronte per UNA sola cifra = errore di lettura
// della scrittura a mano (es. 3685 al posto di 3695): usa quella del fronte
function snapCommessaToFront(comm, commessaFronte) {
  const a = commessaCore(comm), b = commessaCore(commessaFronte);
  if (!a || !b || a === b || a.length !== b.length || a.length < 3) return comm;
  let diff = 0;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) diff++;
  return diff === 1 ? commessaFronte : comm;
}

// Firma di una commessa: sigla iniziale + numeri ("CA 4 3695-25" e "CA 04-3695/25" -> CA [4,3695,25])
function commessaSignature(s) {
  const u = String(s || '').toUpperCase();
  return {
    letters: (u.match(/[A-Z]+/) || [''])[0],
    nums: (u.match(/\d+/g) || []).map(n => String(parseInt(n, 10))),
    plain: u.replace(/[^A-Z0-9]/g, '')
  };
}

// Riconduce la commessa scritta a mano alla voce PIÙ SIMILE dell'elenco commesse del file Excel (rif.cantieri):
// in tabella e in Excel va sempre la voce dell'elenco. Confronta sigla + numeri (l'anno può mancare nella scritta a mano);
// a parità di somiglianza preferisce la commessa del fronte. Resta la lettura originale solo se nessuna voce è vicina.
// Restituisce { value, inList, correctedFrom? } (correctedFrom presente se la voce scelta non coincide esattamente)
function matchCommessaToList(read, list, preferred = '') {
  const raw = String(read || '').trim();
  if (!raw || !Array.isArray(list) || list.length === 0) return { value: raw, inList: false };
  const r = commessaSignature(raw);
  const entries = list.map(e => ({ e, s: commessaSignature(e) }));

  // Voci senza numeri (FERIE, MALATTIA, SMONTO NOTTE...): confronto sul testo
  if (r.nums.length === 0) {
    const hit = entries.find(x => x.s.plain === r.plain);
    return hit ? { value: hit.e, inList: true } : { value: raw, inList: false };
  }

  // Confronto sulla sola sequenza di cifre: trattini e spazi a mano sono irregolari ("CA4-3696" = "CA 4 3696" = "CA43696")
  const readDigits = r.nums.join('');
  const scored = [];
  entries.forEach(x => {
    if (x.s.nums.length === 0) return;
    if (r.letters && x.s.letters && x.s.letters !== r.letters) return; // CA non si confonde con COMM
    // L'anno può mancare nella scritta a mano: prova la voce con e senza l'ultimo numero
    const variants = [x.s.nums.join('')];
    if (x.s.nums.length > 1) variants.push(x.s.nums.slice(0, -1).join(''));
    const dist = Math.min(...variants.map(v => editDistance(readDigits, v)));
    scored.push({ e: x.e, dist, isPreferred: !!preferred && x.e === preferred });
  });
  if (scored.length === 0) return { value: raw, inList: false };

  // Ogni commessa scritta a mano è sempre presente nell'elenco Excel: si sceglie sempre la voce più vicina.
  // A pari merito vince la commessa del fronte; se nemmeno quella decide, si prende la prima e lo si segnala
  const minDist = Math.min(...scored.map(s => s.dist));
  const ties = scored.filter(s => s.dist === minDist);
  const best = ties.find(s => s.isPreferred) || ties[0];
  const ambiguous = !best.isPreferred && ties.length > 1;
  if (minDist === 0 && !ambiguous) return { value: best.e, inList: true };
  return { value: best.e, inList: true, correctedFrom: raw, ambiguous };
}

function sumHours(h) {
  return (h.or || 0) + (h.sd || 0) + (h.on || 0) + (h.sn || 0) + (h.v || 0) + (h.rol || 0);
}

// Unisce lettura del fronte (CHI) e del retro (QUANTE ORE)
function mergeFrontAndBack(front, back, commesseList = []) {
  front = front || {};
  const notes = [];
  // Correzioni di una cifra fatte con l'elenco commesse: sempre segnalate perché la commessa vera potrebbe mancare dall'elenco
  const corrections = new Map();
  let commessaFronte = '';
  const toList = (read) => {
    const m = matchCommessaToList(read, commesseList, commessaFronte);
    if (m.correctedFrom) corrections.set(m.correctedFrom, m);
    return m.value;
  };
  commessaFronte = toList(String(front.commessa || '').trim());

  // 1. Operai presenti: SOLO quelli con X sul fronte
  const righe = Array.isArray(front.righe_con_X_in_colonna_P) ? front.righe_con_X_in_colonna_P : [];
  const workers = [];
  const seen = new Set();
  const unknown = [];
  righe.forEach(r => {
    const nome = String((r && (r.nome || r.nomeLetto)) || '').trim();
    if (!nome) return;
    const w = findWorkerFromScannedText(nome);
    if (!w) { unknown.push(nome); return; }
    if (seen.has(w.index)) return;
    seen.add(w.index);
    workers.push({ worker: w, nome });
  });
  const declared = parseInt(front.totale_righe_con_X);
  if (!isNaN(declared) && declared !== righe.length) {
    notes.push(`Attenzione: dichiarate ${declared} righe con X ma elencate ${righe.length}`);
  }
  if (unknown.length) {
    notes.push(`Nomi con X non riconosciuti nel roster (da aggiungere a mano): ${unknown.join(', ')}`);
  }

  // 2. Blocchi ore dal retro (solo quelli con valori)
  const blocchi = (back && Array.isArray(back.blocchi)) ? back.blocchi : [];
  const mk = t => ({
    or: toNum(t && t.or), sd: toNum(t && t.sd), on: toNum(t && t.on),
    sn: toNum(t && t.sn), v: toNum(t && t.v), rol: toNum(t && t.rol), malattia: 0
  });
  const tables = blocchi.map(b => ({
    posizione: String(b.posizione || '').trim() || '?',
    nomi: Array.isArray(b.nomi_operai) ? b.nomi_operai.map(String) : [],
    localita: Array.isArray(b.localita) ? b.localita.map(String) : [],
    comm1: toList(snapCommessaToFront(String(b.commessa_turno1 || '').trim(), commessaFronte)),
    comm2: toList(snapCommessaToFront(String(b.commessa_turno2 || '').trim(), commessaFronte)),
    t1: mk(b.turno1),
    t2: mk(b.turno2),
    riga1: !!(String(b.entrata_turno1 || '').trim() || String(b.uscita_turno1 || '').trim()),
    riga2: !!(String(b.entrata_turno2 || '').trim() || String(b.uscita_turno2 || '').trim()),
    totale: toNum(b.totale_ore_riquadro)
  })).filter(tb => sumHours(tb.t1) + sumHours(tb.t2) > 0);

  // Turno unico: se la riga 2 della tabella TURNO DI LAVORO non ha orari, le ore scritte nella
  // colonna "2° TURNO" sono una correzione del 1° turno (es. valore cancellato e riscritto accanto)
  tables.forEach(tb => {
    if (tb.riga1 && !tb.riga2 && sumHours(tb.t2) > 0) {
      ['or', 'sd', 'on', 'sn', 'v', 'rol'].forEach(k => { tb.t1[k] += tb.t2[k]; tb.t2[k] = 0; });
      notes.push(`Blocco ${tb.posizione}: ore in colonna 2° turno ma un solo turno con orari → riportate nel 1° turno, verifica`);
    }
  });

  tables.forEach(tb => {
    const somma = sumHours(tb.t1) + sumHours(tb.t2);
    if (tb.totale > 0 && Math.abs(somma - tb.totale) > 0.01) {
      notes.push(`Blocco ${tb.posizione}: somma ore lette ${somma} ≠ TOTALE ORE scritto ${tb.totale} → ricontrolla`);
    }
  });

  // 3. Un record per ogni (operaio con X) x (colonna turno compilata nel suo blocco)
  const records = [];
  workers.forEach(({ worker, nome }) => {
    // Anche le "località": un cognome che è pure un paese (es. SORRENTO) a volte finisce lì.
    // Servono solo ad abbinare gli operai con la X sul fronte, quindi non possono aggiungerne di nuovi
    let table = tables.find(tb => tb.nomi.some(n => nameMatchesWorker(n, worker))) ||
      tables.find(tb => tb.localita.some(n => nameMatchesWorker(n, worker)));
    if (!table) {
      if (tables.length === 1) {
        table = tables[0];
      } else if (tables.length > 1) {
        table = tables.find(tb => tb.nomi.length === 0) || tables[0];
        notes.push(`${worker.fullName}: nome non trovato sul retro, assegnato al blocco ${table.posizione}`);
      }
    }

    const base = { nomeLetto: nome, workerName: worker.name, workerIndex: worker.index };
    if (!table) {
      records.push({ ...base, turno: 1, commessa: commessaFronte, or: 0, sd: 0, on: 0, sn: 0, v: 0, rol: 0, malattia: 0 });
      notes.push(`${worker.fullName}: nessuna ora trovata sul retro`);
      return;
    }
    [1, 2].forEach(t => {
      const h = t === 1 ? table.t1 : table.t2;
      if (sumHours(h) <= 0) return;
      records.push({
        ...base,
        turno: t,
        commessa: (t === 1 ? table.comm1 : table.comm2) || commessaFronte,
        or: h.or, sd: h.sd, on: h.on, sn: h.sn, v: h.v, rol: h.rol, malattia: 0
      });
    });
  });

  if (!back) notes.push('Retro non caricato: ore non lette, inseriscile a mano');
  else if (tables.length === 0) notes.push('Nessun valore ore trovato sul retro');

  if (commesseList.length > 0) {
    const fuoriElenco = [...new Set(records.map(r => r.commessa).filter(c => c && !commesseList.includes(c)))];
    if (fuoriElenco.length) notes.push(`Commessa non presente nell'elenco Excel: ${fuoriElenco.join(', ')} → verifica`);
  }
  corrections.forEach((m, from) => notes.push(m.ambiguous
    ? `Commessa letta "${from}" ugualmente simile a più voci dell'elenco Excel: scelta "${m.value}" → verifica`
    : `Commessa letta "${from}" ricondotta a "${m.value}" (elenco Excel)`));

  return {
    data: front.data,
    giorno: parseInt(front.giorno) || undefined,
    commessa: commessaFronte,
    cantiere: front.cantiere,
    caposquadra: front.caposquadra,
    responsabile: front.responsabile,
    note_particolari: notes.join(' | ') || `${tables.length} blocco/i ore letti sul retro`,
    // Ogni nota è un punto da verificare: nella scansione multipla un rapportino con dubbi aspetta la conferma
    dubbi: notes.slice(),
    operai_con_spunta_X_rilevati: workers.map(w => w.nome),
    records
  };
}

// Chiamata generica a Gemini Vision con fallback tra modelli
// Attesa con possibilità di mostrare il conto alla rovescia
const sleepMs = ms => new Promise(res => setTimeout(res, ms));

// Wrapper: se Google risponde 429 (quota/minuto) attende e riprova automaticamente
async function callGeminiJSON(apiKey, prompt, images, onModelTry) {
  const MAX_ROUNDS = 4;
  for (let round = 1; round <= MAX_ROUNDS; round++) {
    try {
      return await callGeminiJSONOnce(apiKey, prompt, images, onModelTry);
    } catch (e) {
      if (e && e.isRateLimit && round < MAX_ROUNDS) {
        const wait = Math.min(Math.max(e.retryAfterMs || 15000, 8000), 30000);
        console.warn(`Quota Google (429): attendo ${Math.round(wait / 1000)}s e riprovo (tentativo ${round}/${MAX_ROUNDS - 1})`);
        showToast(`Quota Google al limite: nuovo tentativo automatico tra ${Math.round(wait / 1000)}s...`, 'info');
        await sleepMs(wait);
        continue;
      }
      throw e;
    }
  }
}

async function callGeminiJSONOnce(apiKey, prompt, images, onModelTry) {
  const imageParts = images.map(imgDataUrl => {
    const base64Data = imgDataUrl.split(',')[1];
    const mimeType = imgDataUrl.split(';')[0].split(':')[1];
    return { inlineData: { data: base64Data, mimeType: mimeType } };
  });

  const candidateModels = await resolveCandidateModels(apiKey);
  let lastError = null;
  let rateLimited = 0;
  let retryAfterMs = 0;

  for (const modelName of candidateModels) {
    if (onModelTry) onModelTry(modelName);

    // Timeout di sicurezza per singolo tentativo
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 40000);

    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent`;
      const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'x-goog-api-key': apiKey },
        signal: controller.signal,
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }, ...imageParts] }],
          generationConfig: {
            temperature: 0.0,
            responseMimeType: "application/json",
            // Ragionamento ridotto: stessa precisione sui rapportini, risposta molto più rapida (non supportato dai modelli 2.x)
            ...(/^gemini-2\./.test(modelName) ? {} : { thinkingConfig: { thinkingLevel: 'low' } })
          }
        })
      });

      clearTimeout(timeoutId);

      if (response.ok) {
        const jsonResponse = await response.json();
        let textOutput = jsonResponse.candidates?.[0]?.content?.parts?.[0]?.text?.trim();
        if (!textOutput) {
          throw new Error(`Risposta vuota ricevuta da ${modelName}`);
        }
        if (textOutput.startsWith('```')) {
          textOutput = textOutput.replace(/^```[a-z]*\s*/i, '').replace(/```\s*$/i, '').trim();
        }
        return JSON.parse(textOutput);
      } else {
        const errText = await response.text();
        console.warn(`Tentativo con ${modelName} fallito (${response.status}):`, errText);

        let parsedMsg = errText;
        try {
          const errObj = JSON.parse(errText);
          if (errObj.error?.message) parsedMsg = errObj.error.message;
        } catch (_) {}

        if (response.status === 429) {
          // Ogni modello ha una quota propria: prova il successivo prima di arrendersi
          rateLimited++;
          const m = /retry in ([\d.]+)s/i.exec(parsedMsg) || /"retryDelay":\s*"(\d+)s"/i.exec(errText);
          if (m) retryAfterMs = Math.max(retryAfterMs, Math.ceil(parseFloat(m[1]) * 1000) + 1000);
          lastError = new Error(`Limite quota Google raggiunto (Errore 429: Too Many Requests).`);
          continue;
        }
        if (response.status === 400 && parsedMsg.includes('API_KEY_INVALID')) {
          throw new Error(`Chiave Google API non valida. Verifica la chiave nelle Impostazioni.`);
        }

        lastError = new Error(`Errore API [${modelName}] (${response.status}): ${parsedMsg}`);
      }
    } catch (e) {
      clearTimeout(timeoutId);
      if (e.name === 'AbortError') {
        console.warn(`Timeout scaduto (40s) con il modello ${modelName}`);
        lastError = new Error(`Il modello ${modelName} non ha risposto entro 40 secondi.`);
      } else {
        lastError = e;
        // Se è un errore di quota o chiave non valida, non ciclare inutilmente
        if (e.message && (e.message.includes('429') || e.message.includes('non valida'))) {
          throw e;
        }
      }
    }
  }

  if (rateLimited > 0 && rateLimited === candidateModels.length) {
    const err = new Error(`Limite quota Google raggiunto (Errore 429: Too Many Requests) su tutti i modelli. Attendi 1 minuto e riprova: la quota gratuita è al limite.`);
    err.isRateLimit = true;
    err.retryAfterMs = retryAfterMs;
    throw err;
  }
  throw lastError || new Error('Nessun modello Gemini ha risposto con successo');
}

// Ruota un'immagine (dataURL) di 90/180/270 gradi in senso orario
function rotateImage(dataUrl, degrees) {
  degrees = ((degrees % 360) + 360) % 360;
  if (!dataUrl || degrees === 0) return Promise.resolve(dataUrl);
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      const sideways = degrees % 180 !== 0;
      const canvas = document.createElement('canvas');
      canvas.width = sideways ? img.height : img.width;
      canvas.height = sideways ? img.width : img.height;
      const ctx = canvas.getContext('2d');
      ctx.translate(canvas.width / 2, canvas.height / 2);
      ctx.rotate(degrees * Math.PI / 180);
      ctx.drawImage(img, -img.width / 2, -img.height / 2);
      resolve(canvas.toDataURL('image/jpeg', 0.85));
    };
    img.onerror = () => resolve(dataUrl);
    img.src = dataUrl;
  });
}

// Chiamata all'API Gemini Vision: fronte e retro letti separatamente, in parallelo
// Il retro viene raddrizzato prima dell'invio: capovolto viene letto peggio e più lentamente
async function callGeminiVision(apiKey, images, onModelTry, backRotation = 180) {
  const frontImg = images[0];
  const backImg = images[1] ? await rotateImage(images[1], backRotation) : null;

  // L'elenco commesse aiuta a leggere le cifre scritte a mano in modo ambiguo
  const [commesseList] = await Promise.all([loadCommesseList(), loadWorkersRoster()]);
  const hint = commesseList.length
    ? `\n\nELENCO COMMESSE VALIDE (dal file Excel aziendale): ${commesseList.join(' ; ')}\nLa commessa scritta a mano corrisponde di norma a una di queste: usale per interpretare le cifre poco chiare, ma trascrivi ciò che è scritto.`
    : '';

  const [front, back] = await Promise.all([
    callGeminiJSON(apiKey, FRONT_PROMPT + hint, [frontImg], onModelTry),
    backImg ? sleepMs(1500).then(() => callGeminiJSON(apiKey, BACK_PROMPT + hint, [backImg], onModelTry)) : Promise.resolve(null)
  ]);

  console.log('Lettura fronte:', front);
  console.log('Lettura retro:', back);
  return mergeFrontAndBack(front, back, commesseList);
}

// Operai dalla tabella presenze del file Excel attivo (riletti a ogni analisi e al cambio di file:
// assunzioni, uscite e sostituzioni fatte in Excel valgono subito)
async function loadWorkersRoster() {
  if (!cartellaPronta) return;
  // File attivo, oppure il Master se nel file attivo l'elenco non è leggibile
  for (const fileName of [...new Set([state.selectedExcelPath, state.masterFileName].filter(Boolean))]) {
    try {
      const workers = await ExcelXlsm.readWorkers(await CartellaExcel.readBytes(fileName));
      const changed = JSON.stringify(workers) !== JSON.stringify(OPERAI_ROSTER);
      OPERAI_ROSTER = workers;
      if (changed && state.tableData.length > 0) renderPresenzeTable();
      return;
    } catch (err) {
      console.warn(`Elenco operai da ${fileName} non disponibile:`, err.message);
    }
  }
}

// Elenco commesse (intervallo "rif.cantieri") del file attivo unito a quello del Master: un file mese creato
// prima di un aggiornamento dell'elenco non deve far perdere commesse valide. Riletto a ogni analisi
async function loadCommesseList() {
  if (!cartellaPronta) return [];
  const commesse = [];
  for (const fileName of [...new Set([state.selectedExcelPath, state.masterFileName].filter(Boolean))]) {
    try {
      (await ExcelXlsm.readCommesse(await CartellaExcel.readBytes(fileName)))
        .forEach(c => { if (!commesse.includes(c)) commesse.push(c); });
    } catch (err) {
      console.warn(`Elenco commesse da ${fileName} non disponibile:`, err.message);
    }
  }
  return commesse;
}

// Cognome e iniziale del nome dal nome abbreviato di Excel: "DE SANTIS C." -> DESANTIS + C, "D'ANGELO M." -> DANGELO + M
function surnameAndInitial(op) {
  const parts = op.name.toUpperCase().trim().split(/\s+/);
  const initial = parts.length > 1 && /^[A-Z]\.?$/.test(parts[parts.length - 1]) ? parts.pop()[0] : '';
  return { surname: parts.join('').replace(/[^A-Z]/g, ''), initial };
}

// Cerca il cognome di ogni operaio tra le parole lette ("DE SANTIS" anche come "DESANTIS", "D'ANGELO" anche "DANGELO")
// e verifica che la parola successiva inizi con l'iniziale del nome abbreviato di Excel ("ROSSI A.").
// Con fuzzy = true tollera lettere sbagliate o in più (scritta a mano, o cognome scritto in modo diverso
// sul rapportino e in Excel: 1 lettera per cognomi da 5-6 lettere, 2 da 7 in su).
// Restituisce l'operaio solo se la scelta è univoca; altrimenti null (decidono i controlli successivi)
function matchWorkerBySurname(raw, fuzzy = false) {
  const tokens = raw.split(/[^A-Z']+/).map(t => t.replace(/'/g, '')).filter(Boolean);
  const hits = [];
  for (const op of OPERAI_ROSTER) {
    if (/\d/.test(op.name)) continue; // posti di riserva tipo "SCORTA 1"
    const { surname, initial } = surnameAndInitial(op);
    if (surname.length < 3) continue;
    const tol = !fuzzy ? 0 : surname.length >= 7 ? 2 : surname.length >= 5 ? 1 : 0;
    if (fuzzy && tol === 0) continue;

    for (let i = 0; i < tokens.length; i++) {
      let acc = '';
      for (let j = i; j < tokens.length && acc.length <= surname.length + tol; j++) {
        acc += tokens[j];
        if (Math.abs(acc.length - surname.length) > tol) continue;
        const dist = tol ? editDistance(acc, surname) : (acc === surname ? 0 : 1);
        if (dist <= tol) {
          const next = tokens[j + 1];
          if (!initial || !next || next[0] === initial) {
            hits.push({ op, dist, initialOk: !!(initial && next && next[0] === initial) });
          }
          break;
        }
      }
    }
  }
  const confirmed = hits.filter(h => h.initialOk);
  let pool = confirmed.length ? confirmed : hits;
  if (pool.length) {
    const best = Math.min(...pool.map(h => h.dist));
    pool = pool.filter(h => h.dist === best);
  }
  const unique = [...new Set(pool.map(h => h.op))];
  return unique.length === 1 ? unique[0] : null;
}

// Riconciliazione avanzata e deterministica del nome letto sul foglio cartaceo con OPERAI_ROSTER
function findWorkerFromScannedText(text) {
  if (!text || typeof text !== 'string') return null;
  const raw = text.toUpperCase().trim();
  const normalized = raw.replace(/[^A-Z]/g, '');

  // 1. Corrispondenza diretta esatta (completa o abbreviata); le cifre contano ("SCORTA 1" ≠ "SCORTA 2")
  const normDigits = raw.replace(/[^A-Z0-9]/g, '');
  for (const op of OPERAI_ROSTER) {
    const fullNorm = op.fullName.toUpperCase().replace(/[^A-Z0-9]/g, '');
    const shortNorm = op.name.toUpperCase().replace(/[^A-Z0-9]/g, '');
    if (normDigits === fullNorm || normDigits === shortNorm) return op;
  }

  // 1b. Cognome + iniziale del nome, per qualsiasi operaio dell'elenco Excel (anche appena aggiunto):
  // "ROSSI MARIO" -> "ROSSI M.", "BIANCHI MARCO" -> "BIANCHI M." e non "BIANCHI C."
  const bySurname = matchWorkerBySurname(raw);
  if (bySurname) return bySurname;

  // 2. Cognome con qualche lettera diversa (scritta a mano, o cognome scritto diversamente sul rapportino e in Excel)
  const byFuzzySurname = matchWorkerBySurname(raw, true);
  if (byFuzzySurname) return byFuzzySurname;

  // 3. Ultima possibilità: il cognome compare dentro il testo letto
  for (const op of OPERAI_ROSTER) {
    if (/\d/.test(op.name)) continue;
    const { surname } = surnameAndInitial(op);
    if (surname.length > 3 && normalized.includes(surname)) return op;
  }

  return null;
}

// Funzione di ordinamento alfabetico A-Z per gli operai e normalizzazione turni
function sortRecordsAlphabetically(records) {
  // Se un operaio compare più volte (es. doppio turno), garantisce tassativamente che il primo record sia Turno 1 e il secondo Turno 2
  const workerGroups = new Map();
  records.forEach(rec => {
    if (!rec.workerIndex) return; // righe aggiunte a mano senza operaio scelto
    const key = rec.workerIndex;
    if (!workerGroups.has(key)) workerGroups.set(key, []);
    workerGroups.get(key).push(rec);
  });

  workerGroups.forEach(group => {
    if (group.length >= 2) {
      group[0].turno = 1;
      group[1].turno = 2;
    } else if (group.length === 1) {
      group[0].turno = group[0].turno || 1;
    }
  });

  records.sort((a, b) => {
    // Poiché OPERAI_ROSTER è già in ordine alfabetico da 1 a 34, ordinare per workerIndex ordina da A a Z
    // Le righe senza operaio scelto restano in fondo
    const ia = a.workerIndex || Infinity, ib = b.workerIndex || Infinity;
    if (ia !== ib) {
      return ia === Infinity ? 1 : ib === Infinity ? -1 : ia - ib;
    }
    // A parità di operaio (doppio turno), metti prima Turno 1 e poi Turno 2
    return (a.turno || 1) - (b.turno || 1);
  });
}

// Popolamento form e tabella con i dati estratti
function populateExtractedData(data) {
  if (data.data) metaDate.value = data.data;
  if (data.giorno) metaDay.value = data.giorno;
  if (data.commessa) metaCommessa.value = data.commessa;
  if (data.cantiere) metaCantiere.value = data.cantiere;
  if (data.caposquadra) metaCaposquadra.value = data.caposquadra;
  if (data.responsabile) metaRespCommessa.value = data.responsabile;

  const { records, removed } = buildRecordsFromAnalysis(data);
  if (removed > 0) showToast(`Scartati ${removed} record di operai senza X nella colonna P`, 'info');
  state.tableData = records;

  // Mostra banner riassuntivo con operai rilevati e note su turni/cantieri
  if (data.note_particolari || (data.operai_con_spunta_X_rilevati && data.operai_con_spunta_X_rilevati.length > 0)) {
    detectionBanner.style.display = 'flex';
    bannerTitle.textContent = `Analisi AI Completata • ${state.tableData.length} turni/record generati`;
    const operaiStr = (data.operai_con_spunta_X_rilevati || []).join(', ');
    const noteStr = data.note_particolari ? `Note retro: ${data.note_particolari}` : '';
    bannerDesc.textContent = [operaiStr ? `Operai fronte ('X'): ${operaiStr}` : '', noteStr].filter(Boolean).join(' • ');
  } else {
    detectionBanner.style.display = 'none';
  }

  renderPresenzeTable();
}

// Righe della tabella pronte per Excel a partire dal risultato dell'analisi, senza toccare lo schermo:
// la usa anche la scansione multipla, che lavora in sottofondo mentre controlli un altro rapportino
function buildRecordsFromAnalysis(data) {
  let rawRecords = (data.records || []).map(r => ({ ...r }));
  const records = [];
  let removed = 0;

  // 🛡️ FILTRO DI SICUREZZA: tieni SOLO gli operai che l'AI stessa ha elencato con la X nella colonna P del fronte
  const xList = Array.isArray(data.operai_con_spunta_X_rilevati) ? data.operai_con_spunta_X_rilevati : [];
  const allowedIdx = new Set();
  xList.forEach(n => {
    const w = findWorkerFromScannedText(String(n || ''));
    if (w) allowedIdx.add(w.index);
  });
  if (allowedIdx.size > 0) {
    const before = rawRecords.length;
    rawRecords = rawRecords.filter(rec => {
      const w = findWorkerFromScannedText(rec.nomeLetto || rec.nomeLettoSulFoglio || rec.workerName || '');
      return w ? allowedIdx.has(w.index) : false;
    });
    removed = before - rawRecords.length;
  }

  rawRecords.forEach(rec => {
    // 1. Riconcilia il lavoratore basandosi sul testo letto visivamente dal foglio
    const scannedText = rec.nomeLetto || rec.nomeLettoSulFoglio || rec.workerName || '';
    const matched = findWorkerFromScannedText(scannedText);

    if (matched) {
      rec.workerIndex = matched.index;
      rec.workerName = matched.name;
    } else if (rec.workerIndex && OPERAI_ROSTER.find(op => op.index === rec.workerIndex)) {
      // Fallback su indice se non c'è corrispondenza testuale
      rec.workerName = OPERAI_ROSTER.find(op => op.index === rec.workerIndex).name;
    }

    // Assicura turno numerico 1 o 2
    rec.turno = parseInt(rec.turno) === 2 ? 2 : 1;

    // ⛔ DEVE RIPORTARE SOLO LA COMMESSA DEL PRIMO FOGLIO (es. 'CA 4-3695/25'), MAI CANTIERE O ARTEFATTI
    const generalCommessa = (data.commessa || '').trim();
    const siteName = (data.cantiere || '').trim().toUpperCase();

    let commVal = (rec.commessa || rec.cantiere || '').trim();
    
    // Verifica se commVal è una commessa valida specifica del turno o se deve usare generalCommessa del fronte:
    // Deve contenere almeno lettere e numeri tipo 'CA' o 'COMM' o un trattino/barra, e NON essere solo 4 o 5 cifre isolate (tipo '65100' orari o CAP)
    const isInvalidOrArtefact = (val) => {
      if (!val) return true;
      const u = val.toUpperCase().trim();
      if (siteName && (u === siteName || u.includes(siteName) || siteName.includes(u))) return true;
      if (!/\d/.test(u)) return true; // nessun numero -> non è commessa
      if (/^\d{4,5}$/.test(u)) return true; // solo 4-5 cifre tipo 65100 o CAP/orario -> artefatto OCR del retro
      return false;
    };

    if (isInvalidOrArtefact(commVal) || !commVal) {
      commVal = generalCommessa;
    }

    rec.commessa = commVal;
    rec.cantiere = commVal; // sincronizzato con rec.cantiere per la scrittura su colonna C di Excel

    records.push(rec);
  });

  // Ordina sempre rigorosamente in ordine alfabetico A-Z
  sortRecordsAlphabetically(records);
  return { records, removed };
}

// Renderizzazione Tabella Presenze
function renderPresenzeTable() {
  tableBody.innerHTML = '';

  if (state.tableData.length === 0) {
    tableBody.innerHTML = `<tr><td colspan="10" style="text-align: center; color: var(--text-dim); padding: 30px;">Nessun operaio presente nella tabella.</td></tr>`;
    updateSummaryStats();
    return;
  }

  // Numero progressivo per operaio: il secondo turno dello stesso operaio non riceve un nuovo numero
  const workerNumbers = new Map();

  state.tableData.forEach((rec, idx) => {
    const tr = document.createElement('tr');
    tr.classList.toggle('turno-2', rec.turno === 2);

    const isNewWorker = !!rec.workerIndex && !workerNumbers.has(rec.workerIndex);
    if (isNewWorker) workerNumbers.set(rec.workerIndex, workerNumbers.size + 1);

    // Calcolo Totale Ore
    const or = parseFloat(rec.or) || 0;
    const sd = parseFloat(rec.sd) || 0;
    const on = parseFloat(rec.on) || 0;
    const sn = parseFloat(rec.sn) || 0;
    const v = parseFloat(rec.v) || 0;
    const tot = or + sd + on + sn + v;

    let totClass = 'normal';
    if (tot > 8 && tot <= 10) totClass = 'overtime';
    if (tot > 10) totClass = 'high';

    tr.innerHTML = `
      <td>
        <div class="worker-cell">
          <span class="worker-num">${isNewWorker ? workerNumbers.get(rec.workerIndex) : ''}</span>
          <select class="form-input worker-select" data-idx="${idx}" style="font-weight:600;">
            ${rec.workerIndex ? '' : '<option value="0" selected disabled>— Scegli operaio —</option>'}
            ${OPERAI_ROSTER.map(op => `<option value="${op.index}" ${op.index === rec.workerIndex ? 'selected' : ''}>${op.fullName}</option>`).join('')}
          </select>
        </div>
      </td>
      <td style="text-align: center;">
        <select class="form-input turno-select" data-idx="${idx}" style="width: 75px; font-weight: 600;">
          <option value="1" ${rec.turno === 1 ? 'selected' : ''}>Turno 1</option>
          <option value="2" ${rec.turno === 2 ? 'selected' : ''}>Turno 2</option>
        </select>
      </td>
      <td>
        <input type="text" class="form-input cantiere-input" data-idx="${idx}" value="${rec.commessa || rec.cantiere || metaCommessa.value || ''}" placeholder="COMMESSA">
      </td>
      <td style="text-align: center;"><input type="number" step="0.5" class="hour-input val-or" data-idx="${idx}" value="${or || ''}"></td>
      <td style="text-align: center;"><input type="number" step="0.5" class="hour-input val-sd" data-idx="${idx}" value="${sd || ''}"></td>
      <td style="text-align: center;"><input type="number" step="0.5" class="hour-input val-on" data-idx="${idx}" value="${on || ''}"></td>
      <td style="text-align: center;"><input type="number" step="0.5" class="hour-input val-sn" data-idx="${idx}" value="${sn || ''}"></td>
      <td style="text-align: center;"><input type="number" step="0.5" class="hour-input val-v" data-idx="${idx}" value="${v || ''}"></td>
      <td style="text-align: center;"><span class="tot-badge ${totClass}">${tot > 0 ? tot : '0'}</span></td>
      <td style="text-align: center;">
        <button class="tool-btn btn-del" data-idx="${idx}" title="Rimuovi" style="color: var(--danger); width: 28px; height: 28px;">✕</button>
      </td>
    `;
    tableBody.appendChild(tr);
  });

  attachTableEventListeners();
  updateSummaryStats();
}

function attachTableEventListeners() {
  document.querySelectorAll('.worker-select').forEach(sel => {
    sel.addEventListener('change', (e) => {
      const idx = parseInt(e.target.dataset.idx);
      const chosen = OPERAI_ROSTER.find(op => op.index === parseInt(e.target.value));
      if (chosen) {
        state.tableData[idx].workerIndex = chosen.index;
        state.tableData[idx].workerName = chosen.name;
        // Riordina e rinumera: l'operaio scelto va al suo posto alfabetico
        sortRecordsAlphabetically(state.tableData);
        renderPresenzeTable();
      }
    });
  });

  document.querySelectorAll('.turno-select').forEach(sel => {
    sel.addEventListener('change', (e) => {
      const idx = parseInt(e.target.dataset.idx);
      state.tableData[idx].turno = parseInt(e.target.value);
      e.target.closest('tr').classList.toggle('turno-2', state.tableData[idx].turno === 2);
    });
  });

  document.querySelectorAll('.cantiere-input').forEach(inp => {
    inp.addEventListener('input', (e) => {
      const idx = parseInt(e.target.dataset.idx);
      state.tableData[idx].commessa = e.target.value;
      state.tableData[idx].cantiere = e.target.value;
    });
  });

  ['val-or', 'val-sd', 'val-on', 'val-sn', 'val-v'].forEach(cls => {
    document.querySelectorAll(`.${cls}`).forEach(inp => {
      inp.addEventListener('input', (e) => {
        const idx = parseInt(e.target.dataset.idx);
        const field = cls.replace('val-', '');
        state.tableData[idx][field] = parseFloat(e.target.value) || 0;
        
        // Ricalcola badge riga
        const rec = state.tableData[idx];
        const tot = (rec.or || 0) + (rec.sd || 0) + (rec.on || 0) + (rec.sn || 0) + (rec.v || 0);
        const row = e.target.closest('tr');
        const badge = row.querySelector('.tot-badge');
        badge.textContent = tot;
        badge.className = `tot-badge ${tot > 10 ? 'high' : tot > 8 ? 'overtime' : 'normal'}`;
        updateSummaryStats();
      });
    });
  });

  document.querySelectorAll('.btn-del').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const idx = parseInt(e.currentTarget.dataset.idx);
      state.tableData.splice(idx, 1);
      renderPresenzeTable();
    });
  });
}

function updateSummaryStats() {
  statWorkers.textContent = state.tableData.length;
  let totalHours = 0;
  state.tableData.forEach(r => {
    totalHours += (parseFloat(r.or) || 0) + (parseFloat(r.sd) || 0) + (parseFloat(r.on) || 0) + (parseFloat(r.sn) || 0) + (parseFloat(r.v) || 0);
  });
  statHours.textContent = totalHours.toFixed(1) + 'h';
}

// Aggiungi riga manuale
// La riga nasce senza operaio e in fondo alla tabella: con un nome predefinito finiva in cima e passava inosservata
btnAddRow.addEventListener('click', () => {
  state.tableData.push({
    workerIndex: 0,
    workerName: '',
    turno: 1,
    commessa: metaCommessa.value || "COMM. 0",
    cantiere: metaCommessa.value || "COMM. 0",
    or: 8,
    sd: 0,
    on: 0,
    sn: 0,
    v: 0,
    rol: 0,
    malattia: 0
  });
  renderPresenzeTable();
  const selects = tableBody.querySelectorAll('.worker-select');
  const newSelect = selects[selects.length - 1];
  if (newSelect) {
    newSelect.scrollIntoView({ block: 'center', behavior: 'smooth' });
    newSelect.focus();
  }
});

// Svuota tutto: tabella, dati del rapportino e documento caricato (pronto per il rapportino successivo)
btnClearTable.addEventListener('click', () => {
  if (!confirm('Vuoi davvero svuotare tutto? Verranno tolti la tabella, i dati del rapportino e il documento caricato.')) return;

  state.tableData = [];
  renderPresenzeTable();

  [metaDate, metaDay, metaCommessa, metaCantiere, metaCaposquadra, metaRespCommessa].forEach(el => { if (el) el.value = ''; });
  detectionBanner.style.display = 'none';

  // Documento: immagini, anteprima, schede Fronte/Retro e selezione file
  state.currentPdf = null;
  state.pageImages = [];
  state.reviewingBatchItem = null;
  state.currentPageIndex = 0;
  state.rotationDegrees = 0;
  state.backRotation = 180;
  docCanvas.getContext('2d').clearRect(0, 0, docCanvas.width, docCanvas.height);
  docCanvas.width = 0;
  docCanvas.height = 0;
  pageTabs.style.display = 'none';
  fileInput.value = '';
  btnAnalyzeAI.disabled = true;
  btnAnalyzeAI.style.animation = 'none';

  showToast('Tutto svuotato: puoi caricare il prossimo rapportino', 'info');
});

// Scrittura dei turni nel file Excel attivo. onConflict 'ask': se in Excel ci sono già dati diversi per quegli
// operai/turni non scrive e restituisce { conflicts }; 'overwrite': scrive comunque.
// Ordine: file chiuso in Excel -> conflitti -> copia di sicurezza con data e ora -> scrittura
async function writeRecordsToExcel(day, records, onConflict = 'ask') {
  const fileName = state.selectedExcelPath;
  try {
    if (!cartellaPronta) throw new Error('cartella dei file Excel non collegata (barra in alto)');
    if (!fileName) throw new Error('nessun file Excel selezionato alla voce MESE');
    if (await CartellaExcel.isOpenInExcel(fileName)) {
      throw new Error(`Il file ${fileName} è aperto in Excel: chiudilo e riprova la scrittura`);
    }

    const bytes = await CartellaExcel.readBytes(fileName);
    if (onConflict !== 'overwrite') {
      const conflicts = await ExcelXlsm.findConflicts(bytes, day, records);
      if (conflicts.length > 0) return { success: false, conflicts, fileName };
    }

    try {
      await CartellaExcel.datedBackup(fileName, bytes);
    } catch (e) {
      throw new Error(`Copia di sicurezza non riuscita (${e.message}): scrittura annullata, Excel non è stato modificato`);
    }

    const result = await ExcelXlsm.writePresenze(bytes, day, records);
    await CartellaExcel.writeBytes(fileName, result.bytes);
    return { success: true, fileName, countUpdated: result.countUpdated };
  } catch (err) {
    // Excel aperto senza file di blocco visibile: il browser non riesce a sostituire il file
    const msg = /NoModificationAllowed|InvalidState|being used|in uso/i.test(`${err.name} ${err.message}`)
      ? `Il file ${fileName} è in uso (forse aperto in Excel): chiudilo e riprova la scrittura`
      : err.message;
    return { success: false, error: msg, fileName };
  }
}

function formatTurnoValues(v) {
  const labels = { or: 'OR', sd: 'SD', on: 'ON', sn: 'SN', v: 'V', rol: 'ROL', malattia: 'MAL' };
  const hours = Object.entries(labels)
    .filter(([k]) => v[k])
    .map(([k, label]) => `${label} ${String(v[k]).replace('.', ',')}`);
  return [v.commessa || '(senza commessa)', ...hours].join(' · ');
}

// Mostra i conflitti e restituisce la scelta: 'overwrite' | 'skip' | 'cancel'
function askConflictChoice(conflicts, day, fileName) {
  const modal = document.getElementById('conflictModal');
  document.getElementById('conflictIntro').textContent =
    `In ${fileName || 'Excel'}, giorno ${day}, questi turni contengono già dati diversi, ` +
    `probabilmente da un altro rapportino dello stesso giorno. Scegli cosa fare:`;
  document.getElementById('conflictList').innerHTML = conflicts.map(c => `
    <div class="conflict-item">
      <div class="conflict-who">${c.workerName} – Turno ${c.turno}</div>
      <div class="conflict-row"><b>Già in Excel:</b><span>${formatTurnoValues(c.existing)}</span></div>
      <div class="conflict-row"><b>Nuovo:</b><span>${formatTurnoValues(c.nuovo)}</span></div>
    </div>`).join('');
  modal.classList.add('active');

  return new Promise(resolve => {
    const buttons = {
      btnConflictOverwrite: 'overwrite',
      btnConflictSkip: 'skip',
      btnConflictCancel: 'cancel',
      btnCloseConflict: 'cancel'
    };
    const handlers = [];
    const done = choice => {
      handlers.forEach(([el, fn]) => el.removeEventListener('click', fn));
      modal.classList.remove('active');
      resolve(choice);
    };
    Object.entries(buttons).forEach(([id, choice]) => {
      const el = document.getElementById(id);
      const fn = () => done(choice);
      el.addEventListener('click', fn);
      handlers.push([el, fn]);
    });
  });
}

// Scrittura diretta su Excel via API locale
btnWriteExcel.addEventListener('click', async () => {
  if (state.tableData.length === 0) {
    showToast('Nessuna presenza da scrivere su Excel', 'error');
    return;
  }

  const day = parseInt(metaDay.value);
  if (!day || day < 1 || day > 31) {
    showToast('Inserisci un giorno valido tra 1 e 31', 'error');
    metaDay.focus();
    return;
  }

  if (state.tableData.some(r => !r.workerIndex)) {
    showToast('C\'è una riga senza operaio: scegli il nome oppure eliminala con la ✕', 'error');
    return;
  }

  btnWriteExcel.disabled = true;
  btnWriteExcel.innerHTML = '💾 Scrittura su Excel in corso...';

  try {
    let records = state.tableData;
    let result = await writeRecordsToExcel(day, records, 'ask');
    let skipped = 0;

    if (!result.success && Array.isArray(result.conflicts)) {
      const choice = await askConflictChoice(result.conflicts, day, result.fileName);
      if (choice === 'cancel') {
        showToast('Scrittura annullata: Excel non è stato modificato', 'info');
        return;
      }
      if (choice === 'skip') {
        const inConflict = new Set(result.conflicts.map(c => `${c.workerIndex}-${c.turno}`));
        records = records.filter(r => !inConflict.has(`${r.workerIndex}-${parseInt(r.turno) === 2 ? 2 : 1}`));
        skipped = state.tableData.length - records.length;
        if (records.length === 0) {
          showToast('Nessun turno scritto: erano tutti già presenti in Excel', 'info');
          return;
        }
      }
      result = await writeRecordsToExcel(day, records, 'overwrite');
    }

    if (result.success) {
      const targetName = result.fileName || 'Excel';
      showToast(`PRESENZE SCRITTE SU ${targetName.toUpperCase()} (GIORNO ${day})!`, 'success');
      detectionBanner.style.display = 'flex';
      bannerTitle.textContent = `Registrazione Completata su ${targetName} (Giorno ${day})`;
      bannerDesc.textContent = `${records.length} turni scritti e verificati` +
        (skipped ? ` (${skipped} lasciati com'erano perché già presenti)` : '') +
        `. Tutte le formule delle colonne L:R sono aggiornate automaticamente!`;
      finishBatchReview(day, records.length);
    } else {
      showToast(`Errore durante la scrittura: ${result.error}`, 'error');
    }
  } catch (err) {
    console.error('Errore chiamata /api/write-excel:', err);
    showToast(`Errore di connessione: ${err.message}`, 'error');
  } finally {
    btnWriteExcel.disabled = false;
    btnWriteExcel.innerHTML = '✅ CONFERMA E SCRIVI SU EXCEL';
  }
});

// Copia JSON per telecomando o backup
btnExportJson.addEventListener('click', () => {
  const payload = {
    day: parseInt(metaDay.value) || 15,
    records: state.tableData
  };
  navigator.clipboard.writeText(JSON.stringify(payload, null, 2));
  showToast('Dati JSON copiati negli appunti!', 'success');
});

// Toast notification helper
function showToast(message, type = 'info') {
  const container = document.getElementById('toastContainer');
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.innerHTML = `
    <span>${type === 'success' ? '✅' : type === 'error' ? '❌' : 'ℹ️'}</span>
    <span>${message}</span>
  `;
  container.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(10px)';
    setTimeout(() => toast.remove(), 300);
  }, 4000);
}

// ============================================================
// MODULO SCANSIONE MULTIPLA RAPPORTINI (BATCH AUTOMATICO)
// ============================================================
const btnOpenBatch = document.getElementById('btnOpenBatch');
const batchModal = document.getElementById('batchModal');
const btnCloseBatch = document.getElementById('btnCloseBatch');
const btnCancelBatch = document.getElementById('btnCancelBatch');
const batchDropzone = document.getElementById('batchDropzone');
const batchFileInput = document.getElementById('batchFileInput');
const batchListContainer = document.getElementById('batchListContainer');
const batchFilesList = document.getElementById('batchFilesList');
const batchCountBadge = document.getElementById('batchCountBadge');
const btnClearBatchList = document.getElementById('btnClearBatchList');
const btnStartBatch = document.getElementById('btnStartBatch');
const batchProgressArea = document.getElementById('batchProgressArea');
const batchProgressText = document.getElementById('batchProgressText');
const batchProgressPercent = document.getElementById('batchProgressPercent');
const batchProgressBar = document.getElementById('batchProgressBar');
const batchAutoSaveExcel = document.getElementById('batchAutoSaveExcel');
const batchTargetFileName = document.getElementById('batchTargetFileName');

let batchFilesQueue = [];
let isBatchRunning = false;

function updateBatchUI() {
  if (batchTargetFileName) {
    const activeName = (state.selectedExcelPath || 'Master_Operai_2027.xlsm').split(/[/\\]/).pop();
    batchTargetFileName.textContent = activeName;
  }
  if (!batchListContainer || !batchFilesList) return;

  if (batchFilesQueue.length === 0) {
    batchListContainer.style.display = 'none';
    if (btnStartBatch) {
      btnStartBatch.disabled = true;
      btnStartBatch.style.animation = 'none';
    }
    return;
  }

  batchListContainer.style.display = 'block';
  if (btnStartBatch) {
    btnStartBatch.disabled = isBatchRunning;
    // File in attesa e scansione ferma: il pulsante pulsa per indicare che il passo successivo è avviarla
    const hasPending = batchFilesQueue.some(item => item.status === 'pending');
    btnStartBatch.style.animation = (!isBatchRunning && hasPending) ? 'pulse 1.5s infinite' : 'none';
  }
  if (batchCountBadge) batchCountBadge.textContent = `${batchFilesQueue.length} file in coda`;

  batchFilesList.innerHTML = '';
  batchFilesQueue.forEach((item, idx) => {
    const li = document.createElement('li');
    li.style.cssText = 'display:flex; justify-content:space-between; align-items:center; padding: 4px 8px; border-radius: 4px; background: rgba(255,255,255,0.03);';
    
    let statusIcon = '⏳ In attesa';
    let statusColor = 'var(--text-muted)';
    if (item.status === 'processing') {
      statusIcon = '🔄 Elaborazione...';
      statusColor = '#60a5fa';
    } else if (item.status === 'done') {
      statusIcon = `✅ Giorno ${item.day || '?'} (${item.recordsCount || 0} turni)`;
      statusColor = '#34d399';
    } else if (item.status === 'review') {
      const others = item.doubts.length > 1 ? ` (+${item.doubts.length - 1})` : '';
      statusIcon = `⚠️ Da controllare: ${item.doubts[0]}${others}`;
      statusColor = 'var(--warning)';
    } else if (item.status === 'error') {
      statusIcon = `❌ ${item.error || 'Errore'}`;
      statusColor = 'var(--danger)';
    }

    li.innerHTML = `
      <span style="font-family: monospace; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 280px;" title="${item.file.name}">
        ${idx + 1}. ${item.file.name}
      </span>
      <span style="display: flex; align-items: center; gap: 8px; min-width: 0;">
        <span style="color: ${statusColor}; font-size: 0.75rem; font-weight: 500; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 300px;"
          title="${item.status === 'review' ? item.doubts.join('\n').replace(/"/g, '&quot;') : ''}">
          ${statusIcon}
        </span>
        ${item.status === 'review' ? `<button class="btn-batch-review" data-idx="${idx}">Controlla</button>` : ''}
      </span>
    `;
    batchFilesList.appendChild(li);
  });

  batchFilesList.querySelectorAll('.btn-batch-review').forEach(btn => {
    btn.addEventListener('click', () => openBatchItemForReview(batchFilesQueue[parseInt(btn.dataset.idx)]));
  });
}

// Apre un rapportino "da controllare" nella schermata principale: documento, dati letti e note.
// La scansione può continuare in sottofondo; dopo "Conferma e scrivi su Excel" il rapportino risulta scritto
function openBatchItemForReview(item) {
  if (!item || item.status !== 'review') return;
  state.pageImages = item.images;
  state.backRotation = 180;
  pageTabs.style.display = item.images.length > 1 ? 'flex' : 'none';
  switchPage(0);
  btnAnalyzeAI.disabled = false;
  btnAnalyzeAI.style.animation = 'none';
  populateExtractedData(item.extracted);
  state.reviewingBatchItem = item;

  detectionBanner.style.display = 'flex';
  bannerTitle.textContent = `Da controllare: ${item.file.name}`;
  bannerDesc.textContent = `Verifica: ${item.doubts.join(' | ')}. Correggi se serve, poi premi "Conferma e scrivi su Excel".`;

  if (batchModal) batchModal.classList.remove('active');
  showToast('Controlla i dati e premi "Conferma e scrivi su Excel"', 'info');
}

// Dopo la scrittura di un rapportino aperto da "Controlla": lo segna come scritto e passa al successivo
function finishBatchReview(day, recordsCount) {
  const item = state.reviewingBatchItem;
  if (!item) return;
  state.reviewingBatchItem = null;
  item.status = 'done';
  item.day = day;
  item.recordsCount = recordsCount;
  item.images = item.extracted = null;
  updateBatchUI();

  const remaining = batchFilesQueue.filter(it => it.status === 'review').length;
  if (remaining > 0 || isBatchRunning) {
    showToast(remaining > 0 ? `Rapportino scritto: ne restano ${remaining} da controllare` : 'Rapportino scritto: la scansione continua', 'success');
    setTimeout(() => { if (batchModal) batchModal.classList.add('active'); updateBatchUI(); }, 1200);
  }
}

if (btnOpenBatch && batchModal) {
  btnOpenBatch.addEventListener('click', () => {
    batchModal.classList.add('active');
    updateBatchUI();
  });
}

function closeBatchModal() {
  if (isBatchRunning) {
    if (!confirm('La scansione batch è ancora in corso. Vuoi davvero chiudere la finestra?')) return;
  } else if (batchFilesQueue.some(it => it.status === 'review')) {
    // Ci sono rapportini da controllare: la coda resta, si ritrovano riaprendo "Scansione Multipla"
    showToast('Rapportini da controllare ancora in coda: li ritrovi in "Scansione Multipla"', 'info');
  } else {
    // Chiusura a scansione ferma: alla riapertura la finestra riparte pulita
    batchFilesQueue = [];
    if (batchProgressArea) batchProgressArea.style.display = 'none';
    if (batchProgressBar) batchProgressBar.style.width = '0%';
    if (batchProgressPercent) batchProgressPercent.textContent = '0%';
    updateBatchUI();
  }
  if (batchModal) batchModal.classList.remove('active');
}

if (btnCloseBatch) btnCloseBatch.addEventListener('click', closeBatchModal);
if (btnCancelBatch) btnCancelBatch.addEventListener('click', closeBatchModal);

if (batchDropzone && batchFileInput) {
  batchDropzone.addEventListener('click', () => batchFileInput.click());
  batchDropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    batchDropzone.classList.add('dragover');
  });
  batchDropzone.addEventListener('dragleave', () => batchDropzone.classList.remove('dragover'));
  batchDropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    batchDropzone.classList.remove('dragover');
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      addFilesToBatch(Array.from(e.dataTransfer.files));
    }
  });
  batchFileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      addFilesToBatch(Array.from(e.target.files));
      e.target.value = '';
    }
  });
}

function addFilesToBatch(files) {
  const valid = files.filter(f => (f.name && f.name.toLowerCase().endsWith('.pdf')) || f.type.startsWith('image/'));
  if (valid.length === 0) {
    showToast('Nessun file PDF o immagine valido selezionato.', 'error');
    return;
  }
  valid.forEach(f => {
    batchFilesQueue.push({ file: f, status: 'pending', day: null, recordsCount: 0, error: null });
  });
  updateBatchUI();
  showToast(`Aggiunti ${valid.length} file alla coda: premi "Avvia Scansione Batch" per iniziare`, 'info');
}

if (btnClearBatchList) {
  btnClearBatchList.addEventListener('click', () => {
    if (isBatchRunning) return;
    batchFilesQueue = [];
    state.reviewingBatchItem = null;
    updateBatchUI();
  });
}

// Convertitore PDF o Immagine a coppie di immagini (Fronte / Retro)
async function extractImagesFromFile(file) {
  const isPdf = (file.name && file.name.toLowerCase().endsWith('.pdf')) || file.type === 'application/pdf';
  if (isPdf) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(arrayBuffer) }).promise;
    const images = [];
    for (let pageNum = 1; pageNum <= Math.min(pdf.numPages, 2); pageNum++) {
      const page = await pdf.getPage(pageNum);
      const unscaledViewport = page.getViewport({ scale: 1.0 });
      const optimalScale = Math.min(1400 / unscaledViewport.width, 1.5);
      const viewport = page.getViewport({ scale: optimalScale });
      const canvas = document.createElement('canvas');
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      const ctx = canvas.getContext('2d');
      await page.render({ canvasContext: ctx, viewport: viewport }).promise;
      images.push(canvas.toDataURL('image/jpeg', 0.80));
    }
    return images;
  } else {
    const rawUrl = await readFileAsDataURL(file);
    const optimized = await resizeImageIfNeeded(rawUrl, 1400);
    return [optimized];
  }
}

// Avvio processo batch sequenziale
if (btnStartBatch) {
  btnStartBatch.addEventListener('click', async () => {
    const apiKey = localStorage.getItem('gemini_api_key');
    if (!apiKey) {
      showToast('Configura prima la tua API Key nelle impostazioni!', 'error');
      return;
    }
    if (batchFilesQueue.length === 0) return;

    isBatchRunning = true;
    btnStartBatch.disabled = true;
    if (btnClearBatchList) btnClearBatchList.disabled = true;
    if (batchProgressArea) batchProgressArea.style.display = 'block';

    let successCount = 0;
    let totalRecordsWritten = 0;
    const autoSave = !batchAutoSaveExcel || batchAutoSaveExcel.checked;

    for (let i = 0; i < batchFilesQueue.length; i++) {
      const item = batchFilesQueue[i];
      if (item.status === 'done' || item.status === 'review') continue; // già scritti o già letti e in attesa di controllo

      item.status = 'processing';
      updateBatchUI();

      const pct = Math.round(((i) / batchFilesQueue.length) * 100);
      if (batchProgressPercent) batchProgressPercent.textContent = `${pct}%`;
      if (batchProgressBar) batchProgressBar.style.width = `${pct}%`;
      if (batchProgressText) batchProgressText.textContent = `[${i + 1}/${batchFilesQueue.length}] Analisi: ${item.file.name}...`;

      try {
        const images = await extractImagesFromFile(item.file);
        if (images.length === 0) throw new Error('Nessuna pagina trovata');

        // Analisi con AI Vision con le stesse regole ferree attuali
        const extracted = await callGeminiVision(apiKey, images, (modelName) => {
          if (batchProgressText) batchProgressText.textContent = `[${i + 1}/${batchFilesQueue.length}] ${modelName}: ${item.file.name}...`;
        });

        // Stesse trasformazioni e protezioni della lettura singola, ma senza toccare lo schermo principale
        const { records } = buildRecordsFromAnalysis(extracted);
        const day = parseInt(extracted.giorno);
        item.day = day || null;
        item.recordsCount = records.length;

        const doubts = Array.isArray(extracted.dubbi) ? [...extracted.dubbi] : [];
        if (!day) doubts.push('Giorno del mese non letto');
        if (records.length === 0) doubts.push('Nessun operaio letto');
        if (!autoSave) doubts.push('Salvataggio automatico disattivato: conferma tu');

        // Solo i rapportini senza dubbi vengono scritti in automatico (e mai sovrascrivendo dati diversi)
        if (doubts.length === 0) {
          const resJson = await writeRecordsToExcel(day, records, 'ask');
          if (!resJson.success && Array.isArray(resJson.conflicts)) {
            const names = [...new Set(resJson.conflicts.map(c => c.workerName))].join(', ');
            doubts.push(`Dati già presenti in Excel per: ${names}`);
          } else if (!resJson.success) {
            throw new Error(resJson.error || 'Errore salvataggio Excel');
          }
        }

        if (doubts.length === 0) {
          item.status = 'done';
          successCount++;
          totalRecordsWritten += records.length;
        } else {
          // Da controllare: si conservano immagini e lettura per aprirlo nella schermata principale
          item.status = 'review';
          item.doubts = doubts;
          item.images = images;
          item.extracted = extracted;
        }
      } catch (err) {
        console.error(`Errore batch file ${item.file.name}:`, err);
        item.status = 'error';
        item.error = err.message || 'Errore';
      }

      updateBatchUI();
      // Piccola pausa di sicurezza tra un foglio e l'altro per proteggere la quota API gratuita
      await sleepMs(1200);
    }

    const toReview = batchFilesQueue.filter(it => it.status === 'review').length;
    const failed = batchFilesQueue.filter(it => it.status === 'error').length;
    if (batchProgressPercent) batchProgressPercent.textContent = `100%`;
    if (batchProgressBar) batchProgressBar.style.width = `100%`;
    if (batchProgressText) batchProgressText.textContent =
      `Scansione completata: ${successCount} scritti in Excel` +
      (toReview ? `, ${toReview} da controllare (premi "Controlla")` : '') +
      (failed ? `, ${failed} con errore` : '') + '.';

    isBatchRunning = false;
    btnStartBatch.disabled = false;
    if (btnClearBatchList) btnClearBatchList.disabled = false;
    updateBatchUI();

    showToast(`Scansione completata: ${successCount} rapportini (${totalRecordsWritten} turni) scritti in Excel` +
      (toReview ? `, ${toReview} da controllare` : ''), toReview ? 'info' : 'success');
  });
}
