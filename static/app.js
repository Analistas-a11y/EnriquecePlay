const state = {
  currentRecord: null,
  selectedFields: [],
  selectedQueryFields: new Set(),
  queue: [],
  notifications: [],
  currentRole: document.body.dataset.userRole || 'user1',
  activeRequestImportId: null,
  currentAction: 'enrich_only',
  historyCurrentPage: 1,
  historyItemsPerPage: 20,
  stagedFiles: [],
  isBatchIndividual: false,
  observacoes: '',
  exportFormat: 'original',
};

let uploadProgressInterval = null;
const expandedCardIds = new Set();

/* ==========================================================================
   SISTEMA DE MODAL DE DOWNLOAD COM RENOMEAÇÃO PERSONALIZADA
   ========================================================================== */

const downloadModalState = {
  downloadUrl: null,
  defaultFilename: '',
  extension: '.xlsx',
  baseName: '',
  onConfirm: null,
};

function splitFilenameAndExt(filename) {
  if (!filename) return { base: 'arquivo', ext: '.xlsx' };
  const lastDot = filename.lastIndexOf('.');
  if (lastDot <= 0) return { base: filename, ext: '' };
  return {
    base: filename.substring(0, lastDot),
    ext: filename.substring(lastDot),
  };
}

function getStandardExportFilename(itemOrOptions = {}, format = '') {
  let mode = 'enrich_only';
  let queueNumber = '';
  let requester = '';
  let timeStr = '';

  if (typeof itemOrOptions === 'object' && itemOrOptions !== null) {
    mode = itemOrOptions.mode || state.currentAction || 'enrich_only';
    queueNumber = itemOrOptions.queue_number || itemOrOptions.queueNumber || state.queueNumber || '';
    requester = itemOrOptions.requester_name || itemOrOptions.requesterName || state.requesterName || (typeof currentUserName !== 'undefined' ? currentUserName : '') || 'usuario';
    const rawTime = itemOrOptions.completed_at || itemOrOptions.request_time || itemOrOptions.time;
    if (rawTime && typeof rawTime === 'string') {
      const match = rawTime.match(/(\d{2}):(\d{2})/);
      if (match) {
        timeStr = `${match[1]}h${match[2]}`;
      }
    }
  }

  // Se queueNumber ainda estiver vazio, verifica o campo no formulário
  if (!queueNumber && typeof document !== 'undefined') {
    const qEl = document.getElementById('queueNumber');
    if (qEl && qEl.value && qEl.value.trim()) {
      queueNumber = qEl.value.trim();
    }
  }

  // Se o solicitante for padrão/vazio, verifica os campos do formulário
  if ((!requester || requester === 'usuario') && typeof document !== 'undefined') {
    const rEl = document.getElementById('requesterName') || document.getElementById('requesterInputName');
    if (rEl && rEl.value && rEl.value.trim()) {
      requester = rEl.value.trim();
    }
  }

  if (!timeStr) {
    const now = new Date();
    const hh = String(now.getHours()).padStart(2, '0');
    const mm = String(now.getMinutes()).padStart(2, '0');
    timeStr = `${hh}h${mm}`;
  }

  let cleanReq = String(requester || 'usuario').trim();
  const matchInner = cleanReq.match(/\(([^)]+)\)/);
  if (matchInner && matchInner[1].trim()) {
    cleanReq = matchInner[1].trim();
  }
  cleanReq = cleanReq.replace(/[\/\\:\*\?"<>\|\(\)]/g, '').replace(/\s+/g, '_').replace(/^_+|_+$/g, '');
  if (!cleanReq) cleanReq = 'usuario';

  let cleanQueue = String(queueNumber || '').replace(/[\/\\:\*\?"<>\|\(\)\s]/g, '').replace(/^_+|_+$/g, '');
  if (!cleanQueue || cleanQueue === 'TEMP' || cleanQueue === 'ENRICH') {
    cleanQueue = '1';
  }

  const suffix = format === 'full' ? '_completa' : (format === 'standard' || format === '3abas' ? '_3abas' : '');

  // 1. Apenas enriquecer: [nome de quem solicitou e o horario]
  if (mode === 'enrich_only' || mode === 'enrich' || queueNumber === 'ENRICH') {
    return `${cleanReq}_${timeStr}${suffix}.xlsx`;
  }

  // 2. Somente subir: [numero da fila e o horario]
  if (mode === 'queue_only' || mode === 'queue' || mode === 'upload_queue_only') {
    return `fila_${cleanQueue}_${timeStr}${suffix}.xlsx`;
  }

  // 3. Enriquecer e subir: [enri_fila(numero da fila)_hr]
  return `enri_fila_${cleanQueue}_${timeStr}${suffix}.xlsx`;
}

function openDownloadRenameModal(options = {}) {
  const modal = document.getElementById('downloadRenameModal');
  const input = document.getElementById('downloadCustomFilenameInput');
  const extBadge = document.getElementById('downloadCustomExtensionBadge');
  const previewName = document.getElementById('downloadModalDefaultName');
  const fileTypeLabel = document.getElementById('downloadModalFileType');

  let defaultName = options.defaultFilename || '';
  if (!defaultName && options.downloadUrl) {
    const isFull = options.downloadUrl.includes('format=full');
    const is3Abas = options.downloadUrl.includes('format=standard') || options.downloadUrl.includes('format=3abas');
    const dlFmt = isFull ? 'full' : (is3Abas ? '3abas' : '');
    if (options.downloadUrl.includes('/api/request/')) {
      const reqMatch = options.downloadUrl.match(/\/api\/request\/([^\/?#]+)/);
      let foundItem = null;
      if (reqMatch) {
        const reqId = reqMatch[1];
        if (state.requests && Array.isArray(state.requests)) {
          foundItem = state.requests.find(r => String(r.id) === String(reqId));
        }
        if (!foundItem && state.queueHistory && Array.isArray(state.queueHistory)) {
          foundItem = state.queueHistory.find(r => String(r.id) === String(reqId));
        }
        if (!foundItem && state.queue && Array.isArray(state.queue)) {
          foundItem = state.queue.find(r => String(r.id) === String(reqId));
        }
      }
      const domQueue = document.getElementById('queueNumber')?.value?.trim();
      const domReq = document.getElementById('requesterName')?.value?.trim() || document.getElementById('requesterInputName')?.value?.trim();

      if (foundItem) {
        defaultName = getStandardExportFilename(foundItem, dlFmt);
      } else {
        defaultName = getStandardExportFilename({
          mode: state.currentAction,
          queue_number: domQueue || state.queueNumber,
          requester_name: domReq || state.requesterName
        }, dlFmt);
      }
    } else if (options.downloadUrl.includes('/api/converter/')) {
      defaultName = 'documento_convertido.pdf';
    } else {
      defaultName = getStandardExportFilename({}, dlFmt);
    }
  }

  if (!modal || !input) {
    if (options.downloadUrl) {
      triggerBrowserDownload(options.downloadUrl, defaultName || 'arquivo.xlsx');
    }
    return;
  }

  const { base, ext } = splitFilenameAndExt(defaultName);
  downloadModalState.downloadUrl = options.downloadUrl;
  downloadModalState.defaultFilename = defaultName;
  downloadModalState.baseName = base;
  downloadModalState.extension = ext || (options.extension || '.xlsx');
  downloadModalState.onConfirm = options.onConfirm || null;

  input.value = base;
  if (extBadge) extBadge.textContent = downloadModalState.extension;
  if (previewName) {
    previewName.textContent = defaultName;
    previewName.title = defaultName;
  }
  if (fileTypeLabel) {
    fileTypeLabel.textContent = options.fileType || `Arquivo ${downloadModalState.extension.replace('.', '').toUpperCase()}`;
  }

  modal.classList.remove('hidden');
  setTimeout(() => {
    input.focus();
    input.select();
  }, 50);
}

function closeDownloadRenameModal() {
  const modal = document.getElementById('downloadRenameModal');
  if (modal) modal.classList.add('hidden');
  downloadModalState.downloadUrl = null;
  downloadModalState.onConfirm = null;
}

async function confirmDownloadCustomName() {
  const input = document.getElementById('downloadCustomFilenameInput');
  if (!input) return;

  let rawName = input.value.trim();
  if (!rawName) {
    showToast('Por favor, informe um nome para o arquivo.', 'error');
    input.focus();
    return;
  }

  let cleanBase = rawName.replace(/[/\\:*?"<>|]/g, '_').trim();
  if (!cleanBase) cleanBase = 'arquivo';

  const ext = downloadModalState.extension;
  if (ext && cleanBase.toLowerCase().endsWith(ext.toLowerCase())) {
    cleanBase = cleanBase.substring(0, cleanBase.length - ext.length);
  }

  const finalFilename = `${cleanBase}${ext}`;
  const url = downloadModalState.downloadUrl;
  const customCallback = downloadModalState.onConfirm;

  closeDownloadRenameModal();

  if (typeof customCallback === 'function') {
    customCallback(finalFilename);
    return;
  }

  if (url) {
    await triggerBrowserDownload(url, finalFilename);
  }
}

async function triggerBrowserDownload(url, filename) {
  const cleanUrl = url.split('?')[0];
  const separator = cleanUrl.includes('?') ? '&' : '?';
  const customParam = `custom_name=${encodeURIComponent(filename || 'arquivo.xlsx')}`;

  // Se a requisição for para a API de download de solicitação:
  if (url.includes('/api/request/')) {
    const reqMatch = url.match(/\/api\/request\/([^\/?#]+)/);
    if (reqMatch) {
      const requestId = reqMatch[1];
      const formatMatch = url.match(/[?&]format=([^&#]+)/);
      const formatParam = formatMatch ? `&format=${formatMatch[1]}` : '';

      try {
        const infoRes = await fetch(`/api/request/${requestId}/download/info?${customParam}${formatParam}`);
        if (infoRes.ok) {
          const info = await infoRes.json();
          const totalChunks = info.total_chunks || 1;
          const mimeType = info.mime_type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
          const finalFilename = filename || info.filename || 'arquivo.xlsx';

          // Arquivos pesados (> 4 MB) -> Download fatiado para contornar limites de túnel
          if (totalChunks > 1) {
            showToast(`Iniciando download de "${finalFilename}" (${(info.filesize / (1024 * 1024)).toFixed(1)} MB)...`, 'info');
            const chunks = [];
            for (let i = 0; i < totalChunks; i++) {
              const pct = Math.round(((i + 1) / totalChunks) * 100);
              showToast(`Baixando "${finalFilename}": ${pct}% (${i + 1}/${totalChunks})...`, 'info');

              const chunkRes = await fetch(`/api/request/${requestId}/download/chunk?chunk_index=${i}&chunk_size=${info.chunk_size}&${customParam}${formatParam}`);
              if (!chunkRes.ok) {
                throw new Error(`Falha no download da fatia ${i + 1} de ${totalChunks}`);
              }
              const buffer = await chunkRes.arrayBuffer();
              chunks.push(buffer);
            }

            const blob = new Blob(chunks, { type: mimeType });
            const blobUrl = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = blobUrl;
            a.download = finalFilename;
            document.body.appendChild(a);
            a.click();
            setTimeout(() => {
              if (document.body.contains(a)) {
                document.body.removeChild(a);
              }
              URL.revokeObjectURL(blobUrl);
            }, 5000);

            showToast(`Download de "${finalFilename}" concluído com sucesso!`, 'success');
            return;
          } else {
            // Arquivo leve (<= 4 MB) -> Download seguro via Blob
            showToast(`Baixando "${finalFilename}"...`, 'info');
            const fileRes = await fetch(`/api/request/${requestId}/download?${customParam}${formatParam}`);
            if (!fileRes.ok) {
              throw new Error('Falha ao obter arquivo para download.');
            }
            const blob = await fileRes.blob();
            const blobUrl = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = blobUrl;
            a.download = finalFilename;
            document.body.appendChild(a);
            a.click();
            setTimeout(() => {
              if (document.body.contains(a)) document.body.removeChild(a);
              URL.revokeObjectURL(blobUrl);
            }, 5000);

            showToast(`Download de "${finalFilename}" concluído com sucesso!`, 'success');
            return;
          }
        }
      } catch (chunkErr) {
        console.warn('Falha no download via API, tentando fallback direto:', chunkErr);
      }
    }
  }

  // Fallback para download direto via Blob
  try {
    showToast(`Baixando "${filename || 'arquivo.xlsx'}"...`, 'info');
    const downloadUrl = `${cleanUrl}?${customParam}`;
    const directRes = await fetch(downloadUrl);
    if (directRes.ok) {
      const blob = await directRes.blob();
      const blobUrl = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = blobUrl;
      a.download = filename || 'arquivo.xlsx';
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        if (document.body.contains(a)) document.body.removeChild(a);
        URL.revokeObjectURL(blobUrl);
      }, 5000);
      showToast(`Download de "${filename || 'arquivo.xlsx'}" concluído com sucesso!`, 'success');
      return;
    }
  } catch (err) {
    console.error('Erro no fallback de download:', err);
  }
}




function downloadFileDirectly(url, defaultFilename = '') {
  let inferredName = defaultFilename;
  const isFull = url.includes('format=full');
  const is3Abas = url.includes('format=standard') || url.includes('format=3abas');
  const dlFmt = isFull ? 'full' : (is3Abas ? '3abas' : '');
  if (!inferredName) {
    if (url.includes('/api/request/')) {
      const reqMatch = url.match(/\/api\/request\/([^\/?#]+)/);
      let foundItem = null;
      if (reqMatch) {
        const reqId = reqMatch[1];
        if (state.requests && Array.isArray(state.requests)) {
          foundItem = state.requests.find(r => String(r.id) === String(reqId));
        }
        if (!foundItem && state.queueHistory && Array.isArray(state.queueHistory)) {
          foundItem = state.queueHistory.find(r => String(r.id) === String(reqId));
        }
      }
      if (foundItem) {
        inferredName = getStandardExportFilename(foundItem, dlFmt);
      } else {
        inferredName = getStandardExportFilename({
          mode: state.currentAction,
          queue_number: state.queueNumber,
          requester_name: state.requesterName
        }, dlFmt);
      }
    } else if (url.includes('/api/converter/')) {
      inferredName = 'documento_convertido.pdf';
    } else {
      inferredName = getStandardExportFilename({}, dlFmt);
    }
  }

  let fileType = 'Planilha Excel';
  if (inferredName.endsWith('.pdf')) fileType = 'Documento PDF';
  else if (inferredName.endsWith('.docx')) fileType = 'Documento Word';
  else if (inferredName.endsWith('.zip')) fileType = 'Arquivo Compactado (ZIP)';
  else if (inferredName.endsWith('.csv')) fileType = 'Planilha CSV';

  openDownloadRenameModal({
    downloadUrl: url,
    defaultFilename: inferredName,
    fileType: fileType,
  });
}

window.getStandardExportFilename = getStandardExportFilename;
window.openDownloadRenameModal = openDownloadRenameModal;
window.closeDownloadRenameModal = closeDownloadRenameModal;
window.confirmDownloadCustomName = confirmDownloadCustomName;
window.downloadFileDirectly = downloadFileDirectly;
window.downloadDevolutiva = downloadDevolutiva;

const tabButtons = document.querySelectorAll('.tab');
const roleSelector = document.getElementById('roleSwitcher');
const panels = {
  'user-panel': document.getElementById('user-panel'),
  'strategy-panel': document.getElementById('strategy-panel'),
  'logs-panel': document.getElementById('logs-panel'),
};

const uploadForm = document.getElementById('uploadForm');
const validationMessage = document.getElementById('validationMessage');
const previewCard = document.getElementById('previewCard');
const previewText = document.getElementById('previewText');
const availableFields = document.getElementById('availableFields');
const selectedFields = document.getElementById('selectedFields');
const finalizeButton = document.getElementById('finalizeButton');

// Step selectors
const step1Container = document.getElementById('step1Container');
const step2Container = document.getElementById('step2Container');
const step3Container = document.getElementById('step3Container');
const enrichButton = document.getElementById('enrichButton');
const uploadLogsBox = document.getElementById('uploadLogsBox');
const requestModal = document.getElementById('requestModal');
const requestForm = document.getElementById('requestForm');
const queueList = document.getElementById('strategyQueue');
const strategyNotifications = document.getElementById('strategyNotifications');
const userNotifications = document.getElementById('userNotifications');
const logsList = document.getElementById('logsQueue');
const logsModal = document.getElementById('logsModal');
const logsModalTitle = document.getElementById('logsModalTitle');
const logsModalContent = document.getElementById('logsModalContent');
const logCounterTotal = document.getElementById('logCounterTotal');
const logCounterSuccess = document.getElementById('logCounterSuccess');
const logCounterError = document.getElementById('logCounterError');

// Stage Progress Selectors (Step 2)
const stageBoxValidate = document.getElementById('stageBoxValidate');
const stageBarValidate = document.getElementById('stageBarValidate');
const stageTooltipValidate = document.getElementById('stageTooltipValidate');
const stageBadgeValidate = document.getElementById('stageBadgeValidate');

const stageBoxEnrich = document.getElementById('stageBoxEnrich');
const stageBarEnrich = document.getElementById('stageBarEnrich');
const stageTooltipEnrich = document.getElementById('stageTooltipEnrich');
const stageBadgeEnrich = document.getElementById('stageBadgeEnrich');

const stageBoxClean = document.getElementById('stageBoxClean');
const stageBarClean = document.getElementById('stageBarClean');
const stageTooltipClean = document.getElementById('stageTooltipClean');
const stageBadgeClean = document.getElementById('stageBadgeClean');

const stageBoxConcat = document.getElementById('stageBoxConcat');
const stageBarConcat = document.getElementById('stageBarConcat');
const stageTooltipConcat = document.getElementById('stageTooltipConcat');
const stageBadgeConcat = document.getElementById('stageBadgeConcat');

const btnToggleTechLogs = document.getElementById('btnToggleTechLogs');
const uploadLogsBoxContainer = document.getElementById('uploadLogsBoxContainer');
const toggleLogsText = document.getElementById('toggleLogsText');

// Stage Progress Selectors (Logs Modal)
const modalStageBoxValidate = document.getElementById('modalStageBoxValidate');
const modalStageBarValidate = document.getElementById('modalStageBarValidate');
const modalStageTooltipValidate = document.getElementById('modalStageTooltipValidate');
const modalStageBadgeValidate = document.getElementById('modalStageBadgeValidate');

const modalStageBoxEnrich = document.getElementById('modalStageBoxEnrich');
const modalStageBarEnrich = document.getElementById('modalStageBarEnrich');
const modalStageTooltipEnrich = document.getElementById('modalStageTooltipEnrich');
const modalStageBadgeEnrich = document.getElementById('modalStageBadgeEnrich');

const modalStageBoxClean = document.getElementById('modalStageBoxClean');
const modalStageBarClean = document.getElementById('modalStageBarClean');
const modalStageTooltipClean = document.getElementById('modalStageTooltipClean');
const modalStageBadgeClean = document.getElementById('modalStageBadgeClean');

const modalStageBoxConcat = document.getElementById('modalStageBoxConcat');
const modalStageBarConcat = document.getElementById('modalStageBarConcat');
const modalStageTooltipConcat = document.getElementById('modalStageTooltipConcat');
const modalStageBadgeConcat = document.getElementById('modalStageBadgeConcat');

const btnToggleModalTechLogs = document.getElementById('btnToggleModalTechLogs');
const modalLogsBoxContainer = document.getElementById('modalLogsBoxContainer');
const toggleModalLogsText = document.getElementById('toggleModalLogsText');

// Queue active/history elements
const strategyQueueActive = document.getElementById('strategyQueueActive');
const strategyQueueHistory = document.getElementById('strategyQueueHistory');
const activeBadge = document.getElementById('activeBadge');
const historyBadge = document.getElementById('historyBadge');

// Usability selectors
const dropZone = document.getElementById('dropZone');
const validationWarningBox = document.getElementById('validationWarningBox');
const validationWarningText = document.getElementById('validationWarningText');
const btnProceedValidation = document.getElementById('btnProceedValidation');
const btnCancelValidation = document.getElementById('btnCancelValidation');
const errorDetailsList = document.getElementById('errorDetailsList');

const statProcessed = document.getElementById('statProcessed');
const statSuccess = document.getElementById('statSuccess');
const statError = document.getElementById('statError');
const statTimeRemaining = document.getElementById('statTimeRemaining');

const REQUIRED_FIELDS = ['nome', 'telefone'];

function repairMojibake(text) {
  if (text === undefined || text === null) return '';
  let str = String(text);
  if (!str) return '';

  // 1. Tratar double UTF-8 / Mojibake comum (ex: Ã§ -> ç, Ã£ -> ã, etc)
  const mojibakeMarkers = ['\u00c3\u00a7', '\u00c3\u00a3', '\u00c3\u00a9', '\u00c3\u00a1', '\u00c3\u00b3', '\u00c3\u00ba', '\u00c3\u00aa', '\u00c3\u00b4', '\u00c3\u00a2', '\u00c3\u00a0', '\u00c3\u0087', '\u00c3\u0083', '\u00c3\u0089', '\u00c3\u0093', '\u00c3\u009a'];
  if (mojibakeMarkers.some(m => str.includes(m))) {
    try {
      str = decodeURIComponent(escape(str));
    } catch (e) {
      const map = {
        '\u00c3\u00a7': 'ç', '\u00c3\u00a3': 'ã', '\u00c3\u00a9': 'é', '\u00c3\u00a1': 'á',
        '\u00c3\u00b3': 'ó', '\u00c3\u00ba': 'ú', '\u00c3\u00aa': 'ê', '\u00c3\u00b4': 'ô',
        '\u00c3\u00a2': 'â', '\u00c3\u00a0': 'à', '\u00c3\u0087': 'Ç', '\u00c3\u0083': 'Ã',
        '\u00c3\u0089': 'É', '\u00c3\u0093': 'Ó', '\u00c3\u009a': 'Ú'
      };
      for (const [k, v] of Object.entries(map)) {
        str = str.split(k).join(v);
      }
    }
  }

  // 2. Casos de palavras terminadas em ção / ções (Vinculação, Situação, etc)
  str = str.replace(/([a-zA-Z\u00c0-\u00ff]+)(?:[\u00ef\u00bf\u00bd\ufffd\xef\xbf\xbd]|ï¿½|\?{1,2})+([oO]\b)/g, '$1çã$2');
  str = str.replace(/([a-zA-Z\u00c0-\u00ff]+)(?:[\u00ef\u00bf\u00bd\ufffd\xef\xbf\xbd]|ï¿½|\?{1,2})+([eE][sS]\b)/g, '$1çõ$2');

  // Casos de ...ss[bad]o -> ...ssão (Profissão, Emissão, etc)
  str = str.replace(/([a-zA-Z\u00c0-\u00ff]+[sS]{2})çã([oO]\b)/g, '$1ã$2');
  str = str.replace(/([a-zA-Z\u00c0-\u00ff]+(?:is|us|rs|ls))çã([oO]\b)/g, '$1ã$2');
  str = str.replace(/([a-zA-Z\u00c0-\u00ff]+[tT])çã([oO]\b)/g, '$1ã$2');
  str = str.replace(/([a-zA-Z\u00c0-\u00ff]+[rR])çã([oO]\b)/g, '$1ã$2');
  str = str.replace(/(\b[nN])çã([oO]\b)/g, '$1ã$2');
  str = str.replace(/(\b[sS])çã([oO]\b)/g, '$1ã$2');

  // 3. Palavras com 'ç' (Endereço, Serviço, Praça, Braço, etc)
  str = str.replace(/endere(?:[\u00ef\u00bf\u00bd\ufffd\xef\xbf\xbd]|ï¿½|\?{1,2})+o\b/gi, 'endereço');
  str = str.replace(/servi(?:[\u00ef\u00bf\u00bd\ufffd\xef\xbf\xbd]|ï¿½|\?{1,2})+o\b/gi, 'serviço');
  str = str.replace(/pra(?:[\u00ef\u00bf\u00bd\ufffd\xef\xbf\xbd]|ï¿½|\?{1,2})+a\b/gi, 'praça');
  str = str.replace(/bra(?:[\u00ef\u00bf\u00bd\ufffd\xef\xbf\xbd]|ï¿½|\?{1,2})+o\b/gi, 'braço');

  // 4. Palavras acentuadas pontuais
  str = str.replace(/\bn(?:[\u00ef\u00bf\u00bd\ufffd\xef\xbf\xbd]|ï¿½|\?{1,2})+mero\b/gi, 'número');
  str = str.replace(/\bm(?:[\u00ef\u00bf\u00bd\ufffd\xef\xbf\xbd]|ï¿½|\?{1,2})+s\b/gi, 'mês');
  str = str.replace(/\bsal(?:[\u00ef\u00bf\u00bd\ufffd\xef\xbf\xbd]|ï¿½|\?{1,2})+rio\b/gi, 'salário');
  str = str.replace(/\bhor(?:[\u00ef\u00bf\u00bd\ufffd\xef\xbf\xbd]|ï¿½|\?{1,2})+rio\b/gi, 'horário');
  str = str.replace(/\brelat(?:[\u00ef\u00bf\u00bd\ufffd\xef\xbf\xbd]|ï¿½|\?{1,2})+rio\b/gi, 'relatório');
  str = str.replace(/\busu(?:[\u00ef\u00bf\u00bd\ufffd\xef\xbf\xbd]|ï¿½|\?{1,2})+rio\b/gi, 'usuário');
  str = str.replace(/\bbenef(?:[\u00ef\u00bf\u00bd\ufffd\xef\xbf\xbd]|ï¿½|\?{1,2})+cio\b/gi, 'benefício');
  str = str.replace(/\bmunic(?:[\u00ef\u00bf\u00bd\ufffd\xef\xbf\xbd]|ï¿½|\?{1,2})+pio\b/gi, 'município');
  str = str.replace(/\bcr(?:[\u00ef\u00bf\u00bd\ufffd\xef\xbf\xbd]|ï¿½|\?{1,2})+dito\b/gi, 'crédito');
  str = str.replace(/\bd(?:[\u00ef\u00bf\u00bd\ufffd\xef\xbf\xbd]|ï¿½|\?{1,2})+bito\b/gi, 'débito');
  str = str.replace(/\bc(?:[\u00ef\u00bf\u00bd\ufffd\xef\xbf\xbd]|ï¿½|\?{1,2})+digo\b/gi, 'código');
  str = str.replace(/\bper(?:[\u00ef\u00bf\u00bd\ufffd\xef\xbf\xbd]|ï¿½|\?{1,2})+odo\b/gi, 'período');
  str = str.replace(/\bm(?:[\u00ef\u00bf\u00bd\ufffd\xef\xbf\xbd]|ï¿½|\?{1,2})+dico\b/gi, 'médico');
  str = str.replace(/\bsa(?:[\u00ef\u00bf\u00bd\ufffd\xef\xbf\xbd]|ï¿½|\?{1,2})+de\b/gi, 'saúde');
  str = str.replace(/\bv(?:[\u00ef\u00bf\u00bd\ufffd\xef\xbf\xbd]|ï¿½|\?{1,2})+nculo\b/gi, 'vínculo');

  str = str.replace(/\b[a-zA-Z\u00c0-\u00ff]+\b/g, (w) => {
    const uppers = (w.match(/[A-Z]/g) || []).length;
    if (uppers > w.length / 2) return w.toUpperCase();
    if (w[0] === w[0].toUpperCase() && w.length > 1) return w[0].toUpperCase() + w.slice(1).toLowerCase();
    return w;
  });

  str = str.replace(/[\u00ef\u00bf\u00bd\ufffd\xef\xbf\xbd]+/g, '');
  return str.trim();
}

function normalizeFieldName(field) {
  return String(repairMojibake(field || ''))
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function normalizeNameValue(value) {
  if (value === undefined || value === null) return '';
  let text = repairMojibake(String(value));
  text = text.replace(/[\r\n\t]+/g, ' ');

  // Corrigir formato invertido com vírgula (ex: "SILVA, CARLOS EDUARDO" -> "CARLOS EDUARDO SILVA")
  if (text.includes(',')) {
    const parts = text.split(',');
    const part0 = parts[0].trim();
    const part1 = parts[1].trim();
    if (part0 && part1 && !/^\d+$/.test(part0) && !/^\d+$/.test(part1)) {
      text = `${part1} ${part0}`;
    }
  }

  // Remover títulos e prefixos honoríficos (Dr., Dra., Prof., Sr., Sra., etc.)
  text = text.replace(/^(?:DR[AªOº]?\b|DR[AªOº]?\.|PROF[AªOº]?\b|PROF[AªOº]?\.|ME[AªOº]?\b|ME[AªOº]?\.|MS\b|MS\.|SR[AªOº]?\b|SR[AªOº]?\.|ENG[AªOº]?\b|ENG[AªOº]?\.|ILM[AªOº]?\b|ILM[AªOº]?\.|EXM[AªOº]?\b|EXM[AªOº]?\.|ADV[AªOº]?\b|ADV[AªOº]?\.)\s+/i, '');

  text = text.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  return text.replace(/\s+/g, ' ').trim();
}

function isRequiredNameField(field) {
  const normalized = normalizeFieldName(field);
  if (!normalized) return false;
  if (['situacao', 'status', 'resultado', 'parecer', 'justificativa', 'motivo', 'vaga', 'cargo'].some(p => normalized.includes(p))) {
    if (!['candidato', 'cliente', 'titular'].some(p => normalized === `nome_${p}` || normalized.startsWith('nome_'))) {
      return false;
    }
  }
  return (
    normalized === 'nome' ||
    normalized.startsWith('nome_') ||
    normalized.includes('nome') ||
    normalized.includes('candidato') ||
    normalized.includes('cliente') ||
    normalized.includes('razao') ||
    normalized.includes('contato') ||
    normalized.includes('titular') ||
    normalized.includes('pagador') ||
    normalized.includes('proponente') ||
    normalized.includes('solicitante') ||
    normalized.includes('paciente') ||
    normalized.includes('aluno') ||
    normalized.includes('associado') ||
    normalized.includes('motorista') ||
    normalized.includes('funcionario') ||
    normalized.includes('colaborador') ||
    normalized.includes('servidor') ||
    normalized.includes('inscrito') ||
    normalized.includes('participante') ||
    normalized.includes('favorecido') ||
    normalized.includes('pessoa')
  );
}

function isRequiredPhoneField(field) {
  const normalized = normalizeFieldName(field);
  if (!normalized) return false;
  if (normalized === 'ddd' || (normalized.startsWith('ddd_') && !normalized.includes('cel') && !normalized.includes('fone') && !normalized.includes('tel'))) {
    return false;
  }
  return (
    normalized === 'telefone' || normalized.startsWith('telefone_') || normalized.includes('telefone') ||
    normalized === 'celular' || normalized.startsWith('celular_') || normalized.includes('celular') ||
    normalized === 'fone' || normalized.startsWith('fone_') || normalized.includes('fone') ||
    normalized === 'tel' || normalized.startsWith('tel_') || normalized.includes('tel') ||
    normalized === 'cel' || normalized.startsWith('cel_') || normalized.includes('cel') ||
    normalized === 'phone' || normalized.startsWith('phone_') || normalized.includes('phone') ||
    normalized === 'whatsapp' || normalized.includes('whatsapp') || normalized.includes('wpp') ||
    (normalized.startsWith('ddd') && (normalized.includes('cel') || normalized.includes('fone') || normalized.includes('tel')))
  );
}

function isRequiredCpfField(field) {
  const normalized = normalizeFieldName(field);
  if (!normalized) return false;
  return (
    normalized === 'cpf' || normalized.startsWith('cpf_') || normalized.includes('cpf') ||
    normalized === 'documento' || normalized.startsWith('documento_') || normalized.includes('documento') ||
    normalized === 'doc' || normalized.startsWith('doc_') || normalized === 'cnpj' || normalized.includes('cnpj')
  );
}

function isRequiredUfField(field) {
  const normalized = normalizeFieldName(field);
  if (!normalized) return false;
  return (
    normalized === 'uf' || normalized === 'estado' || normalized === 'sigla_uf' || normalized === 'uf_estado' ||
    normalized.startsWith('uf_') || normalized.startsWith('estado_')
  );
}

function isOriginalField(field) {
  if (!state.currentRecord || !Array.isArray(state.currentRecord.columns)) return false;
  const candidate = normalizeFieldName(field);
  return state.currentRecord.columns.some((item) => normalizeFieldName(item) === candidate);
}

function isLockedField(field) {
  if (!field) return false;
  // Permitir que o usuário adicione ou remova colunas livremente na interface
  return false;
}

function normalizeSelectedFields(fields = []) {
  return [...new Set((fields || []).map((field) => String(field).trim()).filter(Boolean))];
}

function renderDataPreview(payload) {
  const container = document.getElementById('dataPreviewContainer');
  const metaEl = document.getElementById('previewFileMeta');
  const badgesEl = document.getElementById('detectedColumnsBadges');
  const headEl = document.getElementById('dataPreviewTableHead');
  const bodyEl = document.getElementById('dataPreviewTableBody');

  if (!container || !payload) return;

  const total = payload.total_in || 0;
  if (metaEl) metaEl.textContent = `(${total.toLocaleString('pt-BR')} linhas carregadas)`;

  const columns = payload.columns || [];
  const sampleRows = payload.sample_rows || [];

  // Identify detected columns
  let detectedName = columns.find(col => isRequiredNameField(col));
  let detectedPhone = columns.find(col => isRequiredPhoneField(col));
  let detectedCpf = columns.find(col => {
    const n = normalizeFieldName(col);
    return n.includes('cpf') || n.includes('documento') || n.includes('doc');
  });

  if (badgesEl) {
    badgesEl.innerHTML = `
      ${detectedName ? `<span class="detected-badge name">Nome: <strong>${detectedName}</strong></span>` : ''}
      ${detectedCpf ? `<span class="detected-badge cpf">CPF: <strong>${detectedCpf}</strong></span>` : ''}
      ${detectedPhone ? `<span class="detected-badge phone">Telefone: <strong>${detectedPhone}</strong></span>` : ''}
    `;
  }

  // Render Table Header (up to first 8 columns)
  const displayCols = columns.slice(0, 8);
  if (headEl) {
    headEl.innerHTML = `
      <tr>
        <th style="width: 40px; text-align: center;">#</th>
        ${displayCols.map(c => `<th>${repairMojibake(c)}</th>`).join('')}
      </tr>
    `;
  }

  // Render Table Body (up to 5 sample rows)
  if (bodyEl) {
    if (sampleRows.length > 0) {
      bodyEl.innerHTML = sampleRows.map((row, idx) => `
        <tr>
          <td style="color: var(--text-muted); font-weight: 700; text-align: center;">${idx + 1}</td>
          ${displayCols.map(c => {
        const rawVal = row[c] !== undefined && row[c] !== null ? row[c] : (row[normalizeFieldName(c)] !== undefined ? row[normalizeFieldName(c)] : '');
        const cleanVal = rawVal !== '' ? repairMojibake(String(rawVal).trim()) : '';
        return `<td>${cleanVal !== '' ? cleanVal : '<span style="color: var(--text-muted);">-</span>'}</td>`;
      }).join('')}
        </tr>
      `).join('');
    } else {
      bodyEl.innerHTML = `<tr><td colspan="${displayCols.length + 1}" style="text-align: center; color: var(--text-muted); padding: 1rem;">Nenhuma linha de amostra disponível.</td></tr>`;
    }
  }

  container.classList.remove('hidden');
}

function labelize(value, stripPrefix = false) {
  if (!value) return 'Campo';
  let safe = repairMojibake(String(value).trim());
  if (safe.startsWith('enriquecido_')) {
    const suffix = safe.replace('enriquecido_', '').trim();
    const normSuffix = normalizeFieldName(suffix);
    let baseLabel = '';
    if (normSuffix === 'cpf') baseLabel = 'CPF';
    else if (normSuffix === 'cbo') baseLabel = 'CBO';
    else if (normSuffix === 'uf' || normSuffix === 'estado') baseLabel = 'Estado';
    else if (normSuffix === 'profissao') baseLabel = 'Profissão';
    else if (normSuffix === 'obito') baseLabel = 'Óbito';
    else if (normSuffix === 'situacao' || normSuffix === 'status') baseLabel = 'Situação';
    else if (normSuffix === 'endereco' || normSuffix === 'logradouro') baseLabel = 'Endereço';
    else if (normSuffix === 'municipio' || normSuffix === 'cidade') baseLabel = 'Cidade';
    else if (normSuffix === 'vinculacao') baseLabel = 'Vinculação';
    else if (normSuffix.startsWith('telefone_')) {
      const num = normSuffix.replace('telefone_', '');
      baseLabel = `Telefone ${num}`;
    } else {
      const cleanSuffix = suffix.replace(/_/g, ' ');
      baseLabel = cleanSuffix.charAt(0).toUpperCase() + cleanSuffix.slice(1);
    }

    if (stripPrefix) {
      return baseLabel;
    }
    return `Enriquecido - ${baseLabel}`;
  }
  const norm = normalizeFieldName(safe);
  if (norm === 'cpf') return 'CPF';
  if (norm === 'cbo') return 'CBO';
  if (norm === 'uf') return 'UF';
  if (norm === 'profissao' && (safe.toLowerCase() === 'profissao' || safe.toLowerCase() === 'profissão')) return 'Profissão';
  if (norm === 'obito' && (safe.toLowerCase() === 'obito' || safe.toLowerCase() === 'óbito')) return 'Óbito';
  if (norm === 'situacao') return 'Situação';
  if (norm === 'endereco') return 'Endereço';
  if (norm === 'vinculacao') return 'Vinculação';
  if (norm.startsWith('telefone_')) {
    const num = norm.replace('telefone_', '');
    return `Telefone ${num}`;
  }
  const clean = safe.replace(/_/g, ' ');
  return clean.charAt(0).toUpperCase() + clean.slice(1);
}

function showToast(message, type = 'success') {
  const container = document.getElementById('toastContainer');
  if (!container) return;

  const toast = document.createElement('div');
  toast.className = `toast ${type}`;

  const content = document.createElement('span');
  content.className = 'toast-content';
  content.textContent = message;

  const closeBtn = document.createElement('button');
  closeBtn.type = 'button';
  closeBtn.className = 'toast-close';
  closeBtn.innerHTML = '&times;';
  closeBtn.onclick = () => {
    toast.classList.add('toast-fade-out');
    toast.addEventListener('animationend', () => toast.remove());
  };

  toast.appendChild(content);
  toast.appendChild(closeBtn);
  container.appendChild(toast);

  // Web Notification API integration (popup notifications when out of tab)
  if ('Notification' in window) {
    if (Notification.permission === 'granted') {
      new Notification('Nova Vida Enriquecimento', {
        body: message,
        icon: '/static/Bookplay.png.webp'
      });
    } else if (Notification.permission !== 'denied') {
      Notification.requestPermission().then(permission => {
        if (permission === 'granted') {
          new Notification('Nova Vida Enriquecimento', {
            body: message,
            icon: '/static/Bookplay.png.webp'
          });
        }
      });
    }
  }

  setTimeout(() => {
    if (toast.parentNode) {
      toast.classList.add('toast-fade-out');
      toast.addEventListener('animationend', () => toast.remove());
    }
  }, 4000);
}

function showMessage(text, type = 'success') {
  validationMessage.textContent = text;
  validationMessage.className = `message ${type}`;
  validationMessage.classList.remove('hidden');
  showToast(text, type);
}

function hideMessage() {
  validationMessage.classList.add('hidden');
}

function formatSelectedFields(fields) {
  const normalizedFields = normalizeSelectedFields(fields);
  const isQueueMode = state.currentAction !== 'enrich_only';

  if (!isQueueMode) {
    return normalizedFields.map((f) => labelize(f)).join(' | ');
  }

  const hasEnrichedName = normalizedFields.some(
    (f) => normalizeFieldName(f) === 'enriquecido_nome' || normalizeFieldName(f) === 'enriquecido_nome_completo'
  );
  const nomeLabel = hasEnrichedName ? labelize('enriquecido_nome').toUpperCase() : labelize('nome').toUpperCase();

  const phoneFields = normalizedFields.filter(
    (f) => isRequiredPhoneField(f) && !normalizeFieldName(f).startsWith('ddd')
  );

  let result = '';
  if (phoneFields.length <= 1) {
    const telLabel = phoneFields.length === 1 ? labelize(phoneFields[0]).toUpperCase() : labelize('telefone').toUpperCase();
    result = `${nomeLabel}|${telLabel}|`;
  } else {
    const phoneSlots = ['', '', '', '', ''];
    for (let i = 0; i < Math.min(5, phoneFields.length); i++) {
      phoneSlots[i] = labelize(phoneFields[i]).toUpperCase();
    }
    result = `${nomeLabel}|${phoneSlots.join('|')}|`;
  }

  const nameFieldsSet = new Set(normalizedFields.filter(isRequiredNameField));
  const phoneFieldsSet = new Set(phoneFields);
  const remaining = normalizedFields.filter(
    (f) => !nameFieldsSet.has(f) && !phoneFieldsSet.has(f) && !normalizeFieldName(f).startsWith('ddd')
  );

  if (remaining.length > 0) {
    result += ' ' + remaining.map((f) => labelize(f, true)).join(' ');
  }
  return result;
}

function formatPreviewFields(fields, previewData) {
  if (!previewData) return '';
  const normalizedFields = normalizeSelectedFields(fields);
  const isQueueMode = state.currentAction !== 'enrich_only';

  if (!isQueueMode) {
    const parts = [];
    for (const field of normalizedFields) {
      const val = previewData[field] !== undefined ? previewData[field] : previewData[normalizeFieldName(field)];
      const valStr = val !== undefined && val !== null && String(val).trim() !== '' ? String(val).trim() : '—';
      parts.push(`${labelize(field)}: ${valStr}`);
    }
    return parts.join(' | ');
  }

  let nomeVal = '';
  const selectedNameFields = normalizedFields.filter(isRequiredNameField);
  const nameFieldsSet = new Set(selectedNameFields);

  // Se o usuário selecionou 'enriquecido_nome', ele tem prioridade e substitui o 'nome'
  const enrichedNameField = selectedNameFields.find(
    (f) => normalizeFieldName(f) === 'enriquecido_nome' || normalizeFieldName(f) === 'enriquecido_nome_completo'
  );

  if (enrichedNameField) {
    const enrVal = previewData[enrichedNameField] !== undefined ? previewData[enrichedNameField] : previewData['enriquecido_nome'];
    if (enrVal !== undefined && enrVal !== null && String(enrVal).trim() !== '') {
      nomeVal = String(enrVal).trim();
    }
  }

  if (!nomeVal) {
    for (const field of selectedNameFields) {
      const val = previewData[field] !== undefined ? previewData[field] : previewData[normalizeFieldName(field)];
      if (val !== undefined && val !== null && String(val).trim() !== '') {
        nomeVal = String(val).trim();
        break;
      }
    }
  }

  if (!nomeVal) {
    nomeVal = previewData['enriquecido_nome'] || previewData['nome'] || previewData['nome_completo'] || previewData['candidato'] || previewData['nome_do_candidato'] || previewData['nome_candidato'] || previewData['nome_cliente'] || previewData['cliente'] || previewData['razao_social'] || previewData['titular'] || previewData['pagador'] || previewData['solicitante'] || previewData['proponente'] || previewData['aluno'] || previewData['paciente'] || previewData['associado'] || previewData['funcionario'] || '—';
  }

  if (nomeVal && nomeVal !== '—') {
    nomeVal = normalizeNameValue(nomeVal);
  }

  const phoneFields = normalizedFields.filter(
    (f) => isRequiredPhoneField(f) && !normalizeFieldName(f).startsWith('ddd')
  );

  let result = '';
  if (phoneFields.length <= 1) {
    let telVal = '';
    if (phoneFields.length === 1) {
      const f = phoneFields[0];
      const val = previewData[f] !== undefined ? previewData[f] : previewData[normalizeFieldName(f)];
      if (val !== undefined && val !== null && String(val).trim() !== '') {
        telVal = String(val).trim().replace(/\D/g, '');
      }
    }
    if (!telVal) {
      telVal = previewData['enriquecido_telefone'] || previewData['celular'] || previewData['cel1'] || previewData['telefone'] || previewData['fone'] || previewData['tel'] || '—';
      telVal = String(telVal).replace(/\D/g, '') || telVal;
    }
    result = `${nomeVal}|${telVal}|`;
  } else {
    const phoneSlots = ['', '', '', '', ''];
    for (let i = 0; i < Math.min(5, phoneFields.length); i++) {
      const f = phoneFields[i];
      const val = previewData[f] !== undefined ? previewData[f] : previewData[normalizeFieldName(f)];
      if (val !== undefined && val !== null && String(val).trim() !== '') {
        const clean = String(val).trim().replace(/\D/g, '');
        phoneSlots[i] = clean || String(val).trim();
      }
    }
    result = `${nomeVal}|${phoneSlots.join('|')}|`;
  }

  const phoneFieldsSet = new Set(phoneFields);
  const remaining = normalizedFields.filter(
    (f) => !nameFieldsSet.has(f) && !phoneFieldsSet.has(f) && !normalizeFieldName(f).startsWith('ddd')
  );

  if (remaining.length > 0) {
    const parts = [];
    for (const field of remaining) {
      const val = previewData[field] !== undefined ? previewData[field] : previewData[normalizeFieldName(field)];
      if (val !== undefined && val !== null && String(val).trim() !== '') {
        parts.push(`${labelize(field, true)}: ${String(val).trim()}`);
      }
    }
    if (parts.length > 0) {
      result += ' ' + parts.join(' ');
    }
  }
  return result;
}

function applyRoleAccess() {
  const roleSelector = document.getElementById('roleSwitcher');
  const userRole = document.body.dataset.userRole;

  if (roleSelector) {
    roleSelector.value = state.currentRole;
    roleSelector.disabled = userRole !== 'admin';
  }

  if (userRole !== 'admin') {
    state.currentRole = 'user1';
  }

  // Switch active panel based on the selected role/tab value
  if (state.currentRole === 'admin') {
    setActiveTab('strategy-panel');
  } else if (state.currentRole === 'admin-user' || state.currentRole === 'user1') {
    setActiveTab('user-panel');
  } else if (state.currentRole === 'admin-logs') {
    setActiveTab('logs-panel');
  }
}

function setActiveTab(targetId) {
  tabButtons.forEach((button) => {
    const shouldBeActive = button.dataset.target === targetId;
    button.classList.toggle('is-active', shouldBeActive);
  });

  Object.entries(panels).forEach(([key, panel]) => {
    if (panel) {
      panel.classList.toggle('hidden', key !== targetId);
      panel.classList.toggle('is-visible', key === targetId);
    }
  });

  if (targetId === 'logs-panel') {
    loadLogsQueue();
  }
}

function renderAvailableFields() {
  if (!state.currentRecord) {
    availableFields.innerHTML = '';
    return;
  }

  const used = new Set(state.selectedFields);
  const fields = state.currentAction === 'queue_only'
    ? (state.currentRecord.columns || [])
    : (state.currentRecord.available_fields || []);
  const items = fields.filter((field) => !used.has(field));

  availableFields.innerHTML = '';
  items.forEach((field) => {
    const tag = document.createElement('button');
    tag.type = 'button';
    tag.className = 'field-tag';
    tag.draggable = true;
    tag.dataset.field = field;
    tag.textContent = labelize(field);
    tag.addEventListener('dragstart', (event) => {
      event.dataTransfer.setData('text/plain', field);
      event.dataTransfer.effectAllowed = 'copy';
    });
    tag.addEventListener('click', () => addField(field));
    availableFields.appendChild(tag);
  });

  availableFields.addEventListener('dragover', (event) => event.preventDefault());
  availableFields.addEventListener('drop', (event) => {
    event.preventDefault();
    const field = event.dataTransfer.getData('text/plain');
    if (field) {
      removeField(field);
    }
  });
}

function renderSelectedFields() {
  if (!state.currentRecord) {
    selectedFields.innerHTML = '';
    return;
  }

  selectedFields.innerHTML = '';
  state.selectedFields.forEach((field) => {
    const tag = document.createElement('button');
    tag.type = 'button';
    tag.className = `field-tag ${isLockedField(field) ? 'required' : ''}`;
    tag.draggable = true;
    tag.dataset.field = field;
    tag.textContent = labelize(field);
    tag.addEventListener('dragstart', (event) => {
      event.dataTransfer.setData('text/plain', field);
      event.dataTransfer.effectAllowed = 'move';
    });
    tag.addEventListener('click', () => removeField(field));
    selectedFields.appendChild(tag);
  });

  selectedFields.addEventListener('dragover', (event) => event.preventDefault());
  selectedFields.addEventListener('drop', (event) => {
    event.preventDefault();
    const field = event.dataTransfer.getData('text/plain');
    if (field) {
      addField(field);
    }
  });
}

function addField(field) {
  if (!state.currentRecord) return;
  const isAlreadySelected = state.selectedFields.includes(field);
  if (isAlreadySelected) return;
  if (field === 'tipo_enriquecimento') return;

  const orderedFields = [...state.selectedFields, field];
  const uniqueFields = [...new Set(orderedFields)];
  state.selectedFields = normalizeSelectedFields(uniqueFields);
  renderAvailableFields();
  renderSelectedFields();
  renderPreview();
}

function removeField(field) {
  if (!state.currentRecord) return;
  if (isLockedField(field)) return;
  state.selectedFields = normalizeSelectedFields(state.selectedFields.filter((item) => item !== field));
  renderAvailableFields();
  renderSelectedFields();
  renderPreview();
}

function renderPreview() {
  if (!state.currentRecord) {
    previewCard.classList.add('hidden');
    return;
  }

  const previewData = state.currentRecord.preview || {};
  const selected = state.selectedFields;

  if (!selected.length) {
    previewText.textContent = 'Selecione os campos para visualizar a prévia.';
    previewCard.classList.remove('hidden');
    return;
  }

  previewText.textContent = formatPreviewFields(selected, previewData);
  previewCard.classList.remove('hidden');
}

function updateRequestButtonState() {
  // Finalize button is always visible in Step 3
}

function updateFinalizeButtonLabel() {
  if (!finalizeButton) return;
  const mode = state.currentAction || 'enrich_only';
  if (state.isBatchIndividual) {
    finalizeButton.textContent = mode === 'enrich_only' ? 'Configurar e Enriquecer Lotes' : 'Configurar e Subir Lotes';
  } else if (mode === 'enrich_only') {
    finalizeButton.textContent = 'Iniciar Enriquecimento e Baixar';
  } else if (mode === 'enrich_and_queue') {
    finalizeButton.textContent = 'Enriquecer e Subir na Fila';
  } else {
    finalizeButton.textContent = 'Importar na fila';
  }
}

function setExportFormat(format) {
  state.exportFormat = (format === 'standard' || format === '3abas') ? 'standard' : 'original';

  // Step 3 / Passo 2
  const rOrig = document.getElementById('exportFormatOriginal');
  const rTabs = document.getElementById('exportFormatTabs');
  const lblOrig = document.getElementById('exportOptionOriginalLabel');
  const lblTabs = document.getElementById('exportOptionTabsLabel');
  if (rOrig) rOrig.checked = (state.exportFormat === 'original');
  if (rTabs) rTabs.checked = (state.exportFormat === 'standard');
  if (lblOrig) lblOrig.classList.toggle('is-active', state.exportFormat === 'original');
  if (lblTabs) lblTabs.classList.toggle('is-active', state.exportFormat === 'standard');

  // Modal Solicitante
  const mOrig = document.getElementById('modalFormatOriginal');
  const mTabs = document.getElementById('modalFormatTabs');
  const mLblOrig = document.getElementById('modalRadioOrigLabel');
  const mLblTabs = document.getElementById('modalRadioTabsLabel');
  if (mOrig) mOrig.checked = (state.exportFormat === 'original');
  if (mTabs) mTabs.checked = (state.exportFormat === 'standard');
  if (mLblOrig) mLblOrig.classList.toggle('is-active', state.exportFormat === 'original');
  if (mLblTabs) mLblTabs.classList.toggle('is-active', state.exportFormat === 'standard');

  // Modal Lotes
  const bOrig = document.getElementById('batchFormatOriginal');
  const bTabs = document.getElementById('batchFormatTabs');
  const bLblOrig = document.getElementById('batchRadioOrigLabel');
  const bLblTabs = document.getElementById('batchRadioTabsLabel');
  if (bOrig) bOrig.checked = (state.exportFormat === 'original');
  if (bTabs) bTabs.checked = (state.exportFormat === 'standard');
  if (bLblOrig) bLblOrig.classList.toggle('is-active', state.exportFormat === 'original');
  if (bLblTabs) bLblTabs.classList.toggle('is-active', state.exportFormat === 'standard');
}

function initExportFormatControls() {
  const radioPairs = [
    ['exportFormatOriginal', 'original'],
    ['exportFormatTabs', 'standard'],
    ['modalFormatOriginal', 'original'],
    ['modalFormatTabs', 'standard'],
    ['batchFormatOriginal', 'original'],
    ['batchFormatTabs', 'standard']
  ];

  radioPairs.forEach(([id, val]) => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener('change', () => {
        if (el.checked) setExportFormat(val);
      });
    }
  });

  const labelPairs = [
    ['exportOptionOriginalLabel', 'original'],
    ['exportOptionTabsLabel', 'standard'],
    ['modalRadioOrigLabel', 'original'],
    ['modalRadioTabsLabel', 'standard'],
    ['batchRadioOrigLabel', 'original'],
    ['batchRadioTabsLabel', 'standard']
  ];

  labelPairs.forEach(([id, val]) => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener('click', (e) => {
        if (e.target && e.target.tagName === 'INPUT') return;
        setExportFormat(val);
      });
    }
  });

  setExportFormat(state.exportFormat || 'original');
}

function clearImportedFileState() {
  // 1. Limpa dados de estado em memória
  state.currentRecord = null;
  state.selectedFields = [];
  state.isBatchIndividual = false;
  state.observacoes = '';
  if (state.selectedQueryFields && typeof state.selectedQueryFields.clear === 'function') {
    state.selectedQueryFields.clear();
  }
  state.activeRequestImportId = null;
  clearStagedFiles();

  // 2. Interrompe qualquer polling ativo de upload/enriquecimento
  if (typeof uploadProgressInterval !== 'undefined' && uploadProgressInterval) {
    clearInterval(uploadProgressInterval);
    uploadProgressInterval = null;
  }

  // 3. Reseta formulário e input de arquivo
  if (uploadForm) {
    uploadForm.reset();
    uploadForm.classList.remove('hidden');
    const submitBtn = uploadForm.querySelector("button[type='submit']");
    if (submitBtn) {
      submitBtn.disabled = false;
      submitBtn.textContent = 'Validar arquivo';
    }
  }
  const excelFileInput = document.getElementById('excelFile');
  if (excelFileInput) {
    excelFileInput.value = '';
  }

  // 4. Reseta a Drop Zone (estilos e texto padrão)
  if (dropZone) {
    dropZone.classList.remove('drop-zone--attached', 'drop-zone--loading', 'drop-zone--over');
    const textEl = dropZone.querySelector('.drop-zone-text');
    if (textEl) {
      const defaultText = 'Arrastar e soltar arquivo aqui (.xlsx, .csv) ou clique para selecionar';
      textEl.textContent = defaultText;
      textEl.dataset.originalText = defaultText;
    }
    const spinner = dropZone.querySelector('.drop-zone-spinner');
    if (spinner) spinner.remove();
  }

  // 5. Oculta e limpa a Prévia da Planilha (Data Preview)
  const dataPreviewContainer = document.getElementById('dataPreviewContainer');
  if (dataPreviewContainer) {
    dataPreviewContainer.classList.add('hidden');
  }
  const previewMeta = document.getElementById('previewFileMeta');
  if (previewMeta) previewMeta.textContent = '(0 linhas identificadas)';
  const detectedBadges = document.getElementById('detectedColumnsBadges');
  if (detectedBadges) detectedBadges.innerHTML = '';
  const previewHead = document.getElementById('dataPreviewTableHead');
  if (previewHead) previewHead.innerHTML = '';
  const previewBody = document.getElementById('dataPreviewTableBody');
  if (previewBody) previewBody.innerHTML = '';

  // 6. Oculta e limpa alertas de validação e mensagens de status
  if (validationWarningBox) {
    validationWarningBox.classList.add('hidden');
  }
  if (errorDetailsList) {
    errorDetailsList.innerHTML = '';
  }
  hideMessage();

  // 7. Oculta e reseta os passos 1, 2 e 3
  const preflightCard = document.getElementById('preflightDiagnosticCard');
  if (preflightCard) {
    preflightCard.classList.add('hidden');
  }

  if (step1Container) {
    step1Container.classList.add('hidden');
    const queryFieldsSel = document.getElementById('queryFieldsSelection');
    if (queryFieldsSel) queryFieldsSel.innerHTML = '';
  }

  if (step2Container) {
    step2Container.classList.add('hidden');
    if (statProcessed) statProcessed.textContent = '0 / 0';
    if (statSuccess) statSuccess.textContent = '0 (0%)';
    if (statError) statError.textContent = '0 (0%)';
    if (statTimeRemaining) statTimeRemaining.textContent = 'Calculando...';

    updateSkillStage(stageBoxValidate, stageBarValidate, stageTooltipValidate, stageBadgeValidate, '', 0, 'Aguardando');
    updateSkillStage(stageBoxEnrich, stageBarEnrich, stageTooltipEnrich, stageBadgeEnrich, '', 0, 'Aguardando');
    updateSkillStage(stageBoxClean, stageBarClean, stageTooltipClean, stageBadgeClean, '', 0, 'Aguardando');
    updateSkillStage(stageBoxConcat, stageBarConcat, stageTooltipConcat, stageBadgeConcat, '', 0, 'Aguardando');

    if (uploadLogsBox) uploadLogsBox.textContent = '';
  }

  if (step3Container) {
    step3Container.classList.add('hidden');
    if (availableFields) availableFields.innerHTML = '';
    if (selectedFields) selectedFields.innerHTML = '';
    if (previewCard) previewCard.classList.add('hidden');
    if (previewText) previewText.textContent = 'Selecione os campos acima para gerar a prévia.';
  }

  // 8. Fecha modais de solicitação se abertos
  if (typeof closeRequestModal === 'function') closeRequestModal();
  if (typeof closeRequesterModal === 'function') closeRequesterModal();
}

function setActionMode(mode) {
  state.currentAction = mode;
  updateFinalizeButtonLabel();
  const efCard = document.getElementById('exportFormatCard');
  if (efCard) {
    efCard.style.display = (mode === 'queue_only') ? 'none' : '';
  }
  document.querySelectorAll('.mode-btn').forEach((button) => {
    const isSelected = button.dataset.actionType === mode;
    button.classList.toggle('is-active', isSelected);
    button.setAttribute('aria-pressed', String(isSelected));
  });

  const stage2TitleEl = stageBoxEnrich ? stageBoxEnrich.querySelector('.title') : null;
  if (stage2TitleEl) {
    stage2TitleEl.textContent = (mode === 'queue_only') ? 'Processando contatos da fila' : 'Enriquecendo o arquivo';
  }

  // Limpa o arquivo e o estado anterior para que o novo modo inicie zerado
  clearImportedFileState();
}

function formatFileSize(bytes) {
  if (!bytes || bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}

function addStagedFiles(newFiles) {
  if (!newFiles || !newFiles.length) return;
  let addedCount = 0;
  let duplicateCount = 0;

  Array.from(newFiles).forEach(file => {
    const lower = file.name.toLowerCase();
    if (!lower.endsWith('.xlsx') && !lower.endsWith('.csv')) {
      showToast(`Formato não suportado: ${file.name}. Envie .xlsx ou .csv`, 'warning');
      return;
    }

    const isDuplicate = state.stagedFiles.some(f => f.name === file.name && f.size === file.size);
    if (isDuplicate) {
      duplicateCount++;
    } else {
      state.stagedFiles.push(file);
      addedCount++;
    }
  });

  if (duplicateCount > 0 && addedCount === 0) {
    showToast('Planilha(s) já adicionada(s) à lista.', 'info');
  } else if (addedCount > 0) {
    showToast(`${addedCount} planilha(s) adicionada(s) à lista.`, 'info');
  }

  renderStagedFiles();
}

function removeStagedFile(index) {
  if (index >= 0 && index < state.stagedFiles.length) {
    const removed = state.stagedFiles.splice(index, 1)[0];
    showToast(`"${removed.name}" removida.`, 'info');
    renderStagedFiles();
  }
}

function clearStagedFiles() {
  state.stagedFiles = [];
  const excelFileInput = document.getElementById('excelFile');
  if (excelFileInput) excelFileInput.value = '';
  renderStagedFiles();
}

function renderStagedFiles() {
  const container = document.getElementById('stagedFilesContainer');
  const list = document.getElementById('stagedFilesList');
  const badge = document.getElementById('stagedFilesCountBadge');
  const totalSizeEl = document.getElementById('stagedTotalSize');
  const singleBtn = document.getElementById('btnSubmitSingleUpload');
  const multiActions = document.getElementById('multiFileActionButtons');
  const textEl = dropZone ? dropZone.querySelector('.drop-zone-text') : null;

  if (!container || !list) return;

  const count = state.stagedFiles.length;

  if (count === 0) {
    container.classList.add('hidden');
    list.innerHTML = '';
    if (dropZone) dropZone.classList.remove('drop-zone--attached');
    if (textEl) {
      const defaultTxt = 'Arrastar e soltar arquivo aqui (.xlsx, .csv) ou clique para selecionar';
      textEl.textContent = defaultTxt;
      textEl.dataset.originalText = defaultTxt;
    }
    if (singleBtn) {
      singleBtn.classList.remove('hidden');
      singleBtn.textContent = 'Validar arquivo';
    }
    if (multiActions) multiActions.classList.add('hidden');
    return;
  }

  container.classList.remove('hidden');
  if (badge) badge.textContent = `${count} ${count === 1 ? 'arquivo' : 'arquivos'}`;

  let totalBytes = 0;
  list.innerHTML = '';

  state.stagedFiles.forEach((file, idx) => {
    totalBytes += file.size;
    const card = document.createElement('div');
    card.className = 'staged-file-card';
    card.innerHTML = `
      <div class="staged-file-info">
        <div class="staged-file-icon">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>
        </div>
        <div class="staged-file-details">
          <span class="staged-file-name" title="${file.name}">${file.name}</span>
          <span class="staged-file-meta">${formatFileSize(file.size)}</span>
        </div>
      </div>
      <button type="button" class="staged-file-remove-btn" onclick="removeStagedFile(${idx})" title="Remover este arquivo">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
        Remover
      </button>
    `;
    list.appendChild(card);
  });

  if (totalSizeEl) {
    totalSizeEl.textContent = `Tamanho total: ${formatFileSize(totalBytes)}`;
  }

  if (dropZone) {
    dropZone.classList.add('drop-zone--attached');
    if (textEl) {
      const attachedText = `+ Arraste mais planilhas aqui ou clique para adicionar`;
      textEl.textContent = attachedText;
      textEl.dataset.originalText = attachedText;
    }
  }

  if (count === 1) {
    if (singleBtn) {
      singleBtn.classList.remove('hidden');
      singleBtn.textContent = 'Validar e processar planilha';
    }
    if (multiActions) multiActions.classList.add('hidden');
  } else {
    if (singleBtn) singleBtn.classList.add('hidden');
    if (multiActions) multiActions.classList.remove('hidden');

    const btnMerge = document.getElementById('btnMergeUpload');
    const btnIndiv = document.getElementById('btnIndividualUpload');

    if (btnMerge) {
      const mergeLabel = state.currentAction === 'enrich_only' ? 'Mesclar enriquecimento' : 'Mesclar planilhas';
      btnMerge.innerHTML = `
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M8 3H5a2 2 0 0 0-2 2v3m18 0V5a2 2 0 0 0-2-2h-3m0 18h3a2 2 0 0 0 2-2v-3M3 16v3a2 2 0 0 0 2 2h3"/></svg>
        <span>${mergeLabel}</span>
      `;
    }

    if (btnIndiv) {
      let indivLabel = 'Subir arquivos individuais';
      if (state.currentAction === 'enrich_only') {
        indivLabel = 'Enriquecer arquivos individuais';
      } else if (state.currentAction === 'enrich_and_queue') {
        indivLabel = 'Filas diferentes (individuais)';
      }
      btnIndiv.innerHTML = `
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>
        <span>${indivLabel}</span>
      `;
    }
  }
}

window.removeStagedFile = removeStagedFile;
window.clearStagedFiles = clearStagedFiles;

async function uploadSpreadsheetFile(file) {
  if (!file) {
    showMessage('Selecione uma planilha primeiro.', 'error');
    return;
  }

  const submitButton = uploadForm ? uploadForm.querySelector("button[type='submit']") : null;

  // Show loading state
  if (dropZone) {
    dropZone.classList.add('drop-zone--loading');
    let spinner = dropZone.querySelector('.drop-zone-spinner');
    if (!spinner) {
      spinner = document.createElement('div');
      spinner.className = 'drop-zone-spinner';
      dropZone.insertBefore(spinner, dropZone.firstChild);
    }
    const textEl = dropZone.querySelector('.drop-zone-text');
    if (textEl) {
      textEl.dataset.originalText = textEl.textContent;
      textEl.textContent = 'Carregando e validando planilha...';
    }
  }
  if (submitButton) {
    submitButton.disabled = true;
    submitButton.textContent = 'Aguarde...';
  }

  const updateProgressText = (pct) => {
    const textEl = dropZone ? dropZone.querySelector('.drop-zone-text') : null;
    if (textEl) {
      textEl.textContent = `Enviando arquivo: ${pct}%...`;
    }
  };

  try {
    let payload = null;
    const CHUNK_SIZE = 2.5 * 1024 * 1024; // 2.5 MB por fatia para estabilidade máxima em túneis

    if (file.size > CHUNK_SIZE) {
      const totalChunks = Math.ceil(file.size / CHUNK_SIZE);
      const uploadId = 'up_' + Date.now() + '_' + Math.random().toString(36).substring(2, 9);

      for (let i = 0; i < totalChunks; i++) {
        const start = i * CHUNK_SIZE;
        const end = Math.min(file.size, start + CHUNK_SIZE);
        const chunkBlob = file.slice(start, end);

        const chunkFormData = new FormData();
        chunkFormData.append('upload_id', uploadId);
        chunkFormData.append('chunk_index', i);
        chunkFormData.append('total_chunks', totalChunks);
        chunkFormData.append('filename', file.name);
        chunkFormData.append('chunk', chunkBlob, file.name);

        const pct = Math.round(((i + 1) / totalChunks) * 100);
        updateProgressText(pct);

        const response = await fetch('/api/upload/chunk', {
          method: 'POST',
          body: chunkFormData,
        });

        const responseText = await response.text();
        let chunkResp = null;
        try {
          chunkResp = JSON.parse(responseText);
        } catch (parseErr) {
          if (response.status === 401) {
            throw new Error('Sua sessão expirou. Atualize a página e faça login novamente.');
          }
          if (response.status === 413) {
            throw new Error('O fragmento enviado excedeu o limite do túnel/servidor.');
          }
          throw new Error(`Erro de resposta do servidor (Código ${response.status}).`);
        }

        if (!response.ok) {
          throw new Error(chunkResp.error || 'Falha no envio da planilha.');
        }

        if (i === totalChunks - 1) {
          payload = chunkResp;
        }
      }
    } else {
      const formData = new FormData();
      formData.append('file', file);

      const response = await fetch('/api/upload', {
        method: 'POST',
        body: formData,
      });

      const responseText = await response.text();
      try {
        payload = JSON.parse(responseText);
      } catch (parseErr) {
        if (response.status === 401) {
          throw new Error('Sua sessão expirou. Atualize a página e faça login novamente.');
        }
        if (response.status === 413) {
          throw new Error('O arquivo enviado é muito grande.');
        }
        if (responseText.includes('<!DOCTYPE') || responseText.includes('<html')) {
          throw new Error('O túnel/servidor retornou uma página HTML em vez de dados. Verifique a conexão com o servidor e tente novamente.');
        }
        throw new Error(`Erro de resposta do servidor (Código ${response.status}).`);
      }

      if (!response.ok) {
        throw new Error(payload.error || 'Arquivo inválido.');
      }
    }

    state.currentRecord = payload;
    state.selectedFields = Array.isArray(payload.columns) ? [...payload.columns] : [];

    hideMessage();
    showMessage(`Arquivo validado: ${payload.detected_type}.`, 'success');

    // Exibe prévia visual da planilha carregada (Data Preview)
    renderDataPreview(payload);

    // Preenche o Diagnóstico Pré-Upload (Pre-flight Check)
    if (payload.preflight) {
      const preflight = payload.preflight;
      const preflightCard = document.getElementById('preflightDiagnosticCard');
      if (preflightCard) {
        const scoreEl = document.getElementById('preflightQualityScore');
        const totalRowsEl = document.getElementById('preflightTotalRows');
        const uniqueRowsEl = document.getElementById('preflightUniqueRows');
        const duplicatesEl = document.getElementById('preflightDuplicates');
        const colsEl = document.getElementById('preflightColumnsCount');

        if (scoreEl) scoreEl.textContent = `${preflight.quality_score || 100}%`;
        if (totalRowsEl) totalRowsEl.textContent = (preflight.total_rows || 0).toLocaleString('pt-BR');
        if (uniqueRowsEl) uniqueRowsEl.textContent = (preflight.unique_rows || 0).toLocaleString('pt-BR');
        if (duplicatesEl) duplicatesEl.textContent = (preflight.duplicate_count || 0).toLocaleString('pt-BR');
        if (colsEl) colsEl.textContent = (preflight.columns_count || 0).toLocaleString('pt-BR');

        preflightCard.classList.remove('hidden');
      }
    }

    // Hide old panels
    step2Container.classList.add('hidden');
    step3Container.classList.add('hidden');

    if (payload.invalid_phone_count && payload.invalid_phone_count > 0) {
      validationWarningText.textContent = `Identificamos ${payload.invalid_phone_count} contatos com telefones sem DDD ou formato inválido na sua planilha. Veja os detalhes abaixo:`;

      if (errorDetailsList) {
        errorDetailsList.innerHTML = '';
        (payload.invalid_phones || []).forEach((item) => {
          const div = document.createElement('div');
          div.className = 'error-item';
          div.innerHTML = `
            <span class="error-item-loc">Linha ${item.row}</span>
            <span class="error-item-info">${item.name} (${item.phone})</span>
            <span class="error-item-reason">${item.reason}</span>
          `;
          errorDetailsList.appendChild(div);
        });
      }

      validationWarningBox.classList.remove('hidden');
      step1Container.classList.add('hidden');
    } else {
      validationWarningBox.classList.add('hidden');
      if (state.currentAction === 'queue_only') {
        step3Container.classList.remove('hidden');
        renderAvailableFields();
        renderSelectedFields();
        renderPreview();
      } else {
        step1Container.classList.remove('hidden');
      }
    }

    updateFinalizeButtonLabel();
    renderQueryFieldsSelection(payload.columns || []);
  } catch (error) {
    state.currentRecord = null;
    state.selectedFields = [];
    const mappingPanel = document.querySelector('.mapping-panel');
    if (mappingPanel) mappingPanel.classList.add('hidden');
    previewCard.classList.add('hidden');

    const dataPrev = document.getElementById('dataPreviewContainer');
    if (dataPrev) dataPrev.classList.add('hidden');

    showMessage(error.message, 'error');
  } finally {
    // Restore state
    if (dropZone) {
      dropZone.classList.remove('drop-zone--loading');
      const textEl = dropZone.querySelector('.drop-zone-text');
      if (textEl && textEl.dataset.originalText) {
        textEl.textContent = textEl.dataset.originalText;
      }
    }
    if (submitButton) {
      submitButton.disabled = false;
      submitButton.textContent = 'Validar arquivo';
    }
  }
}

async function uploadMultipleMerged(files) {
  if (!files || !files.length) {
    showMessage('Nenhuma planilha selecionada.', 'error');
    return;
  }

  const btnMerge = document.getElementById('btnMergeUpload');
  const btnIndiv = document.getElementById('btnIndividualUpload');

  // Ativar loading
  if (dropZone) {
    dropZone.classList.add('drop-zone--loading');
    let spinner = dropZone.querySelector('.drop-zone-spinner');
    if (!spinner) {
      spinner = document.createElement('div');
      spinner.className = 'drop-zone-spinner';
      dropZone.insertBefore(spinner, dropZone.firstChild);
    }
    const textEl = dropZone.querySelector('.drop-zone-text');
    if (textEl) {
      textEl.dataset.originalText = textEl.textContent;
      textEl.textContent = `Mesclando e validando ${files.length} planilhas...`;
    }
  }

  if (btnMerge) {
    btnMerge.disabled = true;
    btnMerge.textContent = 'Mesclando...';
  }
  if (btnIndiv) btnIndiv.disabled = true;

  try {
    const formData = new FormData();
    files.forEach(f => {
      formData.append('files', f);
    });

    const response = await fetch('/api/upload/merge', {
      method: 'POST',
      body: formData,
    });

    const responseText = await response.text();
    let payload = null;
    try {
      payload = JSON.parse(responseText);
    } catch (parseErr) {
      if (response.status === 401) {
        throw new Error('Sua sessão expirou. Atualize a página e faça login novamente.');
      }
      if (response.status === 413) {
        throw new Error('O volume total de arquivos excedeu o limite do servidor.');
      }
      throw new Error(`Erro de resposta do servidor (Código ${response.status}).`);
    }

    if (!response.ok) {
      throw new Error(payload.error || 'Erro ao mesclar planilhas.');
    }

    state.currentRecord = payload;
    state.selectedFields = Array.isArray(payload.columns) ? [...payload.columns] : [];

    hideMessage();
    showMessage(`Sucesso: ${payload.message}`, 'success');

    // Exibe prévia visual da planilha consolidada (Data Preview)
    renderDataPreview(payload);

    // Preenche o Diagnóstico Pré-Upload (Pre-flight Check)
    if (payload.preflight) {
      const preflight = payload.preflight;
      const preflightCard = document.getElementById('preflightDiagnosticCard');
      if (preflightCard) {
        const scoreEl = document.getElementById('preflightQualityScore');
        const totalRowsEl = document.getElementById('preflightTotalRows');
        const uniqueRowsEl = document.getElementById('preflightUniqueRows');
        const duplicatesEl = document.getElementById('preflightDuplicates');
        const colsEl = document.getElementById('preflightColumnsCount');

        if (scoreEl) scoreEl.textContent = `${preflight.quality_score || 100}%`;
        if (totalRowsEl) totalRowsEl.textContent = (preflight.total_rows || 0).toLocaleString('pt-BR');
        if (uniqueRowsEl) uniqueRowsEl.textContent = (preflight.unique_rows || 0).toLocaleString('pt-BR');
        if (duplicatesEl) duplicatesEl.textContent = (preflight.duplicate_count || 0).toLocaleString('pt-BR');
        if (colsEl) colsEl.textContent = (preflight.columns_count || 0).toLocaleString('pt-BR');

        preflightCard.classList.remove('hidden');
      }
    }

    // Hide old panels
    step2Container.classList.add('hidden');
    step3Container.classList.add('hidden');

    if (payload.invalid_phone_count && payload.invalid_phone_count > 0) {
      validationWarningText.textContent = `Identificamos ${payload.invalid_phone_count} contatos com telefones sem DDD ou formato inválido na base mesclada. Veja os detalhes abaixo:`;

      if (errorDetailsList) {
        errorDetailsList.innerHTML = '';
        (payload.invalid_phones || []).forEach((item) => {
          const div = document.createElement('div');
          div.className = 'error-item';
          div.innerHTML = `
            <span class="error-item-loc">Linha ${item.row}</span>
            <span class="error-item-info">${item.name} (${item.phone})</span>
            <span class="error-item-reason">${item.reason}</span>
          `;
          errorDetailsList.appendChild(div);
        });
      }

      validationWarningBox.classList.remove('hidden');
      step1Container.classList.add('hidden');
    } else {
      validationWarningBox.classList.add('hidden');
      if (state.currentAction === 'queue_only') {
        step3Container.classList.remove('hidden');
        renderAvailableFields();
        renderSelectedFields();
        renderPreview();
      } else {
        step1Container.classList.remove('hidden');
      }
    }

    updateFinalizeButtonLabel();
    renderQueryFieldsSelection(payload.columns || []);

  } catch (error) {
    state.currentRecord = null;
    state.selectedFields = [];
    showMessage(error.message, 'error');
  } finally {
    if (dropZone) {
      dropZone.classList.remove('drop-zone--loading');
      const textEl = dropZone.querySelector('.drop-zone-text');
      if (textEl && textEl.dataset.originalText) {
        textEl.textContent = textEl.dataset.originalText;
      }
    }
    if (btnMerge) {
      btnMerge.disabled = false;
      btnMerge.textContent = 'Mesclar planilhas';
    }
    if (btnIndiv) btnIndiv.disabled = false;
  }
}

async function uploadSingleFileForBatch(file) {
  const CHUNK_SIZE = 8 * 1024 * 1024;
  if (file.size > CHUNK_SIZE) {
    const totalChunks = Math.ceil(file.size / CHUNK_SIZE);
    const uploadId = 'up_' + Date.now() + '_' + Math.random().toString(36).substring(2, 9);
    let finalPayload = null;

    for (let i = 0; i < totalChunks; i++) {
      const start = i * CHUNK_SIZE;
      const end = Math.min(file.size, start + CHUNK_SIZE);
      const chunkBlob = file.slice(start, end);

      const chunkFormData = new FormData();
      chunkFormData.append('upload_id', uploadId);
      chunkFormData.append('chunk_index', i);
      chunkFormData.append('total_chunks', totalChunks);
      chunkFormData.append('filename', file.name);
      chunkFormData.append('chunk', chunkBlob, file.name);

      const response = await fetch('/api/upload/chunk', {
        method: 'POST',
        body: chunkFormData,
      });

      if (!response.ok) {
        const errPayload = await response.json().catch(() => ({}));
        throw new Error(errPayload.error || `Falha no envio de ${file.name}`);
      }

      if (i === totalChunks - 1) {
        finalPayload = await response.json();
      }
    }
    return finalPayload;
  } else {
    const formData = new FormData();
    formData.append('file', file);
    const resp = await fetch('/api/upload', {
      method: 'POST',
      body: formData,
    });
    if (!resp.ok) {
      const errPayload = await resp.json().catch(() => ({}));
      throw new Error(errPayload.error || `Falha ao validar ${file.name}`);
    }
    return await resp.json();
  }
}

async function prepareIndividualBatchUpload() {
  if (!state.stagedFiles || state.stagedFiles.length === 0) {
    showMessage('Nenhuma planilha anexada.', 'error');
    return;
  }

  const btnIndiv = document.getElementById('btnIndividualUpload');
  if (btnIndiv) {
    btnIndiv.disabled = true;
    btnIndiv.textContent = 'Analisando arquivos...';
  }

  try {
    const firstFile = state.stagedFiles[0];
    const recordPayload = await uploadSingleFileForBatch(firstFile);

    state.currentRecord = recordPayload;
    state.selectedFields = Array.isArray(recordPayload.columns) ? [...recordPayload.columns] : [];
    state.isBatchIndividual = true;

    renderDataPreview(recordPayload);
    if (recordPayload.preflight) {
      const preflight = recordPayload.preflight;
      const preflightCard = document.getElementById('preflightDiagnosticCard');
      if (preflightCard) {
        const scoreEl = document.getElementById('preflightQualityScore');
        const totalRowsEl = document.getElementById('preflightTotalRows');
        const uniqueRowsEl = document.getElementById('preflightUniqueRows');
        const duplicatesEl = document.getElementById('preflightDuplicates');
        const colsEl = document.getElementById('preflightColumnsCount');

        if (scoreEl) scoreEl.textContent = `${preflight.quality_score || 100}%`;
        if (totalRowsEl) totalRowsEl.textContent = (preflight.total_rows || 0).toLocaleString('pt-BR');
        if (uniqueRowsEl) uniqueRowsEl.textContent = (preflight.unique_rows || 0).toLocaleString('pt-BR');
        if (duplicatesEl) duplicatesEl.textContent = (preflight.duplicate_count || 0).toLocaleString('pt-BR');
        if (colsEl) colsEl.textContent = (preflight.columns_count || 0).toLocaleString('pt-BR');

        preflightCard.classList.remove('hidden');
      }
    }

    if (state.currentAction === 'queue_only') {
      step1Container.classList.add('hidden');
      step2Container.classList.add('hidden');
      step3Container.classList.remove('hidden');
      renderAvailableFields();
      renderSelectedFields();
      renderPreview();
      updateFinalizeButtonLabel();
      showToast(`${state.stagedFiles.length} planilhas carregadas. Veja a prévia e os atributos padrão para subir fila.`, 'info');
    } else {
      step2Container.classList.add('hidden');
      step3Container.classList.add('hidden');
      step1Container.classList.remove('hidden');
      updateFinalizeButtonLabel();
      renderQueryFieldsSelection(recordPayload.columns || []);
      showToast(`${state.stagedFiles.length} planilhas carregadas. Escolha os atributos com que deseja enriquecer.`, 'info');
    }
  } catch (err) {
    showMessage(`Erro ao preparar arquivos individuais: ${err.message}`, 'error');
  } finally {
    if (btnIndiv) {
      btnIndiv.disabled = false;
      btnIndiv.textContent = 'Subir arquivos individuais';
    }
  }
}

function openBatchUploadModal() {
  if (!state.stagedFiles || state.stagedFiles.length === 0) {
    showMessage('Nenhuma planilha anexada.', 'error');
    return;
  }

  setExportFormat(state.exportFormat || 'original');

  const modal = document.getElementById('batchUploadModal');
  const titleEl = document.getElementById('batchUploadModalTitle');
  const countEl = document.getElementById('batchUploadFilesCount');
  const descEl = document.getElementById('batchUploadModalDescription');
  const reqInput = document.getElementById('batchRequesterName');
  const queueSection = document.getElementById('batchQueueSection');
  const filesList = document.getElementById('batchQueueFilesList');
  const confirmBtnText = document.getElementById('btnConfirmBatchUploadText');
  const btnApplySameQueue = document.getElementById('btnApplySameQueue');

  const count = state.stagedFiles.length;
  if (countEl) {
    countEl.textContent = `${count} ${count === 1 ? 'planilha' : 'planilhas'}`;
  }

  if (reqInput) {
    reqInput.value = state.requesterName || '';
  }

  const currentMode = state.currentAction || 'enrich_only';

  if (currentMode === 'enrich_only') {
    if (titleEl) titleEl.textContent = 'Enriquecer Arquivos em Lote';
    if (descEl) descEl.innerHTML = `Você selecionou <strong style="color: var(--text-main);">${count} planilhas</strong> para enriquecimento individual. Informe o nome do solicitante abaixo para iniciar.`;
    if (queueSection) queueSection.classList.add('hidden');
    if (confirmBtnText) confirmBtnText.textContent = 'Iniciar Enriquecimento em Lote';
  } else if (currentMode === 'enrich_and_queue') {
    if (titleEl) titleEl.textContent = 'Enriquecer e Subir Filas Individuais';
    if (descEl) descEl.innerHTML = `Você selecionou <strong style="color: var(--text-main);">${count} planilhas</strong>. Defina a fila de cada arquivo ou clique em <strong>Mesma fila</strong> para preencher automaticamente.`;
    if (queueSection) queueSection.classList.remove('hidden');
    if (confirmBtnText) confirmBtnText.textContent = 'Iniciar Enriquecimento e Subir Filas';
  } else {
    // queue_only
    if (titleEl) titleEl.textContent = 'Subir Arquivos em Lotes Individuais';
    if (descEl) descEl.innerHTML = `Você selecionou <strong style="color: var(--text-main);">${count} planilhas</strong>. Defina a fila de cada arquivo ou clique em <strong>Mesma fila</strong> para preencher automaticamente.`;
    if (queueSection) queueSection.classList.remove('hidden');
    if (confirmBtnText) confirmBtnText.textContent = 'Importar e Baixar Planilhas';
  }

  // Renderiza a lista de arquivos com seus inputs de fila
  if (queueSection && !queueSection.classList.contains('hidden') && filesList) {
    filesList.innerHTML = '';
    state.stagedFiles.forEach((file, idx) => {
      const item = document.createElement('div');
      item.className = 'batch-queue-file-item';
      item.innerHTML = `
        <div class="batch-queue-file-info">
          <span class="batch-queue-file-index">#${idx + 1}</span>
          <span class="batch-queue-file-name" title="${escapeHtml(file.name)}">${escapeHtml(file.name)}</span>
          <span class="batch-queue-file-size">(${formatFileSize(file.size)})</span>
        </div>
        <div class="batch-queue-file-input-wrap">
          <input type="text" class="batch-file-queue-input" data-index="${idx}" maxlength="6" inputmode="numeric" placeholder="Fila 6 dígitos" required />
        </div>
      `;
      filesList.appendChild(item);
    });

    // Configura o botão "Mesma fila"
    if (btnApplySameQueue) {
      btnApplySameQueue.onclick = () => {
        const inputs = Array.from(filesList.querySelectorAll('.batch-file-queue-input'));
        if (inputs.length === 0) return;

        let sourceVal = inputs[0].value.trim();
        if (!sourceVal) {
          const firstFilled = inputs.find(i => i.value.trim().length > 0);
          if (firstFilled) sourceVal = firstFilled.value.trim();
        }

        if (!sourceVal || !/^\d{1,6}$/.test(sourceVal)) {
          showToast('Digite o número da fila no 1º arquivo antes de clicar em Mesma fila.', 'warning');
          inputs[0].focus();
          return;
        }

        inputs.forEach(input => {
          input.value = sourceVal;
        });

        showToast(`Número de fila ${sourceVal} aplicado a todas as ${inputs.length} planilhas!`, 'info');
      };
    }
  }

  if (modal) {
    modal.classList.remove('hidden');
    if (reqInput && !reqInput.value) {
      setTimeout(() => reqInput.focus(), 80);
    } else if (filesList) {
      const firstInput = filesList.querySelector('.batch-file-queue-input');
      if (firstInput) setTimeout(() => firstInput.focus(), 80);
    }
  }
}

function closeBatchUploadModal() {
  const modal = document.getElementById('batchUploadModal');
  if (modal) modal.classList.add('hidden');
}

async function submitBatchUpload() {
  const reqInput = document.getElementById('batchRequesterName');
  const obsInput = document.getElementById('batchObservacoes');
  const requesterName = reqInput ? reqInput.value.trim() : '';
  const observacoes = obsInput ? obsInput.value.trim() : '';
  const currentMode = state.currentAction || 'enrich_only';

  if (!requesterName) {
    showToast('Por favor, informe o nome do solicitante.', 'error');
    if (reqInput) reqInput.focus();
    return;
  }

  const fileQueues = [];
  if (currentMode !== 'enrich_only') {
    const queueInputs = Array.from(document.querySelectorAll('.batch-file-queue-input'));
    for (let i = 0; i < state.stagedFiles.length; i++) {
      const file = state.stagedFiles[i];
      const input = queueInputs[i];
      const queueVal = input ? input.value.trim() : '';
      if (!/^\d{6}$/.test(queueVal)) {
        showToast(`A planilha #${i + 1} (${file.name}) precisa de um número de fila com exatamente 6 dígitos numéricos.`, 'error');
        if (input) input.focus();
        return;
      }
      fileQueues.push(queueVal);
    }
  }

  const batchRadio = document.querySelector('input[name="batchExportFormat"]:checked');
  if (batchRadio) {
    setExportFormat(batchRadio.value);
  }

  state.requesterName = requesterName;
  state.observacoes = observacoes;
  closeBatchUploadModal();

  await executeBatchUpload(state.stagedFiles, fileQueues, requesterName, observacoes);
}

async function executeBatchUpload(files, fileQueues, requesterName, observacoes) {
  if (!files || !files.length) return;

  const total = files.length;
  const currentMode = state.currentAction || 'enrich_only';
  const chosenQueryFields = Array.from(state.selectedQueryFields || []);
  const chosenSelectedFields = normalizeSelectedFields(state.selectedFields);

  const btnMerge = document.getElementById('btnMergeUpload');
  const btnIndiv = document.getElementById('btnIndividualUpload');

  if (btnMerge) btnMerge.disabled = true;
  if (btnIndiv) {
    btnIndiv.disabled = true;
    btnIndiv.textContent = 'Processando lotes...';
  }

  if (dropZone) {
    dropZone.classList.add('drop-zone--loading');
  }

  // Oculta os passos anteriores
  if (step1Container) step1Container.classList.add('hidden');
  if (step3Container) step3Container.classList.add('hidden');

  let successCount = 0;
  let failedFiles = [];
  const batchResults = [];

  try {
    for (let i = 0; i < total; i++) {
      const file = files[i];
      const fileQueue = (fileQueues && fileQueues[i]) ? fileQueues[i] : 'ENRICH';

      let generatedFilename = `fila_${fileQueue}.xlsx`;
      if (currentMode === 'enrich_only') {
        const baseName = file.name.replace(/\.[^/.]+$/, '').replace(/[^\w\s-]/g, '_');
        generatedFilename = `enriquecido_${baseName}.xlsx`;
      } else {
        const previousSameQueueCount = fileQueues.slice(0, i).filter(q => q === fileQueue).length;
        if (previousSameQueueCount > 0) {
          generatedFilename = `fila_${fileQueue}(${previousSameQueueCount}).xlsx`;
        }
      }

      const textEl = dropZone ? dropZone.querySelector('.drop-zone-text') : null;
      if (textEl) {
        textEl.textContent = `Processando arquivo (${i + 1} de ${total}): ${file.name}...`;
      }

      try {
        const recordPayload = await uploadSingleFileForBatch(file);
        let queryFields = chosenQueryFields;
        if (!queryFields || queryFields.length === 0) {
          queryFields = (recordPayload.columns || []).filter(c => isRequiredNameField(c) || isRequiredPhoneField(c) || isRequiredCpfField(c) || isRequiredUfField(c));
          if (queryFields.length === 0) queryFields = recordPayload.columns || [];
        }

        const reqResp = await fetch('/api/request-import', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            record_id: recordPayload.id,
            query_fields: queryFields,
            mode: currentMode,
            requester_name: requesterName
          })
        });

        if (!reqResp.ok) {
          const reqErr = await reqResp.json().catch(() => ({}));
          throw new Error(reqErr.error || 'Falha ao registrar solicitação.');
        }

        const reqPayload = await reqResp.json();
        const requestId = reqPayload.request_id;

        // Aguarda o processamento de fundo (tanto para enriquecimento quanto para subir fila assíncrono em sub-lotes)
        let isProcessing = true;
        let pollAttempts = 0;
        const actionLabel = currentMode === 'queue_only' ? 'Processando fila' : 'Enriquecendo';
        while (isProcessing && pollAttempts < 1200) {
          await new Promise(r => setTimeout(r, 800));
          pollAttempts++;
          const logResp = await fetch(`/api/request/${requestId}/logs`);
          if (logResp.ok) {
            const logData = await logResp.json();
            const pct = logData.total_rows > 0 ? Math.round((logData.processed_count / logData.total_rows) * 100) : 0;
            if (textEl) {
              if (logData.chunk_info && logData.chunk_info.is_chunked) {
                textEl.textContent = `${actionLabel} (${i + 1}/${total}) ${file.name} - Lote ${logData.chunk_info.current_chunk}/${logData.chunk_info.total_chunks}: ${logData.processed_count}/${logData.total_rows} (${pct}%)...`;
              } else {
                textEl.textContent = `${actionLabel} (${i + 1}/${total}) ${file.name}: ${logData.processed_count}/${logData.total_rows} (${pct}%)...`;
              }
            }
            if (logData.status !== 'processing' && logData.status !== 'queued') {
              isProcessing = false;
              if (logData.status === 'error') {
                throw new Error('Falha no processamento do lote.');
              }
            }
          }
        }

        // Finaliza a solicitação com as colunas escolhidas no Passo 2
        let fieldsToFinalize = chosenSelectedFields;
        if (!fieldsToFinalize || fieldsToFinalize.length === 0) {
          fieldsToFinalize = recordPayload.columns || [];
        }

        const chosenFmt = state.exportFormat || 'original';
        const finResp = await fetch(`/api/request/${requestId}/finalize`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            queue_number: fileQueue === 'ENRICH' ? 'ENRICH' : fileQueue,
            requester_name: requesterName,
            observacoes: observacoes,
            selected_fields: fieldsToFinalize,
            mode: currentMode,
            export_format: chosenFmt,
          })
        });

        if (!finResp.ok) {
          const finErr = await finResp.json().catch(() => ({}));
          throw new Error(finErr.error || 'Falha ao finalizar lote.');
        }

        batchResults.push({
          originalName: file.name,
          generatedFilename: generatedFilename,
          queueNumber: fileQueue,
          requestId: requestId,
          downloadUrl: `/api/request/${requestId}/download?format=${chosenFmt}&custom_name=${encodeURIComponent(generatedFilename)}`,
          totalRows: recordPayload.total_in || 0,
        });

        successCount++;
      } catch (errFile) {
        console.error(`Erro ao processar ${file.name}:`, errFile);
        failedFiles.push(`${file.name} (${errFile.message})`);
      }
    }

    if (successCount > 0) {
      clearStagedFiles();
      loadQueue();
      loadLogsQueue();

      const zipBaseName = currentMode === 'enrich_only' ? 'enriquecimentos_lote' : (fileQueues[0] ? `fila_${fileQueues[0]}` : 'lote_filas');
      const zipFilename = `${zipBaseName}.zip`;
      renderBatchResultModal(batchResults, zipFilename);

      showMessage(`${successCount} de ${total} arquivos foram processados com sucesso!`, 'success');
    } else {
      showMessage('Falha ao processar os arquivos do lote.', 'error');
    }

    if (failedFiles.length > 0) {
      showToast(`Falhas: ${failedFiles.join(' | ')}`, 'warning');
    }

  } catch (err) {
    showMessage(`Erro geral no processamento em lote: ${err.message}`, 'error');
  } finally {
    if (dropZone) {
      dropZone.classList.remove('drop-zone--loading');
      const textEl = dropZone.querySelector('.drop-zone-text');
      if (textEl && textEl.dataset.originalText) {
        textEl.textContent = textEl.dataset.originalText;
      }
    }
    if (btnMerge) btnMerge.disabled = false;
    if (btnIndiv) {
      btnIndiv.disabled = false;
      btnIndiv.textContent = 'Subir arquivos individuais';
    }
    renderStagedFiles();
  }
}

function renderBatchResultModal(results, zipFilename) {
  const modal = document.getElementById('batchResultModal');
  const countEl = document.getElementById('batchResultTotalCount');
  const listEl = document.getElementById('batchResultFilesList');
  const btnDownloadZip = document.getElementById('btnDownloadBatchZip');
  const btnDownloadZipText = document.getElementById('btnDownloadBatchZipText');
  const btnDownloadAll = document.getElementById('btnDownloadAllBatchFiles');

  if (!modal || !listEl) return;

  if (countEl) {
    countEl.textContent = `${results.length} ${results.length === 1 ? 'planilha' : 'planilhas'}`;
  }

  listEl.innerHTML = '';
  results.forEach((res) => {
    const card = document.createElement('div');
    card.className = 'batch-result-item-card';
    card.innerHTML = `
      <div class="batch-result-item-info">
        <span class="batch-result-item-badge">${escapeHtml(res.generatedFilename)}</span>
        <div class="batch-result-item-details">
          <span class="batch-result-item-title" title="${escapeHtml(res.originalName)}">${escapeHtml(res.originalName)}</span>
          <span class="batch-result-item-meta">${res.totalRows} contatos formatados</span>
        </div>
      </div>
      <button type="button" class="batch-result-download-btn" onclick="downloadFileDirectly('${res.downloadUrl}', '${escapeHtml(res.generatedFilename)}')">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
        Baixar Planilha
      </button>
    `;
    listEl.appendChild(card);
  });

  // Configuração do Download em ZIP
  if (btnDownloadZip) {
    btnDownloadZip.onclick = async () => {
      try {
        if (btnDownloadZipText) btnDownloadZipText.textContent = 'Gerando ZIP...';
        btnDownloadZip.disabled = true;

        const resp = await fetch('/api/batch/download-zip', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            zip_name: zipFilename,
            items: results.map(r => ({
              request_id: r.requestId,
              filename: r.generatedFilename,
            }))
          })
        });

        if (!resp.ok) {
          const errData = await resp.json().catch(() => ({}));
          throw new Error(errData.error || 'Erro ao gerar arquivo ZIP.');
        }

        const blob = await resp.blob();
        const blobUrl = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = blobUrl;
        a.download = zipFilename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(blobUrl), 1000);

        showToast(`Download de "${zipFilename}" iniciado!`, 'success');
      } catch (err) {
        showToast(err.message || 'Erro ao baixar ZIP.', 'error');
      } finally {
        if (btnDownloadZipText) btnDownloadZipText.textContent = 'Baixar Todas em ZIP';
        btnDownloadZip.disabled = false;
      }
    };
  }

  modal.classList.remove('hidden');
}

function closeBatchResultModal() {
  const modal = document.getElementById('batchResultModal');
  if (modal) modal.classList.add('hidden');
}

window.openBatchUploadModal = openBatchUploadModal;
window.closeBatchUploadModal = closeBatchUploadModal;
window.submitBatchUpload = submitBatchUpload;
window.closeBatchResultModal = closeBatchResultModal;

async function handleUpload(event) {
  event.preventDefault();
  if (state.stagedFiles.length === 0) {
    const file = document.getElementById('excelFile').files[0];
    if (file) {
      await uploadSpreadsheetFile(file);
    } else {
      showMessage('Selecione uma planilha primeiro.', 'error');
    }
  } else if (state.stagedFiles.length === 1) {
    await uploadSpreadsheetFile(state.stagedFiles[0]);
  } else {
    await uploadMultipleMerged(state.stagedFiles);
  }
}

function openRequestModal() {
  if (!state.currentRecord || !state.selectedFields.length) {
    showMessage('Selecione nome e telefone para criar a solicitação.', 'error');
    return;
  }

  const hasName = state.selectedFields.some((field) => isRequiredNameField(field));
  const hasPhone = state.selectedFields.some((field) => isRequiredPhoneField(field));
  if (!hasName || !hasPhone) {
    showMessage('Nome e telefone são obrigatórios para criar a solicitação.', 'error');
    return;
  }

  const reqInput = document.getElementById('requesterName');
  if (reqInput && state.requesterName && !reqInput.value) {
    reqInput.value = state.requesterName;
  }

  requestModal.classList.remove('hidden');
}

function closeRequestModal() {
  requestModal.classList.add('hidden');
  requestForm.reset();
}

async function submitRequest(event) {
  event.preventDefault();

  const submitBtn = event?.target?.querySelector('button[type="submit"]') || document.querySelector('#requestForm button[type="submit"]');
  if (submitBtn && submitBtn.disabled) return;

  const queueNumber = document.getElementById('queueNumber').value.trim();
  const requesterName = document.getElementById('requesterName').value.trim();
  const observacoes = document.getElementById('observacoes').value.trim();

  if (!/^\d{6}$/.test(queueNumber)) {
    showMessage('O número da fila deve conter 6 dígitos.', 'error');
    return;
  }

  if (!requesterName) {
    showMessage('O nome do solicitante é obrigatório.', 'error');
    return;
  }

  if (submitBtn) submitBtn.disabled = true;
  state.queueNumber = queueNumber;
  state.requesterName = requesterName;
  state.observacoes = observacoes;
  closeRequestModal();
  if (submitBtn) setTimeout(() => { submitBtn.disabled = false; }, 1000);

  if (state.currentAction === 'queue_only') {
    // Transição visual imediata para a tela de progresso (step2Container)
    if (step1Container) step1Container.classList.add('hidden');
    if (step3Container) step3Container.classList.add('hidden');
    if (step2Container) step2Container.classList.remove('hidden');

    const progressSubheader = document.getElementById('progressSubheader');
    if (progressSubheader) {
      progressSubheader.textContent = 'Formatando e validando contatos da fila em tempo real.';
    }

    const stage2TitleEl = stageBoxEnrich ? stageBoxEnrich.querySelector('.title') : null;
    if (stage2TitleEl) {
      stage2TitleEl.textContent = 'Processando contatos da fila';
    }

    updateSkillStage(stageBoxValidate, stageBarValidate, stageTooltipValidate, stageBadgeValidate, 'active', 60, 'Validando estrutura...');
    updateSkillStage(stageBoxEnrich, stageBarEnrich, stageTooltipEnrich, stageBadgeEnrich, '', 0, 'Aguardando');
    updateSkillStage(stageBoxClean, stageBarClean, stageTooltipClean, stageBadgeClean, '', 0, 'Aguardando');
    updateSkillStage(stageBoxConcat, stageBarConcat, stageTooltipConcat, stageBadgeConcat, '', 0, 'Aguardando');

    try {
      const selectedFieldsForRequest = normalizeSelectedFields(state.selectedFields);
      const deduplicate = document.getElementById('chkDeduplicateRows')?.checked ?? true;

      const response = await fetch('/api/request-import', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          record_id: state.currentRecord?.id,
          query_fields: Array.from(state.selectedQueryFields || []),
          mode: 'queue_only',
          requester_name: requesterName,
          queue_number: queueNumber,
          observacoes: observacoes,
          selected_fields: selectedFieldsForRequest,
          deduplicate: deduplicate
        })
      });

      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || 'Não foi possível registrar a fila.');
      }

      state.activeRequestImportId = payload.request_id;
      pollUploadProgress(payload.request_id);
    } catch (error) {
      if (step2Container) step2Container.classList.add('hidden');
      if (step3Container) step3Container.classList.remove('hidden');
      showMessage(error.message, 'error');
    }
  } else {
    // Modo enrich_and_queue: inicia enriquecimento e quando concluir finaliza automaticamente com a fila
    startEnrichment(requesterName, queueNumber, observacoes);
  }
}

async function loadQueue() {
  try {
    const response = await fetch('/api/queue');
    if (!response.ok) {
      throw new Error('Não foi possível carregar a fila.');
    }

    const payload = await response.json();
    state.queue = payload.queue || [];
    if (state.currentRole === 'admin') {
      renderQueueActive();
      renderQueueHistory();
    } else {
      renderNormalUserHistory();
    }
    loadAnalyticsDashboard();
  } catch (error) {
    console.error(error);
  }
}

function renderNormalUserHistory() {
  const container = document.getElementById('normalUserHistoryQueue');
  if (!container) return;

  const historyMetricsGrid = document.getElementById('historyMetricsGrid');
  const historyFilterBar = document.querySelector('.history-filter-bar');

  // Identificar todas as solicitações e enriquecimentos reais deste usuário
  const allUserItems = (state.queue || []).filter((item) => {
    const queueNumber = String(item.queue_number || '').trim();
    const isEnrichOnly = item.mode === 'enrich_only' || queueNumber === 'ENRICH';
    const isImportQueue = queueNumber && queueNumber !== 'ENRICH' && queueNumber !== 'TEMP';
    return isEnrichOnly || isImportQueue;
  });

  // Se não houver nenhum histórico real no sistema, oculta os cards de métricas e a barra de filtro
  if (allUserItems.length === 0) {
    if (historyMetricsGrid) historyMetricsGrid.style.display = 'none';
    if (historyFilterBar) historyFilterBar.style.display = 'none';
    container.innerHTML = '<div class="queue-item"><strong>Nenhum histórico disponível no momento.</strong></div>';
    return;
  }

  // Quando houver histórico real, exibe o painel e calcula as métricas com dados 100% reais dos arquivos
  if (historyMetricsGrid) historyMetricsGrid.style.display = 'grid';
  if (historyFilterBar) historyFilterBar.style.display = 'flex';

  const totalRows = allUserItems.reduce((acc, i) => acc + (parseInt(i.total_rows, 10) || 0), 0);
  const totalEnriched = allUserItems.reduce((acc, i) => acc + (parseInt(i.total_rows_enriched || i.success_count, 10) || 0), 0);
  const totalFiles = allUserItems.length;
  const successRate = totalRows > 0 ? ((totalEnriched / totalRows) * 100).toFixed(1).replace('.0', '') + '%' : '0%';

  const metricTotalRows = document.getElementById('metricTotalRows');
  const metricEnrichedRows = document.getElementById('metricEnrichedRows');
  const metricTotalFiles = document.getElementById('metricTotalFiles');
  const metricSuccessRate = document.getElementById('metricSuccessRate');

  if (metricTotalRows) metricTotalRows.textContent = totalRows.toLocaleString('pt-BR');
  if (metricEnrichedRows) metricEnrichedRows.textContent = totalEnriched.toLocaleString('pt-BR');
  if (metricTotalFiles) metricTotalFiles.textContent = totalFiles.toLocaleString('pt-BR');
  if (metricSuccessRate) metricSuccessRate.textContent = successRate;

  const searchText = (state.historySearchText || '').trim().toLowerCase();
  const statusFilter = state.historyStatusFilter || 'all';
  const modeFilter = state.historyModeFilter || 'all';

  const enrichItems = [];
  const importItems = [];

  allUserItems.forEach((item) => {
    // Filtro por modo
    if (modeFilter !== 'all') {
      const isEnrichOnly = item.mode === 'enrich_only' || item.queue_number === 'ENRICH';
      const isQueueOnly = item.mode === 'queue_only';
      const isEnrichAndQueue = item.mode === 'enrich_and_queue';
      if (modeFilter === 'enrich_only' && !isEnrichOnly) return;
      if (modeFilter === 'queue_only' && !isQueueOnly) return;
      if (modeFilter === 'enrich_and_queue' && !isEnrichAndQueue) return;
    }

    // Filtro por status
    if (statusFilter !== 'all') {
      if (statusFilter === 'completed' && item.status !== 'completed') return;
      if (statusFilter === 'processing' && item.status !== 'processing') return;
      if (statusFilter === 'pending' && item.status !== 'pending') return;
      if (statusFilter === 'rejected' && item.status !== 'rejected') return;
    }

    // Filtro por busca de texto
    if (searchText) {
      const matchName = String(item.requester_name || '').toLowerCase().includes(searchText);
      const matchQueue = String(item.queue_number || '').toLowerCase().includes(searchText);
      const matchObs = String(item.observacoes || '').toLowerCase().includes(searchText);
      const matchFile = String(item.filename || '').toLowerCase().includes(searchText);
      if (!matchName && !matchQueue && !matchObs && !matchFile) return;
    }

    const queueNumber = String(item.queue_number || '').trim();
    const isEnrichOnly = item.mode === 'enrich_only' || queueNumber === 'ENRICH';

    if (isEnrichOnly) {
      enrichItems.push(item);
    } else {
      importItems.push(item);
    }
  });

  const buildSection = (title, items, emptyText, variant = 'enrich') => {
    const panel = document.createElement('div');
    panel.className = 'history-panel';
    panel.style.display = 'flex';
    panel.style.flexDirection = 'column';
    panel.style.gap = '0.75rem';
    panel.style.minHeight = '220px';

    const header = document.createElement('div');
    header.className = 'section-title';
    header.style.marginBottom = '0.5rem';
    header.style.fontWeight = '700';
    header.textContent = title;
    panel.appendChild(header);

    if (!items.length) {
      panel.innerHTML += `<div class="queue-item"><strong>${emptyText}</strong></div>`;
      return panel;
    }

    items.forEach((item) => {
      const row = document.createElement('article');
      const isCollapsed = !expandedCardIds.has(item.id);
      row.className = `queue-item ${isCollapsed ? 'collapsed' : ''}`;
      const rotation = isCollapsed ? '0deg' : '-180deg';

      let actionBtnHtml = '';
      if (item.status === 'rejected') {
        const reason = getItemRejectReason(item);
        actionBtnHtml = `<button type="button" class="button secondary" style="background-color: var(--danger-color); color: white; border-color: var(--danger-color); padding: 0.4rem 1rem; border-radius: var(--radius-sm); font-size: 0.8rem;" onclick="openViewReasonModal('${encodeURIComponent(reason)}')">Motivo da Recusa</button>`;
      } else if (item.status === 'processing') {
        actionBtnHtml = `<div class="queue-meta" style="color: var(--text-muted); font-size: 0.85rem;">Processando...</div>`;
      } else if (item.status === 'pending' && variant === 'import') {
        actionBtnHtml = `<div class="queue-meta" style="color: var(--text-muted); font-size: 0.85rem;">Aguardando processamento...</div>`;
      } else {
        if (variant === 'import') {
          let importBtn = `<button type="button" class="download-link-premium" style="display: inline-flex; align-items: center; gap: 6px; border: none; font-family: inherit; font-size: 0.82rem; padding: 6px 12px;" onclick="downloadFileDirectly('/api/request/${item.id}/download')" title="Baixar Planilha Tratada com Coluna 'importar' pronta para copiar e colar">
                  <svg viewBox="0 0 256 256" height="20" width="22" xmlns="http://www.w3.org/2000/svg">
                    <path d="M74.34 85.66a8 8 0 0 1 11.32-11.32L120 108.69V24a8 8 0 0 1 16 0v84.69l34.34-34.35a8 8 0 0 1 11.32 11.32l-48 48a8 8 0 0 1-11.32 0ZM240 136v64a16 16 0 0 1-16 16H32a16 16 0 0 1-16-16v-64a16 16 0 0 1 16-16h52.4a4 4 0 0 1 2.83 1.17L111 145a24 24 0 0 0 34 0l23.8-23.8a4 4 0 0 1 2.8-1.2H224a16 16 0 0 1 16 16m-40 32a12 12 0 1 0-12 12a12 12 0 0 0 12-12" fill="currentColor"></path>
                  </svg>
                  <span>Para Importar</span>
                </button>`;

          let fullBtn = `<button type="button" class="download-link-premium" style="display: inline-flex; align-items: center; gap: 6px; border: none; font-family: inherit; font-size: 0.82rem; padding: 6px 12px; opacity: 0.85;" onclick="downloadFileDirectly('/api/request/${item.id}/download?format=full')" title="Baixar Planilha Completa com todas as informações e colunas">
                  <svg viewBox="0 0 256 256" height="20" width="22" xmlns="http://www.w3.org/2000/svg">
                    <path d="M74.34 85.66a8 8 0 0 1 11.32-11.32L120 108.69V24a8 8 0 0 1 16 0v84.69l34.34-34.35a8 8 0 0 1 11.32 11.32l-48 48a8 8 0 0 1-11.32 0ZM240 136v64a16 16 0 0 1-16 16H32a16 16 0 0 1-16-16v-64a16 16 0 0 1 16-16h52.4a4 4 0 0 1 2.83 1.17L111 145a24 24 0 0 0 34 0l23.8-23.8a4 4 0 0 1 2.8-1.2H224a16 16 0 0 1 16 16m-40 32a12 12 0 1 0-12 12a12 12 0 0 0 12-12" fill="currentColor"></path>
                  </svg>
                  <span>Planilha Completa</span>
                </button>`;

          if (item.summary_name) {
            let devolutivaBtn = `<button type="button" class="primary" style="display: inline-flex; align-items: center; gap: 8px; padding: 0.4rem 1rem; border-radius: var(--radius-sm); font-size: 0.8rem;" onclick="downloadDevolutiva('/api/summary/${item.id}')" title="Baixar Devolutiva">
                  <svg viewBox="0 0 256 256" height="24" width="28" xmlns="http://www.w3.org/2000/svg">
                    <path d="M74.34 85.66a8 8 0 0 1 11.32-11.32L120 108.69V24a8 8 0 0 1 16 0v84.69l34.34-34.35a8 8 0 0 1 11.32 11.32l-48 48a8 8 0 0 1-11.32 0ZM240 136v64a16 16 0 0 1-16 16H32a16 16 0 0 1-16-16v-64a16 16 0 0 1 16-16h52.4a4 4 0 0 1 2.83 1.17L111 145a24 24 0 0 0 34 0l23.8-23.8a4 4 0 0 1 2.8-1.2H224a16 16 0 0 1 16 16m-40 32a12 12 0 1 0-12 12a12 12 0 0 0 12-12" fill="currentColor"></path>
                  </svg>
                  <span>Devolutiva</span>
                </button>`;
            actionBtnHtml = `<div style="display: flex; gap: 0.5rem; align-items: center; flex-wrap: wrap;">${importBtn}${fullBtn}${devolutivaBtn}<button type="button" class="secondary" style="padding: 0.4rem 1rem; border-radius: var(--radius-sm); font-size: 0.8rem;" onclick="openImageViewer('/api/summary/${item.id}')">Visualizar</button></div>`;
          } else {
            actionBtnHtml = `<div style="display: flex; gap: 0.5rem; align-items: center; flex-wrap: wrap;">${importBtn}${fullBtn}</div>`;
          }
        } else {
          actionBtnHtml = `<div style="display: flex; gap: 0.4rem; align-items: center; flex-wrap: wrap;">
            <button type="button" class="download-link-premium" style="display: inline-flex; align-items: center; gap: 6px; border: none; font-family: inherit; font-size: 0.82rem; padding: 6px 12px;" onclick="downloadFileDirectly('/api/request/${item.id}/download?format=original')" title="Baixar Planilha no Layout Original (mesma aba e ordem original)">
              <svg viewBox="0 0 256 256" height="20" width="22" xmlns="http://www.w3.org/2000/svg">
                <path d="M74.34 85.66a8 8 0 0 1 11.32-11.32L120 108.69V24a8 8 0 0 1 16 0v84.69l34.34-34.35a8 8 0 0 1 11.32 11.32l-48 48a8 8 0 0 1-11.32 0ZM240 136v64a16 16 0 0 1-16 16H32a16 16 0 0 1-16-16v-64a16 16 0 0 1 16-16h52.4a4 4 0 0 1 2.83 1.17L111 145a24 24 0 0 0 34 0l23.8-23.8a4 4 0 0 1 2.8-1.2H224a16 16 0 0 1 16 16m-40 32a12 12 0 1 0-12 12a12 12 0 0 0 12-12" fill="currentColor"></path>
              </svg>
              <span>Original</span>
            </button>
            <button type="button" class="download-link-premium" style="display: inline-flex; align-items: center; gap: 6px; border: none; font-family: inherit; font-size: 0.82rem; padding: 6px 12px; opacity: 0.85;" onclick="downloadFileDirectly('/api/request/${item.id}/download?format=standard')" title="Baixar Planilha em 3 Abas (Enriquecido, Falha e Resumo)">
              <svg viewBox="0 0 256 256" height="20" width="22" xmlns="http://www.w3.org/2000/svg">
                <path d="M74.34 85.66a8 8 0 0 1 11.32-11.32L120 108.69V24a8 8 0 0 1 16 0v84.69l34.34-34.35a8 8 0 0 1 11.32 11.32l-48 48a8 8 0 0 1-11.32 0ZM240 136v64a16 16 0 0 1-16 16H32a16 16 0 0 1-16-16v-64a16 16 0 0 1 16-16h52.4a4 4 0 0 1 2.83 1.17L111 145a24 24 0 0 0 34 0l23.8-23.8a4 4 0 0 1 2.8-1.2H224a16 16 0 0 1 16 16m-40 32a12 12 0 1 0-12 12a12 12 0 0 0 12-12" fill="currentColor"></path>
              </svg>
              <span>3 Abas</span>
            </button>
          </div>`;
        }
      }
      const totalRows = item.total_rows || 0;
      const enrichedRows = item.total_rows_enriched || item.success_count || 0;
      let statsText = '';
      if (variant === 'enrich') {
        statsText = `${totalRows} linhas originais → ${enrichedRows} linhas enriquecidas`;
      } else if (variant === 'import') {
        statsText = `Fila ${item.queue_number}: ${totalRows} linhas tratadas e concatenadas`;
      }
      row.innerHTML = `
        <div class="queue-header" style="cursor: pointer; display: flex; justify-content: space-between; align-items: center;">
          <div style="display: flex; align-items: center; gap: 0.5rem;">
            <span class="collapse-icon" style="display: inline-block; transition: transform 0.2s; font-size: 0.8rem; color: var(--text-muted); transform: rotate(${rotation});">▼</span>
            <strong>${item.requester_name}</strong>
          </div>
          <span class="badge ${item.status === 'rejected' ? 'error' : item.status === 'pending' ? 'pending' : 'completed'}">
            ${item.status === 'rejected' ? 'Recusado' : item.status === 'pending' ? 'Pendente' : item.status === 'processing' ? 'Processando' : 'Concluído'}
          </span>
        </div>
        <div class="queue-collapsible-content">
          <div class="queue-meta">Solicitante: ${item.requester_name}</div>
          <div class="queue-meta">Horário: ${item.completed_at || item.request_time}</div>
          ${statsText ? `<div class="queue-meta" style="font-size: 0.9rem; font-weight: 500;">${statsText}</div>` : ''}
          <div class="queue-actions" style="margin-top: 0.5rem;">
            ${actionBtnHtml}
          </div>
        </div>
      `;

      const header = row.querySelector('.queue-header');
      header.addEventListener('click', () => {
        row.classList.toggle('collapsed');
        const icon = header.querySelector('.collapse-icon');
        if (row.classList.contains('collapsed')) {
          icon.style.transform = 'rotate(0deg)';
          expandedCardIds.delete(item.id);
        } else {
          icon.style.transform = 'rotate(-180deg)';
          expandedCardIds.add(item.id);
        }
      });

      panel.appendChild(row);
    });

    return panel;
  };

  if (!enrichItems.length && !importItems.length) {
    container.innerHTML = '<div class="queue-item"><strong>Nenhum histórico disponível no momento.</strong></div>';
    return;
  }

  const columns = document.createElement('div');
  columns.style.display = 'grid';
  columns.style.gridTemplateColumns = 'repeat(2, minmax(260px, 1fr))';
  columns.style.gap = '1rem';

  columns.appendChild(buildSection('Histórico de enriquecimentos', enrichItems, 'Nenhum enriquecimento realizado ainda.', 'enrich'));
  columns.appendChild(buildSection('Importações na fila', importItems, 'Nenhuma importação solicitada ainda.', 'import'));

  container.innerHTML = '';
  container.appendChild(columns);
}

let currentZoom = 1;
let isPanning = false;
let startX = 0;
let startY = 0;
let translateX = 0;
let translateY = 0;

function openImageViewer(src) {
  const modal = document.getElementById('imageViewerModal');
  const img = document.getElementById('viewerImage');
  if (modal && img) {
    img.src = src;
    currentZoom = 1;
    translateX = 0;
    translateY = 0;
    img.style.transform = `scale(${currentZoom}) translate(0px, 0px)`;
    img.style.cursor = 'grab';
    modal.classList.remove('hidden');
  }
}

function closeImageViewer() {
  const modal = document.getElementById('imageViewerModal');
  if (modal) {
    modal.classList.add('hidden');
  }
}

function zoomImage(amount) {
  const img = document.getElementById('viewerImage');
  if (img) {
    currentZoom = Math.max(0.5, Math.min(3, currentZoom + amount));
    img.style.transform = `scale(${currentZoom})`;
  }
}

function resetZoom() {
  const img = document.getElementById('viewerImage');
  if (img) {
    currentZoom = 1;
    img.style.transform = `scale(${currentZoom})`;
  }
}

window.openImageViewer = openImageViewer;
window.closeImageViewer = closeImageViewer;
window.zoomImage = zoomImage;
window.resetZoom = resetZoom;

function renderQueueActive() {
  // Filtrar apenas itens ativos (não completados), excluindo 'enrich_only'
  const activeItems = state.queue.filter(item =>
    item.status !== 'completed' &&
    item.status !== 'rejected' &&
    item.mode !== 'enrich_only'
  );

  if (!activeItems.length) {
    strategyQueueActive.innerHTML = '<div class="queue-item"><strong>Sem solicitações ativas no momento.</strong></div>';
    activeBadge.textContent = '0';
    return;
  }

  activeBadge.textContent = activeItems.length;
  strategyQueueActive.innerHTML = '';
  activeItems.forEach((item) => {
    const row = document.createElement('article');
    const isCollapsed = !expandedCardIds.has(item.id);
    row.className = `queue-item ${isCollapsed ? 'collapsed' : ''}`;
    const rotation = isCollapsed ? '0deg' : '-180deg';
    const formattedFields = formatSelectedFields(item.selected_fields || []);

    // Verificar se está aguardando confirmação (queue_number === "TEMP")
    const isAwaitingConfirmation = item.queue_number === "TEMP";
    let queueNumberDisplay = item.queue_number;
    if (isAwaitingConfirmation) {
      queueNumberDisplay = "TEMP (Aguardando dados)";
    }

    let actionsHtml = '';
    let statusText = 'Pendente';
    let badgeClass = 'pending';

    if (item.status === 'processing') {
      const pct = item.total_rows > 0 ? Math.round((item.processed_count / item.total_rows) * 100) : 0;
      if (item.chunk_info && item.chunk_info.is_chunked) {
        statusText = `Processando Lote ${item.chunk_info.current_chunk}/${item.chunk_info.total_chunks} (${pct}%)`;
      } else {
        statusText = `Processando (${pct}%)`;
      }
      badgeClass = 'processing';
    } else if (item.status === 'queued') {
      statusText = 'Na Fila (Aguardando vez)';
      badgeClass = 'queued';
    } else if (item.status === 'error') {
      statusText = 'Erro';
      badgeClass = 'error';
    } else if (item.status === 'rejected') {
      statusText = 'Recusado';
      badgeClass = 'error';
    }

    const isEnrichOnly = item.queue_number === "ENRICH";

    let refuseBtnHtml = '';
    if (item.status === 'pending' || item.status === 'processing' || item.status === 'queued') {
      refuseBtnHtml = `<button type="button" class="button secondary" onclick="openRefuseModal('${item.id}')" style="background-color: var(--danger-color); color: white; border-color: var(--danger-color); font-size: 0.8rem; padding: 0.4rem 1rem; border-radius: var(--radius-sm);">Recusar</button>`;
    }

    if (item.queue_number !== "TEMP") {
      if (isEnrichOnly) {
        actionsHtml = `
          <div style="display: flex; align-items: center; gap: 0.5rem; flex-wrap: wrap;">
            ${refuseBtnHtml}
            <button type="button" class="download-link-premium" style="display: inline-flex; align-items: center; gap: 6px; border: none; font-family: inherit; font-size: 0.82rem; padding: 6px 12px;" onclick="downloadFileDirectly('/api/request/${item.id}/download?format=original')" title="Baixar Planilha no Layout Original (mesma aba e ordem original)">
              <svg viewBox="0 0 256 256" height="20" width="22" xmlns="http://www.w3.org/2000/svg">
                <path d="M74.34 85.66a8 8 0 0 1 11.32-11.32L120 108.69V24a8 8 0 0 1 16 0v84.69l34.34-34.35a8 8 0 0 1 11.32 11.32l-48 48a8 8 0 0 1-11.32 0ZM240 136v64a16 16 0 0 1-16 16H32a16 16 0 0 1-16-16v-64a16 16 0 0 1 16-16h52.4a4 4 0 0 1 2.83 1.17L111 145a24 24 0 0 0 34 0l23.8-23.8a4 4 0 0 1 2.8-1.2H224a16 16 0 0 1 16 16m-40 32a12 12 0 1 0-12 12a12 12 0 0 0 12-12" fill="currentColor"></path>
              </svg>
              <span>Original</span>
            </button>
            <button type="button" class="download-link-premium" style="display: inline-flex; align-items: center; gap: 6px; border: none; font-family: inherit; font-size: 0.82rem; padding: 6px 12px; opacity: 0.85;" onclick="downloadFileDirectly('/api/request/${item.id}/download?format=standard')" title="Baixar Planilha em 3 Abas (Enriquecido, Falha e Resumo)">
              <svg viewBox="0 0 256 256" height="20" width="22" xmlns="http://www.w3.org/2000/svg">
                <path d="M74.34 85.66a8 8 0 0 1 11.32-11.32L120 108.69V24a8 8 0 0 1 16 0v84.69l34.34-34.35a8 8 0 0 1 11.32 11.32l-48 48a8 8 0 0 1-11.32 0ZM240 136v64a16 16 0 0 1-16 16H32a16 16 0 0 1-16-16v-64a16 16 0 0 1 16-16h52.4a4 4 0 0 1 2.83 1.17L111 145a24 24 0 0 0 34 0l23.8-23.8a4 4 0 0 1 2.8-1.2H224a16 16 0 0 1 16 16m-40 32a12 12 0 1 0-12 12a12 12 0 0 0 12-12" fill="currentColor"></path>
              </svg>
              <span>3 Abas</span>
            </button>
          </div>
        `;
      } else {
        actionsHtml = `
          <div style="display: flex; align-items: center; gap: 0.5rem; flex-wrap: wrap;">
            ${refuseBtnHtml}
            <button type="button" class="download-link-premium" style="display: inline-flex; align-items: center; gap: 6px; border: none; font-family: inherit; font-size: 0.82rem; padding: 6px 12px;" onclick="downloadFileDirectly('/api/request/${item.id}/download')" title="Baixar Planilha Tratada com Coluna 'importar'">
              <svg viewBox="0 0 256 256" height="20" width="22" xmlns="http://www.w3.org/2000/svg">
                <path d="M74.34 85.66a8 8 0 0 1 11.32-11.32L120 108.69V24a8 8 0 0 1 16 0v84.69l34.34-34.35a8 8 0 0 1 11.32 11.32l-48 48a8 8 0 0 1-11.32 0ZM240 136v64a16 16 0 0 1-16 16H32a16 16 0 0 1-16-16v-64a16 16 0 0 1 16-16h52.4a4 4 0 0 1 2.83 1.17L111 145a24 24 0 0 0 34 0l23.8-23.8a4 4 0 0 1 2.8-1.2H224a16 16 0 0 1 16 16m-40 32a12 12 0 1 0-12 12a12 12 0 0 0 12-12" fill="currentColor"></path>
              </svg>
              <span>Para Importar</span>
            </button>
            <button type="button" class="download-link-premium" style="display: inline-flex; align-items: center; gap: 6px; border: none; font-family: inherit; font-size: 0.82rem; padding: 6px 12px; opacity: 0.85;" onclick="downloadFileDirectly('/api/request/${item.id}/download?format=full')" title="Baixar Planilha Completa">
              <svg viewBox="0 0 256 256" height="20" width="22" xmlns="http://www.w3.org/2000/svg">
                <path d="M74.34 85.66a8 8 0 0 1 11.32-11.32L120 108.69V24a8 8 0 0 1 16 0v84.69l34.34-34.35a8 8 0 0 1 11.32 11.32l-48 48a8 8 0 0 1-11.32 0ZM240 136v64a16 16 0 0 1-16 16H32a16 16 0 0 1-16-16v-64a16 16 0 0 1 16-16h52.4a4 4 0 0 1 2.83 1.17L111 145a24 24 0 0 0 34 0l23.8-23.8a4 4 0 0 1 2.8-1.2H224a16 16 0 0 1 16 16m-40 32a12 12 0 1 0-12 12a12 12 0 0 0 12-12" fill="currentColor"></path>
              </svg>
              <span>Completo</span>
            </button>
            <input type="file" accept="image/*" class="summary-upload" data-id="${item.id}" />
            <button type="button" class="primary" data-submit-summary="${item.id}">Check</button>
          </div>
        `;
      }
    } else {
      actionsHtml = `<span class="badge pending">Aguardando dados da fila</span>`;
    }

    let rejectReasonHtml = '';
    if (item.status === 'rejected' && item.reject_reason) {
      rejectReasonHtml = `<div class="queue-meta" style="color: var(--danger-color);"><strong>Motivo da Recusa:</strong> ${item.reject_reason}</div>`;
    }

    row.innerHTML = `
      <div class="queue-header" style="cursor: pointer; display: flex; justify-content: space-between; align-items: center;">
        <div style="display: flex; align-items: center; gap: 0.5rem;">
          <span class="collapse-icon" style="display: inline-block; transition: transform 0.2s; font-size: 0.8rem; color: var(--text-muted); transform: rotate(${rotation});">▼</span>
          <strong>Fila ${queueNumberDisplay}</strong>
        </div>
        <span class="badge ${badgeClass}">${statusText}</span>
      </div>
      <div class="queue-collapsible-content">
        <div class="queue-meta">Solicitante: ${item.requester_name}</div>
        <div class="queue-meta">Horário: ${item.request_time}</div>
        <div class="queue-meta">Campos: ${formattedFields}</div>
        ${rejectReasonHtml}
        <div class="queue-actions">
          ${actionsHtml}
          <button type="button" class="button secondary" onclick="viewDetailedLogs('${item.id}')">Logs</button>
        </div>
      </div>
    `;

    // Add collapse toggle handler
    const header = row.querySelector('.queue-header');
    header.addEventListener('click', (e) => {
      row.classList.toggle('collapsed');
      const icon = header.querySelector('.collapse-icon');
      if (row.classList.contains('collapsed')) {
        icon.style.transform = 'rotate(0deg)';
        expandedCardIds.delete(item.id);
      } else {
        icon.style.transform = 'rotate(-180deg)';
        expandedCardIds.add(item.id);
      }
    });

    strategyQueueActive.appendChild(row);
  });
}

let strategyFilterState = {
  search: '',
  status: 'all',
  mode: 'all',
};

function renderQueueHistory() {
  // Filtrar apenas itens completados ou recusados
  let historyItems = state.queue.filter(item => item.status === 'completed' || item.status === 'rejected');

  const search = (strategyFilterState.search || '').trim().toLowerCase();
  const statusFilter = strategyFilterState.status || 'all';
  const modeFilter = strategyFilterState.mode || 'all';

  if (modeFilter !== 'all') {
    historyItems = historyItems.filter(item => {
      const isEnrichOnly = item.mode === 'enrich_only' || item.queue_number === 'ENRICH';
      const isQueueOnly = item.mode === 'queue_only';
      const isEnrichAndQueue = item.mode === 'enrich_and_queue';
      if (modeFilter === 'enrich_only' && !isEnrichOnly) return false;
      if (modeFilter === 'queue_only' && !isQueueOnly) return false;
      if (modeFilter === 'enrich_and_queue' && !isEnrichAndQueue) return false;
      return true;
    });
  }

  if (statusFilter !== 'all') {
    historyItems = historyItems.filter(item => {
      if (statusFilter === 'completed') return item.status === 'completed';
      if (statusFilter === 'rejected') return item.status === 'rejected';
      return true;
    });
  }

  if (search) {
    historyItems = historyItems.filter(item => {
      const matchStr = `${item.queue_number || ''} ${item.requester_name || ''} ${item.request_time || ''} ${item.completed_at || ''} ${item.observacoes || ''}`.toLowerCase();
      return matchStr.includes(search);
    });
  }

  if (!historyItems.length) {
    strategyQueueHistory.innerHTML = '<div class="queue-item"><strong>Nenhum histórico disponível para os filtros aplicados.</strong></div>';
    historyBadge.textContent = '0';
    return;
  }

  historyBadge.textContent = historyItems.length;

  // Ordenar por data (mais recente primeiro)
  const sortedItems = historyItems.sort((a, b) => {
    const dateA = parseDate(a.completed_at || a.request_time);
    const dateB = parseDate(b.completed_at || b.request_time);
    return dateB - dateA; // Mais recente primeiro
  });

  // Calcular paginação
  const totalPages = Math.ceil(sortedItems.length / state.historyItemsPerPage);
  const startIdx = (state.historyCurrentPage - 1) * state.historyItemsPerPage;
  const endIdx = startIdx + state.historyItemsPerPage;
  const paginatedItems = sortedItems.slice(startIdx, endIdx);

  strategyQueueHistory.innerHTML = '';

  // Renderizar items da página atual
  paginatedItems.forEach((item) => {
    const row = document.createElement('article');
    const isCollapsed = !expandedCardIds.has(item.id);
    row.className = `queue-item ${isCollapsed ? 'collapsed' : ''}`;
    const rotation = isCollapsed ? '0deg' : '-180deg';
    const formattedFields = formatSelectedFields(item.selected_fields || []);

    const isEnrichOnly = item.queue_number === "ENRICH";

    const devolutivaBtnHtml = item.summary_name
      ? `<button type="button" class="primary" style="padding: 0.4rem 1rem; border-radius: var(--radius-sm); font-size: 0.8rem;" onclick="openImageViewer('/api/summary/${item.id}')">Visualizar Devolutiva</button>`
      : '';

    row.innerHTML = `
      <div class="queue-header" style="cursor: pointer; display: flex; justify-content: space-between; align-items: center;">
        <div style="display: flex; align-items: center; gap: 0.5rem;">
          <span class="collapse-icon" style="display: inline-block; transition: transform 0.2s; font-size: 0.8rem; color: var(--text-muted); transform: rotate(${rotation});">▼</span>
          <strong>Fila ${item.queue_number}</strong>
        </div>
        <span class="badge ${item.status === 'rejected' ? 'error' : 'completed'}">
          ${item.status === 'rejected' ? 'Recusado' : (isEnrichOnly ? 'Apenas Enriquecido' : 'Concluído')}
        </span>
      </div>
      <div class="queue-collapsible-content">
        <div class="queue-meta">Solicitante: ${item.requester_name}</div>
        <div class="queue-meta">Horário: ${item.completed_at || item.request_time}</div>
        <div class="queue-meta">Campos: ${formattedFields}</div>
        <div class="queue-actions">
          <span class="badge ${item.status === 'rejected' ? 'error' : 'completed'}">
            ${item.status === 'rejected' ? 'Recusado' : (isEnrichOnly ? 'Apenas Enriquecido' : 'Importado com sucesso')}
          </span>
          ${isEnrichOnly ? `
            <button type="button" class="download-link-premium" style="display: inline-flex; align-items: center; gap: 6px; border: none; font-family: inherit; font-size: 0.82rem; padding: 6px 12px;" onclick="downloadFileDirectly('/api/request/${item.id}/download?format=original')" title="Baixar Planilha no Layout Original (mesma aba e ordem original)">
              <svg viewBox="0 0 256 256" height="20" width="22" xmlns="http://www.w3.org/2000/svg">
                <path d="M74.34 85.66a8 8 0 0 1 11.32-11.32L120 108.69V24a8 8 0 0 1 16 0v84.69l34.34-34.35a8 8 0 0 1 11.32 11.32l-48 48a8 8 0 0 1-11.32 0ZM240 136v64a16 16 0 0 1-16 16H32a16 16 0 0 1-16-16v-64a16 16 0 0 1 16-16h52.4a4 4 0 0 1 2.83 1.17L111 145a24 24 0 0 0 34 0l23.8-23.8a4 4 0 0 1 2.8-1.2H224a16 16 0 0 1 16 16m-40 32a12 12 0 1 0-12 12a12 12 0 0 0 12-12" fill="currentColor"></path>
              </svg>
              <span>Original</span>
            </button>
            <button type="button" class="download-link-premium" style="display: inline-flex; align-items: center; gap: 6px; border: none; font-family: inherit; font-size: 0.82rem; padding: 6px 12px; opacity: 0.85;" onclick="downloadFileDirectly('/api/request/${item.id}/download?format=standard')" title="Baixar Planilha em 3 Abas (Enriquecido, Falha e Resumo)">
              <svg viewBox="0 0 256 256" height="20" width="22" xmlns="http://www.w3.org/2000/svg">
                <path d="M74.34 85.66a8 8 0 0 1 11.32-11.32L120 108.69V24a8 8 0 0 1 16 0v84.69l34.34-34.35a8 8 0 0 1 11.32 11.32l-48 48a8 8 0 0 1-11.32 0ZM240 136v64a16 16 0 0 1-16 16H32a16 16 0 0 1-16-16v-64a16 16 0 0 1 16-16h52.4a4 4 0 0 1 2.83 1.17L111 145a24 24 0 0 0 34 0l23.8-23.8a4 4 0 0 1 2.8-1.2H224a16 16 0 0 1 16 16m-40 32a12 12 0 1 0-12 12a12 12 0 0 0 12-12" fill="currentColor"></path>
              </svg>
              <span>3 Abas</span>
            </button>
          ` : `
            <button type="button" class="download-link-premium" style="display: inline-flex; align-items: center; gap: 6px; border: none; font-family: inherit; font-size: 0.82rem; padding: 6px 12px;" onclick="downloadFileDirectly('/api/request/${item.id}/download')" title="Baixar Planilha Tratada com Coluna 'importar'">
              <svg viewBox="0 0 256 256" height="20" width="22" xmlns="http://www.w3.org/2000/svg">
                <path d="M74.34 85.66a8 8 0 0 1 11.32-11.32L120 108.69V24a8 8 0 0 1 16 0v84.69l34.34-34.35a8 8 0 0 1 11.32 11.32l-48 48a8 8 0 0 1-11.32 0ZM240 136v64a16 16 0 0 1-16 16H32a16 16 0 0 1-16-16v-64a16 16 0 0 1 16-16h52.4a4 4 0 0 1 2.83 1.17L111 145a24 24 0 0 0 34 0l23.8-23.8a4 4 0 0 1 2.8-1.2H224a16 16 0 0 1 16 16m-40 32a12 12 0 1 0-12 12a12 12 0 0 0 12-12" fill="currentColor"></path>
              </svg>
              <span>Para Importar</span>
            </button>
            <button type="button" class="download-link-premium" style="display: inline-flex; align-items: center; gap: 6px; border: none; font-family: inherit; font-size: 0.82rem; padding: 6px 12px; opacity: 0.85;" onclick="downloadFileDirectly('/api/request/${item.id}/download?format=full')" title="Baixar Planilha Completa">
              <svg viewBox="0 0 256 256" height="20" width="22" xmlns="http://www.w3.org/2000/svg">
                <path d="M74.34 85.66a8 8 0 0 1 11.32-11.32L120 108.69V24a8 8 0 0 1 16 0v84.69l34.34-34.35a8 8 0 0 1 11.32 11.32l-48 48a8 8 0 0 1-11.32 0ZM240 136v64a16 16 0 0 1-16 16H32a16 16 0 0 1-16-16v-64a16 16 0 0 1 16-16h52.4a4 4 0 0 1 2.83 1.17L111 145a24 24 0 0 0 34 0l23.8-23.8a4 4 0 0 1 2.8-1.2H224a16 16 0 0 1 16 16m-40 32a12 12 0 1 0-12 12a12 12 0 0 0 12-12" fill="currentColor"></path>
              </svg>
              <span>Completo</span>
            </button>
          `}
          ${item.status === 'rejected' ? `<button type="button" class="button secondary" style="background-color: var(--danger-color); color: white; border-color: var(--danger-color);" onclick="openViewReasonModal('${encodeURIComponent(getItemRejectReason(item))}')">Motivo da Recusa</button>` : ''}
          ${devolutivaBtnHtml}
          <button type="button" class="button secondary" onclick="viewDetailedLogs('${item.id}')">Logs</button>
        </div>
      </div>
    `;

    // Add collapse toggle handler
    const header = row.querySelector('.queue-header');
    header.addEventListener('click', (e) => {
      row.classList.toggle('collapsed');
      const icon = header.querySelector('.collapse-icon');
      if (row.classList.contains('collapsed')) {
        icon.style.transform = 'rotate(0deg)';
        expandedCardIds.delete(item.id);
      } else {
        icon.style.transform = 'rotate(-180deg)';
        expandedCardIds.add(item.id);
      }
    });

    strategyQueueHistory.appendChild(row);
  });

  // Adicionar controles de paginação se houver mais de uma página
  if (totalPages > 1) {
    const paginationContainer = document.createElement('div');
    paginationContainer.className = 'pagination-container';
    paginationContainer.innerHTML = `
      <div class="pagination-controls">
        <button type="button" class="pagination-btn" onclick="goToHistoryPage(1)" ${state.historyCurrentPage === 1 ? 'disabled' : ''}>«</button>
        <button type="button" class="pagination-btn" onclick="goToHistoryPage(${state.historyCurrentPage - 1})" ${state.historyCurrentPage === 1 ? 'disabled' : ''}>‹</button>
        <span class="pagination-info">Página ${state.historyCurrentPage} de ${totalPages}</span>
        <button type="button" class="pagination-btn" onclick="goToHistoryPage(${state.historyCurrentPage + 1})" ${state.historyCurrentPage === totalPages ? 'disabled' : ''}>›</button>
        <button type="button" class="pagination-btn" onclick="goToHistoryPage(${totalPages})" ${state.historyCurrentPage === totalPages ? 'disabled' : ''}>»</button>
      </div>
    `;
    strategyQueueHistory.appendChild(paginationContainer);
  }
}

function parseDate(dateStr) {
  // Esperado formato: "DD/MM/YYYY HH:MM:SS"
  if (!dateStr) return new Date(0);
  try {
    const [datePart, timePart] = dateStr.split(' ');
    const [day, month, year] = datePart.split('/');
    const [hour, minute, second] = timePart ? timePart.split(':') : ['0', '0', '0'];
    return new Date(year, month - 1, day, hour, minute, second);
  } catch (e) {
    return new Date(0);
  }
}

function goToHistoryPage(page) {
  const historyItems = state.queue.filter(item => item.status === 'completed' || item.status === 'rejected');
  const totalPages = Math.ceil(historyItems.length / state.historyItemsPerPage);

  if (page < 1) page = 1;
  if (page > totalPages) page = totalPages;

  state.historyCurrentPage = page;
  renderQueueHistory();
}

function renderQueue() {
  // Manter para compatibilidade, mas agora chama as duas
  renderQueueActive();
  renderQueueHistory();
}

async function submitSummary(event) {
  const button = event.target.closest('[data-submit-summary]');
  if (!button) return;

  const requestId = button.dataset.submitSummary;
  const fileInput = document.querySelector(`.summary-upload[data-id='${requestId}']`);

  if (!fileInput || !fileInput.files || !fileInput.files[0]) {
    showMessage('Anexe uma imagem do resumo antes de concluir.', 'error');
    return;
  }

  const formData = new FormData();
  formData.append('summaryImage', fileInput.files[0]);

  try {
    const response = await fetch(`/api/request/${requestId}/complete`, {
      method: 'POST',
      body: formData,
    });

    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || 'Não foi possível concluir a importação.');
    }

    showMessage(payload.message, 'success');
    await loadNotifications();
    await loadQueue();
  } catch (error) {
    showMessage(error.message, 'error');
  }
}

async function loadNotifications() {
  try {
    const response = await fetch('/api/notifications');
    if (!response.ok) {
      throw new Error('Não foi possível carregar as notificações.');
    }

    const payload = await response.json();
    state.notifications = payload.notifications || [];

    const renderBlock = (target, emptyText) => {
      target.innerHTML = '';
      if (!state.notifications.length) {
        target.innerHTML = `<div class="notification-item"><p>${emptyText}</p></div>`;
        return;
      }

      state.notifications.forEach((notification) => {
        const item = document.createElement('div');
        item.className = 'notification-item';
        item.innerHTML = `<p><strong>${notification.title}</strong> — ${notification.message}</p><div class="queue-meta">${notification.time}</div>`;
        target.appendChild(item);
      });
    };

    renderBlock(strategyNotifications, 'Nenhuma notificação para estratégia no momento.');
    renderBlock(userNotifications, '');
  } catch (error) {
    console.error(error);
  }
}

tabButtons.forEach((button) => {
  button.addEventListener('click', () => setActiveTab(button.dataset.target));
});

if (roleSelector) {
  roleSelector.addEventListener('change', (event) => {
    state.currentRole = event.target.value;
    applyRoleAccess();
  });
}

// Setup user tabs
const userTabButtons = document.querySelectorAll('[data-user-tab]');
const userImportSection = document.getElementById('user-import-section');
const normalUserHistoryContainer = document.getElementById('normalUserHistoryContainer');
const userConverterSection = document.getElementById('user-converter-section');

if (userTabButtons.length > 0) {
  userTabButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      userTabButtons.forEach(b => b.classList.remove('is-active'));
      btn.classList.add('is-active');
      const tabName = btn.getAttribute('data-user-tab');

      // Limpa qualquer arquivo importado/anexado para não constar mais ao trocar de aba
      clearImportedFileState();

      if (tabName === 'import') {
        if (userImportSection) userImportSection.classList.remove('hidden');
        if (normalUserHistoryContainer) normalUserHistoryContainer.classList.add('hidden');
        if (userConverterSection) userConverterSection.classList.add('hidden');
      } else if (tabName === 'history') {
        if (userImportSection) userImportSection.classList.add('hidden');
        if (normalUserHistoryContainer) normalUserHistoryContainer.classList.remove('hidden');
        if (userConverterSection) userConverterSection.classList.add('hidden');
        renderNormalUserHistory(); // refresh the view
      } else if (tabName === 'converter') {
        if (userImportSection) userImportSection.classList.add('hidden');
        if (normalUserHistoryContainer) normalUserHistoryContainer.classList.add('hidden');
        if (userConverterSection) userConverterSection.classList.remove('hidden');
        if (typeof initConverterModule === 'function') {
          initConverterModule();
        }
      }
    });
  });
}

document.querySelectorAll('.mode-btn').forEach((button) => {
  button.addEventListener('click', () => setActionMode(button.dataset.actionType));
});

setActionMode(state.currentAction);

async function handleFinalizeButton() {
  if (state.isBatchIndividual) {
    openBatchUploadModal();
    return;
  }

  if (state.currentAction === 'enrich_only') {
    openRequesterModal();
    return;
  }

  openRequestModal();
}

function openRequesterModal() {
  const modal = document.getElementById('requesterModal');
  if (modal) {
    setExportFormat(state.exportFormat || 'original');
    modal.classList.remove('hidden');
    const input = document.getElementById('requesterInputName');
    if (input) {
      if (state.requesterName) input.value = state.requesterName;
      input.focus();
    }
  }
}

function closeRequesterModal() {
  const modal = document.getElementById('requesterModal');
  if (modal) {
    modal.classList.add('hidden');
    document.getElementById('requesterForm').reset();
  }
}

const requesterForm = document.getElementById('requesterForm');
if (requesterForm) {
  requesterForm.addEventListener('submit', (e) => {
    e.preventDefault();
    const btn = requesterForm.querySelector('button[type="submit"]');
    if (btn && btn.disabled) return;
    const requesterName = document.getElementById('requesterInputName').value.trim();
    if (!requesterName) return;
    const modalRadio = document.querySelector('input[name="modalExportFormat"]:checked');
    if (modalRadio) {
      setExportFormat(modalRadio.value);
    }
    if (btn) btn.disabled = true;
    state.requesterName = requesterName;
    closeRequesterModal();
    if (btn) setTimeout(() => { btn.disabled = false; }, 1000);
    startEnrichment(requesterName);
  });
}

// Expose modal handlers globally for inline onclick
window.closeRequesterModal = closeRequesterModal;

let activeRefuseRequestId = null;

function openRefuseModal(requestId) {
  activeRefuseRequestId = requestId;
  const modal = document.getElementById('refuseModal');
  if (modal) {
    modal.classList.remove('hidden');
    document.getElementById('refuseReasonInput').focus();
  }
}

function closeRefuseModal() {
  activeRefuseRequestId = null;
  const modal = document.getElementById('refuseModal');
  if (modal) {
    modal.classList.add('hidden');
    document.getElementById('refuseForm').reset();
  }
}

const refuseForm = document.getElementById('refuseForm');
if (refuseForm) {
  refuseForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    if (!activeRefuseRequestId) return;

    const reason = document.getElementById('refuseReasonInput').value.trim();
    if (!reason) return;

    try {
      const response = await fetch(`/api/request/${activeRefuseRequestId}/refuse`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ reason })
      });
      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || 'Erro ao recusar solicitação.');
      }
      showMessage(payload.message, 'success');
      closeRefuseModal();
      loadQueue();
    } catch (error) {
      showMessage(error.message, 'error');
    }
  });
}

window.openRefuseModal = openRefuseModal;
window.closeRefuseModal = closeRefuseModal;

function getItemRejectReason(item) {
  if (item && item.reject_reason && String(item.reject_reason).trim() !== '') {
    return String(item.reject_reason).trim();
  }
  if (item && Array.isArray(item.logs)) {
    for (const log of item.logs) {
      if (typeof log === 'string' && log.includes('Solicitação recusada pela Estratégia. Motivo: ')) {
        return log.split('Solicitação recusada pela Estratégia. Motivo: ')[1].trim();
      }
    }
  }
  return 'Sem motivo informado';
}

function openViewReasonModal(encodedReason) {
  const reason = decodeURIComponent(encodedReason);
  const modal = document.getElementById('viewReasonModal');
  const textEl = document.getElementById('viewReasonText');
  if (modal && textEl) {
    textEl.textContent = reason;
    modal.classList.remove('hidden');
  }
}

function closeViewReasonModal() {
  const modal = document.getElementById('viewReasonModal');
  if (modal) {
    modal.classList.add('hidden');
  }
}

window.openViewReasonModal = openViewReasonModal;
window.closeViewReasonModal = closeViewReasonModal;

function downloadDevolutiva(url) {
  openDownloadRenameModal({
    downloadUrl: url,
    defaultFilename: `devolutiva_${Date.now()}.png`,
    fileType: 'Imagem de Devolutiva',
  });
}

uploadForm.addEventListener('submit', handleUpload);
finalizeButton.addEventListener('click', handleFinalizeButton);
enrichButton.addEventListener('click', () => {
  const actionMode = state.currentAction || 'enrich_only';
  const queryFields = Array.from(state.selectedQueryFields || []);

  if (actionMode !== 'queue_only' && queryFields.length === 0) {
    showMessage('Selecione ao menos um atributo para realizar a busca no enriquecimento.', 'error');
    return;
  }

  // Avança para a tela de escolha dos atributos enriquecidos (step3Container)
  step1Container.classList.add('hidden');
  step3Container.classList.remove('hidden');
  renderAvailableFields();
  renderSelectedFields();
  renderPreview();
  updateFinalizeButtonLabel();
});
requestForm.addEventListener('submit', submitRequest);
document.querySelectorAll('.close-modal').forEach((button) => button.addEventListener('click', closeRequestModal));
document.querySelectorAll('.close-logs-modal').forEach((button) => button.addEventListener('click', closeLogsModal));

// Queue tabs event listeners
document.querySelectorAll('.queue-tab').forEach((tab) => {
  tab.addEventListener('click', (e) => {
    e.preventDefault();
    const tabName = tab.dataset.queueTab;

    // Update active tab
    document.querySelectorAll('.queue-tab').forEach((t) => t.classList.remove('is-active'));
    tab.classList.add('is-active');

    // Show/hide content
    strategyQueueActive.classList.toggle('hidden', tabName !== 'active');
    strategyQueueActive.classList.toggle('is-visible', tabName === 'active');
    strategyQueueHistory.classList.toggle('hidden', tabName !== 'history');
    strategyQueueHistory.classList.toggle('is-visible', tabName === 'history');
  });
});

// Add click listeners to both queue containers
[strategyQueueActive, strategyQueueHistory].forEach((container) => {
  container.addEventListener('click', submitSummary);
});

// Logs Queue functions and polling
async function loadLogsQueue() {
  try {
    const response = await fetch('/api/queue');
    if (!response.ok) {
      throw new Error('Não foi possível carregar o histórico.');
    }
    const payload = await response.json();
    state.queue = payload.queue || [];
    renderLogsQueue();
  } catch (error) {
    console.error(error);
  }
}

function renderLogsQueue() {
  if (!logsList) return;
  if (!state.queue.length) {
    logsList.innerHTML = '<div class="queue-item"><strong>Nenhum registro de importação encontrado.</strong></div>';
    return;
  }

  logsList.innerHTML = '';
  state.queue.forEach((item) => {
    const row = document.createElement('article');
    row.className = 'queue-item';

    let statusText = 'Pendente';
    let badgeClass = 'pending';
    if (item.status === 'processing') {
      const pct = item.total_rows > 0 ? Math.round((item.processed_count / item.total_rows) * 100) : 0;
      statusText = `Processando (${pct}%)`;
      badgeClass = 'processing';
    } else if (item.status === 'queued') {
      statusText = 'Na Fila (Aguardando)';
      badgeClass = 'queued';
    } else if (item.status === 'completed') {
      statusText = 'Concluído';
      badgeClass = 'completed';
    } else if (item.status === 'error') {
      statusText = 'Erro';
      badgeClass = 'error';
    }

    const isEnrichOnly = (item.mode === 'enrich_only') || (item.queue_number === 'ENRICH');
    let downloadHtml = '';
    if (item.status === 'completed' || item.status === 'pending') {
      if (isEnrichOnly) {
        downloadHtml = `
          <button type="button" class="download-link-premium" style="display: inline-flex; align-items: center; gap: 6px; border: none; font-family: inherit; font-size: 0.82rem; padding: 6px 12px;" onclick="downloadFileDirectly('/api/request/${item.id}/download?format=original')" title="Baixar Planilha no Layout Original (mesma aba e ordem original)">
            <svg viewBox="0 0 256 256" height="20" width="22" xmlns="http://www.w3.org/2000/svg">
              <path d="M74.34 85.66a8 8 0 0 1 11.32-11.32L120 108.69V24a8 8 0 0 1 16 0v84.69l34.34-34.35a8 8 0 0 1 11.32 11.32l-48 48a8 8 0 0 1-11.32 0ZM240 136v64a16 16 0 0 1-16 16H32a16 16 0 0 1-16-16v-64a16 16 0 0 1 16-16h52.4a4 4 0 0 1 2.83 1.17L111 145a24 24 0 0 0 34 0l23.8-23.8a4 4 0 0 1 2.8-1.2H224a16 16 0 0 1 16 16m-40 32a12 12 0 1 0-12 12a12 12 0 0 0 12-12" fill="currentColor"></path>
            </svg>
            <span>Original</span>
          </button>
          <button type="button" class="download-link-premium" style="display: inline-flex; align-items: center; gap: 6px; border: none; font-family: inherit; font-size: 0.82rem; padding: 6px 12px; opacity: 0.85;" onclick="downloadFileDirectly('/api/request/${item.id}/download?format=standard')" title="Baixar Planilha em 3 Abas (Enriquecido, Falha e Resumo)">
            <svg viewBox="0 0 256 256" height="20" width="22" xmlns="http://www.w3.org/2000/svg">
              <path d="M74.34 85.66a8 8 0 0 1 11.32-11.32L120 108.69V24a8 8 0 0 1 16 0v84.69l34.34-34.35a8 8 0 0 1 11.32 11.32l-48 48a8 8 0 0 1-11.32 0ZM240 136v64a16 16 0 0 1-16 16H32a16 16 0 0 1-16-16v-64a16 16 0 0 1 16-16h52.4a4 4 0 0 1 2.83 1.17L111 145a24 24 0 0 0 34 0l23.8-23.8a4 4 0 0 1 2.8-1.2H224a16 16 0 0 1 16 16m-40 32a12 12 0 1 0-12 12a12 12 0 0 0 12-12" fill="currentColor"></path>
            </svg>
            <span>3 Abas</span>
          </button>
        `;
      } else {
        downloadHtml = `
          <button type="button" class="download-link-premium" style="display: inline-flex; align-items: center; gap: 8px; border: none; font-family: inherit;" onclick="downloadFileDirectly('/api/request/${item.id}/download')" title="Baixar Planilha">
            <svg viewBox="0 0 256 256" height="24" width="28" xmlns="http://www.w3.org/2000/svg">
              <path d="M74.34 85.66a8 8 0 0 1 11.32-11.32L120 108.69V24a8 8 0 0 1 16 0v84.69l34.34-34.35a8 8 0 0 1 11.32 11.32l-48 48a8 8 0 0 1-11.32 0ZM240 136v64a16 16 0 0 1-16 16H32a16 16 0 0 1-16-16v-64a16 16 0 0 1 16-16h52.4a4 4 0 0 1 2.83 1.17L111 145a24 24 0 0 0 34 0l23.8-23.8a4 4 0 0 1 2.8-1.2H224a16 16 0 0 1 16 16m-40 32a12 12 0 1 0-12 12a12 12 0 0 0 12-12" fill="currentColor"></path>
            </svg>
            <span>Download</span>
          </button>
        `;
      }
    }

    row.innerHTML = `
      <div class="queue-header">
        <strong>Fila ${item.queue_number}</strong>
        <span class="badge ${badgeClass}">${statusText}</span>
      </div>
      <div class="queue-meta">Solicitante: ${item.requester_name}</div>
      <div class="queue-meta">Horário: ${item.request_time}</div>
      <div class="queue-meta">Progresso: ${item.processed_count || 0} / ${item.total_rows || 0} (Sucessos: ${item.success_count || 0}, Falhas: ${item.error_count || 0})</div>
      <div class="queue-actions">
        ${downloadHtml}
        <button type="button" class="button secondary" onclick="viewDetailedLogs('${item.id}')">Ver Logs</button>
      </div>
    `;
    logsList.appendChild(row);
  });
}

let logPollInterval = null;
let activeLogRequestId = null;

async function viewDetailedLogs(requestId) {
  activeLogRequestId = requestId;
  logsModal.classList.remove('hidden');

  await fetchAndRenderLogs(requestId);

  if (logPollInterval) {
    clearInterval(logPollInterval);
  }

  logPollInterval = setInterval(async () => {
    const isStillProcessing = await fetchAndRenderLogs(activeLogRequestId);
    if (!isStillProcessing) {
      clearInterval(logPollInterval);
      logPollInterval = null;
      loadLogsQueue();
      loadQueue();
    }
  }, 2000);
}

// Helper function to dynamically update skill stage boxes, bars, tooltips and badges
function updateSkillStage(boxEl, barEl, tooltipEl, badgeEl, stateType, pct, badgeText) {
  if (!boxEl || !barEl || !tooltipEl || !badgeEl) return;

  const validPct = Math.min(100, Math.max(0, Math.round(pct || 0)));
  barEl.style.width = `${validPct}%`;
  tooltipEl.textContent = `${validPct}%`;

  boxEl.classList.remove('active', 'completed');
  if (stateType === 'completed') {
    boxEl.classList.add('completed');
  } else if (stateType === 'active') {
    boxEl.classList.add('active');
  }

  if (badgeText !== undefined && badgeText !== null) {
    badgeEl.textContent = badgeText;
  }
}

// Accordion toggle event listeners
if (btnToggleTechLogs && uploadLogsBoxContainer) {
  btnToggleTechLogs.addEventListener('click', () => {
    const isHidden = uploadLogsBoxContainer.classList.toggle('hidden');
    btnToggleTechLogs.classList.toggle('is-open', !isHidden);
    if (toggleLogsText) {
      toggleLogsText.textContent = isHidden ? 'Mostrar detalhes técnicos do processamento' : 'Ocultar detalhes técnicos';
    }
  });
}

if (btnToggleModalTechLogs && modalLogsBoxContainer) {
  btnToggleModalTechLogs.addEventListener('click', () => {
    const isHidden = modalLogsBoxContainer.classList.toggle('hidden');
    btnToggleModalTechLogs.classList.toggle('is-open', !isHidden);
    if (toggleModalLogsText) {
      toggleModalLogsText.textContent = isHidden ? 'Mostrar logs detalhados' : 'Ocultar logs detalhados';
    }
  });
}

async function fetchAndRenderLogs(requestId) {
  try {
    const response = await fetch(`/api/request/${requestId}/logs`);
    if (!response.ok) {
      throw new Error('Falha ao buscar logs.');
    }
    const data = await response.json();

    let modalStatusHeader = 'Finalizado';
    if (data.status === 'processing') {
      modalStatusHeader = 'Processando';
    } else if (data.status === 'queued') {
      modalStatusHeader = `Na Fila (#${data.queue_position || 1})`;
    }
    logsModalTitle.textContent = `Progresso da Fila (${modalStatusHeader})`;
    logCounterTotal.textContent = data.total_rows || 0;
    logCounterSuccess.textContent = data.success_count || 0;
    logCounterError.textContent = data.error_count || 0;

    const pct = data.total_rows > 0 ? Math.round((data.processed_count / data.total_rows) * 100) : (data.status === 'completed' ? 100 : 0);

    // Modal Stage 1: Validando
    updateSkillStage(modalStageBoxValidate, modalStageBarValidate, modalStageTooltipValidate, modalStageBadgeValidate, 'completed', 100, 'Validado');

    // Modal Stage 2: Enriquecendo
    if (data.status === 'processing') {
      updateSkillStage(modalStageBoxEnrich, modalStageBarEnrich, modalStageTooltipEnrich, modalStageBadgeEnrich, 'active', pct, `${data.processed_count} / ${data.total_rows} (${pct}%)`);
      updateSkillStage(modalStageBoxClean, modalStageBarClean, modalStageTooltipClean, modalStageBadgeClean, '', 0, 'Aguardando');
      updateSkillStage(modalStageBoxConcat, modalStageBarConcat, modalStageTooltipConcat, modalStageBadgeConcat, '', 0, 'Aguardando');
    } else if (data.status === 'queued') {
      updateSkillStage(modalStageBoxEnrich, modalStageBarEnrich, modalStageTooltipEnrich, modalStageBadgeEnrich, 'active', 0, `Na Fila (Posição #${data.queue_position || 1})`);
      updateSkillStage(modalStageBoxClean, modalStageBarClean, modalStageTooltipClean, modalStageBadgeClean, '', 0, 'Aguardando');
      updateSkillStage(modalStageBoxConcat, modalStageBarConcat, modalStageTooltipConcat, modalStageBadgeConcat, '', 0, 'Aguardando');
    } else if (data.status === 'completed') {
      updateSkillStage(modalStageBoxEnrich, modalStageBarEnrich, modalStageTooltipEnrich, modalStageBadgeEnrich, 'completed', 100, 'Concluído');
      updateSkillStage(modalStageBoxClean, modalStageBarClean, modalStageTooltipClean, modalStageBadgeClean, 'completed', 100, 'Concluído');
      updateSkillStage(modalStageBoxConcat, modalStageBarConcat, modalStageTooltipConcat, modalStageBadgeConcat, 'completed', 100, 'Pronto!');
    } else {
      updateSkillStage(modalStageBoxEnrich, modalStageBarEnrich, modalStageTooltipEnrich, modalStageBadgeEnrich, '', pct, data.status === 'error' ? 'Falha' : 'Cancelado');
    }

    if (data.logs && data.logs.length > 0) {
      logsModalContent.textContent = data.logs.join('\n');
      logsModalContent.scrollTop = logsModalContent.scrollHeight;
    } else {
      logsModalContent.textContent = data.status === 'queued' ? 'Aguardando sua vez na fila para iniciar consultas...' : 'Nenhum log registrado ainda.';
    }

    return data.status === 'processing' || data.status === 'queued';
  } catch (err) {
    logsModalContent.textContent = `Erro ao carregar logs: ${err.message}`;
    return false;
  }
}

function closeLogsModal() {
  logsModal.classList.add('hidden');
  activeLogRequestId = null;
  if (logPollInterval) {
    clearInterval(logPollInterval);
    logPollInterval = null;
  }
}

// Step 1: Render Query Fields selection box with assertiveness colors
function renderQueryFieldsSelection(columns) {
  const selectionDiv = document.getElementById('queryFieldsSelection');
  if (!selectionDiv) return;

  selectionDiv.innerHTML = '';

  const detectedType = state.currentRecord?.type || 'nome_telefone';
  const selectedQueryFields = new Set();

  (columns || []).forEach((col) => {
    const tag = document.createElement('button');
    tag.type = 'button';
    tag.className = 'query-tag';
    tag.textContent = col;

    const norm = normalizeFieldName(col);

    let isDefaultSelected = false;

    if (detectedType === 'cpf' && (norm === 'cpf' || norm === 'documento' || norm === 'cnpj')) {
      isDefaultSelected = true;
    } else if (norm === 'cpf' || norm === 'documento' || norm === 'cnpj') {
      isDefaultSelected = true;
    } else if (isRequiredNameField(col)) {
      isDefaultSelected = true;
    } else if (isRequiredPhoneField(col)) {
      isDefaultSelected = true;
    } else if (norm === 'uf' || norm === 'estado' || norm === 'sigla_uf' || norm === 'uf_estado') {
      isDefaultSelected = true;
    } else if (norm === 'cidade' || norm === 'municipio' || norm === 'nome_cidade') {
      isDefaultSelected = true;
    }

    if (isDefaultSelected) {
      selectedQueryFields.add(col);
      tag.classList.add('selected');
    }

    tag.addEventListener('click', () => {
      if (selectedQueryFields.has(col)) {
        selectedQueryFields.delete(col);
        tag.classList.remove('selected');
      } else {
        selectedQueryFields.add(col);
        tag.classList.add('selected');
      }
      state.selectedQueryFields = selectedQueryFields;
    });

    selectionDiv.appendChild(tag);
  });

  state.selectedQueryFields = selectedQueryFields;
}

// Step 2: Start background processing
async function startEnrichment(requesterName, queueNumber = '', observacoes = '') {
  if (!state.currentRecord) return;

  if (requesterName) state.requesterName = requesterName;
  if (queueNumber) state.queueNumber = queueNumber;
  if (observacoes) state.observacoes = observacoes;

  const actionMode = state.currentAction || 'enrich_only';
  const queryFields = Array.from(state.selectedQueryFields || []);
  const deduplicate = document.getElementById('chkDeduplicateRows')?.checked ?? true;

  // Transição visual imediata para a tela de progresso (step2Container)
  if (step1Container) step1Container.classList.add('hidden');
  if (step3Container) step3Container.classList.add('hidden');
  if (step2Container) step2Container.classList.remove('hidden');

  updateSkillStage(stageBoxValidate, stageBarValidate, stageTooltipValidate, stageBadgeValidate, 'active', 50, 'Iniciando...');
  updateSkillStage(stageBoxEnrich, stageBarEnrich, stageTooltipEnrich, stageBadgeEnrich, '', 0, 'Aguardando');
  updateSkillStage(stageBoxClean, stageBarClean, stageTooltipClean, stageBadgeClean, '', 0, 'Aguardando');
  updateSkillStage(stageBoxConcat, stageBarConcat, stageTooltipConcat, stageBadgeConcat, '', 0, 'Aguardando');

  try {
    const chosenFmt = state.exportFormat || 'original';
    const response = await fetch('/api/request-import', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        record_id: state.currentRecord.id,
        query_fields: queryFields,
        mode: actionMode,
        requester_name: requesterName,
        deduplicate: deduplicate,
        export_format: chosenFmt
      })
    });

    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || 'Falha ao iniciar enriquecimento.');
    }

    state.activeRequestImportId = payload.request_id;
    pollUploadProgress(payload.request_id);
  } catch (error) {
    if (step2Container) step2Container.classList.add('hidden');
    if (step3Container) step3Container.classList.remove('hidden');
    showMessage(error.message, 'error');
  }
}


function ensureStepVisibility() {
  if (uploadForm) {
    uploadForm.classList.remove('hidden');
  }
  if (step1Container) {
    step1Container.classList.add('hidden');
  }
  if (step2Container) {
    step2Container.classList.add('hidden');
  }
  if (step3Container) {
    step3Container.classList.add('hidden');
  }
  if (validationWarningBox) {
    validationWarningBox.classList.add('hidden');
  }
}

function formatTimeRemaining(seconds) {
  if (isNaN(seconds) || seconds <= 0) return '0s';
  if (seconds < 60) {
    return `${seconds}s`;
  }
  const minutes = Math.floor(seconds / 60);
  const remSec = seconds % 60;
  if (minutes < 60) {
    return remSec > 0 ? `${minutes}min ${remSec}s` : `${minutes}min`;
  }
  const hours = Math.floor(minutes / 60);
  const remMin = minutes % 60;
  return remMin > 0 ? `${hours}h ${remMin}min` : `${hours}h`;
}

function pollUploadProgress(requestId) {
  if (uploadProgressInterval) {
    clearInterval(uploadProgressInterval);
    uploadProgressInterval = null;
  }

  // Initial State: Todas as barras iniciam zeradas e vão enchendo conforme o fluxo avança
  updateSkillStage(stageBoxValidate, stageBarValidate, stageTooltipValidate, stageBadgeValidate, 'active', 0, 'Validando...');
  updateSkillStage(stageBoxEnrich, stageBarEnrich, stageTooltipEnrich, stageBadgeEnrich, '', 0, 'Aguardando');
  updateSkillStage(stageBoxClean, stageBarClean, stageTooltipClean, stageBadgeClean, '', 0, 'Aguardando');
  updateSkillStage(stageBoxConcat, stageBarConcat, stageTooltipConcat, stageBadgeConcat, '', 0, 'Aguardando');

  if (uploadLogsBox) uploadLogsBox.textContent = '';

  // Animação progressiva da validação do arquivo (0% -> 50% -> 100%)
  setTimeout(() => {
    updateSkillStage(stageBoxValidate, stageBarValidate, stageTooltipValidate, stageBadgeValidate, 'active', 60, 'Validando estrutura...');
    setTimeout(() => {
      updateSkillStage(stageBoxValidate, stageBarValidate, stageTooltipValidate, stageBadgeValidate, 'completed', 100, 'Validado');
      const isQueueOnly = state.currentAction === 'queue_only';
      const stage2TitleEl = stageBoxEnrich ? stageBoxEnrich.querySelector('.title') : null;
      if (stage2TitleEl) {
        stage2TitleEl.textContent = isQueueOnly ? 'Processando contatos da fila' : 'Enriquecendo o arquivo';
      }
      const stage2ActiveLabel = isQueueOnly ? 'Iniciando sub-lotes...' : 'Iniciando enriquecimento...';
      updateSkillStage(stageBoxEnrich, stageBarEnrich, stageTooltipEnrich, stageBadgeEnrich, 'active', 0, stage2ActiveLabel);
    }, 350);
  }, 200);

  // Reset stat fields at start
  if (statProcessed) statProcessed.textContent = '0 / 0';
  if (statSuccess) statSuccess.textContent = '0 (0%)';
  if (statError) statError.textContent = '0 (0%)';
  if (statTimeRemaining) statTimeRemaining.textContent = 'Calculando...';

  const startTime = Date.now();

  // Smooth lerp state for 60fps rendering
  let animState = {
    processed: 0,
    targetProcessed: 0,
    total: 0,
    success: 0,
    targetSuccess: 0,
    error: 0,
    targetError: 0,
    pct: 0,
    targetPct: 0,
    running: true
  };

  function animLoop() {
    if (!animState.running) return;

    const lerpFactor = 0.45;
    const diffProc = animState.targetProcessed - animState.processed;

    const diffSucc = animState.targetSuccess - animState.success;
    const diffErr = animState.targetError - animState.error;
    const diffPct = animState.targetPct - animState.pct;

    if (Math.abs(diffProc) > 0.05) animState.processed += diffProc * lerpFactor;
    else animState.processed = animState.targetProcessed;

    if (Math.abs(diffSucc) > 0.05) animState.success += diffSucc * lerpFactor;
    else animState.success = animState.targetSuccess;

    if (Math.abs(diffErr) > 0.05) animState.error += diffErr * lerpFactor;
    else animState.error = animState.targetError;

    if (Math.abs(diffPct) > 0.05) animState.pct += diffPct * lerpFactor;
    else animState.pct = animState.targetPct;

    const curProc = Math.round(animState.processed);
    const curSucc = Math.round(animState.success);
    const curErr = Math.round(animState.error);
    const curPct = Math.min(100, Math.max(0, Math.round(animState.pct)));

    if (stageBarEnrich) stageBarEnrich.style.width = `${curPct}%`;
    if (stageTooltipEnrich) stageTooltipEnrich.textContent = `${curPct}%`;
    if (stageBadgeEnrich && animState.targetPct < 100) {
      stageBadgeEnrich.textContent = `${curProc} de ${animState.total} (${curPct}%)`;
    }

    if (statProcessed) statProcessed.textContent = `${curProc} / ${animState.total}`;
    if (statSuccess) {
      const pctSuccess = animState.total > 0 ? Math.round((curSucc / animState.total) * 100) : 0;
      statSuccess.textContent = `${curSucc} (${pctSuccess}%)`;
    }
    if (statError) {
      const pctError = animState.total > 0 ? Math.round((curErr / animState.total) * 100) : 0;
      statError.textContent = `${curErr} (${pctError}%)`;
    }

    requestAnimationFrame(animLoop);
  }

  requestAnimationFrame(animLoop);

  let isPolling = false;
  let hasFinalized = false;

  uploadProgressInterval = setInterval(async () => {
    if (isPolling || hasFinalized) return;
    isPolling = true;
    try {
      const response = await fetch(`/api/request/${requestId}/logs`);
      if (!response.ok) {
        throw new Error('Falha ao obter progresso.');
      }
      const data = await response.json();

      animState.total = data.total_rows || 0;
      animState.targetProcessed = data.processed_count || 0;
      animState.targetSuccess = data.success_count || 0;
      animState.targetError = data.error_count || 0;

      const pct = data.total_rows > 0 ? Math.round((data.processed_count / data.total_rows) * 100) : 0;
      animState.targetPct = pct;

      if (data.status === 'processing') {
        stageBoxEnrich.classList.add('active');
        stageBoxEnrich.classList.remove('completed');
        if (data.chunk_info && data.chunk_info.is_chunked) {
          if (stageBadgeEnrich) {
            stageBadgeEnrich.textContent = `Lote ${data.chunk_info.current_chunk}/${data.chunk_info.total_chunks} (${data.processed_count} de ${data.total_rows} - ${pct}%)`;
          }
          if (statProcessed) {
            statProcessed.textContent = `${data.processed_count} / ${data.total_rows} (Lote ${data.chunk_info.current_chunk}/${data.chunk_info.total_chunks})`;
          }
        }
      } else if (data.status === 'queued') {
        stageBoxEnrich.classList.add('active');
        stageBoxEnrich.classList.remove('completed');
        if (stageBadgeEnrich) {
          stageBadgeEnrich.textContent = `Na Fila (Posição #${data.queue_position || 1})`;
        }
        if (stageTooltipEnrich) {
          stageTooltipEnrich.textContent = `Posição #${data.queue_position || 1}`;
        }
      }

      if (statTimeRemaining) {
        const elapsed = (Date.now() - startTime) / 1000;
        if (data.processed_count > 0 && data.status === 'processing') {
          const rate = data.processed_count / elapsed;
          const remaining = data.total_rows - data.processed_count;
          const estSec = Math.ceil(remaining / rate);
          const rateFormatted = Math.round(rate);
          statTimeRemaining.textContent = `${formatTimeRemaining(estSec)} (${rateFormatted}/s)`;
        } else if (data.status === 'queued') {
          statTimeRemaining.textContent = `Aguardando vez (Posição #${data.queue_position || 1})`;
        } else if (data.status !== 'processing') {
          statTimeRemaining.textContent = '0s';
        } else {
          statTimeRemaining.textContent = 'Calculando...';
        }
      }

      if (data.logs && data.logs.length > 0 && uploadLogsBoxContainer && !uploadLogsBoxContainer.classList.contains('hidden')) {
        uploadLogsBox.textContent = data.logs.slice(-30).join('\n');
        uploadLogsBox.scrollTop = uploadLogsBox.scrollHeight;
      }

      if (data.status !== 'processing' && data.status !== 'queued') {
        if (hasFinalized) return;
        hasFinalized = true;
        clearInterval(uploadProgressInterval);
        uploadProgressInterval = null;
        animState.running = false;

        // Force exact end values
        if (stageBarEnrich) stageBarEnrich.style.width = `${pct}%`;
        if (stageTooltipEnrich) stageTooltipEnrich.textContent = `${pct}%`;

        if (data.status === 'error') {
          updateSkillStage(stageBoxEnrich, stageBarEnrich, stageTooltipEnrich, stageBadgeEnrich, '', pct, 'Falha');
          showMessage('Erro ao processar lote no plano de fundo. Veja os logs abaixo.', 'error');
          sendDesktopNotification('Falha no Enriquecimento', 'Ocorreu um erro ao processar sua planilha.');
        } else if (data.status === 'cancelled') {
          updateSkillStage(stageBoxEnrich, stageBarEnrich, stageTooltipEnrich, stageBadgeEnrich, '', pct, 'Cancelado');
          showMessage('O processamento foi cancelado pelo usuário.', 'error');
          setTimeout(() => {
            step2Container.classList.add('hidden');
            step1Container.classList.remove('hidden');
            hideMessage();
          }, 3000);
        } else {
          // Flow: Stage 2 100% -> Stage 3 (Tratamento) -> Stage 4 (Concatenação) -> Download
          const isQueueOnly = state.currentAction === 'queue_only';
          const notifTitle = isQueueOnly ? 'Fila Concluída' : 'Enriquecimento Concluído';
          const notifMsg = isQueueOnly ? `Lote com ${data.total_rows || 0} contatos formatado e pronto para importação!` : `Lote com ${data.total_rows || 0} contatos concluído com sucesso!`;
          sendDesktopNotification(notifTitle, notifMsg);
          updateSkillStage(stageBoxEnrich, stageBarEnrich, stageTooltipEnrich, stageBadgeEnrich, 'completed', 100, 'Concluído');

          // Smooth transition to Stage 3 (Tratando as informações: 0% -> 50% -> 100%)
          updateSkillStage(stageBoxClean, stageBarClean, stageTooltipClean, stageBadgeClean, 'active', 50, 'Tratando dados...');

          setTimeout(() => {
            updateSkillStage(stageBoxClean, stageBarClean, stageTooltipClean, stageBadgeClean, 'completed', 100, 'Concluído');

            // Smooth transition to Stage 4 (Concatenando informações: 0% -> 60% -> 100%)
            const concatLabel = state.currentAction === 'enrich_only' ? 'Estruturando colunas...' : (isQueueOnly ? 'Concatenando contatos...' : 'Concatenando 5 slots...');
            updateSkillStage(stageBoxConcat, stageBarConcat, stageTooltipConcat, stageBadgeConcat, 'active', 60, concatLabel);

            setTimeout(() => {
              updateSkillStage(stageBoxConcat, stageBarConcat, stageTooltipConcat, stageBadgeConcat, 'completed', 100, 'Pronto!');

              setTimeout(async () => {
                state.currentRecord.preview = data.preview || {};
                step2Container.classList.add('hidden');

                const finSelectedFields = normalizeSelectedFields(state.selectedFields);
                const finQueueNumber = state.queueNumber || (state.currentAction === 'enrich_only' ? 'ENRICH' : '1');
                const finRequester = state.requesterName || '';
                const chosenFmt = state.exportFormat || 'original';

                try {
                  const finResp = await fetch(`/api/request/${requestId}/finalize`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                      queue_number: finQueueNumber,
                      requester_name: finRequester,
                      observacoes: state.observacoes || '',
                      selected_fields: finSelectedFields,
                      mode: state.currentAction || 'enrich_only',
                      export_format: chosenFmt,
                    })
                  });

                  const finPayload = await finResp.json();
                  if (!finResp.ok) {
                    throw new Error(finPayload.error || 'Erro ao finalizar.');
                  }

                  const defaultDlName = getStandardExportFilename({
                    mode: state.currentAction || 'enrich_only',
                    queue_number: finQueueNumber,
                    requester_name: finRequester,
                  }, chosenFmt);

                  const dlUrl = (isQueueOnly || chosenFmt === 'original')
                    ? `/api/request/${requestId}/download`
                    : `/api/request/${requestId}/download?format=${chosenFmt}`;

                  downloadFileDirectly(dlUrl, defaultDlName);
                  const successMsg = isQueueOnly
                    ? (finPayload.message || 'Planilha tratada e concatenada com sucesso! O download foi iniciado.')
                    : (finPayload.message || 'Enriquecimento concluído com sucesso! O download foi iniciado.');
                  showMessage(successMsg, 'success');
                  loadQueue();
                  loadLogsQueue();
                } catch (finErr) {
                  showMessage(`Erro ao finalizar arquivo: ${finErr.message}`, 'error');
                }
              }, 400);
            }, 350);
          }, 350);
        }
      }
    } catch (err) {
      console.error(err);
    } finally {
      isPolling = false;
    }
  }, 700);
}

/* ==========================================================================
   Web Notifications (Web Notification API)
   ========================================================================== */
function initBrowserNotifications() {
  const btn = document.getElementById('btnDesktopNotification');
  if (!btn) return;

  if (window.Notification && Notification.permission === 'granted') {
    btn.classList.add('is-active');
    btn.title = 'Notificações na área de trabalho ativadas';
  }

  btn.addEventListener('click', async () => {
    if (!window.Notification) {
      showToast('Seu navegador não suporta notificações de área de trabalho.', 'warning');
      return;
    }
    if (Notification.permission === 'granted') {
      showToast('Notificações de área de trabalho já estão ativadas.', 'info');
      return;
    }
    const perm = await Notification.requestPermission();
    if (perm === 'granted') {
      btn.classList.add('is-active');
      showToast('Notificações ativadas! Você será avisado ao concluir filas.', 'success');
      sendDesktopNotification('Sistema de Enriquecimento', 'Notificações ativadas com sucesso!');
    } else {
      showToast('Permissão para notificações não foi concedida.', 'warning');
    }
  });
}

function sendDesktopNotification(title, message) {
  if (window.Notification && Notification.permission === 'granted') {
    try {
      new Notification(title, {
        body: message,
        icon: '/static/favicon.png',
      });
    } catch (e) {
      console.warn('Erro ao enviar notificação desktop:', e);
    }
  }
}

/* ==========================================================================
   Dashboard de Produtividade e Assertividade (Analytics)
   ========================================================================== */
async function loadAnalyticsDashboard() {
  try {
    const resp = await fetch('/api/analytics');
    if (!resp.ok) return;
    const data = await resp.json();

    const totalRec = document.getElementById('dashTotalRecords');
    const enrRec = document.getElementById('dashEnrichedRecords');
    const assertEl = document.getElementById('dashAssertiveness');
    const compQueues = document.getElementById('dashCompletedQueues');
    const rejQueues = document.getElementById('dashRejectedQueues');

    if (totalRec) totalRec.textContent = (data.total_records_processed || 0).toLocaleString('pt-BR');
    if (enrRec) enrRec.textContent = (data.total_records_enriched || 0).toLocaleString('pt-BR');
    if (assertEl) assertEl.textContent = `${data.avg_assertiveness || 0}%`;
    if (compQueues) compQueues.textContent = data.completed_queues || 0;
    if (rejQueues) rejQueues.textContent = data.rejected_queues || 0;
  } catch (e) {
    console.warn('Erro ao carregar dashboard de analytics:', e);
  }
}

/* ==========================================================================
   Filtros e Busca Instantânea no Histórico da Estratégia
   ========================================================================== */
function initStrategyHistoryFilters() {
  const searchInput = document.getElementById('strategyHistorySearchInput');
  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      strategyFilterState.search = (e.target.value || '').trim().toLowerCase();
      renderQueueHistory();
    });
  }

  const modeSelect = document.getElementById('strategyHistoryModeFilter');
  if (modeSelect) {
    modeSelect.addEventListener('change', (e) => {
      strategyFilterState.mode = e.target.value || 'all';
      renderQueueHistory();
    });
  }

  const filterBtns = document.querySelectorAll('#strategyHistoryStatusFilters .history-filter-btn');
  filterBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      filterBtns.forEach((b) => b.classList.remove('is-active'));
      btn.classList.add('is-active');
      strategyFilterState.status = btn.dataset.strategyFilter || 'all';
      renderQueueHistory();
    });
  });
}

/* ==========================================================================
   Registro de Auditoria (Audit Logs)
   ========================================================================== */
let auditLogsData = [];

function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

async function openAuditModal() {
  const modal = document.getElementById('auditModal');
  if (!modal) return;
  modal.classList.remove('hidden');
  await loadAuditLogs();
}

function closeAuditModal() {
  const modal = document.getElementById('auditModal');
  if (modal) modal.classList.add('hidden');
}

async function loadAuditLogs() {
  const tbody = document.getElementById('auditTableBody');
  if (tbody) {
    tbody.innerHTML = '<tr><td colspan="5" style="text-align:center; padding: 1.5rem; color: var(--text-muted);">Carregando registros de auditoria...</td></tr>';
  }
  try {
    const resp = await fetch('/api/audit-logs');
    if (!resp.ok) throw new Error('Não foi possível carregar os registros.');
    const data = await resp.json();
    auditLogsData = data.audit_logs || [];
    renderAuditTable();
  } catch (e) {
    if (tbody) {
      tbody.innerHTML = `<tr><td colspan="5" style="text-align:center; padding: 1.5rem; color: var(--danger-color);">${e.message}</td></tr>`;
    }
  }
}

function renderAuditTable() {
  const tbody = document.getElementById('auditTableBody');
  const search = (document.getElementById('auditSearchInput')?.value || '').trim().toLowerCase();
  if (!tbody) return;

  const filtered = auditLogsData.filter((log) => {
    if (!search) return true;
    const str = `${log.timestamp} ${log.user_name} ${log.action} ${JSON.stringify(log.details)} ${log.ip}`.toLowerCase();
    return str.includes(search);
  });

  if (filtered.length === 0) {
    tbody.innerHTML = '<tr><td colspan="5" style="text-align:center; padding: 1.5rem; color: var(--text-muted);">Nenhum registro de auditoria encontrado.</td></tr>';
    return;
  }

  tbody.innerHTML = filtered.map((item) => {
    let badgeClass = 'login';
    if (item.action.includes('UPLOAD')) badgeClass = 'upload';
    else if (item.action.includes('FINALIZED') || item.action.includes('ENRICH')) badgeClass = 'finalize';
    else if (item.action.includes('REJECTED')) badgeClass = 'rejected';
    else if (item.action.includes('DOWNLOAD')) badgeClass = 'download';

    const detailsStr = typeof item.details === 'object' ? Object.entries(item.details).map(([k, v]) => `${k}: ${v}`).join(' | ') : String(item.details || '');

    return `
      <tr>
        <td style="white-space: nowrap; font-size: 0.8rem; color: var(--text-muted);">${escapeHtml(item.timestamp)}</td>
        <td><strong>${escapeHtml(item.user_name || 'Sistema')}</strong></td>
        <td><span class="audit-badge-action ${badgeClass}">${escapeHtml(item.action)}</span></td>
        <td style="max-width: 320px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${escapeHtml(detailsStr)}">${escapeHtml(detailsStr)}</td>
        <td style="font-size: 0.8rem; color: var(--text-muted);">${escapeHtml(item.ip || '127.0.0.1')}</td>
      </tr>
    `;
  }).join('');
}

window.openAuditModal = openAuditModal;
window.closeAuditModal = closeAuditModal;

const btnOpenAudit = document.getElementById('btnOpenAuditLogs');
if (btnOpenAudit) {
  btnOpenAudit.addEventListener('click', openAuditModal);
}

const btnRefreshAudit = document.getElementById('btnRefreshAudit');
if (btnRefreshAudit) {
  btnRefreshAudit.addEventListener('click', loadAuditLogs);
}

const auditSearch = document.getElementById('auditSearchInput');
if (auditSearch) {
  auditSearch.addEventListener('input', renderAuditTable);
}

const auditModalEl = document.getElementById('auditModal');
if (auditModalEl) {
  auditModalEl.addEventListener('click', (e) => {
    if (e.target === auditModalEl) {
      closeAuditModal();
    }
  });
}

// Background auto-refresh for queue and logs based on visible panel
setInterval(() => {
  const visiblePanel = Object.entries(panels).find(([key, panel]) => panel && panel.classList.contains('is-visible'));
  if (visiblePanel) {
    const target = visiblePanel[0];
    const userRole = document.body.dataset.userRole;
    if (target === 'strategy-panel' && userRole === 'admin') {
      loadQueue();
    } else if (target === 'user-panel') {
      loadQueue();
    } else if (target === 'logs-panel') {
      loadLogsQueue();
    }
  }
}, 10000);

// Global viewDetailedLogs registration
window.viewDetailedLogs = viewDetailedLogs;

applyRoleAccess();
// Activate initial panel based on the starting role
if (state.currentRole === 'admin') {
  setActiveTab('strategy-panel');
} else {
  setActiveTab('user-panel');
}
ensureStepVisibility();
renderAvailableFields();
renderSelectedFields();
renderPreview();
updateRequestButtonState();
loadQueue();
loadNotifications();
loadAnalyticsDashboard();
initBrowserNotifications();
initStrategyHistoryFilters();

async function cancelEnrichment() {
  const requestId = state.activeRequestImportId;
  if (!requestId) return;

  const btn = document.getElementById('cancelEnrichButton');
  try {
    if (btn) {
      btn.disabled = true;
      btn.textContent = 'Cancelando...';
    }
    const response = await fetch(`/api/request/${requestId}/cancel`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || 'Erro ao cancelar enriquecimento.');
    }
    showMessage(payload.message, 'success');
  } catch (error) {
    showMessage(error.message, 'error');
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.textContent = 'Cancelar Enriquecimento';
    }
  }
}

const cancelEnrichButton = document.getElementById('cancelEnrichButton');
if (cancelEnrichButton) {
  cancelEnrichButton.addEventListener('click', cancelEnrichment);
}

// Drag & Drop File Upload listeners
const excelFile = document.getElementById('excelFile');
if (dropZone && excelFile) {
  ['dragenter', 'dragover'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      dropZone.classList.add('drop-zone--over');
    }, false);
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      dropZone.classList.remove('drop-zone--over');
    }, false);
  });

  dropZone.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    const files = dt.files;
    if (files && files.length) {
      addStagedFiles(files);
    }
  }, false);

  // Listener to accumulate attached files from explorer dialog
  excelFile.addEventListener('change', () => {
    if (excelFile.files && excelFile.files.length) {
      addStagedFiles(excelFile.files);
      excelFile.value = '';
    }
  });
}

// Botões de Ação para Arquivos Staged
const btnClearStagedFiles = document.getElementById('btnClearStagedFiles');
if (btnClearStagedFiles) {
  btnClearStagedFiles.addEventListener('click', () => {
    clearStagedFiles();
    showToast('Lista de planilhas limpa.', 'info');
  });
}

const btnMergeUpload = document.getElementById('btnMergeUpload');
if (btnMergeUpload) {
  btnMergeUpload.addEventListener('click', () => {
    uploadMultipleMerged(state.stagedFiles);
  });
}

const btnIndividualUpload = document.getElementById('btnIndividualUpload');
if (btnIndividualUpload) {
  btnIndividualUpload.addEventListener('click', () => {
    prepareIndividualBatchUpload();
  });
}

// Warning box buttons click listeners
if (btnProceedValidation) {
  btnProceedValidation.addEventListener('click', () => {
    if (validationWarningBox) validationWarningBox.classList.add('hidden');
    if (state.currentAction === 'queue_only') {
      if (step3Container) {
        step3Container.classList.remove('hidden');
        renderAvailableFields();
        renderSelectedFields();
        renderPreview();
      }
    } else {
      if (step1Container) step1Container.classList.remove('hidden');
    }
    showToast("Planilha aceita pelo usuário.", "info");
  });
}

if (btnCancelValidation) {
  btnCancelValidation.addEventListener('click', () => {
    clearImportedFileState();
    showToast("Envie uma planilha corrigida.", "warning");
  });
}

// ==========================================================================
// Modern Theme Switcher Logic (8 Themes, Swatches, Rainbow & RGB Auto-Cycle)
// ==========================================================================
const THEME_META = {
  classic: { name: 'Cyber Dark', color: '#00d2ff' },
  light: { name: 'Modo Claro', color: '#0ea5e9' },
  emerald: { name: 'Esmeralda Pro', color: '#34d399' },
  purple: { name: 'Roxo Velvet', color: '#c084fc' },
  sunset: { name: 'Sunset', color: '#fbbf24' },
  oled: { name: 'Midnight OLED', color: '#38bdf8' },
  rainbow: { name: 'Arco-íris', color: '#f97316' },
  rgb_cycle: { name: 'Ciclo RGB', color: '#f97316' },
};

let rgbCycleTimer = null;
const RGB_THEME_CYCLE = ['classic', 'emerald', 'purple', 'sunset', 'rainbow', 'oled', 'light'];
let rgbCycleIndex = 0;

function startRgbCycle() {
  stopRgbCycle();
  document.documentElement.classList.add('theme-rgb-active');
  document.body.classList.add('theme-rgb-active');

  rgbCycleTimer = setInterval(() => {
    rgbCycleIndex = (rgbCycleIndex + 1) % RGB_THEME_CYCLE.length;
    const nextTheme = RGB_THEME_CYCLE[rgbCycleIndex];
    document.documentElement.setAttribute('data-theme', nextTheme);
    document.body.setAttribute('data-theme', nextTheme);
    const swatch = document.getElementById('themeCurrentSwatch');
    if (swatch) {
      swatch.style.background = THEME_META[nextTheme]?.color || '#ff007f';
      swatch.style.boxShadow = `0 0 10px ${THEME_META[nextTheme]?.color || '#ff007f'}`;
    }
  }, 10000);
}

function stopRgbCycle() {
  if (rgbCycleTimer) {
    clearInterval(rgbCycleTimer);
    rgbCycleTimer = null;
  }
  document.documentElement.classList.remove('theme-rgb-active');
  document.body.classList.remove('theme-rgb-active');
}

function initThemeSystem() {
  const dropdownWrapper = document.getElementById('themeDropdownWrapper');
  const dropdownTrigger = document.getElementById('btnThemeDropdown');
  const dropdownMenu = document.getElementById('themeDropdownMenu');
  const currentSwatch = document.getElementById('themeCurrentSwatch');
  const currentName = document.getElementById('themeCurrentName');
  const themeOptions = document.querySelectorAll('.theme-option-item');

  function setTheme(themeName) {
    const validTheme = THEME_META[themeName] ? themeName : 'classic';
    localStorage.setItem('selected-theme', validTheme);

    const meta = THEME_META[validTheme];
    if (currentName) currentName.textContent = meta.name;

    if (validTheme === 'rgb_cycle') {
      if (currentSwatch) {
        currentSwatch.className = 'theme-current-swatch rgb-animated-swatch';
        currentSwatch.style.background = '';
        currentSwatch.style.boxShadow = '';
      }
      startRgbCycle();
    } else {
      stopRgbCycle();
      document.documentElement.setAttribute('data-theme', validTheme);
      document.body.setAttribute('data-theme', validTheme);
      if (currentSwatch) {
        currentSwatch.className = 'theme-current-swatch';
        currentSwatch.style.background = meta.color;
        currentSwatch.style.boxShadow = `0 0 8px ${meta.color}`;
      }
    }

    themeOptions.forEach(opt => {
      if (opt.dataset.themeName === validTheme) {
        opt.classList.add('is-active');
      } else {
        opt.classList.remove('is-active');
      }
    });
  }

  if (dropdownTrigger && dropdownMenu) {
    dropdownTrigger.onclick = (e) => {
      e.preventDefault();
      e.stopPropagation();
      const isOpen = dropdownMenu.classList.contains('is-open');
      if (isOpen) {
        dropdownMenu.classList.remove('is-open');
        dropdownTrigger.setAttribute('aria-expanded', 'false');
      } else {
        dropdownMenu.classList.add('is-open');
        dropdownTrigger.setAttribute('aria-expanded', 'true');
      }
    };

    document.addEventListener('click', (e) => {
      if (dropdownWrapper && !dropdownWrapper.contains(e.target)) {
        dropdownMenu.classList.remove('is-open');
        dropdownTrigger.setAttribute('aria-expanded', 'false');
      }
    });

    themeOptions.forEach(opt => {
      opt.onclick = (e) => {
        e.preventDefault();
        e.stopPropagation();
        const themeName = opt.dataset.themeName;
        setTheme(themeName);
        dropdownMenu.classList.remove('is-open');
        dropdownTrigger.setAttribute('aria-expanded', 'false');
      };
    });
  }

  // Load saved theme on startup
  const savedTheme = localStorage.getItem('selected-theme') || 'classic';
  setTheme(savedTheme);
}

initThemeSystem();

// Add mouse wheel zoom listener to viewerImage and initialize history filters
function initHistoryFilters() {
  const searchInput = document.getElementById('historySearchInput');
  const filterBtns = document.querySelectorAll('.history-filter-btn');
  const modeSelect = document.getElementById('historyModeFilter');

  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      state.historySearchText = e.target.value;
      renderNormalUserHistory();
    });
  }

  if (modeSelect) {
    modeSelect.addEventListener('change', (e) => {
      state.historyModeFilter = e.target.value || 'all';
      renderNormalUserHistory();
    });
  }

  filterBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      filterBtns.forEach(b => b.classList.remove('is-active'));
      btn.classList.add('is-active');
      state.historyStatusFilter = btn.dataset.status || 'all';
      renderNormalUserHistory();
    });
  });
}

initHistoryFilters();
initExportFormatControls();

document.addEventListener('DOMContentLoaded', () => {
  initThemeSystem();
  initHistoryFilters();
  initExportFormatControls();
  const img = document.getElementById('viewerImage');
  if (img) {
    const applyTransform = () => {
      const container = img.parentElement;
      if (container) {
        // Safe boundaries to prevent image from sliding completely off screen
        const maxTranslateX = Math.max(0, (img.clientWidth * currentZoom - container.clientWidth) / 2) + (container.clientWidth * 0.15);
        const maxTranslateY = Math.max(0, (img.clientHeight * currentZoom - container.clientHeight) / 2) + (container.clientHeight * 0.15);
        translateX = Math.max(-maxTranslateX, Math.min(maxTranslateX, translateX));
        translateY = Math.max(-maxTranslateY, Math.min(maxTranslateY, translateY));
      }
      img.style.transform = `scale(${currentZoom}) translate(${translateX}px, ${translateY}px)`;
    };

    img.addEventListener('wheel', (e) => {
      e.preventDefault();
      const zoomSpeed = 0.15;
      if (e.deltaY < 0) {
        currentZoom = Math.min(3, currentZoom + zoomSpeed);
      } else {
        currentZoom = Math.max(0.5, currentZoom - zoomSpeed);
      }
      applyTransform();
    });

    img.addEventListener('mousedown', (e) => {
      isPanning = true;
      img.style.cursor = 'grabbing';
      startX = e.clientX - translateX * currentZoom;
      startY = e.clientY - translateY * currentZoom;
    });

    window.addEventListener('mousemove', (e) => {
      if (!isPanning) return;
      translateX = (e.clientX - startX) / currentZoom;
      translateY = (e.clientY - startY) / currentZoom;
      applyTransform();
    });

    window.addEventListener('mouseup', () => {
      isPanning = false;
      if (img) {
        img.style.cursor = 'grab';
      }
    });
  }
});

/* ==========================================================================
   LÓGICA DO CONVERSOR DE ARQUIVOS (Estilo iLovePDF)
   ========================================================================== */

const converterState = {
  selectedTool: null,
  toolMeta: null,
  selectedFiles: [],
  options: {
    splitMode: 'all',
    splitRange: '',
    pageRangeMode: 'all',
    pageStart: '',
    pageEnd: '',
    pageRange: '',
    rotateAngle: 90,
    protectPassword: '',
    unlockPassword: '',
    imageFormat: 'png',
    excelMode: 'smart_enrich',
  },
  isInitialized: false,
  isProcessing: false,
};

function initConverterModule() {
  if (converterState.isInitialized) return;
  converterState.isInitialized = true;

  const toolsView = document.getElementById('converterToolsView');
  const workspaceContainer = document.getElementById('converterWorkspaceContainer');
  const toolsGrid = document.getElementById('converterToolsGrid');
  const filterPills = document.querySelectorAll('[data-converter-filter]');
  const btnBackToTools = document.getElementById('btnBackToTools');
  const dropZone = document.getElementById('converterDropZone');
  const fileInput = document.getElementById('converterFileInput');
  const btnAddMoreFiles = document.getElementById('btnAddMoreFiles');
  const btnExecute = document.getElementById('btnExecuteConversion');
  const btnConvertAnother = document.getElementById('btnConvertAnother');

  // Filtros de Categoria
  filterPills.forEach(pill => {
    pill.addEventListener('click', () => {
      filterPills.forEach(p => p.classList.remove('is-active'));
      pill.classList.add('is-active');
      const filter = pill.getAttribute('data-converter-filter');
      const cards = toolsGrid.querySelectorAll('.tool-card');

      cards.forEach(card => {
        const cat = card.getAttribute('data-category');
        if (filter === 'all' || cat === filter) {
          card.style.display = 'flex';
        } else {
          card.style.display = 'none';
        }
      });
    });
  });

  // Clique nos Cards de Ferramenta
  const toolCards = toolsGrid.querySelectorAll('.tool-card');
  toolCards.forEach(card => {
    card.addEventListener('click', () => {
      const tool = card.getAttribute('data-tool');
      const title = card.getAttribute('data-title');
      const desc = card.getAttribute('data-desc');
      const accept = card.getAttribute('data-accept');
      const isSingle = card.getAttribute('data-single') === 'true';
      const icon = card.querySelector('.tool-icon').innerHTML;

      openConverterWorkspace({ tool, title, desc, accept, isSingle, icon });
    });
  });

  // Botão Voltar para Todas as Ferramentas
  if (btnBackToTools) {
    btnBackToTools.addEventListener('click', () => {
      closeConverterWorkspace();
    });
  }

  // Drag & Drop na DropZone
  if (dropZone && fileInput) {
    dropZone.addEventListener('click', (e) => {
      if (e.target.tagName !== 'INPUT') {
        fileInput.click();
      }
    });

    ['dragenter', 'dragover'].forEach(eventName => {
      dropZone.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.add('drag-over');
      });
    });

    ['dragleave', 'drop'].forEach(eventName => {
      dropZone.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove('drag-over');
      });
    });

    dropZone.addEventListener('drop', (e) => {
      const dt = e.dataTransfer;
      if (dt && dt.files && dt.files.length > 0) {
        handleConverterFiles(Array.from(dt.files));
      }
    });

    fileInput.addEventListener('change', () => {
      if (fileInput.files && fileInput.files.length > 0) {
        handleConverterFiles(Array.from(fileInput.files));
      }
    });
  }

  if (btnAddMoreFiles && fileInput) {
    btnAddMoreFiles.addEventListener('click', () => {
      fileInput.click();
    });
  }

  // Opções Dinâmicas: Dividir PDF
  const splitRadios = document.querySelectorAll('input[name="splitMode"]');
  const splitRangeWrapper = document.getElementById('splitRangeInputWrapper');
  const splitRangeInput = document.getElementById('splitPageRangeInput');
  splitRadios.forEach(radio => {
    radio.addEventListener('change', () => {
      converterState.options.splitMode = radio.value;
      if (radio.value === 'range') {
        splitRangeWrapper.classList.remove('hidden');
      } else {
        splitRangeWrapper.classList.add('hidden');
      }
    });
  });
  if (splitRangeInput) {
    splitRangeInput.addEventListener('input', (e) => {
      converterState.options.splitRange = e.target.value.trim();
    });
  }

  // Opções Dinâmicas: Seleção de Intervalo de Páginas (De X até Y / Intervalos livres)
  const pageModeRadios = document.querySelectorAll('input[name="converterPageMode"]');
  const pageRangeInputsWrapper = document.getElementById('converterPageRangeInputs');
  const pageStartInput = document.getElementById('converterPageStartInput');
  const pageEndInput = document.getElementById('converterPageEndInput');
  const pageCustomRangeInput = document.getElementById('converterPageCustomRangeInput');

  function syncPageRangeFromInputs() {
    const s = converterState.options.pageStart;
    const e = converterState.options.pageEnd;
    if (s && e) {
      if (pageCustomRangeInput) pageCustomRangeInput.value = `${s}-${e}`;
      converterState.options.pageRange = `${s}-${e}`;
    } else if (s && !e) {
      if (pageCustomRangeInput) pageCustomRangeInput.value = `${s}-`;
      converterState.options.pageRange = `${s}-`;
    } else if (!s && e) {
      if (pageCustomRangeInput) pageCustomRangeInput.value = `1-${e}`;
      converterState.options.pageRange = `1-${e}`;
    }
  }

  pageModeRadios.forEach(radio => {
    radio.addEventListener('change', () => {
      converterState.options.pageRangeMode = radio.value;
      if (radio.value === 'range') {
        if (pageRangeInputsWrapper) pageRangeInputsWrapper.classList.remove('hidden');
      } else {
        if (pageRangeInputsWrapper) pageRangeInputsWrapper.classList.add('hidden');
      }
    });
  });

  if (pageStartInput) {
    pageStartInput.addEventListener('input', (e) => {
      converterState.options.pageStart = e.target.value.trim();
      syncPageRangeFromInputs();
    });
  }

  if (pageEndInput) {
    pageEndInput.addEventListener('input', (e) => {
      converterState.options.pageEnd = e.target.value.trim();
      syncPageRangeFromInputs();
    });
  }

  if (pageCustomRangeInput) {
    pageCustomRangeInput.addEventListener('input', (e) => {
      converterState.options.pageRange = e.target.value.trim();
    });
  }

  // Opções Dinâmicas: Girar PDF
  const rotateButtons = document.querySelectorAll('#optRotatePdf .opt-btn');
  rotateButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      rotateButtons.forEach(b => b.classList.remove('is-active'));
      btn.classList.add('is-active');
      converterState.options.rotateAngle = parseInt(btn.getAttribute('data-angle'), 10) || 90;
    });
  });

  // Opções Dinâmicas: PDF para Imagens
  const imgFormatButtons = document.querySelectorAll('#optPdfToImages .opt-btn');
  imgFormatButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      imgFormatButtons.forEach(b => b.classList.remove('is-active'));
      btn.classList.add('is-active');
      converterState.options.imageFormat = btn.getAttribute('data-format') || 'png';
    });
  });

  // Opções Dinâmicas: PDF para Excel
  const excelModeButtons = document.querySelectorAll('#optPdfToExcel .opt-btn');
  excelModeButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      excelModeButtons.forEach(b => b.classList.remove('is-active'));
      btn.classList.add('is-active');
      converterState.options.excelMode = btn.getAttribute('data-excel-mode') || 'smart_enrich';
    });
  });

  // Opções Dinâmicas: Senha Proteger / Desbloquear
  const protectInput = document.getElementById('protectPasswordInput');
  if (protectInput) {
    protectInput.addEventListener('input', (e) => {
      converterState.options.protectPassword = e.target.value;
    });
  }
  const unlockInput = document.getElementById('unlockPasswordInput');
  if (unlockInput) {
    unlockInput.addEventListener('input', (e) => {
      converterState.options.unlockPassword = e.target.value;
    });
  }

  // Execução da Conversão
  if (btnExecute) {
    btnExecute.addEventListener('click', () => {
      executeConverterProcess();
    });
  }

  // Converter outro arquivo
  if (btnConvertAnother) {
    btnConvertAnother.addEventListener('click', () => {
      resetConverterWorkspaceState();
    });
  }

  // Enriquecer Planilha Convertida (1-clique)
  const btnEnrichConverted = document.getElementById('btnEnrichConvertedFile');
  if (btnEnrichConverted) {
    btnEnrichConverted.addEventListener('click', async () => {
      if (!converterState.lastConvertedResult) {
        showToast('Nenhum arquivo convertido encontrado.', 'error');
        return;
      }
      const { download_id, filename } = converterState.lastConvertedResult;
      const originalHtml = btnEnrichConverted.innerHTML;
      btnEnrichConverted.disabled = true;
      btnEnrichConverted.innerHTML = `<span>Carregando...</span>`;

      try {
        const res = await fetch(`/api/converter/download/${download_id}`);
        if (!res.ok) throw new Error('Não foi possível obter o arquivo convertido para enriquecimento.');
        const blob = await res.blob();
        const file = new File([blob], filename, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

        setActiveTab('user-panel');
        const userTabImport = document.querySelector('[data-user-tab="import"]');
        if (userTabImport) {
          userTabButtons.forEach(b => b.classList.remove('is-active'));
          userTabImport.classList.add('is-active');
          if (userImportSection) userImportSection.classList.remove('hidden');
          if (normalUserHistoryContainer) normalUserHistoryContainer.classList.add('hidden');
          if (userConverterSection) userConverterSection.classList.add('hidden');
        }
        clearImportedFileState();
        await uploadSpreadsheetFile(file);
        showToast('Planilha transferida para enriquecimento com sucesso!', 'success');
      } catch (err) {
        console.error(err);
        showToast(err.message || 'Erro ao carregar planilha no enriquecimento.', 'error');
      } finally {
        btnEnrichConverted.disabled = false;
        btnEnrichConverted.innerHTML = originalHtml;
      }
    });
  }
}

function openConverterWorkspace(meta) {
  converterState.selectedTool = meta.tool;
  converterState.toolMeta = meta;
  converterState.selectedFiles = [];

  const toolsView = document.getElementById('converterToolsView');
  const workspaceContainer = document.getElementById('converterWorkspaceContainer');
  const toolTitle = document.getElementById('workspaceToolTitle');
  const toolDesc = document.getElementById('workspaceToolDesc');
  const toolIcon = document.getElementById('workspaceToolIcon');
  const fileInput = document.getElementById('converterFileInput');
  const acceptedBadge = document.getElementById('converterDropAccepted');
  const btnExecuteText = document.getElementById('btnExecuteText');

  toolTitle.textContent = meta.title;
  toolDesc.textContent = meta.desc;
  toolIcon.innerHTML = meta.icon;

  if (fileInput) {
    fileInput.value = '';
    fileInput.accept = meta.accept;
    fileInput.multiple = true;
  }

  if (acceptedBadge) {
    acceptedBadge.textContent = `Formatos aceitos: ${meta.accept.split(',').join(', ')}`;
  }

  // Nome dinâmico para o botão de conversão
  if (btnExecuteText) {
    if (meta.tool === 'merge_pdf') {
      btnExecuteText.textContent = 'Juntar PDFs';
    } else if (meta.tool === 'split_pdf') {
      btnExecuteText.textContent = 'Dividir PDF';
    } else if (meta.tool === 'compress_pdf') {
      btnExecuteText.textContent = 'Comprimir PDF';
    } else if (meta.tool === 'rotate_pdf') {
      btnExecuteText.textContent = 'Girar PDF';
    } else if (meta.tool === 'protect_pdf') {
      btnExecuteText.textContent = 'Proteger PDF';
    } else if (meta.tool === 'unlock_pdf') {
      btnExecuteText.textContent = 'Desbloquear PDF';
    } else {
      btnExecuteText.textContent = `Converter para ${meta.title.split(' para ')[1] || 'Documento'}`;
    }
  }

  resetConverterWorkspaceState();

  toolsView.classList.add('hidden');
  workspaceContainer.classList.remove('hidden');
}

function closeConverterWorkspace() {
  const toolsView = document.getElementById('converterToolsView');
  const workspaceContainer = document.getElementById('converterWorkspaceContainer');

  converterState.selectedTool = null;
  converterState.toolMeta = null;
  converterState.selectedFiles = [];

  workspaceContainer.classList.add('hidden');
  toolsView.classList.remove('hidden');
}

function resetConverterWorkspaceState() {
  const dropZone = document.getElementById('converterDropZone');
  const fileListContainer = document.getElementById('converterFileListContainer');
  const optionsPanel = document.getElementById('converterDynamicOptions');
  const actionRow = document.getElementById('converterActionRow');
  const progressContainer = document.getElementById('converterProgressContainer');
  const resultContainer = document.getElementById('converterResultContainer');
  const fileInput = document.getElementById('converterFileInput');

  if (fileInput) fileInput.value = '';
  converterState.selectedFiles = [];

  dropZone.classList.remove('hidden');
  fileListContainer.classList.add('hidden');
  actionRow.classList.add('hidden');
  progressContainer.classList.add('hidden');
  resultContainer.classList.add('hidden');

  // Mostrar opções específicas da ferramenta
  setupToolDynamicOptions(converterState.selectedTool);
}

function setupToolDynamicOptions(tool) {
  const optionsPanel = document.getElementById('converterDynamicOptions');
  const optSplit = document.getElementById('optSplitPdf');
  const optRotate = document.getElementById('optRotatePdf');
  const optProtect = document.getElementById('optProtectPdf');
  const optUnlock = document.getElementById('optUnlockPdf');
  const optPdfToImages = document.getElementById('optPdfToImages');
  const optPdfToExcel = document.getElementById('optPdfToExcel');
  const optPageSelection = document.getElementById('optPageSelection');

  // Reset page range options
  const pageModeAll = document.getElementById('pageModeAll');
  if (pageModeAll) pageModeAll.checked = true;
  converterState.options.pageRangeMode = 'all';
  converterState.options.pageStart = '';
  converterState.options.pageEnd = '';
  converterState.options.pageRange = '';
  const pageRangeInputsWrapper = document.getElementById('converterPageRangeInputs');
  if (pageRangeInputsWrapper) pageRangeInputsWrapper.classList.add('hidden');
  const pageStartInput = document.getElementById('converterPageStartInput');
  if (pageStartInput) pageStartInput.value = '';
  const pageEndInput = document.getElementById('converterPageEndInput');
  if (pageEndInput) pageEndInput.value = '';
  const pageCustomRangeInput = document.getElementById('converterPageCustomRangeInput');
  if (pageCustomRangeInput) pageCustomRangeInput.value = '';

  // Esconder todas as opções primeiro
  [optSplit, optRotate, optProtect, optUnlock, optPdfToImages, optPdfToExcel, optPageSelection].forEach(opt => {
    if (opt) opt.classList.add('hidden');
  });

  let hasCustomOptions = false;

  if (tool === 'split_pdf') {
    optSplit.classList.remove('hidden');
    hasCustomOptions = true;
  } else if (tool === 'rotate_pdf') {
    optRotate.classList.remove('hidden');
    if (optPageSelection) optPageSelection.classList.remove('hidden');
    hasCustomOptions = true;
  } else if (tool === 'protect_pdf') {
    optProtect.classList.remove('hidden');
    hasCustomOptions = true;
  } else if (tool === 'unlock_pdf') {
    optUnlock.classList.remove('hidden');
    hasCustomOptions = true;
  } else if (tool === 'pdf_to_images') {
    optPdfToImages.classList.remove('hidden');
    if (optPageSelection) optPageSelection.classList.remove('hidden');
    hasCustomOptions = true;
  } else if (tool === 'pdf_to_excel') {
    optPdfToExcel.classList.remove('hidden');
    if (optPageSelection) optPageSelection.classList.remove('hidden');
    hasCustomOptions = true;
  } else if (tool === 'word_to_excel') {
    optPdfToExcel.classList.remove('hidden');
    hasCustomOptions = true;
  } else if (tool === 'pdf_to_word' || tool === 'pdf_to_text') {
    if (optPageSelection) optPageSelection.classList.remove('hidden');
    hasCustomOptions = true;
  }


  if (hasCustomOptions) {
    optionsPanel.classList.remove('hidden');
  } else {
    optionsPanel.classList.add('hidden');
  }
}

function handleConverterFiles(files) {
  if (!files || files.length === 0) return;
  const meta = converterState.toolMeta;
  if (!meta) return;

  const validExts = meta.accept.split(',').map(e => e.trim().toLowerCase());
  const filteredFiles = files.filter(f => {
    const ext = '.' + f.name.split('.').pop().toLowerCase();
    return validExts.includes(ext);
  });

  if (filteredFiles.length === 0) {
    showToast(`Formato não suportado. Por favor, envie arquivos com extensão: ${meta.accept}`, 'error');
    return;
  }

  // Adiciona múltiplos arquivos sem duplicatas de nome/tamanho
  filteredFiles.forEach(f => {
    const exists = converterState.selectedFiles.some(cur => cur.name === f.name && cur.size === f.size);
    if (!exists) {
      converterState.selectedFiles.push(f);
    }
  });

  renderConverterFilesList();
}

function renderConverterFilesList() {
  const fileListContainer = document.getElementById('converterFileListContainer');
  const fileList = document.getElementById('converterFileList');
  const filesCount = document.getElementById('converterFilesCount');
  const actionRow = document.getElementById('converterActionRow');
  const dropZone = document.getElementById('converterDropZone');
  const btnAddMore = document.getElementById('btnAddMoreFiles');

  if (converterState.selectedFiles.length === 0) {
    fileListContainer.classList.add('hidden');
    actionRow.classList.add('hidden');
    dropZone.classList.remove('hidden');
    return;
  }

  filesCount.textContent = converterState.selectedFiles.length;
  fileList.innerHTML = '';

  const formatFileSize = (bytes) => {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
  };

  const fileIconSvg = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>`;

  converterState.selectedFiles.forEach((file, index) => {
    const chip = document.createElement('div');
    chip.className = 'converter-file-chip';

    chip.innerHTML = `
      <div class="file-chip-info">
        <span class="file-chip-icon">${fileIconSvg}</span>
        <span class="file-chip-name" title="${file.name}">${file.name}</span>
        <span class="file-chip-size">(${formatFileSize(file.size)})</span>
      </div>
      <button type="button" class="file-chip-remove" title="Remover arquivo" data-index="${index}">×</button>
    `;

    chip.querySelector('.file-chip-remove').addEventListener('click', (e) => {
      e.stopPropagation();
      converterState.selectedFiles.splice(index, 1);
      renderConverterFilesList();
    });

    fileList.appendChild(chip);
  });

  if (btnAddMore) {
    btnAddMore.style.display = 'inline-block';
  }

  fileListContainer.classList.remove('hidden');
  actionRow.classList.remove('hidden');
}

let converterTimerInterval = null;
let converterPollInterval = null;
let converterStartTime = null;

function formatConverterDuration(seconds) {
  const s = Math.max(0, Math.floor(seconds));
  const mins = Math.floor(s / 60);
  const secs = s % 60;
  if (mins >= 60) {
    const hrs = Math.floor(mins / 60);
    const remMins = mins % 60;
    return `${String(hrs).padStart(2, '0')}:${String(remMins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
  }
  return `${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
}

function startConverterTimer() {
  stopConverterTimer();
  converterStartTime = Date.now();
  const statElapsed = document.getElementById('statConverterTimeElapsed');
  if (statElapsed) statElapsed.textContent = '00:00';

  converterTimerInterval = setInterval(() => {
    if (!converterStartTime) return;
    const elapsedSec = (Date.now() - converterStartTime) / 1000;
    if (statElapsed) {
      statElapsed.textContent = formatConverterDuration(elapsedSec);
    }
  }, 500);
}

function stopConverterTimer() {
  if (converterTimerInterval) {
    clearInterval(converterTimerInterval);
    converterTimerInterval = null;
  }
  if (converterPollInterval) {
    clearInterval(converterPollInterval);
    converterPollInterval = null;
  }
}

async function executeConverterProcess() {
  if (converterState.isProcessing) return;
  if (!converterState.selectedTool || converterState.selectedFiles.length === 0) {
    showToast('Por favor, selecione ao menos um arquivo para conversão.', 'error');
    return;
  }

  // Validações específicas
  const tool = converterState.selectedTool;
  const options = { ...converterState.options };

  if (tool === 'protect_pdf') {
    const pwd = document.getElementById('protectPasswordInput')?.value.trim() || '';
    if (!pwd) {
      showToast('Por favor, defina uma senha para proteger o documento.', 'error');
      return;
    }
    options.password = pwd;
  } else if (tool === 'unlock_pdf') {
    const pwd = document.getElementById('unlockPasswordInput')?.value.trim() || '';
    options.password = pwd;
  } else if (tool === 'split_pdf') {
    if (options.splitMode === 'range' && !options.splitRange) {
      showToast('Por favor, informe o intervalo de páginas (ex: 1-3, 5, 8).', 'error');
      return;
    }
    options.mode = options.splitMode;
    options.range = options.splitRange;
  } else if (tool === 'rotate_pdf') {
    options.angle = options.rotateAngle || 90;
  } else if (tool === 'pdf_to_images') {
    options.format = options.imageFormat || 'png';
  } else if (tool === 'pdf_to_excel' || tool === 'word_to_excel') {
    options.excelMode = options.excelMode || 'smart_enrich';
  }


  // Validação do intervalo de páginas caso selecionado
  if (options.pageRangeMode === 'range') {
    const s = parseInt(options.pageStart, 10);
    const e = parseInt(options.pageEnd, 10);
    if (!isNaN(s) && !isNaN(e) && s > e) {
      showToast('A página inicial não pode ser maior que a página final.', 'error');
      return;
    }
    if (!options.pageRange && !options.pageStart && !options.pageEnd) {
      showToast('Por favor, defina o intervalo de páginas desejado.', 'error');
      return;
    }
  }

  const dropZone = document.getElementById('converterDropZone');
  const fileListContainer = document.getElementById('converterFileListContainer');
  const optionsPanel = document.getElementById('converterDynamicOptions');
  const actionRow = document.getElementById('converterActionRow');
  const progressContainer = document.getElementById('converterProgressContainer');
  const resultContainer = document.getElementById('converterResultContainer');

  const statPages = document.getElementById('statConverterPages');
  const statRemaining = document.getElementById('statConverterTimeRemaining');
  const progressFill = document.getElementById('converterProgressFill');
  const progressTooltip = document.getElementById('converterProgressTooltip');
  const currentStepText = document.getElementById('converterCurrentStepText');

  // Reset UI elements
  if (statPages) statPages.textContent = 'Iniciando...';
  if (statRemaining) statRemaining.textContent = 'Calculando...';
  if (progressFill) progressFill.style.width = '0%';
  if (progressTooltip) progressTooltip.textContent = '0%';
  if (currentStepText) currentStepText.textContent = 'Iniciando leitura do documento...';

  // Transição para estado de carregamento
  converterState.isProcessing = true;
  dropZone.classList.add('hidden');
  fileListContainer.classList.add('hidden');
  optionsPanel.classList.add('hidden');
  actionRow.classList.add('hidden');
  resultContainer.classList.add('hidden');
  progressContainer.classList.remove('hidden');

  startConverterTimer();

  const CHUNK_SIZE = 4 * 1024 * 1024; // 4 MB por fatia - previne erro 413 em qualquer proxy / tunnel
  const totalBytes = converterState.selectedFiles.reduce((acc, f) => acc + (f.size || 0), 0);
  const requiresChunking = converterState.selectedFiles.some(f => f.size > CHUNK_SIZE) || totalBytes > CHUNK_SIZE;

  try {
    let startPayload;

    if (requiresChunking) {
      const uploadId = 'conv_' + Date.now() + '_' + Math.random().toString(36).substring(2, 9);
      let bytesUploaded = 0;

      for (let fIdx = 0; fIdx < converterState.selectedFiles.length; fIdx++) {
        const file = converterState.selectedFiles[fIdx];
        const totalChunks = Math.ceil(file.size / CHUNK_SIZE) || 1;

        for (let cIdx = 0; cIdx < totalChunks; cIdx++) {
          const start = cIdx * CHUNK_SIZE;
          const end = Math.min(file.size, start + CHUNK_SIZE);
          const chunkBlob = file.slice(start, end);

          const chunkFormData = new FormData();
          chunkFormData.append('upload_id', uploadId);
          chunkFormData.append('file_index', fIdx);
          chunkFormData.append('filename', file.name);
          chunkFormData.append('chunk_index', cIdx);
          chunkFormData.append('total_chunks', totalChunks);
          chunkFormData.append('chunk', chunkBlob, file.name);

          bytesUploaded += (end - start);
          const uploadPct = Math.min(99, Math.round((bytesUploaded / totalBytes) * 100));

          if (progressFill) progressFill.style.width = `${uploadPct}%`;
          if (progressTooltip) progressTooltip.textContent = `${uploadPct}%`;
          if (currentStepText) {
            currentStepText.textContent = `Enviando arquivo em partes (${uploadPct}%): fatia ${cIdx + 1} de ${totalChunks}...`;
          }

          const chunkRes = await fetch('/api/converter/chunk', {
            method: 'POST',
            body: chunkFormData,
          });

          if (!chunkRes.ok) {
            const errTxt = await chunkRes.text();
            let parsedErr;
            try { parsedErr = JSON.parse(errTxt); } catch (e) { }
            if (chunkRes.status === 413) {
              throw new Error('O fragmento enviado excedeu o limite do servidor.');
            }
            throw new Error((parsedErr && parsedErr.error) || `Falha no envio da parte ${cIdx + 1} de ${totalChunks}.`);
          }
        }
      }

      if (currentStepText) {
        currentStepText.textContent = 'Arquivos enviados. Iniciando conversão...';
      }

      // Inicia conversão passando apenas o upload_id (payload leve, imune a 413)
      const startFormData = new FormData();
      startFormData.append('tool', tool);
      startFormData.append('options', JSON.stringify(options));
      startFormData.append('upload_id', uploadId);

      const response = await fetch('/api/converter/process', {
        method: 'POST',
        body: startFormData,
      });

      const rawText = await response.text();
      try {
        startPayload = JSON.parse(rawText);
      } catch (parseErr) {
        if (response.status === 401 || rawText.includes('login') || rawText.includes('<!doctype')) {
          throw new Error('Sessão expirada ou servidor reiniciado. Por favor, recarregue a página ou faça login novamente.');
        }
        throw new Error(`Erro na resposta do servidor (${response.status}): ${rawText.slice(0, 100)}`);
      }

      if (!response.ok || !startPayload.success || !startPayload.task_id) {
        throw new Error(startPayload.error || 'Erro ao iniciar conversão.');
      }

    } else {
      // Fluxo direto para arquivos pequenos (<= 4 MB)
      const formData = new FormData();
      formData.append('tool', tool);
      formData.append('options', JSON.stringify(options));

      converterState.selectedFiles.forEach(file => {
        formData.append('files', file);
      });

      const response = await fetch('/api/converter/process', {
        method: 'POST',
        body: formData,
      });

      const rawText = await response.text();
      try {
        startPayload = JSON.parse(rawText);
      } catch (parseErr) {
        if (response.status === 401 || rawText.includes('login') || rawText.includes('<!doctype')) {
          throw new Error('Sessão expirada ou servidor reiniciado. Por favor, recarregue a página ou faça login novamente.');
        }
        throw new Error(`Erro na resposta do servidor (${response.status}): ${rawText.slice(0, 100)}`);
      }

      if (!response.ok || !startPayload.success || !startPayload.task_id) {
        throw new Error(startPayload.error || 'Erro ao iniciar conversão.');
      }
    }

    const taskId = startPayload.task_id;

    // Loop de Polling em Background com tolerancia a instabilidade de rede / devtunnels
    await new Promise((resolve, reject) => {
      let consecutiveErrors = 0;
      const MAX_CONSECUTIVE_ERRORS = 15;

      converterPollInterval = setInterval(async () => {
        try {
          const statusRes = await fetch(`/api/converter/status/${taskId}`);
          if (!statusRes.ok) {
            consecutiveErrors++;
            console.warn(`[Conversor] Falha temporária ao consultar status (${statusRes.status}). Tentativa ${consecutiveErrors}/${MAX_CONSECUTIVE_ERRORS}`);
            if (currentStepText) {
              currentStepText.textContent = `Aguardando resposta do servidor (tentativa ${consecutiveErrors}/${MAX_CONSECUTIVE_ERRORS})...`;
            }
            if (consecutiveErrors >= MAX_CONSECUTIVE_ERRORS) {
              stopConverterTimer();
              reject(new Error(`Falha ao consultar status da conversão após ${MAX_CONSECUTIVE_ERRORS} tentativas (${statusRes.status}).`));
            }
            return;
          }

          consecutiveErrors = 0;
          const taskData = await statusRes.json();

          if (taskData.status === 'completed') {
            stopConverterTimer();
            if (progressFill) progressFill.style.width = '100%';
            if (progressTooltip) progressTooltip.textContent = '100%';
            if (statPages) statPages.textContent = taskData.page_info || 'Concluído';
            if (statRemaining) statRemaining.textContent = '0s';
            if (currentStepText) currentStepText.textContent = taskData.current_step || 'Conversão concluída com sucesso!';
            resolve(taskData);
          } else if (taskData.status === 'error') {
            stopConverterTimer();
            reject(new Error(taskData.error || 'Erro durante o processamento do arquivo.'));
          } else {
            // Atualização de progresso
            const pct = Math.min(99, Math.max(0, taskData.progress || 0));
            if (progressFill) progressFill.style.width = `${pct}%`;
            if (progressTooltip) progressTooltip.textContent = `${pct}%`;

            if (taskData.page_info) {
              if (statPages) statPages.textContent = taskData.page_info;
            } else if (taskData.total_pages > 0) {
              if (statPages) statPages.textContent = `Página ${taskData.processed_pages || 0} de ${taskData.total_pages}`;
            }

            if (taskData.current_step && currentStepText) {
              currentStepText.textContent = taskData.current_step;
            }

            // Estimativa de tempo restante
            const elapsed = (Date.now() - (converterStartTime || Date.now())) / 1000;
            if (pct > 5 && elapsed > 1) {
              const totalEst = elapsed / (pct / 100);
              const remaining = Math.max(0, Math.round(totalEst - elapsed));
              if (statRemaining) {
                statRemaining.textContent = remaining < 60 ? `~${remaining}s` : `~${Math.floor(remaining / 60)}m ${remaining % 60}s`;
              }
            } else {
              if (statRemaining) statRemaining.textContent = 'Calculando...';
            }
          }
        } catch (pollErr) {
          consecutiveErrors++;
          console.warn('[Conversor] Erro de conexao no polling:', pollErr);
          if (consecutiveErrors >= MAX_CONSECUTIVE_ERRORS) {
            stopConverterTimer();
            reject(pollErr);
          }
        }
      }, 1000);
    }).then((payload) => {
      // Sucesso na conversão
      progressContainer.classList.add('hidden');
      resultContainer.classList.remove('hidden');

      converterState.lastConvertedResult = payload;

      const resultFileName = document.getElementById('resultFileName');
      const resultFileSize = document.getElementById('resultFileSize');
      const btnDownload = document.getElementById('btnDownloadConvertedFile');
      const btnEnrich = document.getElementById('btnEnrichConvertedFile');

      if (resultFileName) resultFileName.textContent = payload.filename;
      if (resultFileSize) resultFileSize.textContent = payload.filesize;

      if (btnDownload) {
        btnDownload.href = `/api/converter/download/${payload.download_id}`;
        btnDownload.setAttribute('download', payload.filename);

        btnDownload.onclick = (e) => {
          e.preventDefault();
          let fileType = 'Arquivo Convertido';
          if (payload.filename.endsWith('.xlsx')) fileType = 'Planilha Excel';
          else if (payload.filename.endsWith('.docx')) fileType = 'Documento Word';
          else if (payload.filename.endsWith('.pdf')) fileType = 'Documento PDF';
          else if (payload.filename.endsWith('.zip')) fileType = 'Arquivo Compactado (ZIP)';
          else if (payload.filename.endsWith('.txt')) fileType = 'Arquivo de Texto';

          openDownloadRenameModal({
            downloadUrl: `/api/converter/download/${payload.download_id}`,
            defaultFilename: payload.filename,
            fileType: fileType,
          });
        };
      }

      const isSpreadsheet = (payload.filename.toLowerCase().endsWith('.xlsx') || payload.filename.toLowerCase().endsWith('.csv') || payload.filename.toLowerCase().endsWith('.xls')) && !payload.filename.toLowerCase().endsWith('.zip');
      if (btnEnrich) {
        if (isSpreadsheet) {
          btnEnrich.classList.remove('hidden');
          btnEnrich.style.display = 'inline-flex';
        } else {
          btnEnrich.classList.add('hidden');
          btnEnrich.style.display = 'none';
        }
      }

      showToast('Conversão finalizada com sucesso!', 'success');
    });

  } catch (error) {
    stopConverterTimer();
    console.error('Erro na conversão:', error);
    showToast(error.message || 'Falha ao processar arquivo.', 'error');

    // Restaurar visão
    progressContainer.classList.add('hidden');
    fileListContainer.classList.remove('hidden');
    optionsPanel.classList.remove('hidden');
    actionRow.classList.remove('hidden');
  } finally {
    converterState.isProcessing = false;
  }
}



