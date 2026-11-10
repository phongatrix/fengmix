/**
 * YoungMix – app.js
 * Vanilla JavaScript frontend logic for the exam shuffler UI.
 * No external libraries required.
 */

'use strict';

/* ============================================================
   CONFIG
   ============================================================ */
const API_BASE = '';   // Same origin; empty string = relative URLs

/* ============================================================
   STATE
   ============================================================ */
let selectedFile  = null;   // File object from drag-drop / file picker
let jobId         = null;   // Job ID returned by the API after processing
let progressTimer = null;   // rAF/setTimeout handle for fake progress
let isProcessing  = false;  // Guard against double-submit

/* ============================================================
   DOM REFERENCES  (resolved after DOMContentLoaded)
   ============================================================ */
let dropzone, dropzoneContent, fileInfo, fileNameEl, fileSizeEl,
    clearFileBtn, fileInput, mixForm, soMaDe, batDauTu, kieuTron,
    errorMsg, progressWrap, progressPct, progressBar, progressStatus,
    processBtn, resultsPanel, resultsTitle, resultsList, downloadBtn,
    resetBtn, ctaBtn, navbar;

/* ============================================================
   INIT
   ============================================================ */
document.addEventListener('DOMContentLoaded', () => {
  resolveRefs();
  initNavbar();
  initDragDrop();
  initFileInput();
  initForm();
  initCTA();
  initReset();
  initScrollAnimations();
  initQuickSample();
  initYmWorkspace();
});

/* ---- Resolve all DOM refs in one shot ---- */
function resolveRefs() {
  dropzone        = document.getElementById('dropzone');
  dropzoneContent = document.getElementById('dropzone-content');
  fileInfo        = document.getElementById('file-info');
  fileNameEl      = document.getElementById('file-name');
  fileSizeEl      = document.getElementById('file-size');
  clearFileBtn    = document.getElementById('clear-file');
  fileInput       = document.getElementById('file-input');
  mixForm         = document.getElementById('mix-form');
  soMaDe          = document.getElementById('so-ma-de');
  batDauTu        = document.getElementById('bat-dau-tu');
  kieuTron        = document.getElementById('kieu-tron');
  errorMsg        = document.getElementById('error-msg');
  progressWrap    = document.getElementById('progress-wrap');
  progressPct     = document.getElementById('progress-pct');
  progressBar     = document.getElementById('progress-bar');
  progressStatus  = document.getElementById('progress-status');
  processBtn      = document.getElementById('process-btn');
  resultsPanel    = document.getElementById('results-panel');
  resultsTitle    = document.getElementById('results-title');
  resultsList     = document.getElementById('results-list');
  downloadBtn     = document.getElementById('download-btn');
  resetBtn        = document.getElementById('reset-btn');
  ctaBtn          = document.getElementById('cta-btn');
  navbar          = document.getElementById('navbar');
}

/* ============================================================
   NAVBAR — add "scrolled" class for shadow on scroll
   ============================================================ */
function initNavbar() {
  window.addEventListener('scroll', () => {
    navbar.classList.toggle('scrolled', window.scrollY > 10);
  }, { passive: true });
}

/* ============================================================
   DRAG & DROP
   ============================================================ */
function initDragDrop() {
  // Prevent browser default open-file behaviour when dropping anywhere
  document.addEventListener('dragover',  e => e.preventDefault());
  document.addEventListener('drop',      e => e.preventDefault());

  // Dropzone events
  dropzone.addEventListener('dragover',  handleDragOver);
  dropzone.addEventListener('dragleave', handleDragLeave);
  dropzone.addEventListener('drop',      handleDrop);

  // Keyboard / click to open file picker
  dropzone.addEventListener('click',  openFilePicker);
  dropzone.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openFilePicker(); }
  });

  clearFileBtn.addEventListener('click', e => {
    e.stopPropagation();   // Don't re-open file picker
    clearFile();
  });
}

function handleDragOver(e) {
  e.preventDefault();
  if (!isProcessing) dropzone.classList.add('dragover');
}

function handleDragLeave() {
  dropzone.classList.remove('dragover');
}

function handleDrop(e) {
  e.preventDefault();
  dropzone.classList.remove('dragover');
  if (isProcessing) return;

  const file = e.dataTransfer.files[0];
  if (file) acceptFile(file);
}

function openFilePicker() {
  if (isProcessing) return;
  fileInput.click();
}

/* ============================================================
   FILE INPUT (hidden <input type="file">)
   ============================================================ */
function initFileInput() {
  fileInput.addEventListener('change', () => {
    const file = fileInput.files[0];
    if (file) acceptFile(file);
    fileInput.value = '';   // Reset so same file can be picked again
  });
}

/* ---- Validate & store a chosen file ---- */
function acceptFile(file) {
  hideError();

  // Validate extension
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext !== 'docx') {
    showError('⚠️ Chỉ chấp nhận file .docx (Microsoft Word). Vui lòng chọn đúng định dạng.');
    return;
  }

  selectedFile = file;
  showFileInfo(file);
  loadAndShowPreview(file);
}

function showFileInfo(file) {
  dropzoneContent.style.display = 'none';
  fileInfo.style.display        = 'flex';
  fileNameEl.textContent        = file.name;
  fileSizeEl.textContent        = formatFileSize(file.size);
  dropzone.classList.add('has-file');
  hideResults();
}

function clearFile() {
  selectedFile = null;
  fileInfo.style.display        = 'none';
  dropzoneContent.style.display = '';
  dropzone.classList.remove('has-file');
  hideError();
  hideResults();
  const ws = document.getElementById('ym-workspace');
  if (ws) ws.style.display = 'none';
  initialLoadedExamData = null;
  workingExamData       = null;
  modifiedAnswers       = { p1: {}, p2: {}, p3: {} };
  modifiedQuestions     = { p1: {}, p2: {}, p3: {} };
  hasUnsavedChanges     = false;
  setChangeBadge(false);
}

/* ============================================================
   FORM SUBMISSION
   ============================================================ */
function initForm() {
  if (mixForm) mixForm.addEventListener('submit', handleSubmit);
}

async function handleSubmit(e) {
  if (e) e.preventDefault();
  if (isProcessing) return;

  hideError();
  hideResults();

  /* ----- Validate ----- */
  if (!selectedFile) {
    showError('⚠️ Vui lòng chọn file Word (.docx) trước khi trộn đề.');
    return;
  }

  // Sync soMaDe from ym-opt-count if selected in workspace
  const countSelect = document.getElementById('ym-opt-count');
  let soMaDeVal = countSelect ? parseInt(countSelect.value, 10) : (soMaDe ? parseInt(soMaDe.value, 10) : 4);
  if (!soMaDeVal || isNaN(soMaDeVal) || soMaDeVal < 1 || soMaDeVal > 100) {
    soMaDeVal = (soMaDe ? parseInt(soMaDe.value, 10) : 4) || 4;
  }
  if (soMaDe) soMaDe.value = soMaDeVal;
  if (countSelect) countSelect.value = String(soMaDeVal);

  const startInput = document.getElementById('ym-opt-start');
  const batDauTuVal = startInput ? (parseInt(startInput.value, 10) || 101) : (batDauTu ? parseInt(batDauTu.value, 10) || 101 : 101);

  if (isNaN(soMaDeVal) || soMaDeVal < 1 || soMaDeVal > 100) {
    showError('⚠️ Số mã đề phải là số nguyên từ 1 đến 100.');
    return;
  }
  if (isNaN(batDauTuVal) || batDauTuVal < 1) {
    showError('⚠️ Mã đề bắt đầu phải là số nguyên dương (≥ 1).');
    return;
  }

  /* ----- Start processing ----- */
  isProcessing = true;
  if (processBtn) {
    processBtn.disabled = true;
    const textEl = processBtn.querySelector('.btn__text');
    if (textEl) textEl.textContent = 'Đang trộn...';
  }

  showProgress(0);
  updateProgressStatus(`Đang tải file và chuẩn bị trộn ${soMaDeVal} mã đề...`);

  // Start fake progress animation (0 → 85%) while waiting for the API
  const fakeStop = startFakeProgress(0, 85, 2000);

  /* ----- Build FormData ----- */
  const formData = new FormData();
  formData.append('file',         selectedFile);
  formData.append('so_ma_de',     soMaDeVal);
  formData.append('num_codes',    soMaDeVal);
  formData.append('bat_dau_tu',   batDauTuVal);
  const randomCheck = document.getElementById('ym-opt-random');
  const isRandom = randomCheck ? randomCheck.checked : false;
  formData.append('random_codes', isRandom ? 'true' : 'false');
  formData.append('shuffle_mode', kieuTron.value);
  if (modifiedAnswers && (Object.keys(modifiedAnswers.p1 || {}).length || Object.keys(modifiedAnswers.p2 || {}).length || Object.keys(modifiedAnswers.p3 || {}).length)) {
    formData.append('modified_answers', JSON.stringify(modifiedAnswers));
  }
  if (modifiedQuestions && (Object.keys(modifiedQuestions.p1 || {}).length || Object.keys(modifiedQuestions.p2 || {}).length || Object.keys(modifiedQuestions.p3 || {}).length)) {
    formData.append('modified_questions', JSON.stringify(modifiedQuestions));
  }

  /* ----- POST to API ----- */
  let response, data;
  try {
    response = await fetch(`${API_BASE}/api/process`, {
      method: 'POST',
      body:   formData,
    });
    data = await response.json();
  } catch (err) {
    fakeStop();
    handleApiError('Không thể kết nối đến máy chủ. Vui lòng kiểm tra kết nối mạng và thử lại.');
    return;
  }

  if (!response.ok) {
    fakeStop();
    const msg = data?.detail || data?.message || data?.error || `Lỗi máy chủ (HTTP ${response.status}).`;
    handleApiError(`❌ ${msg}`);
    return;
  }

  /* ----- Success ----- */
  fakeStop();
  jobId = data.job_id ?? data.id ?? null;

  updateProgressStatus('Đang hoàn tất và đóng gói kết quả...');
  await animateProgress(getCurrentProgress(), 100, 600);

  setTimeout(() => {
    hideProgress();
    showResults(data);
    resetProcessButton();
    isProcessing = false;
  }, 400);
}

/* ============================================================
   PROGRESS BAR
   ============================================================ */
let _currentProgress = 0;

function showProgress(pct) {
  progressWrap.style.display = '';
  setProgress(pct);
}

function hideProgress() {
  progressWrap.style.display = 'none';
  setProgress(0);
}

function setProgress(pct) {
  _currentProgress       = pct;
  progressBar.style.width = `${pct}%`;
  progressPct.textContent = `${Math.round(pct)}%`;
}

function getCurrentProgress() { return _currentProgress; }

function updateProgressStatus(msg) {
  if (progressStatus) progressStatus.textContent = msg;
}

/**
 * Animate progress from `from` to `to` over `durationMs`.
 * Returns a Promise that resolves when animation is complete.
 */
function animateProgress(from, to, durationMs) {
  return new Promise(resolve => {
    const start  = performance.now();
    const range  = to - from;

    function step(now) {
      const elapsed  = now - start;
      const fraction = Math.min(elapsed / durationMs, 1);
      // Ease-out cubic
      const eased    = 1 - Math.pow(1 - fraction, 3);
      setProgress(from + range * eased);
      if (fraction < 1) {
        requestAnimationFrame(step);
      } else {
        resolve();
      }
    }
    requestAnimationFrame(step);
  });
}

/**
 * Fake progress: smoothly advances from `from` → `to` over `durationMs`.
 * Returns a stop function; call it once the real response arrives.
 */
function startFakeProgress(from, to, durationMs) {
  setProgress(from);
  let stopped = false;

  const statusMessages = [
    { pct: 10,  msg: 'Đang phân tích cấu trúc đề thi...' },
    { pct: 30,  msg: 'Đang xáo trộn câu hỏi và đáp án...' },
    { pct: 55,  msg: 'Đang tạo các mã đề...' },
    { pct: 72,  msg: 'Đang xuất file Word và Excel...' },
  ];
  let msgIdx = 0;

  const startTime = performance.now();

  function step(now) {
    if (stopped) return;
    const elapsed   = now - startTime;
    const fraction  = Math.min(elapsed / durationMs, 1);
    const eased     = 1 - Math.pow(1 - fraction, 2);  // Ease-out quad
    const newPct    = from + (to - from) * eased;

    setProgress(newPct);

    // Update status message at key milestones
    while (msgIdx < statusMessages.length && newPct >= statusMessages[msgIdx].pct) {
      updateProgressStatus(statusMessages[msgIdx].msg);
      msgIdx++;
    }

    if (fraction < 1) requestAnimationFrame(step);
  }

  requestAnimationFrame(step);

  return function stop() {
    stopped = true;
  };
}

/* ============================================================
   RESULTS
   ============================================================ */
function showResults(dataOrTotal) {
  let total = 2;
  let zipFilename = null;
  if (typeof dataOrTotal === 'object' && dataOrTotal !== null) {
    total = dataOrTotal.total || dataOrTotal.total_exams || 2;
    if (dataOrTotal.job_id) jobId = dataOrTotal.job_id;
    if (dataOrTotal.zip_filename) zipFilename = dataOrTotal.zip_filename;
  } else if (typeof dataOrTotal === 'number' && !isNaN(dataOrTotal)) {
    total = dataOrTotal;
  }

  // Định dạng file zip: fengmix_ngày_giờ_Tên file (lấy trong phần Thông tin đề thi gốc - bỏ .docx)
  if (!zipFilename) {
    const infoNameEl = document.getElementById('ym-info-filename');
    const rawName = (infoNameEl && infoNameEl.textContent && infoNameEl.textContent !== '...')
      ? infoNameEl.textContent.trim()
      : (selectedFile?.name || 'de_thi');
    const cleanName = rawName.replace(/\.docx$/i, '').trim() || 'de_thi';
    const now = new Date();
    const pad = n => String(n).padStart(2, '0');
    const dStr = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}`;
    const tStr = `${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
    zipFilename = `fengmix_${dStr}_${tStr}_${cleanName}.zip`;
  }

  resultsTitle.innerHTML = `Hoàn thành! Đã trộn <strong>${total}</strong> mã đề`;

  // Build file list
  resultsList.innerHTML = '';
  const files = [
    `📄 ${total} file Word (.docx) đề đã trộn`,
    `📊 1 file Excel đáp án (.xlsx) chuẩn app chấm thi QM 2025`,
    `📱 ${total} ảnh QR code đáp án quét vào app QM`,
  ];
  files.forEach(text => {
    const li = document.createElement('li');
    li.textContent = text;
    resultsList.appendChild(li);
  });

  // Set download link
  if (jobId) {
    downloadBtn.href = `${API_BASE}/api/download/${jobId}`;
    downloadBtn.setAttribute('download', zipFilename);
  } else {
    downloadBtn.href = '#';
    downloadBtn.addEventListener('click', e => e.preventDefault(), { once: true });
  }

  resultsPanel.style.display = '';

  // Smooth scroll to results
  setTimeout(() => {
    resultsPanel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }, 100);
}

function hideResults() {
  resultsPanel.style.display = 'none';
}

function syncExamCodeOptions(isRandom) {
  const startRow    = document.getElementById('row-opt-start');
  const startInput  = document.getElementById('ym-opt-start');
  const randomRow   = document.getElementById('row-opt-random');
  const randomCheck = document.getElementById('ym-opt-random');

  if (isRandom) {
    // Mã đề ngẫu nhiên: SÁNG
    if (randomCheck) randomCheck.checked = true;
    if (randomRow) {
      randomRow.classList.remove('ym-option-row--dimmed');
      randomRow.classList.add('ym-option-row--active');
    }
    // Mã đề bắt đầu: MỜ
    if (startInput) {
      startInput.disabled = true;
      startInput.title = 'Đang chọn chế độ mã đề ngẫu nhiên';
    }
    if (startRow) {
      startRow.classList.add('ym-option-row--dimmed');
      startRow.classList.remove('ym-option-row--active');
    }
  } else {
    // Mã đề ngẫu nhiên: MỜ
    if (randomCheck) randomCheck.checked = false;
    if (randomRow) {
      randomRow.classList.add('ym-option-row--dimmed');
      randomRow.classList.remove('ym-option-row--active');
    }
    // Mã đề bắt đầu: SÁNG
    if (startInput) {
      startInput.disabled = false;
      startInput.title = 'Mã đề đầu tiên (ví dụ: 101, 102...)';
    }
    if (startRow) {
      startRow.classList.remove('ym-option-row--dimmed');
      startRow.classList.add('ym-option-row--active');
    }
  }
}

/* ============================================================
   RESET
   ============================================================ */
function initReset() {
  resetBtn.addEventListener('click', fullReset);
}

function fullReset() {
  clearFile();
  hideError();
  hideProgress();
  hideResults();
  if (soMaDe) soMaDe.value  = 4;
  const countSelect = document.getElementById('ym-opt-count');
  if (countSelect) countSelect.value = '4';
  if (batDauTu) batDauTu.value = 101;
  const startInput = document.getElementById('ym-opt-start');
  if (startInput) {
    startInput.value = '101';
  }
  syncExamCodeOptions(false);
  if (kieuTron) kieuTron.selectedIndex = 0;
  jobId = null;
  isProcessing = false;
  resetProcessButton();

  // Scroll back up to file upload section
  const appEl = document.getElementById('app');
  if (appEl) appEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function resetProcessButton() {
  if (processBtn) {
    processBtn.disabled = false;
    const textEl = processBtn.querySelector('.btn__text');
    if (textEl) textEl.textContent = 'Trộn đề ngay';
  }
}

/* ============================================================
   ERROR HANDLING
   ============================================================ */
function handleApiError(msg) {
  showError(msg);
  hideProgress();
  resetProcessButton();
  isProcessing = false;
}

function showError(msg) {
  if (!errorMsg) return;
  errorMsg.textContent      = msg;
  errorMsg.style.display    = '';
  errorMsg.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function hideError() {
  if (!errorMsg) return;
  errorMsg.style.display = 'none';
  errorMsg.textContent   = '';
}

/* ============================================================
   CTA BUTTON — smooth scroll to #app
   ============================================================ */
function initCTA() {
  if (!ctaBtn) return;
  ctaBtn.addEventListener('click', e => {
    e.preventDefault();
    document.getElementById('app').scrollIntoView({ behavior: 'smooth', block: 'start' });
  });
}

/* ============================================================
   SCROLL ANIMATIONS (IntersectionObserver)
   ============================================================ */
function initScrollAnimations() {
  const targets = document.querySelectorAll('.fade-in-up');

  if (!('IntersectionObserver' in window)) {
    // Fallback: show all immediately
    targets.forEach(el => el.classList.add('visible'));
    return;
  }

  const observer = new IntersectionObserver(
    entries => {
      entries.forEach(entry => {
        if (entry.isIntersecting) {
          // Respect existing CSS animation-delay
          const delay = getComputedStyle(entry.target).animationDelay;
          const ms    = parseFloat(delay) * 1000 || 0;
          setTimeout(() => entry.target.classList.add('visible'), ms);
          observer.unobserve(entry.target);
        }
      });
    },
    { threshold: 0.12, rootMargin: '0px 0px -40px 0px' }
  );

  targets.forEach(el => observer.observe(el));
}

/* ============================================================
   UTILITIES
   ============================================================ */

/**
 * Format a byte count into a human-readable string.
 * @param {number} bytes
 * @returns {string}
 */
function formatFileSize(bytes) {
  if (bytes === 0) return '0 B';
  const k    = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i     = Math.floor(Math.log(bytes) / Math.log(k));
  return `${parseFloat((bytes / Math.pow(k, i)).toFixed(1))} ${sizes[i]}`;
}


/* ============================================================
   SAMPLE FILES (DANH SÁCH DẠNG BẢNG)
   ============================================================ */

const SAMPLE_CATALOG = {
  '2025_De Tin Hoc Tham Khao YM.docx': {
    title: 'Đề tham khảo Tin học 2025 (Đề gốc)',
    category: 'degoc',
    icon: '💻',
    badge: { label: 'Word .docx', class: 'badge--docx' },
    desc: 'Đề thi minh họa môn Tin học cấu trúc 2025 (Phần I: 24 câu trắc nghiệm, Phần II: 4 câu đúng sai, Phần III: 6 câu trả lời ngắn).',
    isInputDocx: true,
    order: 1,
  },
  '2025_De Dia li Tham Khao YM.docx': {
    title: 'Đề tham khảo Địa lí 2025 (Đề gốc)',
    category: 'degoc',
    icon: '🌍',
    badge: { label: 'Word .docx', class: 'badge--docx' },
    desc: 'Đề thi minh họa môn Địa lí cấu trúc 2025 theo chuẩn Bộ GD&ĐT (đầy đủ 3 phần trắc nghiệm, đúng/sai, trả lời ngắn).',
    isInputDocx: true,
    order: 2,
  },
};

let allSampleItems = [];
let currentFilter = 'all';
let currentSearch = '';

/**
 * Nạp trực tiếp file mẫu vào giao diện xem trước & trộn đề
 */
async function useSampleDirectly(sampleName, downloadUrl, btnEl) {
  if (isProcessing) return;
  const originalHtml = btnEl ? btnEl.innerHTML : '';
  if (btnEl) {
    btnEl.disabled = true;
    btnEl.innerHTML = '⏳ Đang nạp...';
  }

  try {
    const res = await fetch(downloadUrl);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const blob = await res.blob();
    const file = new File([blob], sampleName, {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });

    acceptFile(file);

    setTimeout(() => {
      const target = document.getElementById('ym-workspace') || document.getElementById('mix-section');
      if (target) {
        target.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    }, 120);

    if (btnEl) {
      btnEl.innerHTML = '✅ Đã nạp!';
      setTimeout(() => {
        btnEl.innerHTML = originalHtml;
        btnEl.disabled = false;
      }, 2000);
    }
  } catch (err) {
    showError('Không thể nạp file mẫu: ' + err.message);
    if (btnEl) {
      btnEl.innerHTML = originalHtml;
      btnEl.disabled = false;
    }
  }
}

/**
 * Chuẩn hóa metadata của file mẫu (chỉ nhận đúng 2 file đề tham khảo chuẩn)
 */
function getSampleMetadata(item) {
  const filename = item.name;
  if (SAMPLE_CATALOG[filename]) {
    return { ...SAMPLE_CATALOG[filename], ...item };
  }
  return null;
}

/**
 * Render một hàng (tr) trong bảng danh sách file mẫu
 */
function renderSampleRow(sample, index) {
  const tr = document.createElement('tr');
  tr.dataset.category = sample.category || 'degoc';
  tr.dataset.name = `${sample.title} ${sample.name} ${sample.desc}`.toLowerCase();

  const isZip = sample.category === 'zip';
  const downloadBtnClass = isZip ? 'btn-sample-dl btn-sample-dl--full' : 'btn-sample-dl';

  const actionHtml = `
    <a class="${downloadBtnClass}" href="${sample.download_url}" download="${sample.name}" title="Tải file về máy tính">
      ${isZip ? '📥 Tải về trọn bộ' : '📥 Tải về'}
    </a>
  `;

  tr.innerHTML = `
    <td class="col-stt">
      <span class="sample-stt-badge">${index}</span>
    </td>
    <td class="col-name">
      <div class="sample-file-cell">
        <span class="sample-file-icon">${sample.icon || '📄'}</span>
        <div class="sample-file-meta">
          <span class="sample-file-title">${sample.title || sample.name}</span>
          <span class="sample-file-rawname">${sample.name}</span>
        </div>
      </div>
    </td>
    <td class="col-type" style="text-align: center;">
      <span class="sample-type-badge ${sample.badge ? sample.badge.class : ''}">${sample.badge ? sample.badge.label : 'FILE'}</span>
    </td>
    <td class="col-desc">
      <div class="sample-desc-text">${sample.desc || sample.name}</div>
    </td>
    <td class="col-size">
      <span class="sample-size-text">${formatFileSize(sample.size)}</span>
    </td>
    <td class="col-action">
      <div class="sample-actions-cell">
        ${actionHtml}
      </div>
    </td>
  `;

  const useBtn = tr.querySelector('.btn-sample-use');
  if (useBtn) {
    useBtn.addEventListener('click', (e) => {
      e.preventDefault();
      useSampleDirectly(sample.name, sample.download_url, useBtn);
    });
  }

  return tr;
}

/**
 * Lọc và hiển thị dữ liệu bảng theo Tab và Tìm kiếm
 */
function applySampleFilters() {
  const tbody = document.getElementById('samples-table-body');
  if (!tbody) return;

  const filtered = allSampleItems.filter(item => {
    if (currentFilter !== 'all' && item.category !== currentFilter) {
      return false;
    }
    if (currentSearch) {
      const q = currentSearch.toLowerCase().trim();
      const text = `${item.title} ${item.name} ${item.desc}`.toLowerCase();
      if (!text.includes(q)) return false;
    }
    return true;
  });

  if (filtered.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="6" class="samples-empty">
          🔍 Không tìm thấy file mẫu phù hợp với từ khóa "${currentSearch}".
        </td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = '';
  filtered.forEach((sample, idx) => {
    tbody.appendChild(renderSampleRow(sample, idx + 1));
  });
}

/**
 * Tải danh sách file mẫu từ API và render bảng
 */
async function loadSamples() {
  const tbody = document.getElementById('samples-table-body');
  if (!tbody) return;

  try {
    const res = await fetch('/api/samples');
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();

    const rawFiles = data.samples || [];
    const items = [];

    rawFiles.forEach(f => {
      const meta = getSampleMetadata(f);
      if (meta) items.push(meta);
    });

    items.sort((a, b) => (a.order || 99) - (b.order || 99));
    allSampleItems = items;

    // Cập nhật số lượng
    const countAllEl = document.getElementById('count-all');
    if (countAllEl) countAllEl.textContent = allSampleItems.length;

    // Render bảng
    applySampleFilters();

    // Khởi tạo sự kiện Tabs lọc
    const tabButtons = document.querySelectorAll('.sample-tab-btn');
    tabButtons.forEach(btn => {
      btn.addEventListener('click', () => {
        tabButtons.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        currentFilter = btn.dataset.filter || 'all';
        applySampleFilters();
      });
    });

    // Khởi tạo sự kiện tìm kiếm
    const searchInput = document.getElementById('samples-search-input');
    if (searchInput) {
      searchInput.addEventListener('input', (e) => {
        currentSearch = e.target.value;
        applySampleFilters();
      });
    }

  } catch (err) {
    tbody.innerHTML = `
      <tr>
        <td colspan="6" class="samples-loading">
          ⚠️ Không thể tải danh sách file mẫu (${err.message}).
        </td>
      </tr>
    `;
  }
}

// Auto-load when page is ready
document.addEventListener('DOMContentLoaded', loadSamples);

/**
 * Initialize the "Thử ngay với Đề gốc mẫu" quick-load button.
 */
function initQuickSample() {
  const btn = document.getElementById('btn-use-sample');
  if (!btn) return;

  btn.addEventListener('click', async (e) => {
    e.stopPropagation();
    btn.disabled = true;
    const oldText = btn.textContent;
    btn.textContent = '⏳ Đang nạp đề mẫu...';

    try {
      const res = await fetch('/api/samples/download/de_goc_mau.docx');
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const blob = await res.blob();
      const file = new File([blob], 'de_goc_mau.docx', {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });
      acceptFile(file);
      btn.textContent = '✅ Đã nạp file mẫu!';
      setTimeout(() => {
        btn.textContent = '⚡ Đổi lại đề mẫu';
        btn.disabled = false;
      }, 2000);
    } catch (err) {
      showError('Không thể nạp file mẫu: ' + err.message);
      btn.textContent = oldText;
      btn.disabled = false;
    }
  });
}


/* ============================================================
   YOUNGMIX WORKSPACE: PREVIEW & CONFIG (SAU INPUT, TRƯỚC OUTPUT)
   ============================================================ */

function escapeHtml(str) {
  if (!str) return '';
  return String(str).replace(/[&<>"']/g, m => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;',
  })[m]);
}

/* ============================================================
   EXAM PREVIEW & INTERACTIVE EDITING STATE
   ============================================================ */
let initialLoadedExamData = null;  // Pristine parsed data as originally uploaded
let workingExamData       = null;  // Current working copy with user edits
let modifiedAnswers       = { p1: {}, p2: {}, p3: {} }; // Track user modifications to answers
let modifiedQuestions     = { p1: {}, p2: {}, p3: {} }; // Track user modifications to question content and options
let hasUnsavedChanges     = false; // Dirty flag

function setChangeBadge(hasChanges) {
  const badge = document.getElementById('ym-change-badge');
  if (badge) badge.style.display = hasChanges ? 'inline-block' : 'none';
}

function showToolbarToast(message, type = 'success') {
  const toast = document.getElementById('ym-toolbar-toast');
  if (!toast) return;
  toast.textContent = message;
  toast.className = `ym-toolbar-toast ${type}`;
  toast.style.display = 'block';
  setTimeout(() => {
    toast.style.display = 'none';
  }, 3500);
}

function renderWarningsCard() {
  const warnList = document.getElementById('ym-warnings-list');
  if (!warnList) return;
  const warnings = workingExamData?.warnings || [];
  if (warnings.length > 0) {
    warnList.style.display = 'flex';
    warnList.innerHTML = warnings
      .map(w => `<li class="ym-warn-li"><span class="ym-warn-bullet">•</span> <span class="ym-warn-icon">⚠️</span> <span class="ym-warn-item-msg">${escapeHtml(w)}</span></li>`)
      .join('');
  } else {
    warnList.style.display = 'none';
    warnList.innerHTML = '';
  }
}

function recalcWarnings() {
  if (!workingExamData || !workingExamData.groups) return;
  const newWarnings = [];
  workingExamData.groups.forEach(g => {
    (g.questions || []).forEach(q => {
      if (q.type === 'mcq') {
        const hasCorrect = (q.options || []).some(o => o.is_correct);
        if (!hasCorrect) {
          const w = `Câu ${q.num}: Chưa đánh dấu đáp án đúng`;
          q.warning = w;
          newWarnings.push(w);
        } else {
          q.warning = null;
        }
      } else if (q.type === 'sa') {
        const hasAns = q.options && q.options[0] && q.options[0].text && q.options[0].text.trim().length > 0;
        if (!hasAns) {
          const w = `Câu ${q.num}: Chưa có đáp án`;
          q.warning = w;
          newWarnings.push(w);
        } else {
          q.warning = null;
        }
      }
    });
  });
  workingExamData.warnings = newWarnings;
  renderWarningsCard();
}

function renderPreviewContent() {
  const content = document.getElementById('ym-preview-content');
  if (!content || !workingExamData) return;
  content.innerHTML = '';

  (workingExamData.groups || []).forEach(group => {
    const gBlock = document.createElement('div');
    gBlock.className = 'ym-group-block';

    const gBadge = document.createElement('div');
    gBadge.className = 'ym-group-badge';
    gBadge.textContent = `Nhóm ${group.group_num}`;
    gBlock.appendChild(gBadge);

    if (group.title) {
      const gTitle = document.createElement('div');
      gTitle.className = 'ym-group-title ym-editable';
      gTitle.contentEditable = 'true';
      gTitle.spellcheck = false;
      gTitle.title = 'Bấm vào để sửa tiêu đề nhóm';
      gTitle.textContent = group.title;
      gTitle.addEventListener('paste', (e) => {
        e.preventDefault();
        const text = (e.clipboardData || window.clipboardData).getData('text/plain');
        document.execCommand('insertText', false, text);
      });
      gTitle.addEventListener('input', () => {
        group.title = gTitle.innerText.trim();
        hasUnsavedChanges = true;
        setChangeBadge(true);
      });
      gBlock.appendChild(gTitle);
    }

    (group.questions || []).forEach(q => {
      const partKey = q.type === 'mcq' ? 'p1' : (q.type === 'tf' ? 'p2' : 'p3');
      const qNumStr = String(q.num);

      const qBlock = document.createElement('div');
      qBlock.className = 'ym-question-block';

      // Question Title Row: [Câu X.] [Nội dung câu hỏi editable]
      const qTitle = document.createElement('div');
      qTitle.className = 'ym-question-title';

      const qNum = document.createElement('strong');
      qNum.className = 'ym-q-label';
      qNum.textContent = `Câu ${q.num}. `;
      qTitle.appendChild(qNum);

      const qText = document.createElement('span');
      qText.className = 'ym-question-text ym-editable';
      qText.contentEditable = 'true';
      qText.spellcheck = false;
      qText.title = 'Bấm vào để sửa nội dung câu hỏi';
      qText.setAttribute('data-placeholder', '(Nhập nội dung câu hỏi...)');
      qText.textContent = q.text || '';

      qText.addEventListener('paste', (e) => {
        e.preventDefault();
        const text = (e.clipboardData || window.clipboardData).getData('text/plain');
        document.execCommand('insertText', false, text);
      });

      qText.addEventListener('input', () => {
        const val = qText.innerText.trim();
        q.text = val;
        if (!modifiedQuestions[partKey]) modifiedQuestions[partKey] = {};
        if (!modifiedQuestions[partKey][qNumStr]) modifiedQuestions[partKey][qNumStr] = {};
        modifiedQuestions[partKey][qNumStr].text = val;
        hasUnsavedChanges = true;
        setChangeBadge(true);
      });

      qTitle.appendChild(qText);
      qBlock.appendChild(qTitle);

      const optsContainer = document.createElement('div');
      optsContainer.className = 'ym-options-container';

      if (q.type === 'mcq') {
        (q.options || []).forEach(opt => {
          const optLine = document.createElement('div');
          optLine.className = `ym-option-line ${opt.is_correct ? 'is-correct' : ''}`;
          optLine.title = 'Bấm để chọn đây là đáp án đúng (hoặc sửa chữ)';

          const optIndicator = document.createElement('span');
          optIndicator.className = `opt-indicator ${opt.is_correct ? 'opt-check' : 'opt-dash'}`;
          optIndicator.textContent = opt.is_correct ? '✓' : '-';

          const optText = document.createElement('span');
          optText.className = 'ym-option-text ym-editable';
          optText.contentEditable = 'true';
          optText.spellcheck = false;
          optText.title = 'Bấm vào để sửa nội dung phương án';
          optText.setAttribute('data-placeholder', '(Nhập nội dung phương án...)');
          optText.textContent = opt.text || '';

          const selectCorrectMCQ = () => {
            q.options.forEach(o => { o.is_correct = false; });
            opt.is_correct = true;
            modifiedAnswers.p1[qNumStr] = opt.key;
            hasUnsavedChanges = true;
            setChangeBadge(true);

            // Update UI elements in this question block directly without destroying DOM
            optsContainer.querySelectorAll('.ym-option-line').forEach((l, i) => {
              const o = q.options[i];
              l.className = `ym-option-line ${o.is_correct ? 'is-correct' : ''}`;
              const ind = l.querySelector('.opt-indicator');
              if (ind) {
                ind.className = `opt-indicator ${o.is_correct ? 'opt-check' : 'opt-dash'}`;
                ind.textContent = o.is_correct ? '✓' : '-';
              }
            });
            recalcWarnings();
          };

          optIndicator.addEventListener('click', () => {
            selectCorrectMCQ();
          });

          optLine.addEventListener('click', () => {
            selectCorrectMCQ();
          });

          optText.addEventListener('paste', (e) => {
            e.preventDefault();
            const text = (e.clipboardData || window.clipboardData).getData('text/plain');
            document.execCommand('insertText', false, text);
          });

          optText.addEventListener('input', () => {
            const val = optText.innerText.trim();
            opt.text = val;
            if (!modifiedQuestions.p1) modifiedQuestions.p1 = {};
            if (!modifiedQuestions.p1[qNumStr]) modifiedQuestions.p1[qNumStr] = {};
            if (!modifiedQuestions.p1[qNumStr].options) modifiedQuestions.p1[qNumStr].options = {};
            modifiedQuestions.p1[qNumStr].options[opt.key] = val;
            hasUnsavedChanges = true;
            setChangeBadge(true);
          });

          optLine.appendChild(optIndicator);
          optLine.appendChild(optText);
          optsContainer.appendChild(optLine);
        });
      } else if (q.type === 'tf') {
        (q.options || []).forEach(opt => {
          const optLine = document.createElement('div');
          optLine.className = `ym-option-line ${opt.is_correct ? 'is-correct' : ''}`;
          optLine.title = 'Bấm để đổi Đúng (✓) / Sai (-)';

          const optIndicator = document.createElement('span');
          optIndicator.className = `opt-indicator ${opt.is_correct ? 'opt-check' : 'opt-dash'}`;
          optIndicator.title = 'Bấm để đổi Đúng (✓) / Sai (-)';
          optIndicator.textContent = opt.is_correct ? '✓' : '-';

          const optLabel = document.createElement('span');
          optLabel.className = 'opt-key-label';
          optLabel.title = 'Bấm để đổi Đúng (✓) / Sai (-)';
          optLabel.innerHTML = `<strong>${escapeHtml(opt.key)})</strong>`;

          const optText = document.createElement('span');
          optText.className = 'ym-option-text ym-editable';
          optText.contentEditable = 'true';
          optText.spellcheck = false;
          optText.title = 'Bấm vào để sửa nội dung phát biểu';
          optText.setAttribute('data-placeholder', '(Nhập nội dung phát biểu...)');
          optText.textContent = opt.text || '';

          const toggleTF = () => {
            opt.is_correct = !opt.is_correct;
            opt.val = opt.is_correct ? 'D' : 'S';
            if (!modifiedAnswers.p2[qNumStr]) modifiedAnswers.p2[qNumStr] = {};
            modifiedAnswers.p2[qNumStr][opt.key] = opt.val;
            hasUnsavedChanges = true;
            setChangeBadge(true);

            // Update line UI directly without destroying DOM
            optLine.className = `ym-option-line ${opt.is_correct ? 'is-correct' : ''}`;
            optIndicator.className = `opt-indicator ${opt.is_correct ? 'opt-check' : 'opt-dash'}`;
            optIndicator.textContent = opt.is_correct ? '✓' : '-';
          };

          optIndicator.addEventListener('click', (e) => {
            e.stopPropagation();
            toggleTF();
          });
          optLabel.addEventListener('click', (e) => {
            e.stopPropagation();
            toggleTF();
          });
          optLine.addEventListener('click', (e) => {
            if (e.target !== optText) {
              toggleTF();
            }
          });

          optText.addEventListener('paste', (e) => {
            e.preventDefault();
            const text = (e.clipboardData || window.clipboardData).getData('text/plain');
            document.execCommand('insertText', false, text);
          });

          optText.addEventListener('input', () => {
            const val = optText.innerText.trim();
            opt.text = val;
            if (!modifiedQuestions.p2) modifiedQuestions.p2 = {};
            if (!modifiedQuestions.p2[qNumStr]) modifiedQuestions.p2[qNumStr] = {};
            if (!modifiedQuestions.p2[qNumStr].statements) modifiedQuestions.p2[qNumStr].statements = {};
            modifiedQuestions.p2[qNumStr].statements[opt.key] = val;
            hasUnsavedChanges = true;
            setChangeBadge(true);
          });

          optLine.appendChild(optIndicator);
          optLine.appendChild(optLabel);
          optLine.appendChild(optText);
          optsContainer.appendChild(optLine);
        });
      } else if (q.type === 'sa') {
        const saRow = document.createElement('div');
        saRow.className = 'ym-sa-row';
        saRow.innerHTML = `<span class="ym-sa-label">Đáp án:</span>`;

        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'ym-sa-input';
        input.value = q.options && q.options[0] ? q.options[0].text : '';
        input.placeholder = 'Nhập đáp án...';
        input.addEventListener('input', (e) => {
          const val = e.target.value.trim();
          if (!q.options || q.options.length === 0) {
            q.options = [{ key: 'A', text: val, is_correct: true }];
          } else {
            q.options[0].text = val;
            q.options[0].is_correct = true;
          }
          modifiedAnswers.p3[qNumStr] = val;
          hasUnsavedChanges = true;
          setChangeBadge(true);
          recalcWarnings();
        });

        saRow.appendChild(input);
        optsContainer.appendChild(saRow);
      }

      qBlock.appendChild(optsContainer);
      gBlock.appendChild(qBlock);
    });

    content.appendChild(gBlock);
  });
}

/**
 * Load file via POST /api/parse and display the exact YoungMix pre-output view.
 */
async function loadAndShowPreview(file) {
  const ws = document.getElementById('ym-workspace');
  const toolbar = document.getElementById('ym-preview-toolbar');
  const loading = document.getElementById('ym-preview-loading');
  const content = document.getElementById('ym-preview-content');
  if (!ws || !loading || !content) return;

  // Show workspace, show loading
  ws.style.display = 'block';
  if (toolbar) toolbar.style.display = 'flex';
  loading.style.display = 'block';
  content.style.display = 'none';

  // Populate immediate file info
  const fileNameEl = document.getElementById('ym-info-filename');
  const groupsEl   = document.getElementById('ym-info-groups');
  const qCountEl   = document.getElementById('ym-info-questions');

  if (fileNameEl) fileNameEl.textContent = file.name;
  if (groupsEl)   groupsEl.textContent   = '...';
  if (qCountEl)   qCountEl.textContent   = '...';

  // Smooth scroll down to the workspace
  ws.scrollIntoView({ behavior: 'smooth', block: 'start' });

  try {
    const formData = new FormData();
    formData.append('file', file);

    const res = await fetch(`${API_BASE}/api/parse`, {
      method: 'POST',
      body:   formData,
    });

    if (!res.ok) {
      const errData = await res.json().catch(() => ({}));
      throw new Error(errData.detail || `Lỗi máy chủ (${res.status})`);
    }

    const data = await res.json();

    // Store baseline and working states
    initialLoadedExamData = JSON.parse(JSON.stringify(data));
    workingExamData       = JSON.parse(JSON.stringify(data));
    modifiedAnswers       = { p1: {}, p2: {}, p3: {} };
    modifiedQuestions     = { p1: {}, p2: {}, p3: {} };
    hasUnsavedChanges     = false;
    setChangeBadge(false);

    // Populate Left Sidebar Cards
    if (fileNameEl) fileNameEl.textContent = data.filename || file.name;
    if (groupsEl)   groupsEl.textContent   = data.total_groups || '0';
    if (qCountEl)   qCountEl.textContent   = data.total_questions || '0';

    // Warnings Card
    renderWarningsCard();

    // Render Right Preview Pane
    renderPreviewContent();

    loading.style.display = 'none';
    content.style.display = 'block';

  } catch (err) {
    loading.innerHTML = `<div style="color:#ef4444; padding:20px; font-weight:600;">❌ Không thể phân tích đề: ${escapeHtml(err.message)}</div>`;
  }
}

/**
 * Initialize workspace actions (mix button, back button, save button, reset button).
 */
function initYmWorkspace() {
  const btnMix      = document.getElementById('btn-ym-mix');
  const btnBack     = document.getElementById('btn-back-upload');
  const btnSave     = document.getElementById('btn-save-answers');
  const btnReset    = document.getElementById('btn-reset-answers');
  const countSelect = document.getElementById('ym-opt-count');
  const startInput  = document.getElementById('ym-opt-start');
  const randomCheck = document.getElementById('ym-opt-random');

  // Synchronize countSelect with soMaDe input
  if (countSelect) {
    countSelect.addEventListener('change', () => {
      if (soMaDe) soMaDe.value = countSelect.value;
    });
  }
  if (soMaDe) {
    soMaDe.addEventListener('input', () => {
      if (countSelect) countSelect.value = soMaDe.value;
    });
  }

  // Synchronize startInput with batDauTu input
  if (startInput && batDauTu) {
    startInput.addEventListener('input', () => {
      batDauTu.value = startInput.value;
    });
    batDauTu.addEventListener('input', () => {
      startInput.value = batDauTu.value;
    });
  }

  // Handle random toggle & mutual exclusive sáng/mờ with startInput
  const startRow  = document.getElementById('row-opt-start');
  const randomRow = document.getElementById('row-opt-random');

  syncExamCodeOptions(false);

  if (randomCheck) {
    randomCheck.addEventListener('change', () => {
      syncExamCodeOptions(randomCheck.checked);
    });
  }

  if (startRow && startInput) {
    startRow.addEventListener('click', () => {
      if (randomCheck && randomCheck.checked) {
        syncExamCodeOptions(false);
        startInput.focus();
      }
    });
  }

  // Back button
  if (btnBack) {
    btnBack.addEventListener('click', () => {
      const ws = document.getElementById('ym-workspace');
      if (ws) ws.style.display = 'none';
      clearFile();
      const appEl = document.getElementById('app');
      if (appEl) appEl.scrollIntoView({ behavior: 'smooth' });
    });
  }

  // Save changes button
  if (btnSave) {
    btnSave.addEventListener('click', () => {
      hasUnsavedChanges = false;
      setChangeBadge(false);
      showToolbarToast('💾 Đã lưu thay đổi nội dung câu hỏi và đáp án thành công! Khi bấm "Trộn đề", hệ thống sẽ sử dụng các thông tin đã cập nhật này.', 'success');
    });
  }

  // Reset to original button
  if (btnReset) {
    btnReset.addEventListener('click', () => {
      if (!initialLoadedExamData) return;
      workingExamData   = JSON.parse(JSON.stringify(initialLoadedExamData));
      modifiedAnswers   = { p1: {}, p2: {}, p3: {} };
      modifiedQuestions = { p1: {}, p2: {}, p3: {} };
      hasUnsavedChanges = false;
      setChangeBadge(false);
      renderPreviewContent();
      renderWarningsCard();
      showToolbarToast('🔄 Đã phục hồi toàn bộ nội dung câu hỏi và đáp án gốc ban đầu của file Word!', 'info');
    });
  }

  // Mix button
  if (btnMix) {
    btnMix.addEventListener('click', async () => {
      if (!selectedFile) {
        alert('Vui lòng chọn file đề thi trước khi trộn.');
        return;
      }

      const countSelect = document.getElementById('ym-opt-count');
      const startInput  = document.getElementById('ym-opt-start');
      const randomCheck = document.getElementById('ym-opt-random');

      const soMaDeVal = countSelect ? (parseInt(countSelect.value, 10) || 4) : (parseInt(soMaDe?.value, 10) || 4);
      if (soMaDe) soMaDe.value = soMaDeVal;

      const isRandom = randomCheck ? randomCheck.checked : false;
      const batDauTuVal = (!isRandom && startInput) ? (parseInt(startInput.value, 10) || 101) : (parseInt(batDauTu?.value, 10) || 101);
      if (batDauTu) batDauTu.value = batDauTuVal;

      const progressWrap   = document.getElementById('ym-progress-wrap');
      const progressBar    = document.getElementById('ym-progress-bar');
      const progressStatus = document.getElementById('ym-progress-status');

      btnMix.disabled = true;
      btnMix.innerHTML = '⏳ Đang trộn đề...';

      if (progressWrap) progressWrap.style.display = 'block';
      if (progressBar)  progressBar.style.width = '35%';
      if (progressStatus) progressStatus.textContent = `Đang trộn ${soMaDeVal} mã đề...`;

      const formData = new FormData();
      formData.append('file',         selectedFile);
      formData.append('so_ma_de',     soMaDeVal);
      formData.append('num_codes',    soMaDeVal);
      formData.append('bat_dau_tu',   batDauTuVal);
      formData.append('random_codes', isRandom ? 'true' : 'false');
      formData.append('shuffle_mode', 'ca_cau_va_dap_an');

      // Append modified answers if any
      const modsPayload = {
        p1: modifiedAnswers.p1 || {},
        p2: modifiedAnswers.p2 || {},
        p3: modifiedAnswers.p3 || {}
      };
      if (Object.keys(modsPayload.p1).length || Object.keys(modsPayload.p2).length || Object.keys(modsPayload.p3).length) {
        formData.append('modified_answers', JSON.stringify(modsPayload));
      }

      // Append modified questions if any
      const modsQPayload = {
        p1: modifiedQuestions.p1 || {},
        p2: modifiedQuestions.p2 || {},
        p3: modifiedQuestions.p3 || {}
      };
      if (Object.keys(modsQPayload.p1).length || Object.keys(modsQPayload.p2).length || Object.keys(modsQPayload.p3).length) {
        formData.append('modified_questions', JSON.stringify(modsQPayload));
      }

      try {
        if (progressBar) progressBar.style.width = '70%';

        const res = await fetch(`${API_BASE}/api/process`, {
          method: 'POST',
          body:   formData,
        });

        if (!res.ok) {
          const errData = await res.json().catch(() => ({}));
          throw new Error(errData.detail || 'Lỗi khi trộn đề thi');
        }

        const data = await res.json();
        jobId = data.job_id || data.id;

        if (progressBar) progressBar.style.width = '100%';
        if (progressStatus) progressStatus.textContent = `Đã trộn xong ${soMaDeVal} mã đề!`;

        setTimeout(() => {
          if (progressWrap) progressWrap.style.display = 'none';
          btnMix.disabled = false;
          btnMix.innerHTML = '<span class="btn-ym-icon">⬇</span> Trộn đề';
          showResults(data);
          const resultsPanel = document.getElementById('results-panel');
          if (resultsPanel) resultsPanel.scrollIntoView({ behavior: 'smooth' });
        }, 350);

      } catch (err) {
        alert('Lỗi: ' + err.message);
        btnMix.disabled = false;
        btnMix.innerHTML = '<span class="btn-ym-icon">⬇</span> Trộn đề';
        if (progressWrap) progressWrap.style.display = 'none';
      }
    });
  }
}
