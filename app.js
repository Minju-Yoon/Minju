/**
 * ExcelStat Pro - Pure JavaScript Statistical Analysis Engine
 * Handles multi-file parsing, descriptive stats, cross-file comparison,
 * correlation matrices, Chart.js visualizations, and Excel export.
 */

(function () {
  'use strict';

  // =========================================================================
  // Global Application State
  // =========================================================================
  const state = {
    files: [], // Array of { id, name, size, workbook, sheets, activeSheet, rows, headers, colTypes, enabled }
    activeTab: 'tab-merged',
    theme: 'dark',
    preview: {
      dataset: '__merged__',
      search: '',
      page: 1,
      pageSize: 50,
      filteredRows: []
    },
    charts: {
      mergedDist: null,
      mergedCat: null,
      compare: null,
      corrScatter: null
    }
  };

  // Color Palettes for Chart.js
  const chartColors = [
    '#6366f1', '#10b981', '#f59e0b', '#ec4899', '#3b82f6', 
    '#8b5cf6', '#14b8a6', '#f97316', '#06b6d4', '#84cc16'
  ];

  // =========================================================================
  // DOM Elements Selection
  // =========================================================================
  const dom = {
    uploadSection: document.getElementById('upload-section'),
    dropZone: document.getElementById('drop-zone'),
    fileInput: document.getElementById('file-input'),
    btnLoadSample: document.getElementById('btn-load-sample'),
    btnEmptySample: document.getElementById('btn-empty-sample'),
    btnExportExcel: document.getElementById('btn-export-excel'),
    btnClearAll: document.getElementById('btn-clear-all'),
    btnToggleTheme: document.getElementById('btn-toggle-theme'),
    iconMoon: document.getElementById('theme-icon-moon'),
    iconSun: document.getElementById('theme-icon-sun'),

    fileManagerContainer: document.getElementById('file-manager-container'),
    fileCountBadge: document.getElementById('file-count'),
    fileCardsList: document.getElementById('file-cards-list'),
    btnSelectAllFiles: document.getElementById('btn-select-all-files'),
    btnDeselectAllFiles: document.getElementById('btn-deselect-all-files'),

    emptyState: document.getElementById('empty-state'),
    dashboardContent: document.getElementById('dashboard-content'),

    // KPI Elements
    kpiFilesCount: document.getElementById('kpi-files-count'),
    kpiFilesDetail: document.getElementById('kpi-files-detail'),
    kpiTotalRows: document.getElementById('kpi-total-rows'),
    kpiAvgRows: document.getElementById('kpi-avg-rows'),
    kpiNumericCols: document.getElementById('kpi-numeric-cols'),
    kpiTotalCols: document.getElementById('kpi-total-cols'),
    kpiMissingRate: document.getElementById('kpi-missing-rate'),
    kpiMissingCount: document.getElementById('kpi-missing-count'),

    // Tabs
    tabButtons: document.querySelectorAll('.tab-btn'),
    tabPanes: document.querySelectorAll('.tab-pane'),

    // Tab 1: Merged
    mergedColSelect: document.getElementById('merged-col-select'),
    mergedDistCanvas: document.getElementById('merged-dist-chart'),
    mergedColumnKpis: document.getElementById('merged-column-kpis'),
    statsSearchInput: document.getElementById('stats-search-input'),
    mergedStatsTbody: document.getElementById('merged-stats-tbody'),
    mergedCatColSelect: document.getElementById('merged-cat-col-select'),
    mergedCatCanvas: document.getElementById('merged-cat-chart'),
    mergedCatTbody: document.getElementById('merged-cat-tbody'),

    // Tab 2: Compare
    compareColSelect: document.getElementById('compare-col-select'),
    compareMetricSelect: document.getElementById('compare-metric-select'),
    compareChartType: document.getElementById('compare-chart-type'),
    compareCanvas: document.getElementById('compare-chart'),
    compareMatrixTbody: document.getElementById('compare-matrix-tbody'),

    // Tab 3: Individual
    individualFileSelect: document.getElementById('individual-file-select'),
    individualFileMeta: document.getElementById('individual-file-meta'),
    individualStatsTbody: document.getElementById('individual-stats-tbody'),

    // Tab 4: Correlation
    corrMatrixTable: document.getElementById('correlation-matrix-table'),
    corrXCol: document.getElementById('corr-x-col'),
    corrYCol: document.getElementById('corr-y-col'),
    corrScatterCanvas: document.getElementById('correlation-scatter-chart'),

    // Tab 5: Preview
    previewDatasetSelect: document.getElementById('preview-dataset-select'),
    previewSearchInput: document.getElementById('preview-search-input'),
    previewThead: document.getElementById('preview-thead'),
    previewTbody: document.getElementById('preview-tbody'),
    previewPageInfo: document.getElementById('preview-page-info'),
    previewPageNum: document.getElementById('preview-page-num'),
    btnPrevPage: document.getElementById('btn-prev-page'),
    btnNextPage: document.getElementById('btn-next-page')
  };

  // =========================================================================
  // Math & Statistics Helper Functions
  // =========================================================================

  /**
   * Safely parse a value into a float, stripping commas, currency symbols, and percentage signs.
   */
  function cleanNumber(val) {
    if (val === null || val === undefined || val === '') return null;
    if (typeof val === 'number') return isNaN(val) ? null : val;
    if (typeof val === 'string') {
      let trimmed = val.trim();
      if (trimmed === '' || trimmed === '-' || trimmed === 'N/A' || trimmed === 'null') return null;
      // Handle percentage
      const isPercent = trimmed.endsWith('%');
      let cleaned = trimmed.replace(/[₩$,]/g, '').trim();
      if (isPercent) cleaned = cleaned.replace('%', '').trim();
      const num = parseFloat(cleaned);
      if (isNaN(num)) return null;
      return num;
    }
    return null;
  }

  /**
   * Calculate comprehensive descriptive statistics for an array of values.
   */
  function calculateDescriptiveStats(rawValues) {
    const totalCount = rawValues.length;
    const numericValues = [];
    let missingCount = 0;

    for (let i = 0; i < totalCount; i++) {
      const num = cleanNumber(rawValues[i]);
      if (num !== null) {
        numericValues.push(num);
      } else {
        missingCount++;
      }
    }

    const n = numericValues.length;
    const missingRate = totalCount > 0 ? (missingCount / totalCount) * 100 : 0;

    if (n === 0) {
      return {
        n: 0,
        missing: missingCount,
        missingRate: missingRate,
        sum: 0,
        mean: 0,
        median: 0,
        mode: '-',
        variance: 0,
        stdDev: 0,
        min: 0,
        max: 0,
        range: 0,
        q1: 0,
        q3: 0,
        iqr: 0,
        cv: 0,
        values: []
      };
    }

    // Sort ascending for quartiles and median
    numericValues.sort((a, b) => a - b);

    const sum = numericValues.reduce((acc, curr) => acc + curr, 0);
    const mean = sum / n;

    // Median
    let median = 0;
    const mid = Math.floor(n / 2);
    if (n % 2 === 0) {
      median = (numericValues[mid - 1] + numericValues[mid]) / 2;
    } else {
      median = numericValues[mid];
    }

    // Quartiles (Linear interpolation / rank)
    function getPercentile(arr, p) {
      if (arr.length === 1) return arr[0];
      const index = (arr.length - 1) * p;
      const lower = Math.floor(index);
      const upper = Math.ceil(index);
      const weight = index - lower;
      return arr[lower] * (1 - weight) + arr[upper] * weight;
    }

    const q1 = getPercentile(numericValues, 0.25);
    const q3 = getPercentile(numericValues, 0.75);
    const iqr = q3 - q1;

    // Mode (most frequent)
    const freq = new Map();
    let maxFreq = 0;
    let modeVal = numericValues[0];
    for (let i = 0; i < n; i++) {
      const v = numericValues[i];
      const count = (freq.get(v) || 0) + 1;
      freq.set(v, count);
      if (count > maxFreq) {
        maxFreq = count;
        modeVal = v;
      }
    }
    const mode = maxFreq > 1 ? modeVal : '유일값 없음';

    // Variance & Sample Standard Deviation
    let variance = 0;
    if (n > 1) {
      const sumSquaredDiff = numericValues.reduce((acc, v) => acc + Math.pow(v - mean, 2), 0);
      variance = sumSquaredDiff / (n - 1);
    }
    const stdDev = Math.sqrt(variance);

    // Coefficient of Variation (%)
    const cv = mean !== 0 ? (stdDev / Math.abs(mean)) * 100 : 0;

    const min = numericValues[0];
    const max = numericValues[n - 1];
    const range = max - min;

    return {
      n,
      missing: missingCount,
      missingRate,
      sum,
      mean,
      median,
      mode,
      variance,
      stdDev,
      min,
      max,
      range,
      q1,
      q3,
      iqr,
      cv,
      values: numericValues
    };
  }

  /**
   * Calculate categorical frequencies.
   */
  function calculateCategoricalStats(rawValues) {
    const counts = new Map();
    let validCount = 0;
    let missingCount = 0;

    for (let i = 0; i < rawValues.length; i++) {
      const val = rawValues[i];
      if (val === null || val === undefined || String(val).trim() === '') {
        missingCount++;
      } else {
        const strVal = String(val).trim();
        counts.set(strVal, (counts.get(strVal) || 0) + 1);
        validCount++;
      }
    }

    const sortedEntries = Array.from(counts.entries())
      .map(([value, count]) => ({
        value,
        count,
        percent: validCount > 0 ? (count / validCount) * 100 : 0
      }))
      .sort((a, b) => b.count - a.count);

    return {
      total: rawValues.length,
      valid: validCount,
      missing: missingCount,
      uniqueCount: counts.size,
      frequencies: sortedEntries
    };
  }

  /**
   * Calculate Pearson Correlation Coefficient between two arrays.
   */
  function calculatePearsonCorrelation(xArr, yArr) {
    const pairs = [];
    const len = Math.min(xArr.length, yArr.length);

    for (let i = 0; i < len; i++) {
      const x = cleanNumber(xArr[i]);
      const y = cleanNumber(yArr[i]);
      if (x !== null && y !== null) {
        pairs.push({ x, y });
      }
    }

    const n = pairs.length;
    if (n < 3) return null;

    let sumX = 0;
    let sumY = 0;
    for (let i = 0; i < n; i++) {
      sumX += pairs[i].x;
      sumY += pairs[i].y;
    }
    const meanX = sumX / n;
    const meanY = sumY / n;

    let numerator = 0;
    let varX = 0;
    let varY = 0;

    for (let i = 0; i < n; i++) {
      const dx = pairs[i].x - meanX;
      const dy = pairs[i].y - meanY;
      numerator += dx * dy;
      varX += dx * dx;
      varY += dy * dy;
    }

    const denom = Math.sqrt(varX * varY);
    if (denom === 0) return 0;
    return numerator / denom;
  }

  /**
   * Format numbers with commas and reasonable decimals.
   */
  function formatNumber(num, decimals = 2) {
    if (num === null || num === undefined || isNaN(num)) return '-';
    if (typeof num === 'string') return num;
    if (Number.isInteger(num)) return num.toLocaleString('ko-KR');
    return num.toLocaleString('ko-KR', {
      minimumFractionDigits: 0,
      maximumFractionDigits: decimals
    });
  }

  // =========================================================================
  // File Parsing & Data Handling
  // =========================================================================

  /**
   * Detect column types (numeric vs categorical).
   */
  function inferColumnTypes(rows, headers) {
    const types = {};
    const sampleSize = Math.min(rows.length, 100);

    for (const h of headers) {
      let numCount = 0;
      let nonNullCount = 0;

      for (let i = 0; i < sampleSize; i++) {
        const val = rows[i][h];
        if (val !== null && val !== undefined && String(val).trim() !== '') {
          nonNullCount++;
          if (cleanNumber(val) !== null) {
            numCount++;
          }
        }
      }

      // If at least 70% of non-null values are numbers, treat as numeric
      if (nonNullCount > 0 && numCount / nonNullCount >= 0.7) {
        types[h] = 'numeric';
      } else {
        types[h] = 'categorical';
      }
    }
    return types;
  }

  /**
   * Parse workbook array buffer using SheetJS.
   */
  function parseWorkbook(buffer, fileName, fileSize) {
    try {
      const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
      const sheets = workbook.SheetNames;
      if (!sheets.length) return null;

      const activeSheet = sheets[0];
      const worksheet = workbook.Sheets[activeSheet];
      const rows = XLSX.utils.sheet_to_json(worksheet, { defval: null });

      let headers = [];
      if (rows.length > 0) {
        headers = Object.keys(rows[0]);
      } else {
        // Sheet might have headers only
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
        for (let C = range.s.c; C <= range.e.c; ++C) {
          const cell = worksheet[XLSX.utils.encode_cell({ r: range.s.r, c: C })];
          if (cell && cell.v) headers.push(String(cell.v));
        }
      }

      const colTypes = inferColumnTypes(rows, headers);

      return {
        id: 'file_' + Math.random().toString(36).substring(2, 9),
        name: fileName,
        size: fileSize,
        workbook,
        sheets,
        activeSheet,
        rows,
        headers,
        colTypes,
        enabled: true
      };
    } catch (err) {
      console.error('엑셀 파싱 에러:', err);
      alert(`"${fileName}" 파싱 중 오류가 발생했습니다: ${err.message}`);
      return null;
    }
  }

  /**
   * Process multiple File objects.
   */
  async function handleFiles(fileList) {
    if (!fileList || !fileList.length) return;

    for (let i = 0; i < fileList.length; i++) {
      const file = fileList[i];
      // Check if file already added
      const exists = state.files.some(f => f.name === file.name && f.size === file.size);
      if (exists) continue;

      const buffer = await readFileAsArrayBuffer(file);
      const parsed = parseWorkbook(buffer, file.name, file.size);
      if (parsed) {
        state.files.push(parsed);
      }
    }

    renderFileManager();
    refreshAnalysis();
  }

  function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = (e) => reject(e);
      reader.readAsArrayBuffer(file);
    });
  }

  // =========================================================================
  // Sample Data Generator
  // =========================================================================

  function generateSampleFiles() {
    const quarters = [
      { name: '2024년_1분기_지점실적.xlsx', factor: 1.0 },
      { name: '2024년_2분기_지점실적.xlsx', factor: 1.12 },
      { name: '2024년_3분기_지점실적.xlsx', factor: 1.25 }
    ];

    const branches = [
      { code: 'B01', name: '강남본점', region: '수도권', baseSales: 9500 },
      { code: 'B02', name: '서초지점', region: '수도권', baseSales: 8200 },
      { code: 'B03', name: '판교테크노점', region: '수도권', baseSales: 9100 },
      { code: 'B04', name: '여의도IFC점', region: '수도권', baseSales: 8800 },
      { code: 'B05', name: '송도센트럴점', region: '수도권', baseSales: 7400 },
      { code: 'B06', name: '해운대마린점', region: '경상권', baseSales: 8300 },
      { code: 'B07', name: '부산서면점', region: '경상권', baseSales: 7100 },
      { code: 'B08', name: '대구동성로점', region: '경상권', baseSales: 6900 },
      { code: 'B09', name: '대전둔산점', region: '충청권', baseSales: 6500 },
      { code: 'B10', name: '광주상무점', region: '전라권', baseSales: 6200 },
      { code: 'B11', name: '수원광교점', region: '수도권', baseSales: 7800 },
      { code: 'B12', name: '일산호수점', region: '수도권', baseSales: 6700 },
      { code: 'B13', name: '울산삼산점', region: '경상권', baseSales: 6400 },
      { code: 'B14', name: '천안불당점', region: '충청권', baseSales: 5900 },
      { code: 'B15', name: '전주에코점', region: '전라권', baseSales: 5600 }
    ];

    state.files = []; // clear existing

    quarters.forEach((q, qIndex) => {
      const rows = [];
      branches.forEach(b => {
        // Generate monthly records for each quarter (3 months)
        for (let m = 1; m <= 3; m++) {
          const monthNum = qIndex * 3 + m;
          const monthStr = `2024년 ${monthNum}월`;
          
          const noise = 0.92 + Math.random() * 0.16;
          const sales = Math.round(b.baseSales * q.factor * noise);
          const profitMargin = 0.18 + (Math.random() * 0.08 - 0.04);
          const profit = Math.round(sales * profitMargin);
          const newCustomers = Math.round((sales / 15) * (0.9 + Math.random() * 0.2));
          const satisfaction = +(82 + Math.random() * 15).toFixed(1);
          const discountRate = +(8 + Math.random() * 10).toFixed(1);

          rows.push({
            '지점코드': b.code,
            '지점명': b.name,
            '권역': b.region,
            '실적월': monthStr,
            '월매출액(만원)': sales,
            '영업이익(만원)': profit,
            '신규고객수(명)': newCustomers,
            '고객만족도(점)': satisfaction,
            '할인율(%)': discountRate
          });
        }
      });

      // Build SheetJS Workbook
      const ws = XLSX.utils.json_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, '지점실적');
      const arrayBuffer = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });

      const parsed = parseWorkbook(arrayBuffer, q.name, arrayBuffer.byteLength);
      if (parsed) {
        state.files.push(parsed);
      }
    });

    renderFileManager();
    refreshAnalysis();
  }

  // =========================================================================
  // Merged & Comparative Query Helpers
  // =========================================================================

  function getActiveFiles() {
    return state.files.filter(f => f.enabled);
  }

  /**
   * Merged rows of all active files with `_sourceFile` attached.
   */
  function getMergedRows() {
    const active = getActiveFiles();
    const merged = [];
    active.forEach(file => {
      file.rows.forEach(row => {
        merged.push({
          ...row,
          _sourceFile: file.name
        });
      });
    });
    return merged;
  }

  /**
   * Get all common or union column names across active files.
   */
  function getAllHeaders() {
    const active = getActiveFiles();
    const set = new Set();
    active.forEach(f => {
      f.headers.forEach(h => set.add(h));
    });
    return Array.from(set);
  }

  /**
   * Get numeric columns based on active files.
   */
  function getNumericHeaders() {
    const active = getActiveFiles();
    const candidates = getAllHeaders();
    const numericCols = [];

    candidates.forEach(col => {
      // Check if majority of files recognize this as numeric
      let numCount = 0;
      let totalAssessed = 0;
      active.forEach(f => {
        if (f.headers.includes(col)) {
          totalAssessed++;
          if (f.colTypes[col] === 'numeric') numCount++;
        }
      });
      if (totalAssessed > 0 && numCount / totalAssessed >= 0.5) {
        numericCols.push(col);
      }
    });

    return numericCols;
  }

  /**
   * Get categorical columns.
   */
  function getCategoricalHeaders() {
    const all = getAllHeaders();
    const numeric = new Set(getNumericHeaders());
    return all.filter(h => !numeric.has(h));
  }

  // =========================================================================
  // UI Renderers
  // =========================================================================

  function renderFileManager() {
    const count = state.files.length;
    dom.fileCountBadge.textContent = count;

    if (count === 0) {
      if (dom.uploadSection) dom.uploadSection.classList.remove('has-files');
      dom.fileManagerContainer.classList.add('hidden');
      dom.emptyState.classList.remove('hidden');
      dom.dashboardContent.classList.add('hidden');
      dom.btnExportExcel.disabled = true;
      return;
    }

    if (dom.uploadSection) dom.uploadSection.classList.add('has-files');
    dom.fileManagerContainer.classList.remove('hidden');
    dom.emptyState.classList.add('hidden');
    dom.dashboardContent.classList.remove('hidden');
    dom.btnExportExcel.disabled = false;

    dom.fileCardsList.innerHTML = '';

    state.files.forEach((file, index) => {
      const card = document.createElement('div');
      card.className = `file-card-item ${file.enabled ? '' : 'disabled'}`;
      
      const sizeKB = (file.size / 1024).toFixed(1);
      const rowCount = file.rows.length;
      const colCount = file.headers.length;

      // Sheet options
      const sheetOptions = file.sheets
        .map(s => `<option value="${s}" ${s === file.activeSheet ? 'selected' : ''}>${s}</option>`)
        .join('');

      card.innerHTML = `
        <div class="fc-top">
          <div class="fc-info">
            <input type="checkbox" class="file-toggle-cb" data-index="${index}" ${file.enabled ? 'checked' : ''} title="통계 계산 포함 여부 토글">
            <svg class="fc-icon" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
              <polyline points="14 2 14 8 20 8"></polyline>
            </svg>
            <span class="fc-name" title="${file.name}">${file.name}</span>
          </div>
          <div class="fc-actions">
            <button class="btn-remove-file" data-index="${index}" title="파일 삭제">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <line x1="18" y1="6" x2="6" y2="18"></line>
                <line x1="6" y1="6" x2="18" y2="18"></line>
              </svg>
            </button>
          </div>
        </div>
        <div class="fc-bottom">
          <span>${rowCount.toLocaleString()}행 × ${colCount}열 (${sizeKB} KB)</span>
          ${file.sheets.length > 1 ? `
            <select class="fc-sheet-select" data-index="${index}">
              ${sheetOptions}
            </select>
          ` : `<span>시트: ${file.activeSheet}</span>`}
        </div>
      `;

      dom.fileCardsList.appendChild(card);
    });

    // Attach File Manager event listeners
    dom.fileCardsList.querySelectorAll('.file-toggle-cb').forEach(cb => {
      cb.addEventListener('change', (e) => {
        const idx = parseInt(e.target.dataset.index);
        state.files[idx].enabled = e.target.checked;
        renderFileManager();
        refreshAnalysis();
      });
    });

    dom.fileCardsList.querySelectorAll('.btn-remove-file').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const idx = parseInt(e.currentTarget.dataset.index);
        state.files.splice(idx, 1);
        renderFileManager();
        refreshAnalysis();
      });
    });

    dom.fileCardsList.querySelectorAll('.fc-sheet-select').forEach(sel => {
      sel.addEventListener('change', (e) => {
        const idx = parseInt(e.target.dataset.index);
        const newSheet = e.target.value;
        switchFileSheet(idx, newSheet);
      });
    });
  }

  function switchFileSheet(fileIndex, sheetName) {
    const file = state.files[fileIndex];
    if (!file) return;

    try {
      const ws = file.workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
      const headers = rows.length > 0 ? Object.keys(rows[0]) : [];
      const colTypes = inferColumnTypes(rows, headers);

      file.activeSheet = sheetName;
      file.rows = rows;
      file.headers = headers;
      file.colTypes = colTypes;

      renderFileManager();
      refreshAnalysis();
    } catch (err) {
      alert(`시트 변경 실패: ${err.message}`);
    }
  }

  /**
   * Main Refresh trigger when files change or active sheet updates.
   */
  function refreshAnalysis() {
    const activeFiles = getActiveFiles();
    if (activeFiles.length === 0) {
      updateKpisEmpty();
      return;
    }

    renderKpiSummary();
    populateDropdowns();
    renderMergedAnalysis();
    renderCompareAnalysis();
    renderIndividualAnalysis();
    renderCorrelationAnalysis();
    renderDataPreview();
  }

  function updateKpisEmpty() {
    dom.kpiFilesCount.innerHTML = `0 <span class="unit">개</span>`;
    dom.kpiFilesDetail.textContent = `활성화 0 / 전체 ${state.files.length}`;
    dom.kpiTotalRows.innerHTML = `0 <span class="unit">행</span>`;
    dom.kpiAvgRows.textContent = `파일당 평균 0행`;
    dom.kpiNumericCols.innerHTML = `0 <span class="unit">개</span>`;
    dom.kpiTotalCols.textContent = `전체 컬럼 중 수치형 계산`;
    dom.kpiMissingRate.innerHTML = `0.0 <span class="unit">%</span>`;
    dom.kpiMissingCount.textContent = `총 0건 결측 셀`;
  }

  // =========================================================================
  // KPI Top Bar Calculation
  // =========================================================================

  function renderKpiSummary() {
    const active = getActiveFiles();
    const merged = getMergedRows();
    const totalRows = merged.length;
    const avgRows = active.length > 0 ? Math.round(totalRows / active.length) : 0;
    const numericCols = getNumericHeaders();
    const allCols = getAllHeaders();

    // Calculate total cells & missing cells
    let totalCells = 0;
    let missingCells = 0;

    merged.forEach(row => {
      allCols.forEach(col => {
        totalCells++;
        const val = row[col];
        if (val === null || val === undefined || String(val).trim() === '') {
          missingCells++;
        }
      });
    });

    const missingRate = totalCells > 0 ? ((missingCells / totalCells) * 100).toFixed(1) : '0.0';

    dom.kpiFilesCount.innerHTML = `${active.length} <span class="unit">개</span>`;
    dom.kpiFilesDetail.textContent = `활성화 ${active.length} / 전체 ${state.files.length}`;
    dom.kpiTotalRows.innerHTML = `${totalRows.toLocaleString()} <span class="unit">행</span>`;
    dom.kpiAvgRows.textContent = `파일당 평균 ${avgRows.toLocaleString()}행`;
    dom.kpiNumericCols.innerHTML = `${numericCols.length} <span class="unit">개</span>`;
    dom.kpiTotalCols.textContent = `전체 ${allCols.length}개 컬럼 중 수치형`;
    dom.kpiMissingRate.innerHTML = `${missingRate} <span class="unit">%</span>`;
    dom.kpiMissingCount.textContent = `총 ${missingCells.toLocaleString()}건 결측 셀`;
  }

  // =========================================================================
  // Dropdown Populators
  // =========================================================================

  function populateDropdowns() {
    const numericCols = getNumericHeaders();
    const catCols = getCategoricalHeaders();
    const activeFiles = getActiveFiles();

    // Helper to refill a select
    function refillSelect(sel, options, keepVal = true) {
      const current = sel.value;
      sel.innerHTML = options.map(opt => `<option value="${opt}">${opt}</option>`).join('');
      if (keepVal && options.includes(current)) {
        sel.value = current;
      } else if (options.length > 0) {
        sel.value = options[0];
      }
    }

    refillSelect(dom.mergedColSelect, numericCols);
    refillSelect(dom.mergedCatColSelect, catCols);
    refillSelect(dom.compareColSelect, numericCols);

    // Individual File selector
    const currentFileId = dom.individualFileSelect.value;
    dom.individualFileSelect.innerHTML = activeFiles
      .map(f => `<option value="${f.id}">${f.name} (${f.rows.length.toLocaleString()}행)</option>`)
      .join('');
    if (activeFiles.some(f => f.id === currentFileId)) {
      dom.individualFileSelect.value = currentFileId;
    }

    // Correlation X & Y
    refillSelect(dom.corrXCol, numericCols);
    refillSelect(dom.corrYCol, numericCols, false);
    if (numericCols.length > 1) {
      dom.corrYCol.value = numericCols[1];
    }

    // Preview Dataset select
    const currentPreviewDs = dom.previewDatasetSelect.value;
    let previewOptions = `<option value="__merged__">통합 데이터 전체 (${getMergedRows().length.toLocaleString()}행)</option>`;
    activeFiles.forEach(f => {
      previewOptions += `<option value="${f.id}">${f.name} (${f.rows.length.toLocaleString()}행)</option>`;
    });
    dom.previewDatasetSelect.innerHTML = previewOptions;
    if (currentPreviewDs) dom.previewDatasetSelect.value = currentPreviewDs;
  }

  // =========================================================================
  // TAB 1: 통합 기술 통계 (Merged Descriptive Stats)
  // =========================================================================

  function renderMergedAnalysis() {
    const mergedRows = getMergedRows();
    const numericCols = getNumericHeaders();
    const catCols = getCategoricalHeaders();

    // 1. Render Table
    const filterQuery = dom.statsSearchInput.value.trim().toLowerCase();
    dom.mergedStatsTbody.innerHTML = '';

    const statsCache = {};

    numericCols.forEach(col => {
      if (filterQuery && !col.toLowerCase().includes(filterQuery)) return;

      const rawValues = mergedRows.map(r => r[col]);
      const stats = calculateDescriptiveStats(rawValues);
      statsCache[col] = stats;

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td class="text-left">
          <strong>${col}</strong>
          <span class="col-badge">수치형</span>
        </td>
        <td>${formatNumber(stats.n, 0)}</td>
        <td>${stats.missing.toLocaleString()} (${stats.missingRate.toFixed(1)}%)</td>
        <td>${formatNumber(stats.sum, 1)}</td>
        <td><strong>${formatNumber(stats.mean, 2)}</strong></td>
        <td>${formatNumber(stats.median, 2)}</td>
        <td>${typeof stats.mode === 'number' ? formatNumber(stats.mode, 2) : stats.mode}</td>
        <td>${formatNumber(stats.stdDev, 2)}</td>
        <td>${formatNumber(stats.min, 2)}</td>
        <td>${formatNumber(stats.q1, 2)}</td>
        <td>${formatNumber(stats.q3, 2)}</td>
        <td>${formatNumber(stats.max, 2)}</td>
        <td>${formatNumber(stats.iqr, 2)}</td>
        <td>${stats.cv ? stats.cv.toFixed(1) + '%' : '-'}</td>
      `;
      dom.mergedStatsTbody.appendChild(tr);
    });

    // 2. Render Selected Column Histogram Chart
    const selectedCol = dom.mergedColSelect.value || (numericCols.length > 0 ? numericCols[0] : null);
    if (selectedCol && statsCache[selectedCol]) {
      renderHistogramChart(selectedCol, statsCache[selectedCol]);
    } else if (selectedCol) {
      const rawValues = mergedRows.map(r => r[selectedCol]);
      const stats = calculateDescriptiveStats(rawValues);
      renderHistogramChart(selectedCol, stats);
    }

    // 3. Render Categorical Distribution
    renderMergedCategorical();
  }

  function renderHistogramChart(colName, stats) {
    // Render rich 3x3 sub-kpi matrix beside the chart
    dom.mergedColumnKpis.innerHTML = `
      <div class="sub-kpi-item">
        <span class="sub-kpi-title">평균 (Mean)</span>
        <span class="sub-kpi-value">${formatNumber(stats.mean, 2)}</span>
      </div>
      <div class="sub-kpi-item">
        <span class="sub-kpi-title">중앙값 (Median)</span>
        <span class="sub-kpi-value">${formatNumber(stats.median, 2)}</span>
      </div>
      <div class="sub-kpi-item">
        <span class="sub-kpi-title">최빈값 (Mode)</span>
        <span class="sub-kpi-value">${typeof stats.mode === 'number' ? formatNumber(stats.mode, 2) : stats.mode}</span>
      </div>
      <div class="sub-kpi-item">
        <span class="sub-kpi-title">표준편차 (Std Dev)</span>
        <span class="sub-kpi-value">${formatNumber(stats.stdDev, 2)}</span>
      </div>
      <div class="sub-kpi-item">
        <span class="sub-kpi-title">표본 분산 (Variance)</span>
        <span class="sub-kpi-value">${formatNumber(stats.variance, 2)}</span>
      </div>
      <div class="sub-kpi-item">
        <span class="sub-kpi-title">변동계수 (CV)</span>
        <span class="sub-kpi-value">${stats.cv ? stats.cv.toFixed(1) + '%' : '-'}</span>
      </div>
      <div class="sub-kpi-item">
        <span class="sub-kpi-title">사분위범위 (IQR)</span>
        <span class="sub-kpi-value">${formatNumber(stats.iqr, 2)}</span>
      </div>
      <div class="sub-kpi-item">
        <span class="sub-kpi-title">Q1 / Q3 (사분위수)</span>
        <span class="sub-kpi-value">${formatNumber(stats.q1, 1)} / ${formatNumber(stats.q3, 1)}</span>
      </div>
      <div class="sub-kpi-item">
        <span class="sub-kpi-title">최솟값 ~ 최댓값 (범위)</span>
        <span class="sub-kpi-value">${formatNumber(stats.min, 1)} ~ ${formatNumber(stats.max, 1)} (${formatNumber(stats.range, 1)})</span>
      </div>
    `;

    if (stats.n === 0) return;

    // Create 10-15 Bins for Histogram
    const numBins = Math.min(15, Math.max(7, Math.round(1 + 3.322 * Math.log10(stats.n))));
    const binWidth = stats.range > 0 ? stats.range / numBins : 1;
    const bins = Array(numBins).fill(0);
    const binLabels = [];

    for (let i = 0; i < numBins; i++) {
      const start = stats.min + i * binWidth;
      const end = start + binWidth;
      binLabels.push(`${formatNumber(start, 1)} ~ ${formatNumber(end, 1)}`);
    }

    stats.values.forEach(val => {
      let binIdx = Math.floor((val - stats.min) / binWidth);
      if (binIdx >= numBins) binIdx = numBins - 1;
      if (binIdx < 0) binIdx = 0;
      bins[binIdx]++;
    });

    if (state.charts.mergedDist) {
      state.charts.mergedDist.destroy();
    }

    const ctx = dom.mergedDistCanvas.getContext('2d');
    const isDark = document.body.classList.contains('dark-theme');

    state.charts.mergedDist = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: binLabels,
        datasets: [{
          label: `${colName} 빈도수 (Frequency)`,
          data: bins,
          backgroundColor: 'rgba(99, 102, 241, 0.65)',
          borderColor: '#6366f1',
          borderWidth: 1.5,
          borderRadius: 4
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            labels: { color: isDark ? '#e5e7eb' : '#374151' }
          },
          tooltip: {
            callbacks: {
              afterLabel: function (context) {
                const count = context.parsed.y;
                const pct = ((count / stats.n) * 100).toFixed(1);
                return `비율: ${pct}% (전체 ${stats.n}건 중)`;
              }
            }
          }
        },
        scales: {
          x: {
            grid: { color: isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)' },
            ticks: {
              color: isDark ? '#9ca3af' : '#6b7280',
              font: { size: 11 },
              maxRotation: 45
            }
          },
          y: {
            grid: { color: isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)' },
            ticks: {
              color: isDark ? '#9ca3af' : '#6b7280',
              precision: 0
            },
            title: {
              display: true,
              text: '데이터 건수',
              color: isDark ? '#9ca3af' : '#6b7280'
            }
          }
        }
      }
    });
  }

  function renderMergedCategorical() {
    const mergedRows = getMergedRows();
    const catCols = getCategoricalHeaders();
    const selectedCatCol = dom.mergedCatColSelect.value || (catCols.length > 0 ? catCols[0] : null);

    if (!selectedCatCol) {
      dom.mergedCatTbody.innerHTML = '<tr><td colspan="3" style="text-align:center;color:var(--text-muted);">범주형 컬럼이 없습니다.</td></tr>';
      if (state.charts.mergedCat) state.charts.mergedCat.destroy();
      return;
    }

    const rawValues = mergedRows.map(r => r[selectedCatCol]);
    const catStats = calculateCategoricalStats(rawValues);

    // Populate Table
    dom.mergedCatTbody.innerHTML = '';
    catStats.frequencies.forEach(item => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td class="text-left">${item.value}</td>
        <td>${item.count.toLocaleString()}</td>
        <td><strong>${item.percent.toFixed(1)}%</strong></td>
      `;
      dom.mergedCatTbody.appendChild(tr);
    });

    // Chart.js Donut
    if (state.charts.mergedCat) state.charts.mergedCat.destroy();

    const isDark = document.body.classList.contains('dark-theme');
    const topCategories = catStats.frequencies.slice(0, 8);
    let chartLabels = topCategories.map(c => c.value);
    let chartData = topCategories.map(c => c.count);

    if (catStats.frequencies.length > 8) {
      const othersCount = catStats.frequencies.slice(8).reduce((acc, c) => acc + c.count, 0);
      chartLabels.push('기타');
      chartData.push(othersCount);
    }

    const ctx = dom.mergedCatCanvas.getContext('2d');
    state.charts.mergedCat = new Chart(ctx, {
      type: 'doughnut',
      data: {
        labels: chartLabels,
        datasets: [{
          data: chartData,
          backgroundColor: chartColors,
          borderWidth: 2,
          borderColor: isDark ? '#111827' : '#ffffff'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'right',
            labels: { color: isDark ? '#e5e7eb' : '#374151', font: { size: 12 } }
          }
        }
      }
    });
  }

  // =========================================================================
  // TAB 2: 파일 간 비교 분석 (Comparative Analysis)
  // =========================================================================

  function renderCompareAnalysis() {
    const active = getActiveFiles();
    const colName = dom.compareColSelect.value;
    const metric = dom.compareMetricSelect.value;
    const chartType = dom.compareChartType.value;

    if (!colName || active.length === 0) return;

    // Compute stats for each file
    const fileStats = active.map(file => {
      const vals = file.rows.map(r => r[colName]);
      const stats = calculateDescriptiveStats(vals);
      return {
        fileName: file.name,
        stats
      };
    });

    // Calculate overall average of selected metric across files for deviation reference
    const totalMean = fileStats.reduce((sum, f) => sum + (f.stats.mean || 0), 0) / fileStats.length;

    // Populate Matrix Table
    dom.compareMatrixTbody.innerHTML = '';
    fileStats.forEach(f => {
      const s = f.stats;
      const deviation = totalMean !== 0 ? ((s.mean - totalMean) / totalMean) * 100 : 0;
      const devColor = deviation > 0 ? 'var(--success)' : deviation < 0 ? 'var(--danger)' : 'inherit';
      const devSign = deviation > 0 ? '+' : '';

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td class="text-left"><strong>${f.fileName}</strong></td>
        <td>${s.n.toLocaleString()}</td>
        <td>${formatNumber(s.sum, 1)}</td>
        <td><strong>${formatNumber(s.mean, 2)}</strong></td>
        <td>${formatNumber(s.median, 2)}</td>
        <td>${formatNumber(s.stdDev, 2)}</td>
        <td>${formatNumber(s.min, 2)}</td>
        <td>${formatNumber(s.max, 2)}</td>
        <td style="color: ${devColor}; font-weight: 700;">${devSign}${deviation.toFixed(1)}%</td>
      `;
      dom.compareMatrixTbody.appendChild(tr);
    });

    // Chart.js Multi-file Comparison
    if (state.charts.compare) state.charts.compare.destroy();

    const isDark = document.body.classList.contains('dark-theme');
    const labels = fileStats.map(f => f.fileName);
    const metricValues = fileStats.map(f => {
      switch (metric) {
        case 'mean': return f.stats.mean;
        case 'sum': return f.stats.sum;
        case 'median': return f.stats.median;
        case 'std': return f.stats.stdDev;
        case 'min': return f.stats.min;
        case 'max': return f.stats.max;
        case 'count': return f.stats.n;
        default: return f.stats.mean;
      }
    });

    const metricNames = {
      mean: '평균', sum: '합계', median: '중앙값', std: '표준편차', min: '최솟값', max: '최댓값', count: '데이터 건수'
    };

    const ctx = dom.compareCanvas.getContext('2d');
    state.charts.compare = new Chart(ctx, {
      type: chartType,
      data: {
        labels: labels,
        datasets: [{
          label: `${colName} - ${metricNames[metric]} 비교`,
          data: metricValues,
          backgroundColor: chartType === 'bar' ? chartColors.slice(0, labels.length) : 'rgba(99, 102, 241, 0.2)',
          borderColor: '#6366f1',
          borderWidth: 2,
          fill: chartType === 'line',
          tension: 0.25,
          borderRadius: 6
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            labels: { color: isDark ? '#e5e7eb' : '#374151' }
          }
        },
        scales: {
          x: {
            grid: { color: isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)' },
            ticks: { color: isDark ? '#9ca3af' : '#6b7280' }
          },
          y: {
            grid: { color: isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)' },
            ticks: { color: isDark ? '#9ca3af' : '#6b7280' }
          }
        }
      }
    });
  }

  // =========================================================================
  // TAB 3: 개별 파일 상세 분석 (Individual File Deep-Dive)
  // =========================================================================

  function renderIndividualAnalysis() {
    const fileId = dom.individualFileSelect.value;
    const file = state.files.find(f => f.id === fileId);

    if (!file) {
      dom.individualStatsTbody.innerHTML = '<tr><td colspan="10" style="text-align:center;color:var(--text-muted);">선택된 파일이 없습니다.</td></tr>';
      dom.individualFileMeta.innerHTML = '';
      return;
    }

    // Render Meta Banner
    dom.individualFileMeta.innerHTML = `
      <div>파일명: <strong>${file.name}</strong></div>
      <div>활성 시트: <strong>${file.activeSheet}</strong></div>
      <div>데이터 규모: <strong>${file.rows.length.toLocaleString()}행 × ${file.headers.length}개 컬럼</strong></div>
      <div>파일 크기: <strong>${(file.size / 1024).toFixed(1)} KB</strong></div>
    `;

    // Populate Table
    dom.individualStatsTbody.innerHTML = '';

    file.headers.forEach(header => {
      const type = file.colTypes[header];
      const rawVals = file.rows.map(r => r[header]);

      const tr = document.createElement('tr');
      if (type === 'numeric') {
        const s = calculateDescriptiveStats(rawVals);
        tr.innerHTML = `
          <td class="text-left"><strong>${header}</strong></td>
          <td><span class="col-badge">수치형</span></td>
          <td>${s.n.toLocaleString()}</td>
          <td>${s.missing.toLocaleString()} (${s.missingRate.toFixed(1)}%)</td>
          <td><strong>${formatNumber(s.mean, 2)}</strong></td>
          <td>${formatNumber(s.stdDev, 2)}</td>
          <td>${formatNumber(s.median, 2)}</td>
          <td>${formatNumber(s.min, 2)}</td>
          <td>${formatNumber(s.max, 2)}</td>
          <td>${formatNumber(s.iqr, 2)}</td>
        `;
      } else {
        const cat = calculateCategoricalStats(rawVals);
        const topItem = cat.frequencies.length > 0 ? cat.frequencies[0] : null;
        const topText = topItem ? `${topItem.value} (${topItem.percent.toFixed(1)}%)` : '-';
        tr.innerHTML = `
          <td class="text-left"><strong>${header}</strong></td>
          <td><span class="col-badge" style="background: rgba(16, 185, 129, 0.15); color: #10b981;">범주형</span></td>
          <td>${cat.valid.toLocaleString()}</td>
          <td>${cat.missing.toLocaleString()}</td>
          <td colspan="6" style="text-align:left; color: var(--text-secondary); font-family: var(--font-sans);">
            고유값: <strong>${cat.uniqueCount}개</strong> / 최다 빈도: <strong>${topText}</strong>
          </td>
        `;
      }
      dom.individualStatsTbody.appendChild(tr);
    });
  }

  // =========================================================================
  // TAB 4: 변수 간 상관관계 분석 (Correlation Matrix & Scatter)
  // =========================================================================

  function renderCorrelationAnalysis() {
    const mergedRows = getMergedRows();
    const numericCols = getNumericHeaders();

    if (numericCols.length < 2) {
      dom.corrMatrixTable.innerHTML = '<tr><td style="padding:2rem;color:var(--text-muted);">상관관계를 계산하려면 최소 2개 이상의 수치형 변수가 필요합니다.</td></tr>';
      if (state.charts.corrScatter) state.charts.corrScatter.destroy();
      return;
    }

    // Build Correlation Matrix
    let tableHtml = '<thead><tr><th class="text-left">변수명</th>';
    numericCols.forEach(col => {
      tableHtml += `<th>${col}</th>`;
    });
    tableHtml += '</tr></thead><tbody>';

    numericCols.forEach(col1 => {
      tableHtml += `<tr><td class="text-left"><strong>${col1}</strong></td>`;
      const vals1 = mergedRows.map(r => r[col1]);

      numericCols.forEach(col2 => {
        if (col1 === col2) {
          tableHtml += `<td class="corr-cell" style="background: rgba(99, 102, 241, 0.2); font-weight:700;">1.00</td>`;
        } else {
          const vals2 = mergedRows.map(r => r[col2]);
          const r = calculatePearsonCorrelation(vals1, vals2);
          if (r === null) {
            tableHtml += `<td>-</td>`;
          } else {
            let bgColor = 'transparent';
            if (r > 0) {
              bgColor = `rgba(16, 185, 129, ${Math.min(0.85, Math.max(0.12, r))})`;
            } else {
              bgColor = `rgba(239, 68, 68, ${Math.min(0.85, Math.max(0.12, Math.abs(r)))})`;
            }
            tableHtml += `<td class="corr-cell" data-x="${col1}" data-y="${col2}" style="background: ${bgColor};" title="${col1} vs ${col2} 상관계수: ${r.toFixed(3)}">${r.toFixed(2)}</td>`;
          }
        }
      });
      tableHtml += '</tr>';
    });
    tableHtml += '</tbody>';

    dom.corrMatrixTable.innerHTML = tableHtml;

    // Click on cell sets X and Y for scatter
    dom.corrMatrixTable.querySelectorAll('.corr-cell').forEach(cell => {
      cell.addEventListener('click', () => {
        const x = cell.dataset.x;
        const y = cell.dataset.y;
        if (x && y) {
          dom.corrXCol.value = x;
          dom.corrYCol.value = y;
          renderScatterPlot();
        }
      });
    });

    renderScatterPlot();
  }

  function renderScatterPlot() {
    const xCol = dom.corrXCol.value;
    const yCol = dom.corrYCol.value;
    const mergedRows = getMergedRows();

    if (!xCol || !yCol) return;

    const dataPoints = [];
    mergedRows.forEach(r => {
      const x = cleanNumber(r[xCol]);
      const y = cleanNumber(r[yCol]);
      if (x !== null && y !== null) {
        dataPoints.push({ x, y });
      }
    });

    if (state.charts.corrScatter) state.charts.corrScatter.destroy();

    const isDark = document.body.classList.contains('dark-theme');
    const ctx = dom.corrScatterCanvas.getContext('2d');

    state.charts.corrScatter = new Chart(ctx, {
      type: 'scatter',
      data: {
        datasets: [{
          label: `${xCol} vs ${yCol}`,
          data: dataPoints,
          backgroundColor: 'rgba(99, 102, 241, 0.65)',
          borderColor: '#6366f1',
          pointRadius: 4.5,
          pointHoverRadius: 7
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            labels: { color: isDark ? '#e5e7eb' : '#374151' }
          },
          tooltip: {
            callbacks: {
              label: function (ctx) {
                return `(${xCol}: ${formatNumber(ctx.parsed.x, 2)}, ${yCol}: ${formatNumber(ctx.parsed.y, 2)})`;
              }
            }
          }
        },
        scales: {
          x: {
            title: { display: true, text: xCol, color: isDark ? '#9ca3af' : '#6b7280' },
            grid: { color: isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)' },
            ticks: { color: isDark ? '#9ca3af' : '#6b7280' }
          },
          y: {
            title: { display: true, text: yCol, color: isDark ? '#9ca3af' : '#6b7280' },
            grid: { color: isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)' },
            ticks: { color: isDark ? '#9ca3af' : '#6b7280' }
          }
        }
      }
    });
  }

  // =========================================================================
  // TAB 5: 데이터 미리보기 (Data Preview)
  // =========================================================================

  function renderDataPreview() {
    const dsKey = dom.previewDatasetSelect.value;
    let targetRows = [];

    if (dsKey === '__merged__') {
      targetRows = getMergedRows();
    } else {
      const file = state.files.find(f => f.id === dsKey);
      targetRows = file ? file.rows : [];
    }

    const searchQuery = dom.previewSearchInput.value.trim().toLowerCase();
    if (searchQuery) {
      targetRows = targetRows.filter(row => {
        return Object.values(row).some(v => v !== null && String(v).toLowerCase().includes(searchQuery));
      });
    }

    state.preview.filteredRows = targetRows;

    const total = targetRows.length;
    const pageSize = state.preview.pageSize;
    const maxPage = Math.max(1, Math.ceil(total / pageSize));

    if (state.preview.page > maxPage) state.preview.page = maxPage;
    if (state.preview.page < 1) state.preview.page = 1;

    const page = state.preview.page;
    const startIdx = (page - 1) * pageSize;
    const endIdx = Math.min(startIdx + pageSize, total);
    const pagedRows = targetRows.slice(startIdx, endIdx);

    // Update Pagination labels
    dom.previewPageInfo.textContent = `${total === 0 ? 0 : startIdx + 1} - ${endIdx} / 전체 ${total.toLocaleString()}건`;
    dom.previewPageNum.textContent = `${page} / ${maxPage}`;
    dom.btnPrevPage.disabled = page <= 1;
    dom.btnNextPage.disabled = page >= maxPage;

    // Render Headers
    const headers = targetRows.length > 0 ? Object.keys(targetRows[0]) : getAllHeaders();
    dom.previewThead.innerHTML = `<tr>${headers.map(h => `<th class="text-left">${h}</th>`).join('')}</tr>`;

    // Render Rows
    if (pagedRows.length === 0) {
      dom.previewTbody.innerHTML = `<tr><td colspan="${headers.length}" style="text-align:center;padding:2rem;color:var(--text-muted);">표시할 데이터가 없습니다.</td></tr>`;
      return;
    }

    let rowsHtml = '';
    pagedRows.forEach(row => {
      rowsHtml += '<tr>';
      headers.forEach(h => {
        const val = row[h];
        const isNum = typeof val === 'number';
        rowsHtml += `<td class="${isNum ? '' : 'text-left'}">${val === null || val === undefined ? '-' : val}</td>`;
      });
      rowsHtml += '</tr>';
    });
    dom.previewTbody.innerHTML = rowsHtml;
  }

  // =========================================================================
  // Excel Export Feature
  // =========================================================================

  function exportStatisticsToExcel() {
    const active = getActiveFiles();
    if (active.length === 0) {
      alert('내보낼 데이터가 없습니다.');
      return;
    }

    const mergedRows = getMergedRows();
    const numericCols = getNumericHeaders();

    const wb = XLSX.utils.book_new();

    // 1. Sheet 1: Merged Descriptive Statistics
    const statsSheetData = [
      ['컬럼명', '유효 표본수(N)', '결측치수', '결측률(%)', '합계', '평균', '중앙값', '최빈값', '표준편차', '최솟값', 'Q1(25%)', 'Q3(75%)', '최댓값', 'IQR', '변동계수(%)']
    ];

    numericCols.forEach(col => {
      const vals = mergedRows.map(r => r[col]);
      const s = calculateDescriptiveStats(vals);
      statsSheetData.push([
        col,
        s.n,
        s.missing,
        +s.missingRate.toFixed(2),
        +s.sum.toFixed(2),
        +s.mean.toFixed(2),
        +s.median.toFixed(2),
        typeof s.mode === 'number' ? +s.mode.toFixed(2) : String(s.mode),
        +s.stdDev.toFixed(2),
        +s.min.toFixed(2),
        +s.q1.toFixed(2),
        +s.q3.toFixed(2),
        +s.max.toFixed(2),
        +s.iqr.toFixed(2),
        +s.cv.toFixed(2)
      ]);
    });

    const wsStats = XLSX.utils.aoa_to_sheet(statsSheetData);
    XLSX.utils.book_append_sheet(wb, wsStats, '통합_기술통계');

    // 2. Sheet 2: Cross-file Comparison for Numeric Columns
    const compareSheetData = [
      ['컬럼명', '파일명', '데이터건수', '평균', '중앙값', '표준편차', '최솟값', '최댓값', '합계']
    ];

    numericCols.forEach(col => {
      active.forEach(file => {
        const vals = file.rows.map(r => r[col]);
        const s = calculateDescriptiveStats(vals);
        compareSheetData.push([
          col,
          file.name,
          s.n,
          +s.mean.toFixed(2),
          +s.median.toFixed(2),
          +s.stdDev.toFixed(2),
          +s.min.toFixed(2),
          +s.max.toFixed(2),
          +s.sum.toFixed(2)
        ]);
      });
    });

    const wsCompare = XLSX.utils.aoa_to_sheet(compareSheetData);
    XLSX.utils.book_append_sheet(wb, wsCompare, '파일간_비교분석');

    // 3. Sheet 3: Pearson Correlation Matrix
    if (numericCols.length >= 2) {
      const corrData = [['상관계수(r)', ...numericCols]];
      numericCols.forEach(c1 => {
        const row = [c1];
        const v1 = mergedRows.map(r => r[c1]);
        numericCols.forEach(c2 => {
          if (c1 === c2) {
            row.push(1.0);
          } else {
            const v2 = mergedRows.map(r => r[c2]);
            const r = calculatePearsonCorrelation(v1, v2);
            row.push(r !== null ? +r.toFixed(3) : '-');
          }
        });
        corrData.push(row);
      });

      const wsCorr = XLSX.utils.aoa_to_sheet(corrData);
      XLSX.utils.book_append_sheet(wb, wsCorr, '피어슨_상관행렬');
    }

    // Save File
    const now = new Date();
    const dateStr = now.toISOString().slice(0, 10).replace(/-/g, '');
    const timeStr = now.toTimeString().slice(0, 8).replace(/:/g, '');
    const exportFileName = `ExcelStat_분석결과_${dateStr}_${timeStr}.xlsx`;

    XLSX.writeFile(wb, exportFileName);
  }

  // =========================================================================
  // Event Listeners Registration
  // =========================================================================

  function setupEventListeners() {
    // Drag & Drop
    const dz = dom.dropZone;
    ['dragenter', 'dragover'].forEach(name => {
      dz.addEventListener(name, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dz.classList.add('dragover');
      });
    });

    ['dragleave', 'drop'].forEach(name => {
      dz.addEventListener(name, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dz.classList.remove('dragover');
      });
    });

    dz.addEventListener('drop', (e) => {
      const files = e.dataTransfer.files;
      if (files && files.length) {
        handleFiles(files);
      }
    });

    dz.addEventListener('click', () => {
      dom.fileInput.click();
    });

    dom.fileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files.length) {
        handleFiles(e.target.files);
        dom.fileInput.value = ''; // reset so same files can be re-selected if removed
      }
    });

    // Sample Data Buttons
    dom.btnLoadSample.addEventListener('click', generateSampleFiles);
    dom.btnEmptySample.addEventListener('click', generateSampleFiles);

    // Export Button
    dom.btnExportExcel.addEventListener('click', exportStatisticsToExcel);

    // Clear All
    dom.btnClearAll.addEventListener('click', () => {
      if (!state.files.length) return;
      if (confirm('모든 업로드 파일과 분석 결과를 초기화하시겠습니까?')) {
        state.files = [];
        renderFileManager();
        refreshAnalysis();
      }
    });

    // Theme Toggle
    dom.btnToggleTheme.addEventListener('click', () => {
      const body = document.body;
      if (body.classList.contains('dark-theme')) {
        body.classList.remove('dark-theme');
        body.classList.add('light-theme');
        dom.iconMoon.classList.add('hidden');
        dom.iconSun.classList.remove('hidden');
        state.theme = 'light';
      } else {
        body.classList.remove('light-theme');
        body.classList.add('dark-theme');
        dom.iconSun.classList.add('hidden');
        dom.iconMoon.classList.remove('hidden');
        state.theme = 'dark';
      }
      // Re-render active charts to update axis colors
      refreshAnalysis();
    });

    // File Manager bulk select/deselect
    dom.btnSelectAllFiles.addEventListener('click', () => {
      state.files.forEach(f => (f.enabled = true));
      renderFileManager();
      refreshAnalysis();
    });

    dom.btnDeselectAllFiles.addEventListener('click', () => {
      state.files.forEach(f => (f.enabled = false));
      renderFileManager();
      refreshAnalysis();
    });

    // Navigation Tabs
    dom.tabButtons.forEach(btn => {
      btn.addEventListener('click', () => {
        const targetTab = btn.dataset.tab;
        state.activeTab = targetTab;

        dom.tabButtons.forEach(b => b.classList.remove('active'));
        dom.tabPanes.forEach(p => p.classList.remove('active'));

        btn.classList.add('active');
        const targetPane = document.getElementById(targetTab);
        if (targetPane) targetPane.classList.add('active');
      });
    });

    // Tab 1 Controls
    dom.mergedColSelect.addEventListener('change', () => {
      const col = dom.mergedColSelect.value;
      const mergedRows = getMergedRows();
      const rawVals = mergedRows.map(r => r[col]);
      const stats = calculateDescriptiveStats(rawVals);
      renderHistogramChart(col, stats);
    });

    dom.statsSearchInput.addEventListener('input', () => {
      renderMergedAnalysis();
    });

    dom.mergedCatColSelect.addEventListener('change', () => {
      renderMergedCategorical();
    });

    // Tab 2 Controls
    dom.compareColSelect.addEventListener('change', renderCompareAnalysis);
    dom.compareMetricSelect.addEventListener('change', renderCompareAnalysis);
    dom.compareChartType.addEventListener('change', renderCompareAnalysis);

    // Tab 3 Controls
    dom.individualFileSelect.addEventListener('change', renderIndividualAnalysis);

    // Tab 4 Controls
    dom.corrXCol.addEventListener('change', renderScatterPlot);
    dom.corrYCol.addEventListener('change', renderScatterPlot);

    // Tab 5 Controls
    dom.previewDatasetSelect.addEventListener('change', () => {
      state.preview.page = 1;
      renderDataPreview();
    });

    dom.previewSearchInput.addEventListener('input', () => {
      state.preview.page = 1;
      renderDataPreview();
    });

    dom.btnPrevPage.addEventListener('click', () => {
      if (state.preview.page > 1) {
        state.preview.page--;
        renderDataPreview();
      }
    });

    dom.btnNextPage.addEventListener('click', () => {
      state.preview.page++;
      renderDataPreview();
    });
  }

  // =========================================================================
  // Application Initialization
  // =========================================================================
  function init() {
    setupEventListeners();
    renderFileManager();
  }

  // Start when DOM is ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

})();
