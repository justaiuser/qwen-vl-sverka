// Frontend logic for Accounting Reconciliation Dashboard
let currentReport = null;
let activeFilter = 'all';

document.addEventListener('DOMContentLoaded', () => {
  setupDropzone('card-1c', 'file-1c', 'name-1c', 'dropzone-1c');
  setupDropzone('card-cp', 'file-cp', 'name-cp', 'dropzone-cp');

  document.getElementById('btn-reconcile').addEventListener('click', runReconciliation);
  document.getElementById('btn-demo').addEventListener('click', runDemo);
  document.getElementById('search-input').addEventListener('input', renderTable);

  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      activeFilter = btn.dataset.filter;
      renderTable();
    });
  });
});

function setupDropzone(cardId, inputId, nameId, zoneId) {
  const card = document.getElementById(cardId);
  const input = document.getElementById(inputId);
  const nameEl = document.getElementById(nameId);
  const zone = document.getElementById(zoneId);

  zone.addEventListener('click', () => input.click());

  input.addEventListener('change', () => {
    if (input.files.length > 0) {
      nameEl.textContent = input.files[0].name;
      nameEl.style.color = '#1e40af';
    }
  });

  ['dragenter', 'dragover'].forEach(eventName => {
    zone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      zone.classList.add('dragover');
    });
  });

  ['dragleave', 'drop'].forEach(eventName => {
    zone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      zone.classList.remove('dragover');
    });
  });

  zone.addEventListener('drop', (e) => {
    if (e.dataTransfer.files.length > 0) {
      input.files = e.dataTransfer.files;
      nameEl.textContent = input.files[0].name;
      nameEl.style.color = '#1e40af';
    }
  });
}

async function runDemo() {
  showLoader("Запуск демонстрационного примера...", "Загрузка подготовленных образцов (1С Excel + Скан акта с расхождениями)...");
  try {
    const res = await fetch('/api/reconcile-demo', { credentials: 'omit' });
    const data = await res.json();
    if (data.error) {
      alert("Ошибка: " + data.error);
    } else {
      document.getElementById('name-1c').textContent = "act_1c_sample.xlsx";
      document.getElementById('name-cp').textContent = "act_counterparty_sample.pdf";
      displayResults(data);
    }
  } catch (err) {
    alert("Ошибка соединения с сервером: " + err.message);
  } finally {
    hideLoader();
  }
}

async function runReconciliation() {
  const file1c = document.getElementById('file-1c').files[0];
  const fileCp = document.getElementById('file-cp').files[0];

  if (!file1c || !fileCp) {
    alert("Пожалуйста, выберите оба файла: файл 1С (.xlsx) и скан/PDF контрагента!");
    return;
  }

  const model = document.getElementById('model-select').value;
  const tolerance = document.getElementById('date-tolerance').value;
  const forceOcr = document.getElementById('force-ocr').checked;

  const formData = new FormData();
  formData.append('file_1c', file1c);
  formData.append('file_cp', fileCp);
  formData.append('model', model);
  formData.append('tolerance', tolerance);
  formData.append('force_ocr', forceOcr);

  showLoader("Обработка актов сверки...", "Qwen-VL распознает скан на GPU AMD Radeon. Это может занять несколько секунд...");

  try {
    const res = await fetch('/api/reconcile', {
      method: 'POST',
      body: formData,
      credentials: 'omit'
    });
    const data = await res.json();
    if (data.error) {
      alert("Ошибка: " + data.error);
    } else {
      displayResults(data);
    }
  } catch (err) {
    alert("Ошибка запроса: " + err.message);
  } finally {
    hideLoader();
  }
}

function showLoader(title, desc) {
  document.getElementById('loader').style.display = 'block';
  document.getElementById('loader-title').textContent = title;
  document.getElementById('loader-desc').textContent = desc;
}

function hideLoader() {
  document.getElementById('loader').style.display = 'none';
}

function formatRub(val) {
  if (val === null || val === undefined) return "—";
  return new Intl.NumberFormat('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(val) + " ₽";
}

function displayResults(data) {
  currentReport = data;
  document.getElementById('results').style.display = 'block';

  // 1. Баннер вердикта
  const banner = document.getElementById('verdict-banner');
  const vIcon = document.getElementById('verdict-icon');
  const vTitle = document.getElementById('verdict-title');
  const vDesc = document.getElementById('verdict-details');

  if (!data.has_discrepancies) {
    banner.className = 'verdict-banner success';
    vIcon.textContent = '✅';
    vTitle.textContent = 'СВЕРКА ПРОЙДЕНА УСПЕШНО';
    vDesc.textContent = 'Все операции и сальдо полностью совпадают с данными контрагента!';
  } else {
    banner.className = 'verdict-banner danger';
    vIcon.textContent = '⚠️';
    vTitle.textContent = 'ОБНАРУЖЕНЫ РАСХОЖДЕНИЯ';
    vDesc.textContent = `Выявлено расхождений: ${data.discrepancy_count} шт. Разница конечного сальдо: ${formatRub(data.diff_final_balance)}`;
  }

  // Кнопка скачивания Excel
  const btnDown = document.getElementById('btn-download-excel');
  btnDown.onclick = () => {
    window.location.href = `/api/download/${data.report_file}`;
  };

  // Метаданные
  document.getElementById('meta-org').textContent = data.statement_1c.organization || 'Наша организация (по 1С)';
  document.getElementById('meta-cp').textContent = data.statement_cp.counterparty || data.statement_1c.counterparty || 'Контрагент';
  document.getElementById('meta-period').textContent = data.statement_1c.period_str || data.statement_cp.period_str || 'Не указан';

  // KPI
  setupKpi('kpi-init-1c', 'kpi-init-cp', 'kpi-init-diff', data.statement_1c.initial_balance, data.statement_cp.initial_balance, data.diff_initial_balance);
  setupKpi('kpi-dt-1c', 'kpi-dt-cp', 'kpi-dt-diff', data.statement_1c.turnover_debit, data.statement_cp.turnover_credit, data.diff_turnover_debit);
  setupKpi('kpi-kt-1c', 'kpi-kt-cp', 'kpi-kt-diff', data.statement_1c.turnover_credit, data.statement_cp.turnover_debit, data.diff_turnover_credit);
  setupKpi('kpi-final-1c', 'kpi-final-cp', 'kpi-final-diff', data.statement_1c.final_balance, data.statement_cp.final_balance, data.diff_final_balance);

  // Счетчики фильтров
  const totalPairs = data.pairs.length;
  const matchedPairs = data.pairs.filter(p => isPairMatched(p.status)).length;
  const diffPairs = totalPairs - matchedPairs;

  document.getElementById('cnt-all').textContent = totalPairs;
  document.getElementById('cnt-diff').textContent = diffPairs;
  document.getElementById('cnt-matched').textContent = matchedPairs;

  renderTable();

  // Плавный скролл к результатам
  document.getElementById('results').scrollIntoView({ behavior: 'smooth' });
}

function setupKpi(id1c, idCp, idDiff, v1c, vCp, diff) {
  document.getElementById(id1c).textContent = formatRub(v1c);
  document.getElementById(idCp).textContent = formatRub(vCp);
  const diffEl = document.getElementById(idDiff);
  diffEl.textContent = "Разница: " + formatRub(diff);
  if (Math.abs(diff) < 0.01) {
    diffEl.className = "kpi-diff match";
  } else {
    diffEl.className = "kpi-diff mismatch";
  }
}

function isPairMatched(status) {
  return status === 'EXACT' || status === 'AMOUNT_AND_DOC' || status === 'AMOUNT_AND_DATE';
}

function renderTable() {
  if (!currentReport) return;
  const tbody = document.getElementById('table-body');
  tbody.innerHTML = '';

  const q = document.getElementById('search-input').value.toLowerCase().trim();

  currentReport.pairs.forEach(p => {
    const isMatched = isPairMatched(p.status);

    // Фильтрация по вкладкам
    if (activeFilter === 'diff' && isMatched) return;
    if (activeFilter === 'matched' && !isMatched) return;

    // Поиск
    if (q) {
      const rowSearch = `${p.item_1c?.doc_raw || ''} ${p.item_cp?.doc_raw || ''} ${p.explanation || ''} ${p.amount_1c} ${p.amount_cp}`.toLowerCase();
      if (!rowSearch.includes(q)) return;
    }

    const tr = document.createElement('tr');

    // Статус-пилл
    const statusPill = getStatusPill(p.status);
    const diffClass = Math.abs(p.diff_amount) > 0.01 ? 'diff-danger' : 'diff-zero';

    tr.innerHTML = `
      <td>${statusPill}</td>
      <td class="num-cell ${diffClass}">${p.diff_amount !== 0 ? (p.diff_amount > 0 ? '+' : '') + formatRub(p.diff_amount) : '0.00 ₽'}</td>
      
      <td>${p.item_1c?.date_str || '—'}</td>
      <td><strong>${p.item_1c?.doc_raw || '—'}</strong></td>
      <td class="num-cell">${p.item_1c?.debit ? formatRub(p.item_1c.debit) : ''}</td>
      <td class="num-cell">${p.item_1c?.credit ? formatRub(p.item_1c.credit) : ''}</td>

      <td>${p.item_cp?.date_str || '—'}</td>
      <td><strong>${p.item_cp?.doc_raw || '—'}</strong></td>
      <td class="num-cell">${p.item_cp?.debit ? formatRub(p.item_cp.debit) : ''}</td>
      <td class="num-cell">${p.item_cp?.credit ? formatRub(p.item_cp.credit) : ''}</td>

      <td style="color: #475569; font-size: 11px;">${p.explanation}</td>
    `;

    tbody.appendChild(tr);
  });
}

function getStatusPill(status) {
  switch (status) {
    case 'EXACT':
      return '<span class="status-pill exact">✅ Сверено</span>';
    case 'AMOUNT_AND_DOC':
      return '<span class="status-pill date-shift">🕒 Сдвиг даты</span>';
    case 'AMOUNT_AND_DATE':
      return '<span class="status-pill date-shift">🔎 Номер отличен</span>';
    case 'AMOUNT_DIFF':
      return '<span class="status-pill amt-diff">⚠️ Разница сумм</span>';
    case 'ONLY_1C':
      return '<span class="status-pill only-1c">❌ Только в 1С</span>';
    case 'ONLY_CP':
      return '<span class="status-pill only-cp">❌ Только у КП</span>';
    default:
      return `<span class="status-pill">${status}</span>`;
  }
}
