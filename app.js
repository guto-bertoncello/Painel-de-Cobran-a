import {
  applyMainFilters,
  buildActionCounts,
  buildAgentReport,
  buildExportAoA,
  cleanCell,
  createDefaultFilters,
  decodeCsvBuffer,
  deduplicateMainRows,
  detectDelimiter,
  formatCurrencyBRL,
  formatDateTime,
  formatMonthLabel,
  formatNumber,
  getAvailableActions,
  matrixToMainRows,
  matrixToNovacaoRows,
  splitRowsByOperators,
  summariseNovacao,
} from "./logic.js";
import { deleteStoredImport, getStoredImport, saveStoredImport } from "./storage.js";

const FILTER_STORAGE_KEY = "site-cobranca-filters-v1";
const ACTION_CHART_COLORS = [
  "#4dc6a6",
  "#f6bd4f",
  "#4b8cff",
  "#ff7e7e",
  "#67d4ff",
  "#8dd58d",
  "#ff9f6b",
  "#ca93ff",
  "#f4e27a",
  "#7fd2c8",
  "#b7c5d7",
  "#36a2eb",
];

const state = {
  section: "dashboard",
  mainRows: [],
  mainPreparedRows: [],
  mainMeta: null,
  mainWarning: "",
  novacaoRows: [],
  novacaoMeta: null,
  novacaoWarning: "",
  filters: createDefaultFilters(),
  savedRun: null,
  hasUnsavedChanges: false,
  chartSelection: new Set(),
  chartSelectionInitialized: false,
  charts: {
    actions: null,
    novacao: null,
  },
};

const dom = {};
let toastTimeout = null;

window.addEventListener("DOMContentLoaded", init);

async function init() {
  await waitForLibraries();
  cacheDom();
  bindEvents();
  updateCurrentMonthLabel();
  await loadPersistedState();
  renderAll();
}

function cacheDom() {
  dom.navLinks = [...document.querySelectorAll(".nav-link")];
  dom.jumpButtons = [...document.querySelectorAll("[data-jump]")];
  dom.views = [...document.querySelectorAll(".view")];
  dom.mainFileInput = document.querySelector("#main-file-input");
  dom.mainImportTrigger = document.querySelector("#main-import-trigger");
  dom.mainExportButton = document.querySelector("#main-export-button");
  dom.mainClearButton = document.querySelector("#main-clear-button");
  dom.mainFileName = document.querySelector("#main-file-name");
  dom.mainFileDetail = document.querySelector("#main-file-detail");
  dom.mainSavedName = document.querySelector("#main-saved-name");
  dom.mainSavedDate = document.querySelector("#main-saved-date");
  dom.mainSavedStatus = document.querySelector("#main-saved-status");
  dom.previewTableBody = document.querySelector("#preview-table-body");
  dom.previewBadge = document.querySelector("#preview-badge");
  dom.statImported = document.querySelector("#stat-imported");
  dom.statImportedNote = document.querySelector("#stat-imported-note");
  dom.statDeduped = document.querySelector("#stat-deduped");
  dom.statDedupedNote = document.querySelector("#stat-deduped-note");
  dom.statActions = document.querySelector("#stat-actions");
  dom.statOperators = document.querySelector("#stat-operators");
  dom.statOperatorsNote = document.querySelector("#stat-operators-note");
  dom.sidebarMainStatus = document.querySelector("#sidebar-main-status");
  dom.sidebarMainSubtitle = document.querySelector("#sidebar-main-subtitle");
  dom.sidebarNovacaoStatus = document.querySelector("#sidebar-novacao-status");
  dom.sidebarNovacaoSubtitle = document.querySelector("#sidebar-novacao-subtitle");
  dom.currentMonthLabel = document.querySelector("#current-month-label");
  dom.saveFiltersButton = document.querySelector("#save-filters-button");
  dom.tempoMinInput = document.querySelector("#tempo-min-input");
  dom.tempoMaxInput = document.querySelector("#tempo-max-input");
  dom.valorVencidoSelect = document.querySelector("#valor-vencido-select");
  dom.valorRiscoSelect = document.querySelector("#valor-risco-select");
  dom.hotSelect = document.querySelector("#hot-select");
  dom.operadoresInput = document.querySelector("#operadores-input");
  dom.padraoToggle = document.querySelector("#padrao-toggle");
  dom.actionsFilterGrid = document.querySelector("#actions-filter-grid");
  dom.savedSummary = document.querySelector("#saved-summary");
  dom.chartSourceBadge = document.querySelector("#chart-source-badge");
  dom.chartActionsGrid = document.querySelector("#chart-actions-grid");
  dom.actionsChartCanvas = document.querySelector("#actions-chart");
  dom.novacaoFileInput = document.querySelector("#novacao-file-input");
  dom.novacaoImportTrigger = document.querySelector("#novacao-import-trigger");
  dom.novacaoClearButton = document.querySelector("#novacao-clear-button");
  dom.novacaoFileName = document.querySelector("#novacao-file-name");
  dom.novacaoFileDetail = document.querySelector("#novacao-file-detail");
  dom.novacaoTotalNov = document.querySelector("#novacao-total-nov");
  dom.novacaoTotalNote = document.querySelector("#novacao-total-note");
  dom.novacaoAtrasoTotal = document.querySelector("#novacao-atraso-total");
  dom.novacaoAtrasoNote = document.querySelector("#novacao-atraso-note");
  dom.novacaoReceberTotal = document.querySelector("#novacao-receber-total");
  dom.novacaoReceberNote = document.querySelector("#novacao-receber-note");
  dom.novacaoReferenceNote = document.querySelector("#novacao-reference-note");
  dom.novacaoChartCanvas = document.querySelector("#novacao-chart");
  dom.agentsList = document.querySelector("#agents-list");
  dom.toast = document.querySelector("#toast");
  dom.loadingOverlay = document.querySelector("#loading-overlay");
  dom.loadingText = document.querySelector("#loading-text");
}

function bindEvents() {
  dom.navLinks.forEach((button) => {
    button.addEventListener("click", () => setSection(button.dataset.target));
  });

  dom.jumpButtons.forEach((button) => {
    button.addEventListener("click", () => setSection(button.dataset.jump));
  });

  dom.mainImportTrigger.addEventListener("click", () => dom.mainFileInput.click());
  dom.mainFileInput.addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    if (file) {
      await importMainFile(file);
      dom.mainFileInput.value = "";
    }
  });

  dom.mainExportButton.addEventListener("click", exportCurrentWorkbook);
  dom.mainClearButton.addEventListener("click", clearMainStorage);

  [
    dom.tempoMinInput,
    dom.tempoMaxInput,
    dom.valorVencidoSelect,
    dom.valorRiscoSelect,
    dom.hotSelect,
    dom.operadoresInput,
    dom.padraoToggle,
  ].forEach((element) => {
    const eventName = element.tagName === "SELECT" || element.type === "checkbox" ? "change" : "input";
    element.addEventListener(eventName, handleScalarFilterChange);
  });

  dom.actionsFilterGrid.addEventListener("change", handleActionFilterChange);
  dom.actionsFilterGrid.addEventListener("input", handleActionFilterChange);
  dom.saveFiltersButton.addEventListener("click", saveFiltersRun);

  dom.chartActionsGrid.addEventListener("change", handleChartSelectionChange);

  dom.novacaoImportTrigger.addEventListener("click", () => dom.novacaoFileInput.click());
  dom.novacaoFileInput.addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    if (file) {
      await importNovacaoFile(file);
      dom.novacaoFileInput.value = "";
    }
  });

  dom.novacaoClearButton.addEventListener("click", clearNovacaoStorage);
}

async function loadPersistedState() {
  try {
    const [mainRecord, novacaoRecord] = await Promise.all([
      getStoredImport("main"),
      getStoredImport("novacao"),
    ]);

    if (mainRecord) {
      restoreMainState(mainRecord);
    }

    if (novacaoRecord) {
      restoreNovacaoState(novacaoRecord);
    }

    const savedFilters = loadSavedFilters();
    state.filters = createDefaultFilters(state.mainPreparedRows, savedFilters);
    state.chartSelection = new Set(getAvailableActions(getChartSourceRows()));
    state.chartSelectionInitialized = false;
  } catch (error) {
    console.error(error);
    showToast("Não foi possível carregar os arquivos salvos localmente.");
  }
}

async function importMainFile(file) {
  showLoading("Importando base principal...");

  try {
    await nextFrame();
    const matrix = await readMatrixFromFile(file);
    const parsed = matrixToMainRows(matrix);
    const deduped = deduplicateMainRows(parsed.rows);
    const importedAt = new Date().toISOString();

    state.mainRows = parsed.rows;
    state.mainPreparedRows = deduped.rows;
    state.mainMeta = {
      name: file.name,
      importedAt,
      dedupedCount: deduped.rows.length,
      duplicateCount: deduped.removedCount,
      rowCount: parsed.rows.length,
    };
    state.mainWarning = parsed.warning || "";
    state.savedRun = null;
    state.hasUnsavedChanges = false;
    state.filters = createDefaultFilters(state.mainPreparedRows, loadSavedFilters());
    state.chartSelection = new Set(getAvailableActions(getChartSourceRows()));
    state.chartSelectionInitialized = false;

    await saveStoredImport({
      slot: "main",
      name: file.name,
      importedAt,
      rows: state.mainRows,
      meta: state.mainMeta,
      warning: state.mainWarning,
    });

    renderAll();

    if (state.mainWarning) {
      showToast(state.mainWarning);
    } else {
      showToast("Base principal importada e salva localmente.");
    }
  } catch (error) {
    console.error(error);
    showToast("Não foi possível importar a base principal.");
  } finally {
    hideLoading();
  }
}

async function importNovacaoFile(file) {
  showLoading("Importando base de novação...");

  try {
    await nextFrame();
    const matrix = await readMatrixFromFile(file);
    const parsed = matrixToNovacaoRows(matrix);
    const importedAt = new Date().toISOString();

    state.novacaoRows = parsed.rows;
    state.novacaoMeta = {
      name: file.name,
      importedAt,
      rowCount: parsed.rows.length,
    };
    state.novacaoWarning = parsed.warning || "";

    await saveStoredImport({
      slot: "novacao",
      name: file.name,
      importedAt,
      rows: state.novacaoRows,
      meta: state.novacaoMeta,
      warning: state.novacaoWarning,
    });

    renderAll();

    if (state.novacaoWarning) {
      showToast(state.novacaoWarning);
    } else {
      showToast("Base de novação importada e salva localmente.");
    }
  } catch (error) {
    console.error(error);
    showToast("Não foi possível importar a base de novação.");
  } finally {
    hideLoading();
  }
}

async function clearMainStorage() {
  if (!state.mainMeta && !state.mainRows.length) {
    showToast("Nenhuma base principal salva para remover.");
    return;
  }

  if (!window.confirm("Remover a base principal salva localmente? O arquivo original não será apagado.")) {
    return;
  }

  try {
    await deleteStoredImport("main");
    state.mainRows = [];
    state.mainPreparedRows = [];
    state.mainMeta = null;
    state.mainWarning = "";
    state.savedRun = null;
    state.hasUnsavedChanges = false;
    state.filters = createDefaultFilters();
    state.chartSelection = new Set();
    state.chartSelectionInitialized = false;
    saveFilterState(state.filters);
    renderAll();
    showToast("Base principal removida do armazenamento local.");
  } catch (error) {
    console.error(error);
    showToast("Não foi possível remover a base principal salva.");
  }
}

async function clearNovacaoStorage() {
  if (!state.novacaoMeta && !state.novacaoRows.length) {
    showToast("Nenhuma base de novação salva para remover.");
    return;
  }

  if (!window.confirm("Remover a base de novação salva localmente? O arquivo original não será apagado.")) {
    return;
  }

  try {
    await deleteStoredImport("novacao");
    state.novacaoRows = [];
    state.novacaoMeta = null;
    state.novacaoWarning = "";
    renderAll();
    showToast("Base de novação removida do armazenamento local.");
  } catch (error) {
    console.error(error);
    showToast("Não foi possível remover a base de novação salva.");
  }
}

function restoreMainState(record) {
  state.mainRows = record.rows ?? [];
  const deduped = deduplicateMainRows(state.mainRows);
  state.mainPreparedRows = deduped.rows;
  state.mainMeta = {
    name: record.name,
    importedAt: record.importedAt,
    rowCount: record.meta?.rowCount ?? state.mainRows.length,
    dedupedCount: record.meta?.dedupedCount ?? deduped.rows.length,
    duplicateCount: record.meta?.duplicateCount ?? deduped.removedCount,
  };
  state.mainWarning = record.warning ?? "";
}

function restoreNovacaoState(record) {
  state.novacaoRows = record.rows ?? [];
  state.novacaoMeta = {
    name: record.name,
    importedAt: record.importedAt,
    rowCount: record.meta?.rowCount ?? state.novacaoRows.length,
  };
  state.novacaoWarning = record.warning ?? "";
}

function handleScalarFilterChange() {
  syncScalarFiltersFromInputs();
  state.hasUnsavedChanges = true;
  renderSummary();
}

function handleActionFilterChange(event) {
  const card = event.target.closest("[data-action]");
  if (!card) {
    return;
  }

  const action = card.dataset.action;
  const filter = state.filters.actionFilters[action];
  if (!filter) {
    return;
  }

  const role = event.target.dataset.role;
  if (role === "selected") {
    filter.selected = event.target.checked;
  }

  if (role === "min") {
    filter.min = event.target.value.trim();
  }

  if (role === "max") {
    filter.max = event.target.value.trim();
  }

  state.hasUnsavedChanges = true;
  renderSummary();
}

function handleChartSelectionChange(event) {
  const target = event.target;
  if (!target.matches("[data-chart-action]")) {
    return;
  }

  const action = target.dataset.chartAction;
  if (target.checked) {
    state.chartSelection.add(action);
  } else {
    state.chartSelection.delete(action);
  }

  state.chartSelectionInitialized = true;
  renderActionsSection();
}

function saveFiltersRun() {
  if (!state.mainPreparedRows.length) {
    showToast("Importe a base principal antes de salvar os filtros.");
    return;
  }

  syncScalarFiltersFromInputs();

  const result = applyMainFilters(state.mainPreparedRows, state.filters);
  const operators = Math.max(1, Number(state.filters.operadores) || 1);
  const split = splitRowsByOperators(result.rows, operators);

  state.savedRun = {
    ...result,
    operators,
    split,
    savedAt: new Date().toISOString(),
  };
  state.hasUnsavedChanges = false;
  state.chartSelection = new Set(getAvailableActions(getChartSourceRows()));
  state.chartSelectionInitialized = false;

  saveFilterState(state.filters);
  renderAll();

  showToast(
    `Filtros salvos. ${formatNumber(result.totalClients)} clientes distribuídos em ${operators} aba(s).`,
  );
}

function exportCurrentWorkbook() {
  if (!window.XLSX) {
    showToast("A biblioteca de exportação ainda não está pronta.");
    return;
  }

  const rows = state.savedRun?.rows?.length ? state.savedRun.rows : state.mainPreparedRows;
  if (!rows.length) {
    showToast("Não há linhas para exportar.");
    return;
  }

  const operators = state.savedRun?.operators ?? 1;
  const sheets = splitRowsByOperators(rows, operators);
  const workbook = window.XLSX.utils.book_new();

  sheets.forEach((sheetRows, index) => {
    const aoa = buildExportAoA(sheetRows);
    const sheet = window.XLSX.utils.aoa_to_sheet(aoa);
    window.XLSX.utils.book_append_sheet(workbook, sheet, `Operador ${index + 1}`);
  });

  const fileName = `cobranca_filtrada_${buildTimestamp()}.xlsx`;
  window.XLSX.writeFile(workbook, fileName, { compression: true });

  if (state.savedRun && state.hasUnsavedChanges) {
    showToast("Planilha exportada com base na última configuração salva.");
    return;
  }

  showToast(`Planilha exportada com ${operators} aba(s) de operadores.`);
}

function renderAll() {
  renderSection();
  renderSidebar();
  renderDashboard();
  renderFiltersSection();
  renderActionsSection();
  renderNovacaoSection();
  renderAgentsSection();
}

function renderSection() {
  dom.navLinks.forEach((button) => {
    button.classList.toggle("is-active", button.dataset.target === state.section);
  });

  dom.views.forEach((view) => {
    view.classList.toggle("is-active", view.dataset.section === state.section);
  });
}

function renderSidebar() {
  if (state.mainMeta) {
    dom.sidebarMainStatus.textContent = state.mainMeta.name;
    dom.sidebarMainSubtitle.textContent = `${formatNumber(state.mainPreparedRows.length)} cliente(s) após deduplicação.`;
  } else {
    dom.sidebarMainStatus.textContent = "Nenhuma base importada";
    dom.sidebarMainSubtitle.textContent = "Importe a planilha principal para começar.";
  }

  if (state.novacaoMeta) {
    dom.sidebarNovacaoStatus.textContent = state.novacaoMeta.name;
    dom.sidebarNovacaoSubtitle.textContent = `${formatNumber(state.novacaoRows.length)} linha(s) carregadas.`;
  } else {
    dom.sidebarNovacaoStatus.textContent = "Sem arquivo salvo";
    dom.sidebarNovacaoSubtitle.textContent =
      "A tela de novação usa uma importação separada.";
  }
}

function renderDashboard() {
  dom.mainFileName.textContent = state.mainMeta?.name ?? "Nenhum arquivo carregado";
  dom.mainFileDetail.textContent = state.mainMeta
    ? `Importado em ${formatDateTime(state.mainMeta.importedAt)}`
    : "Aceita .xlsx, .xls, .csv e .tsv";
  dom.mainSavedName.textContent = state.mainMeta?.name ?? "Nenhum arquivo salvo";
  dom.mainSavedDate.textContent = state.mainMeta?.importedAt
    ? formatDateTime(state.mainMeta.importedAt)
    : "-";

  if (state.mainWarning) {
    dom.mainSavedStatus.textContent = state.mainWarning;
  } else if (state.mainMeta) {
    dom.mainSavedStatus.textContent = "Pronto para filtrar e exportar";
  } else {
    dom.mainSavedStatus.textContent = "Aguardando importação";
  }

  dom.statImported.textContent = formatNumber(state.mainRows.length);
  dom.statImportedNote.textContent = state.mainMeta
    ? `${formatNumber(state.mainMeta.duplicateCount ?? 0)} duplicata(s) removidas pela coluna CPF.`
    : "Sem base ativa.";
  dom.statDeduped.textContent = formatNumber(state.mainPreparedRows.length);
  dom.statDedupedNote.textContent = state.mainPreparedRows.length
    ? "CPF duplicado removido mantendo a primeira linha encontrada."
    : "CPF é deduplicado mantendo a primeira ocorrência.";
  dom.statActions.textContent = formatNumber(getAvailableActions(state.mainPreparedRows).length);
  dom.statOperators.textContent = formatNumber(state.savedRun?.operators ?? 0);
  dom.statOperatorsNote.textContent = state.savedRun
    ? `${formatNumber(state.savedRun.totalClients)} cliente(s) na última simulação salva.`
    : "Atualizado quando você salva os filtros.";

  const previewRows = (state.savedRun?.rows?.length ? state.savedRun.rows : state.mainPreparedRows).slice(0, 8);
  const previewSource = state.savedRun?.rows?.length ? "Filtros salvos" : "Base deduplicada";
  dom.previewBadge.textContent = previewRows.length ? previewSource : "Sem dados";

  if (!previewRows.length) {
    dom.previewTableBody.innerHTML = `
      <tr>
        <td colspan="9" class="empty-row">Importe a base principal para visualizar a prévia.</td>
      </tr>
    `;
    return;
  }

  dom.previewTableBody.innerHTML = previewRows
    .map(
      (row) => `
        <tr>
          <td>${escapeHtml(row.cpfDisplay || row.cpf)}</td>
          <td>${escapeHtml(row.nome)}</td>
          <td>${escapeHtml(row.contrato)}</td>
          <td>${escapeHtml(valueOrDash(row.tempoAtrasoRaw || row.tempoAtraso))}</td>
          <td>${escapeHtml(formatPreviewCurrency(row.valorVencidoRaw, row.valorVencido))}</td>
          <td>${escapeHtml(formatPreviewCurrency(row.valorRiscoRaw, row.valorRisco))}</td>
          <td>${escapeHtml(row.telefone || "-")}</td>
          <td>${escapeHtml(valueOrDash(row.defasagemRaw || row.defasagem))}</td>
          <td>${escapeHtml(row.acao || "-")}</td>
        </tr>
      `,
    )
    .join("");
}

function renderFiltersSection() {
  renderFilterInputs();
  renderActionFilters();
  renderSummary();
}

function renderFilterInputs() {
  dom.tempoMinInput.value = state.filters.tempoMin;
  dom.tempoMaxInput.value = state.filters.tempoMax;
  dom.valorVencidoSelect.value = state.filters.valorVencidoBand;
  dom.valorRiscoSelect.value = state.filters.valorRiscoBand;
  dom.hotSelect.value = state.filters.hot;
  dom.operadoresInput.value = String(state.filters.operadores);
  dom.padraoToggle.checked = Boolean(state.filters.padrao);
}

function renderActionFilters() {
  if (!state.filters.actions.length) {
    dom.actionsFilterGrid.innerHTML = `
      <div class="empty-state">Importe a base principal para habilitar a lista de ações.</div>
    `;
    return;
  }

  dom.actionsFilterGrid.innerHTML = state.filters.actions
    .map((action) => {
      const filter = state.filters.actionFilters[action];
      return `
        <article class="action-card" data-action="${escapeHtml(action)}">
          <div class="action-card__head">
            <label class="action-card__toggle">
              <input data-role="selected" type="checkbox" ${filter.selected ? "checked" : ""} />
              <span>${escapeHtml(action)}</span>
            </label>
            <span class="badge">Coluna L</span>
          </div>
          <div class="action-card__range">
            <label class="range-field">
              <span>Defasagem mínima</span>
              <input data-role="min" type="number" min="0" value="${escapeAttribute(filter.min)}" />
            </label>
            <label class="range-field">
              <span>Defasagem máxima</span>
              <input data-role="max" type="number" min="0" value="${escapeAttribute(filter.max)}" />
            </label>
          </div>
        </article>
      `;
    })
    .join("");
}

function renderSummary() {
  if (!state.mainPreparedRows.length) {
    dom.savedSummary.innerHTML =
      'Importe a base principal na tela <strong>Lar</strong> para liberar a simulação.';
    return;
  }

  if (!state.savedRun) {
    dom.savedSummary.innerHTML =
      'Clique em <strong>Salvar</strong> para calcular o total de clientes e a divisão por operadores.';
    return;
  }

  const splitDescription = state.savedRun.split
    .map((chunk, index) => `Aba ${index + 1}: ${formatNumber(chunk.length)} cliente(s)`)
    .join(" · ");

  const dirtyNote = state.hasUnsavedChanges
    ? '<br /><span class="hint-inline">Existem alterações nos filtros que ainda não foram salvas.</span>'
    : "";

  dom.savedSummary.innerHTML = `
    <strong>Total de clientes:</strong> ${formatNumber(state.savedRun.totalClients)}<br />
    <strong>Operadores:</strong> ${formatNumber(state.savedRun.operators)}<br />
    <strong>Divisão:</strong> ${splitDescription}<br />
    <strong>Último cálculo:</strong> ${formatDateTime(state.savedRun.savedAt)}
    ${dirtyNote}
  `;
}

function renderActionsSection() {
  const rows = getChartSourceRows();
  const counts = buildActionCounts(rows);
  const availableActions = getAvailableActions(rows).filter((action) => counts.has(action));

  if (!state.chartSelectionInitialized && availableActions.length) {
    state.chartSelection = new Set(availableActions);
    state.chartSelectionInitialized = true;
  }

  state.chartSelection.forEach((action) => {
    if (!availableActions.includes(action)) {
      state.chartSelection.delete(action);
    }
  });

  dom.chartSourceBadge.textContent = state.savedRun?.rows?.length
    ? "Usando filtros salvos"
    : "Base deduplicada";

  if (!availableActions.length) {
    dom.chartActionsGrid.innerHTML = `
      <div class="empty-state">Importe a base principal para habilitar o gráfico de ações.</div>
    `;
    renderPieChart(dom.actionsChartCanvas, "actions", [], [], "Nenhuma ação disponível");
    return;
  }

  dom.chartActionsGrid.innerHTML = availableActions
    .map((action) => {
      const count = counts.get(action) ?? 0;
      return `
        <label class="selection-row">
          <input
            data-chart-action="${escapeHtml(action)}"
            type="checkbox"
            ${state.chartSelection.has(action) ? "checked" : ""}
          />
          <span>${escapeHtml(action)}</span>
          <span>${formatNumber(count)}</span>
        </label>
      `;
    })
    .join("");

  const selectedActions = availableActions.filter((action) => state.chartSelection.has(action));
  const labels = selectedActions;
  const data = labels.map((action) => counts.get(action) ?? 0);

  renderPieChart(
    dom.actionsChartCanvas,
    "actions",
    labels,
    data,
    labels.length ? "Ações selecionadas" : "Nenhuma ação selecionada",
  );
}

function renderNovacaoSection() {
  dom.novacaoFileName.textContent = state.novacaoMeta?.name ?? "Nenhum arquivo carregado";
  dom.novacaoFileDetail.textContent = state.novacaoMeta
    ? `Importado em ${formatDateTime(state.novacaoMeta.importedAt)}`
    : "A base fica salva localmente para futuras leituras";

  const summary = summariseNovacao(state.novacaoRows, new Date());
  dom.novacaoTotalNov.textContent = formatNumber(summary.totalNov);
  dom.novacaoTotalNote.textContent = state.novacaoWarning
    ? state.novacaoWarning
    : `${formatNumber(summary.currentMonthNov)} linha(s) NOV no mês atual.`;
  dom.novacaoAtrasoTotal.textContent = formatCurrencyBRL(summary.atrasoTotal);
  dom.novacaoAtrasoNote.textContent = `${formatNumber(summary.atrasoCount)} título(s) em atraso no mês atual.`;
  dom.novacaoReceberTotal.textContent = formatCurrencyBRL(summary.receberTotal);
  dom.novacaoReceberNote.textContent = `${formatNumber(summary.receberCount)} título(s) a receber no mês atual.`;
  dom.novacaoReferenceNote.textContent = `Referência automática: ${formatMonthLabel(new Date())}.`;

  renderPieChart(
    dom.novacaoChartCanvas,
    "novacao",
    ["Em atraso", "A receber"],
    [summary.atrasoTotal, summary.receberTotal],
    "Novação do mês atual",
  );
}

function renderAgentsSection() {
  const report = buildAgentReport(state.mainPreparedRows);
  if (!report.length) {
    dom.agentsList.innerHTML =
      '<div class="empty-state">Importe a base principal para gerar o relatório por agente.</div>';
    return;
  }

  dom.agentsList.innerHTML = report
    .map(
      (agent) => `
        <article class="agent-card">
          <p class="eyebrow">Operador</p>
          <h3>${escapeHtml(formatAgentLabel(agent.operator))}</h3>
          <div class="agent-card__metrics">
            <div class="agent-card__metric">
              <span>Acionamentos</span>
              <strong>${formatNumber(agent.acionamentos)}</strong>
            </div>
            <div class="agent-card__metric">
              <span>Acordos</span>
              <strong>${formatNumber(agent.agreements)}</strong>
            </div>
          </div>
        </article>
      `,
    )
    .join("");
}

function renderPieChart(canvas, chartKey, labels, values, title) {
  if (!window.Chart) {
    return;
  }

  const hasPositiveValues = values.some((value) => value > 0);
  const safeLabels = labels.length && hasPositiveValues ? labels : ["Sem dados"];
  const safeValues = labels.length && hasPositiveValues ? values : [1];

  if (state.charts[chartKey]) {
    state.charts[chartKey].destroy();
  }

  state.charts[chartKey] = new window.Chart(canvas, {
    type: "pie",
    data: {
      labels: safeLabels,
      datasets: [
        {
          label: title,
          data: safeValues,
          backgroundColor: ACTION_CHART_COLORS.slice(0, safeLabels.length),
          borderColor: "rgba(8, 17, 27, 0.88)",
          borderWidth: 3,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: "bottom",
          labels: {
            color: "#c7d2df",
            boxWidth: 14,
            padding: 18,
          },
        },
        title: {
          display: true,
          text: title,
          color: "#f3f7fb",
          font: {
            family: "Sora",
            size: 18,
            weight: "600",
          },
          padding: {
            bottom: 18,
          },
        },
        tooltip: {
          callbacks: {
            label(context) {
              const value = context.raw ?? 0;
              const isCurrencyChart = chartKey === "novacao";
              return `${context.label}: ${
                isCurrencyChart ? formatCurrencyBRL(value) : formatNumber(value)
              }`;
            },
          },
        },
      },
    },
  });
}

function setSection(section) {
  state.section = section;
  renderSection();
}

function syncScalarFiltersFromInputs() {
  state.filters.tempoMin = dom.tempoMinInput.value.trim();
  state.filters.tempoMax = dom.tempoMaxInput.value.trim();
  state.filters.valorVencidoBand = dom.valorVencidoSelect.value;
  state.filters.valorRiscoBand = dom.valorRiscoSelect.value;
  state.filters.hot = dom.hotSelect.value;
  state.filters.operadores = Math.max(1, Number(dom.operadoresInput.value) || 1);
  state.filters.padrao = dom.padraoToggle.checked;
}

function getChartSourceRows() {
  return state.savedRun?.rows?.length ? state.savedRun.rows : state.mainPreparedRows;
}

async function readMatrixFromFile(file) {
  const arrayBuffer = await file.arrayBuffer();
  const extension = file.name.split(".").pop()?.toLowerCase();

  if (extension === "csv" || extension === "tsv" || extension === "txt") {
    const text = decodeCsvBuffer(arrayBuffer);
    const delimiter = extension === "tsv" ? "\t" : detectDelimiter(text);
    const workbook = window.XLSX.read(text, {
      type: "string",
      raw: false,
      FS: delimiter,
    });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    return window.XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: false,
      defval: "",
    });
  }

  const workbook = window.XLSX.read(arrayBuffer, {
    type: "array",
    raw: false,
  });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return window.XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: "",
  });
}

function updateCurrentMonthLabel() {
  dom.currentMonthLabel.textContent = formatMonthLabel(new Date());
}

function showLoading(message) {
  dom.loadingText.textContent = message;
  dom.loadingOverlay.hidden = false;
}

function hideLoading() {
  dom.loadingOverlay.hidden = true;
}

function showToast(message) {
  dom.toast.textContent = message;
  dom.toast.classList.add("is-visible");
  window.clearTimeout(toastTimeout);
  toastTimeout = window.setTimeout(() => {
    dom.toast.classList.remove("is-visible");
  }, 3600);
}

function loadSavedFilters() {
  try {
    const serialized = localStorage.getItem(FILTER_STORAGE_KEY);
    return serialized ? JSON.parse(serialized) : null;
  } catch (error) {
    console.error(error);
    return null;
  }
}

function saveFilterState(filters) {
  try {
    localStorage.setItem(FILTER_STORAGE_KEY, JSON.stringify(filters));
  } catch (error) {
    console.error(error);
  }
}

function buildTimestamp() {
  const now = new Date();
  const values = [
    now.getFullYear(),
    String(now.getMonth() + 1).padStart(2, "0"),
    String(now.getDate()).padStart(2, "0"),
    String(now.getHours()).padStart(2, "0"),
    String(now.getMinutes()).padStart(2, "0"),
    String(now.getSeconds()).padStart(2, "0"),
  ];
  return values.join("");
}

function escapeHtml(value) {
  return cleanCell(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function escapeAttribute(value) {
  return escapeHtml(value);
}

function valueOrDash(value) {
  const cleaned = cleanCell(value);
  return cleaned || "-";
}

function formatPreviewCurrency(rawValue, parsedValue) {
  const raw = cleanCell(rawValue);
  if (raw) {
    return raw;
  }

  return Number.isFinite(parsedValue) ? formatCurrencyBRL(parsedValue) : "-";
}

function formatAgentLabel(operator) {
  const cleaned = cleanCell(operator);
  if (/^\d+$/.test(cleaned)) {
    return `Operador ${cleaned}`;
  }

  return cleaned || "Sem agente";
}

function nextFrame() {
  return new Promise((resolve) => window.requestAnimationFrame(resolve));
}

async function waitForLibraries() {
  for (let attempt = 0; attempt < 40; attempt += 1) {
    if (window.XLSX && window.Chart) {
      return;
    }}

    await new Promise((resolve) => window.setTimeout(resolve, 100));
  }
