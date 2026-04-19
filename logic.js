export const KNOWN_ACTIONS = [
  "FNV",
  "SEM",
  "WPP",
  "DES",
  "AED",
  "CPC",
  "NEG",
  "PRE",
  "ATR",
  "REC",
  "ACR",
  "ACP",
  "STE",
  "SEI",
  "PRO",
  "PRV",
  "PES",
  "DSD",
  "FDB",
  "FLC",
  "ALE",
  "CXP",
];

export const MAIN_EXPORT_COLUMNS = [
  { label: "CPF do Cliente", value: (row) => row.cpfDisplay || row.cpf },
  { label: "Nome do Cliente", value: (row) => row.nome },
  { label: "Contrato do Cliente", value: (row) => row.contrato },
  { label: "Tempo de Atraso", value: (row) => row.tempoAtrasoRaw || valueOrBlank(row.tempoAtraso) },
  { label: "Valor Vencido", value: (row) => row.valorVencidoRaw || valueOrBlank(row.valorVencido) },
  { label: "Valor Risco", value: (row) => row.valorRiscoRaw || valueOrBlank(row.valorRisco) },
  { label: "Telefone", value: (row) => row.telefone },
  { label: "Defasagem", value: (row) => row.defasagemRaw || valueOrBlank(row.defasagem) },
  { label: "Acionamento", value: (row) => row.acaoRaw || row.acao },
];

const MAIN_HEADER_KEYWORDS = [
  "cpf",
  "cliente",
  "contrato",
  "carne",
  "atraso",
  "vencido",
  "risco",
  "telefone",
  "hot",
  "defasagem",
  "acionamento",
  "agente",
];

const NOVACAO_HEADER_KEYWORDS = [
  "credor",
  "filial",
  "cod",
  "matricula",
  "nome",
  "cpf",
  "tipo",
  "titulo",
  "parcela",
  "vencimento",
  "valor",
  "atraso",
  "inclusao",
  "sit",
  "fase",
  "uf",
];

const PADRAO_EXCLUDED_ACTIONS = new Set(["ACR", "ACP", "ATR", "PRO", "PRV", "FDB", "FLC"]);
const PADRAO_DEFASAGEM_LIMITS = {
  CPC: 3,
  NEG: 4,
  PRE: 7,
  ALE: 3,
  DES: 0,
  SEM: 2,
  WPP: 4,
  REC: 2,
  PES: 2,
  AED: 2,
};

export function normalizeToken(value) {
  return cleanCell(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

export function cleanCell(value) {
  if (value === null || value === undefined) {
    return "";
  }

  return String(value).trim();
}

export function sanitizeCpf(value) {
  const digits = cleanCell(value).replace(/\D+/g, "");
  return digits;
}

export function valueOrBlank(value) {
  return Number.isFinite(value) ? value : "";
}

export function parseBrazilianNumber(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : NaN;
  }

  let text = cleanCell(value);
  if (!text) {
    return NaN;
  }

  text = text
    .replace(/R\$/gi, "")
    .replace(/\s+/g, "")
    .replace(/[^\d,.-]/g, "");

  if (!text) {
    return NaN;
  }

  const lastComma = text.lastIndexOf(",");
  const lastDot = text.lastIndexOf(".");

  if (lastComma !== -1 && lastDot !== -1) {
    if (lastComma > lastDot) {
      text = text.replace(/\./g, "").replace(",", ".");
    } else {
      text = text.replace(/,/g, "");
    }
  } else if (lastComma !== -1) {
    text = text.replace(/\./g, "").replace(",", ".");
  }

  const parsed = Number(text);
  return Number.isFinite(parsed) ? parsed : NaN;
}

export function parseIntegerValue(value) {
  const parsed = parseBrazilianNumber(value);
  return Number.isFinite(parsed) ? Math.trunc(parsed) : null;
}

export function parseBrazilianDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value;
  }

  const text = cleanCell(value);
  if (!text) {
    return null;
  }

  if (/^\d{2}\/\d{2}\/\d{4}$/.test(text)) {
    const [dayText, monthText, yearText] = text.split("/");
    const day = Number(dayText);
    const month = Number(monthText);
    const year = Number(yearText);
    const parsed = new Date(year, month - 1, day);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }

  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

export function formatCurrencyBRL(value) {
  return new Intl.NumberFormat("pt-BR", {
    style: "currency",
    currency: "BRL",
  }).format(Number.isFinite(value) ? value : 0);
}

export function formatNumber(value) {
  return new Intl.NumberFormat("pt-BR").format(Number.isFinite(value) ? value : 0);
}

export function formatDateTime(value) {
  if (!value) {
    return "-";
  }

  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) {
    return "-";
  }

  return new Intl.DateTimeFormat("pt-BR", {
    dateStyle: "short",
    timeStyle: "short",
  }).format(date);
}

export function formatMonthLabel(value) {
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) {
    return "mês atual";
  }

  return new Intl.DateTimeFormat("pt-BR", {
    month: "long",
    year: "numeric",
  }).format(date);
}

export function decodeCsvBuffer(arrayBuffer) {
  const utf8 = new TextDecoder("utf-8", { fatal: false }).decode(arrayBuffer);
  if (!looksMisdecoded(utf8)) {
    return utf8;
  }

  return new TextDecoder("windows-1252").decode(arrayBuffer);
}

export function detectDelimiter(text) {
  const sample = text.split(/\r?\n/).slice(0, 8).join("\n");
  const delimiters = [";", "\t", ","];
  const ranked = delimiters
    .map((delimiter) => ({
      delimiter,
      score: sample.split(delimiter).length,
    }))
    .sort((left, right) => right.score - left.score);

  return ranked[0]?.delimiter || ";";
}

export function detectPlaceholderMainMatrix(matrix) {
  const meaningfulRows = matrix.filter((row) => row.some((cell) => cleanCell(cell)));
  if (meaningfulRows.length < 18) {
    return false;
  }

  return meaningfulRows.every((row) => {
    const first = normalizeToken(row[0]);
    const second = normalizeToken(row[1]);
    return first === "coluna" && /^[a-z]$/.test(second);
  });
}

export function matrixToMainRows(matrix) {
  const compactRows = compactMatrix(matrix);
  if (!compactRows.length) {
    return { rows: [], warning: "Nenhuma linha válida foi encontrada na planilha principal." };
  }

  if (detectPlaceholderMainMatrix(compactRows)) {
    return {
      rows: [],
      warning:
        "O arquivo principal parece ser apenas um modelo de colunas, sem linhas reais de clientes.",
    };
  }

  const firstRow = compactRows[0];
  const hasHeader = rowLooksLikeHeader(firstRow, MAIN_HEADER_KEYWORDS, 4);
  const startIndex = hasHeader ? 1 : 0;
  const rows = [];

  for (let index = startIndex; index < compactRows.length; index += 1) {
    const row = compactRows[index];
    if (!row.some((cell) => cleanCell(cell))) {
      continue;
    }

    rows.push(toMainRowObject(row, index + 1));
  }

  return {
    rows,
    warning: rows.length
      ? ""
      : "Nenhuma linha de cliente foi identificada depois da leitura do arquivo principal.",
  };
}

export function matrixToNovacaoRows(matrix) {
  const compactRows = compactMatrix(matrix);
  if (!compactRows.length) {
    return { rows: [], warning: "Nenhuma linha válida foi encontrada na planilha de novação." };
  }

  const hasHeader = rowLooksLikeHeader(compactRows[0], NOVACAO_HEADER_KEYWORDS, 4);
  const startIndex = hasHeader ? 1 : 0;
  const rows = [];

  for (let index = startIndex; index < compactRows.length; index += 1) {
    const row = compactRows[index];
    if (!row.some((cell) => cleanCell(cell))) {
      continue;
    }

    rows.push(toNovacaoRowObject(row, index + 1));
  }

  return {
    rows,
    warning: rows.length
      ? ""
      : "Nenhuma linha de novação foi identificada depois da leitura do arquivo importado.",
  };
}

export function deduplicateMainRows(rows) {
  const seen = new Set();
  const deduped = [];
  let removedCount = 0;

  rows.forEach((row) => {
    if (!row.cpf) {
      deduped.push(row);
      return;
    }

    if (seen.has(row.cpf)) {
      removedCount += 1;
      return;
    }

    seen.add(row.cpf);
    deduped.push(row);
  });

  return { rows: deduped, removedCount };
}

export function getAvailableActions(rows = []) {
  const discovered = new Set();

  rows.forEach((row) => {
    if (row.acao) {
      discovered.add(row.acao);
    }
  });

  const customActions = [...discovered].filter((action) => !KNOWN_ACTIONS.includes(action)).sort();
  return [...KNOWN_ACTIONS, ...customActions];
}

export function hydrateActionFilters(rows = [], existingFilters = {}) {
  const actions = getAvailableActions(rows);
  const nextFilters = {};

  actions.forEach((action) => {
    const current = existingFilters[action] ?? {};
    nextFilters[action] = {
      selected: current.selected ?? true,
      min: current.min ?? "",
      max: current.max ?? "",
    };
  });

  return { actions, filters: nextFilters };
}

export function createDefaultFilters(rows = [], savedFilters = null) {
  const { actions, filters } = hydrateActionFilters(rows, savedFilters?.actions ?? {});

  return {
    tempoMin: savedFilters?.tempoMin ?? "",
    tempoMax: savedFilters?.tempoMax ?? "",
    valorVencidoBand: savedFilters?.valorVencidoBand ?? "todos",
    valorRiscoBand: savedFilters?.valorRiscoBand ?? "todos",
    hot: savedFilters?.hot ?? "todos",
    operadores: savedFilters?.operadores ?? 1,
    padrao: savedFilters?.padrao ?? false,
    actions,
    actionFilters: filters,
  };
}

export function applyMainFilters(rows, filters) {
  const thresholdsVencido = buildBandThresholds(rows, "valorVencido");
  const thresholdsRisco = buildBandThresholds(rows, "valorRisco");
  const selectedActions = new Set(
    filters.actions.filter((action) => filters.actionFilters[action]?.selected),
  );

  const filteredRows = rows.filter((row) => {
    if (filters.tempoMin !== "" && (row.tempoAtraso === null || row.tempoAtraso < Number(filters.tempoMin))) {
      return false;
    }

    if (filters.tempoMax !== "" && (row.tempoAtraso === null || row.tempoAtraso > Number(filters.tempoMax))) {
      return false;
    }

    if (filters.hot !== "todos" && normalizeHot(row.hotRaw || row.hot) !== filters.hot) {
      return false;
    }

    if (!matchesBand(row.valorVencido, filters.valorVencidoBand, thresholdsVencido)) {
      return false;
    }

    if (!matchesBand(row.valorRisco, filters.valorRiscoBand, thresholdsRisco)) {
      return false;
    }

    if (filters.padrao && !passesPadraoFilter(row)) {
      return false;
    }

    if (!selectedActions.has(row.acao)) {
      return false;
    }

    const actionFilter = filters.actionFilters[row.acao];
    if (!passesDefasagemRange(row, actionFilter)) {
      return false;
    }

    return true;
  });

  return {
    rows: filteredRows,
    totalClients: filteredRows.length,
    thresholds: {
      vencido: thresholdsVencido,
      risco: thresholdsRisco,
    },
  };
}

export function splitRowsByOperators(rows, operators) {
  const safeOperators = Math.max(1, Number(operators) || 1);
  const chunks = [];
  const baseSize = Math.floor(rows.length / safeOperators);
  const remainder = rows.length % safeOperators;
  let offset = 0;

  for (let index = 0; index < safeOperators; index += 1) {
    const size = baseSize + (index < remainder ? 1 : 0);
    chunks.push(rows.slice(offset, offset + size));
    offset += size;
  }

  return chunks;
}

export function buildExportAoA(rows) {
  const headers = MAIN_EXPORT_COLUMNS.map((column) => column.label);
  const body = rows.map((row) => MAIN_EXPORT_COLUMNS.map((column) => column.value(row)));
  return [headers, ...body];
}

export function buildActionCounts(rows) {
  const counts = new Map();

  rows.forEach((row) => {
    if (!row.acao) {
      return;
    }

    counts.set(row.acao, (counts.get(row.acao) ?? 0) + 1);
  });

  return counts;
}

export function buildAgentReport(rows) {
  const report = new Map();
  const agreementActions = new Set(["ACO", "ACR", "ACP"]);

  rows.forEach((row) => {
    const key = cleanCell(row.agente) || "Sem agente";
    const current = report.get(key) ?? {
      operator: key,
      agreements: 0,
      acionamentos: 0,
    };

    if (agreementActions.has(row.acao)) {
      current.agreements += 1;
    } else {
      current.acionamentos += 1;
    }

    report.set(key, current);
  });

  return [...report.values()].sort(compareOperators);
}

export function summariseNovacao(rows, referenceDate = new Date()) {
  const month = referenceDate.getMonth();
  const year = referenceDate.getFullYear();
  const summary = {
    totalNov: 0,
    currentMonthNov: 0,
    atrasoCount: 0,
    atrasoTotal: 0,
    receberCount: 0,
    receberTotal: 0,
  };

  rows.forEach((row) => {
    if (row.tipo !== "NOV") {
      return;
    }

    summary.totalNov += 1;

    const vencimento = row.vencimentoDate;
    if (!vencimento || vencimento.getMonth() !== month || vencimento.getFullYear() !== year) {
      return;
    }

    summary.currentMonthNov += 1;

    if (Number.isFinite(row.atraso) && row.atraso > 0) {
      summary.atrasoCount += 1;
      summary.atrasoTotal += Number.isFinite(row.valor) ? row.valor : 0;
    } else if (Number.isFinite(row.atraso) && row.atraso < 1) {
      summary.receberCount += 1;
      summary.receberTotal += Number.isFinite(row.valor) ? row.valor : 0;
    }
  });

  return summary;
}

function compactMatrix(matrix) {
  return matrix
    .map((row) => (Array.isArray(row) ? row : [row]))
    .map((row) => row.map((cell) => cleanCell(cell)))
    .filter((row) => row.some((cell) => cell));
}

function rowLooksLikeHeader(row, keywords, minMatches) {
  const normalizedCells = row.map((cell) => normalizeToken(cell));
  const matches = keywords.filter((keyword) =>
    normalizedCells.some((cell) => cell.includes(keyword)),
  );

  return matches.length >= minMatches;
}

function toMainRowObject(sourceRow, lineNumber) {
  const cells = [...sourceRow];
  while (cells.length < 18) {
    cells.push("");
  }

  return {
    lineNumber,
    cpfDisplay: cleanCell(cells[0]),
    cpf: sanitizeCpf(cells[0]),
    nome: cleanCell(cells[1]),
    contrato: cleanCell(cells[2]),
    carne: cleanCell(cells[3]),
    tempoAtrasoRaw: cleanCell(cells[4]),
    tempoAtraso: parseIntegerValue(cells[4]),
    valorVencidoRaw: cleanCell(cells[5]),
    valorVencido: parseBrazilianNumber(cells[5]),
    valorRiscoRaw: cleanCell(cells[6]),
    valorRisco: parseBrazilianNumber(cells[6]),
    telefone: cleanCell(cells[7]),
    hotRaw: cleanCell(cells[8]),
    hot: normalizeHot(cells[8]),
    defasagemRaw: cleanCell(cells[9]),
    defasagem: parseIntegerValue(cells[9]),
    ultimoContato: cleanCell(cells[10]),
    acaoRaw: cleanCell(cells[11]).toUpperCase(),
    acao: cleanCell(cells[11]).toUpperCase(),
    agente: cleanCell(cells[12]),
    historico: cleanCell(cells[13]),
    real: cleanCell(cells[14]),
    vencimento: cleanCell(cells[15]),
    protestado: cleanCell(cells[16]),
    email: cleanCell(cells[17]),
  };
}

function toNovacaoRowObject(sourceRow, lineNumber) {
  const cells = [...sourceRow];
  while (cells.length < 19) {
    cells.push("");
  }

  return {
    lineNumber,
    credor: cleanCell(cells[0]),
    filial: cleanCell(cells[1]),
    codigo: cleanCell(cells[2]),
    matricula: cleanCell(cells[3]),
    nome: cleanCell(cells[4]),
    cpfCnpj: cleanCell(cells[5]),
    tipoRaw: cleanCell(cells[6]),
    tipo: cleanCell(cells[6]).toUpperCase(),
    tituloContrato: cleanCell(cells[7]),
    parcela: cleanCell(cells[8]),
    vencimentoRaw: cleanCell(cells[9]),
    vencimentoDate: parseBrazilianDate(cells[9]),
    valorRaw: cleanCell(cells[10]),
    valor: parseBrazilianNumber(cells[10]),
    atrasoRaw: cleanCell(cells[11]),
    atraso: parseBrazilianNumber(cells[11]),
    inclusao: cleanCell(cells[12]),
    retirada: cleanCell(cells[13]),
    fila: cleanCell(cells[14]),
    sit: cleanCell(cells[15]),
    mercadoria: cleanCell(cells[16]),
    fase: cleanCell(cells[17]),
    uf: cleanCell(cells[18]),
  };
}

function normalizeHot(value) {
  const token = normalizeToken(value);
  if (token === "sim") {
    return "sim";
  }

  if (token === "nao" || token === "não") {
    return "nao";
  }

  return "todos";
}

function buildBandThresholds(rows, fieldName) {
  const values = rows
    .map((row) => row[fieldName])
    .filter((value) => Number.isFinite(value))
    .sort((left, right) => left - right);

  if (!values.length) {
    return null;
  }

  return {
    low: quantile(values, 1 / 3),
    high: quantile(values, 2 / 3),
  };
}

function matchesBand(value, selectedBand, thresholds) {
  if (!thresholds || selectedBand === "todos") {
    return true;
  }

  if (!Number.isFinite(value)) {
    return false;
  }

  if (selectedBand === "baixos") {
    return value <= thresholds.low;
  }

  if (selectedBand === "altos") {
    return value >= thresholds.high;
  }

  return value > thresholds.low && value < thresholds.high;
}

function passesDefasagemRange(row, actionFilter) {
  if (!actionFilter) {
    return false;
  }

  if (actionFilter.min !== "") {
    if (row.defasagem === null || row.defasagem <= Number(actionFilter.min)) {
      return false;
    }
  }

  if (actionFilter.max !== "") {
    if (row.defasagem === null || row.defasagem >= Number(actionFilter.max)) {
      return false;
    }
  }

  return true;
}

function passesPadraoFilter(row) {
  if (PADRAO_EXCLUDED_ACTIONS.has(row.acao)) {
    return false;
  }

  const limit = PADRAO_DEFASAGEM_LIMITS[row.acao];
  if (limit === undefined) {
    return true;
  }

  if (row.defasagem === null) {
    return true;
  }

  return row.defasagem > limit;
}

function quantile(sortedValues, fraction) {
  if (!sortedValues.length) {
    return 0;
  }

  const position = (sortedValues.length - 1) * fraction;
  const baseIndex = Math.floor(position);
  const remainder = position - baseIndex;
  const baseValue = sortedValues[baseIndex];
  const nextValue = sortedValues[baseIndex + 1] ?? baseValue;
  return baseValue + remainder * (nextValue - baseValue);
}

function compareOperators(left, right) {
  const leftNumber = Number(left.operator);
  const rightNumber = Number(right.operator);

  if (Number.isFinite(leftNumber) && Number.isFinite(rightNumber)) {
    return leftNumber - rightNumber;
  }

  return left.operator.localeCompare(right.operator, "pt-BR");
}

function looksMisdecoded(text) {
  if (!text) {
    return false;
  }

  const suspiciousTokens = ["Ã", "â", "�"];
  const matches = suspiciousTokens.reduce(
    (count, token) => count + (text.match(new RegExp(token, "g"))?.length ?? 0),
    0,
  );

  return matches > 4;
}
