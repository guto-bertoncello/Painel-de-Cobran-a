import test from "node:test";
import assert from "node:assert/strict";
import {
  applyMainFilters,
  createDefaultFilters,
  deduplicateMainRows,
  matrixToMainRows,
  summariseNovacao,
} from "../assets/logic.js";

function makeMainRow(overrides = {}) {
  return {
    cpfDisplay: overrides.cpfDisplay ?? "123.456.789-00",
    cpf: overrides.cpf ?? "12345678900",
    nome: overrides.nome ?? "Cliente Teste",
    contrato: overrides.contrato ?? "CONTRATO-1",
    carne: overrides.carne ?? "",
    tempoAtrasoRaw: overrides.tempoAtrasoRaw ?? "20",
    tempoAtraso: overrides.tempoAtraso ?? 20,
    valorVencidoRaw: overrides.valorVencidoRaw ?? "R$ 100,00",
    valorVencido: overrides.valorVencido ?? 100,
    valorRiscoRaw: overrides.valorRiscoRaw ?? "R$ 80,00",
    valorRisco: overrides.valorRisco ?? 80,
    telefone: overrides.telefone ?? "11999999999",
    hotRaw: overrides.hotRaw ?? "SIM",
    hot: overrides.hot ?? "sim",
    defasagemRaw: overrides.defasagemRaw ?? "5",
    defasagem: overrides.defasagem ?? 5,
    ultimoContato: overrides.ultimoContato ?? "",
    acaoRaw: overrides.acaoRaw ?? "CPC",
    acao: overrides.acao ?? "CPC",
    agente: overrides.agente ?? "1",
    historico: overrides.historico ?? "",
    real: overrides.real ?? "",
    vencimento: overrides.vencimento ?? "",
    protestado: overrides.protestado ?? "",
    email: overrides.email ?? "",
  };
}

test("detecta arquivo principal de exemplo vazio por colunas", () => {
  const matrix = Array.from({ length: 21 }, (_, index) => ["Coluna", String.fromCharCode(97 + index)]);
  const parsed = matrixToMainRows(matrix);
  assert.equal(parsed.rows.length, 0);
  assert.match(parsed.warning, /modelo de colunas/i);
});

test("remove CPF duplicado mantendo a primeira ocorrência", () => {
  const rows = [
    makeMainRow({ cpf: "11111111111", nome: "Primeiro" }),
    makeMainRow({ cpf: "11111111111", nome: "Duplicado" }),
    makeMainRow({ cpf: "22222222222", nome: "Segundo" }),
  ];

  const result = deduplicateMainRows(rows);
  assert.equal(result.rows.length, 2);
  assert.equal(result.removedCount, 1);
  assert.equal(result.rows[0].nome, "Primeiro");
  assert.equal(result.rows[1].nome, "Segundo");
});

test("aplica defasagem exclusiva por ação e filtro padrão", () => {
  const rows = [
    makeMainRow({ cpf: "1", acao: "CPC", defasagem: 3, defasagemRaw: "3" }),
    makeMainRow({ cpf: "2", acao: "CPC", defasagem: 5, defasagemRaw: "5" }),
    makeMainRow({ cpf: "3", acao: "NEG", defasagem: 4, defasagemRaw: "4" }),
    makeMainRow({ cpf: "4", acao: "NEG", defasagem: 5, defasagemRaw: "5" }),
  ];

  const filters = createDefaultFilters(rows);
  filters.actions = ["CPC", "NEG"];
  filters.actionFilters.CPC = { selected: true, min: "3", max: "6" };
  filters.actionFilters.NEG = { selected: true, min: "", max: "" };

  let result = applyMainFilters(rows, filters);
  assert.deepEqual(
    result.rows.map((row) => row.cpf),
    ["2", "3", "4"],
  );

  filters.padrao = true;
  result = applyMainFilters(rows, filters);
  assert.deepEqual(
    result.rows.map((row) => row.cpf),
    ["2", "4"],
  );
});

test("soma novação do mês atual entre atraso e a receber", () => {
  const referenceDate = new Date(2026, 3, 19);
  const rows = [
    {
      tipo: "NOV",
      vencimentoDate: new Date(2026, 3, 10),
      valor: 120,
      atraso: 3,
    },
    {
      tipo: "NOV",
      vencimentoDate: new Date(2026, 3, 11),
      valor: 80,
      atraso: 0,
    },
    {
      tipo: "NOV",
      vencimentoDate: new Date(2026, 2, 11),
      valor: 999,
      atraso: 2,
    },
    {
      tipo: "CDC",
      vencimentoDate: new Date(2026, 3, 11),
      valor: 999,
      atraso: 2,
    },
  ];

  const summary = summariseNovacao(rows, referenceDate);
  assert.equal(summary.totalNov, 3);
  assert.equal(summary.currentMonthNov, 2);
  assert.equal(summary.atrasoCount, 1);
  assert.equal(summary.atrasoTotal, 120);
  assert.equal(summary.receberCount, 1);
  assert.equal(summary.receberTotal, 80);
});
