const SPREADSHEET_ID   = "1bYkrUeoGiJrT75r3T9l7EuSmt4HpP7t6dyZVQ0C3svs";
const SHEET_CARTEIRA   = "Carteira";
const SHEET_INTERACOES = "Interacoes";

const COLS = {
  etapa:          "Etapa",
  canal:          "Canal",
  produto:        "Produto",
  status:         "Status Canal",
  respOps:        "Resp. Pa.OPS",
  respOnb:        "Resp. Onboarding",
  inicioHml:      "Início da HML",
  fimHml:         "Fim da HML",
  statusHml:      "Status HML",
  inicioProducao: "Data Início Produção",
  conexao:        "Forma de conexão",
  cm:             "Client Manager",
  vertical:       "Vertical",
  sub:            "Subcategoria",
  update:         "Update geral",
};

// ------------------------------------------------------------
// ENTRY POINT
// ------------------------------------------------------------
function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("CRM Canais Parceiros")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ------------------------------------------------------------
// CONTROLE DE ACESSO — retorna e-mail do usuário logado
// ------------------------------------------------------------
function getEmail() {
  return Session.getActiveUser().getEmail();
}

// ------------------------------------------------------------
// DASHBOARD
// ------------------------------------------------------------
function getDashboard() {
  const rows = getParceiros();
  const countContains = (arr, key, val) =>
    arr.filter(r => (r[key] || "").includes(val)).length;
  const byKey = (arr, key) =>
    arr.reduce((acc, r) => {
      const v = r[key] || "Sem info";
      acc[v] = (acc[v] || 0) + 1;
      return acc;
    }, {});

  return {
    total:        rows.length,
    ativos:       rows.filter(r => r["Status Canal"] === "Ativo").length,
    onboarding:   rows.filter(r => r["Etapa"] === "Onboarding").length,
    hmlAndamento: rows.filter(r => r["Status HML"] === "Em andamento").length,
    porStatus:    byKey(rows, "Status Canal"),
    porEtapa:     byKey(rows, "Etapa"),
    porHml:       byKey(rows, "Status HML"),
    porVertical:  byKey(rows, "Vertical"),
    porSub:       byKey(rows, "Subcategoria"),
    porOps:       byKey(rows, "Resp. Pa.OPS"),
    porProduto: {
      "Prestamista":        countContains(rows, "Produto", "Prestamista"),
      "Vida":               countContains(rows, "Produto", "Vida"),
      "Residencial":        countContains(rows, "Produto", "Residencial"),
      "Celular":            countContains(rows, "Produto", "Celular"),
      "Riscos Diversos":    countContains(rows, "Produto", "Riscos Diversos"),
      "Acidentes Pessoais": countContains(rows, "Produto", "Acidentes Pessoais"),
      "GAE":                countContains(rows, "Produto", "GAE"),
      "Empresarial":        countContains(rows, "Produto", "Empresarial"),
      "Bem-estar":          countContains(rows, "Produto", "Bem-estar"),
      "Cuida+":             countContains(rows, "Produto", "Cuida+"),
    }
  };
}

// ------------------------------------------------------------
// LEITURA — Carteira
// ------------------------------------------------------------
function getParceiros() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_CARTEIRA);
  if (!sh) throw new Error('Aba "' + SHEET_CARTEIRA + '" não encontrada na planilha.');
  const rows = sh.getDataRange().getValues();
  const headers = rows[0];
  return rows.slice(1)
    .filter(r => r[headers.indexOf(COLS.canal)])
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = r[i] !== undefined ? String(r[i]) : "");
      return obj;
    });
}

// ------------------------------------------------------------
// LEITURA — Interações por canal
// ------------------------------------------------------------
function getInteracoes(nomeCanal) {
  if (!nomeCanal) return [];
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_INTERACOES);
  if (!sh) return [];
  const rows = sh.getDataRange().getValues();
  const headers = rows[0];
  return rows.slice(1)
    .map(r => { const o = {}; headers.forEach((h, i) => o[h] = String(r[i] || "")); return o; })
    .filter(r => r["Canal"] === nomeCanal);
}

// ------------------------------------------------------------
// ESCRITA — Parceiro (novo ou edição)
// ------------------------------------------------------------
function salvarParceiro(dados) {
  if (!dados) return { ok: false, erro: "Nenhum dado recebido." };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_CARTEIRA);
  const rows = sh.getDataRange().getValues();
  const headers = rows[0];

  // _rowIndex é 0-based (índice no array de dados, sem contar o header)
  // Linha real na sheet = _rowIndex + 2 (linha 1 = header, dados a partir da linha 2)
  if (dados._rowIndex !== null && dados._rowIndex !== undefined && dados._rowIndex !== "") {
    const sheetRow = parseInt(dados._rowIndex) + 2;
    Object.entries(COLS).forEach(([, colName]) => {
      const j = headers.indexOf(colName);
      if (j >= 0 && dados[colName] !== undefined) {
        sh.getRange(sheetRow, j + 1).setValue(dados[colName]);
      }
    });
    return { ok: true };
  }

  // Novo canal
  const row = headers.map(h => dados[h] || "");
  var lastRow = sh.getLastRow();
sh.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);
  return { ok: true };
}

// ------------------------------------------------------------
// EXCLUSÃO — Remove a linha do parceiro na planilha
// ------------------------------------------------------------
function excluirParceiro(rowIndex) {
  if (rowIndex === undefined || rowIndex === null) {
    return { ok: false, erro: "Índice não informado." };
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_CARTEIRA);

  // rowIndex é 0-based (índice no array de dados, sem contar o header)
  // Linha real na sheet = rowIndex + 2 (linha 1 = header, dados a partir da linha 2)
  const sheetRow = parseInt(rowIndex) + 2;
  const lastRow = sh.getLastRow();

  if (sheetRow < 2 || sheetRow > lastRow) {
    return { ok: false, erro: "Linha fora dos limites: " + sheetRow };
  }

  sh.deleteRow(sheetRow);
  return { ok: true };
}

// ------------------------------------------------------------
// ESCRITA — Interação
// ------------------------------------------------------------
function salvarInteracao(dados) {
  if (!dados || !dados["Canal"]) return { ok: false, erro: "Dados inválidos." };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_INTERACOES);
  if (!sh) {
    sh = ss.insertSheet(SHEET_INTERACOES);
    sh.appendRow(["ID", "Canal", "Data", "Tipo", "Descrição", "Próximo Passo", "Data Follow-up", "Registrado por"]);
  }

  sh.appendRow([
    "I" + Date.now(),
    dados["Canal"],
    new Date().toLocaleDateString("pt-BR"),
    dados["Tipo"] || "",
    dados["Descrição"] || "",
    dados["Próximo Passo"] || "",
    dados["Data Follow-up"] || "",
    Session.getActiveUser().getEmail() || "—"
  ]);
  return { ok: true };
}
