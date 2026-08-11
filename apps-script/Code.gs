// ═══════════════════════════════════════════════════════════
// CREDIX TIME TRACKING — Google Apps Script Backend
// ═══════════════════════════════════════════════════════════
// Abas usadas nesta planilha (criadas automaticamente quando necessário):
//   Records  — um registro por linha (batida de ponto ou ocorrência)
//   Users    — colaboradores e quem é admin
//   Settings — hash da senha do admin e da palavra de segurança
//
// Depois de editar este arquivo: Ctrl+S e
// Implantar › Gerenciar implantações › ✏️ › Versão: Nova versão › Implantar

const SHEET_NAME = "Records";
const RECORD_COLUMNS = ["id", "user", "name", "type", "date", "time", "notes", "timestamp"];

const DEFAULT_USERS = [
  { key: "evelyn", name: "Evelyn", admin: false },
  { key: "patrick", name: "Patrick", admin: false },
  { key: "elisangela", name: "Elisangela", admin: false },
  { key: "enzo", name: "Enzo", admin: true }
];

function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  // impede que duas pessoas gravando ao mesmo tempo embaralhem as linhas
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    return json_({ success: false, error: "Planilha ocupada, tente novamente." });
  }

  try {
    const params = (e && e.parameter) || {};
    const body = (e && e.postData && e.postData.contents)
      ? JSON.parse(e.postData.contents)
      : {};
    const action = params.action || body.action || null;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getRecordSheet_(ss);

    // ── REGISTROS ──────────────────────────────────────────
    if (action === "getRecords") {
      const data = sheet.getDataRange().getValues();
      const headerRow = data[0];
      const records = data.slice(1)
        .filter(function (row) { return String(row[0]).trim() !== ""; })
        .map(function (row) {
          const obj = {};
          headerRow.forEach(function (h, i) { obj[h] = row[i]; });
          return obj;
        });
      return json_({ success: true, records: records });
    }

    if (action === "addRecord") {
      const rec = body.record || {};
      if (!rec.id) return json_({ success: false, error: "Record without id" });
      sheet.appendRow(recordToRow_(rec));
      return json_({ success: true });
    }

    if (action === "updateRecord") {
      const rec = body.record || {};
      const rowIndex = findRowById_(sheet, rec.id);
      if (rowIndex === -1) return json_({ success: false, error: "Record not found" });
      sheet.getRange(rowIndex, 1, 1, RECORD_COLUMNS.length).setValues([recordToRow_(rec)]);
      return json_({ success: true });
    }

    if (action === "deleteRecord") {
      const rowIndex = findRowById_(sheet, body.id);
      if (rowIndex === -1) return json_({ success: false, error: "Record not found" });
      sheet.deleteRow(rowIndex);
      return json_({ success: true });
    }

    // ── COLABORADORES ──────────────────────────────────────
    if (action === "getUsers") {
      return json_({ success: true, users: readUsers_(ss) });
    }

    if (action === "saveUsers") {
      const users = body.users;
      // nunca apagar a lista inteira por causa de uma requisição vazia
      if (!Array.isArray(users) || !users.length) {
        return json_({ success: false, error: "Empty user list refused" });
      }
      let userSheet = ss.getSheetByName("Users");
      if (!userSheet) userSheet = ss.insertSheet("Users");
      userSheet.clearContents();
      userSheet.appendRow(["key", "name", "admin"]);
      users.forEach(function (u) {
        userSheet.appendRow([String(u.key), String(u.name), u.admin === true]);
      });
      return json_({ success: true });
    }

    // ── SENHA DO ADMIN / PALAVRA DE SEGURANÇA ──────────────
    // Chegam já como hash SHA-256 — a senha em texto puro nunca é gravada.
    if (action === "getSettings") {
      return json_({ success: true, settings: readSettings_(ss) });
    }

    if (action === "saveSettings") {
      writeSettings_(ss, body.settings || {});
      return json_({ success: true });
    }

    return json_({ success: false, error: "Unknown action" });

  } catch (err) {
    return json_({ success: false, error: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

// ═══════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════
function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getRecordSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(RECORD_COLUMNS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// O apóstrofo força a célula a ficar como texto. Sem isso o Sheets converte
// "11/08/2026" conforme o idioma da planilha e "20:42:07" virava um horário
// serial (1899-12-30T...), que era o que chegava torto no app.
function asText_(value) {
  const s = (value === null || value === undefined) ? "" : String(value);
  return s === "" ? "" : "'" + s;
}

function recordToRow_(rec) {
  return [
    asText_(rec.id),
    asText_(rec.user),
    asText_(rec.name),
    asText_(rec.type),
    asText_(rec.date),
    asText_(rec.time),
    asText_(rec.notes || ""),
    asText_(rec.timestamp)
  ];
}

// devolve o número da linha na planilha (1-based) ou -1
function findRowById_(sheet, id) {
  const wanted = String(id || "").trim();
  if (!wanted) return -1;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === wanted) return i + 1;
  }
  return -1;
}

const TRUTHY = ["true", "verdadeiro", "sim", "yes", "1", "x"];

function readUsers_(ss) {
  const sh = ss.getSheetByName("Users");
  if (!sh) return DEFAULT_USERS;
  const users = sh.getDataRange().getValues().slice(1)
    .filter(function (row) { return String(row[0]).trim() !== ""; })
    .map(function (row) {
      return {
        key: String(row[0]).trim(),
        name: String(row[1] || row[0]).trim(),
        admin: TRUTHY.indexOf(String(row[2]).trim().toLowerCase()) > -1
      };
    });
  return users.length ? users : DEFAULT_USERS;
}

function readSettings_(ss) {
  const out = { pwdHash: "", wordHash: "" };
  const sh = ss.getSheetByName("Settings");
  if (!sh) return out;
  sh.getDataRange().getValues().slice(1).forEach(function (row) {
    if (row[0]) out[String(row[0]).trim()] = String(row[1] || "").trim();
  });
  return out;
}

function writeSettings_(ss, settings) {
  const current = readSettings_(ss);
  // grava só o que veio preenchido, para trocar a senha não apagar a palavra
  const next = {
    pwdHash: settings.pwdHash ? String(settings.pwdHash) : current.pwdHash,
    wordHash: settings.wordHash ? String(settings.wordHash) : current.wordHash
  };
  let sh = ss.getSheetByName("Settings");
  if (!sh) sh = ss.insertSheet("Settings");
  sh.clear();
  sh.appendRow(["key", "value"]);
  sh.appendRow(["pwdHash", next.pwdHash]);
  sh.appendRow(["wordHash", next.wordHash]);
}

// Roda esta função no editor (▶ Run) para descobrir a planilha ligada ao script.
function qualPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Nome: " + ss.getName());
  Logger.log("Link: " + ss.getUrl());
}
