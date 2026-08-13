// ═══════════════════════════════════════════════════════════
// CREDIX TIME TRACKING — Google Apps Script Backend
// ═══════════════════════════════════════════════════════════
// Abas usadas nesta planilha (criadas automaticamente quando necessário):
//   Records  — um registro por linha (batida de ponto ou ocorrência)
//   Users    — colaboradores, quem é admin, e os hashes de senha
//   Settings — legado: senha do admin antes de cada um ter a sua
//
// Os hashes NUNCA são devolvidos ao navegador. A conferência de senha
// acontece aqui dentro, na ação "login".
//
// Depois de editar este arquivo: Ctrl+S e
// Implantar › Gerenciar implantações › ✏️ › Versão: Nova versão › Implantar

const SHEET_NAME = "Records";
const RECORD_COLUMNS = ["id", "user", "name", "type", "date", "time", "notes", "timestamp"];
const USER_COLUMNS = ["key", "name", "admin", "pwdHash", "wordHash"];

const DEFAULT_USERS = [
  { key: "evelyn", name: "Evelyn", admin: false, pwdHash: "", wordHash: "" },
  { key: "patrick", name: "Patrick", admin: false, pwdHash: "", wordHash: "" },
  { key: "elisangela", name: "Elisangela", admin: false, pwdHash: "", wordHash: "" },
  { key: "enzo", name: "Enzo", admin: true, pwdHash: "", wordHash: "" }
];

function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

// Só as gravações precisam da fila. Antes tudo esperava, e uma leitura durante
// uma gravação de outra pessoa ficava parada sem motivo.
const WRITE_ACTIONS = ["addRecord", "updateRecord", "deleteRecord", "saveUsers",
                       "setPassword", "setSecurityWord", "resetPassword"];

function handleRequest(e) {
  const params = (e && e.parameter) || {};
  let body;
  try {
    body = (e && e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : {};
  } catch (err) {
    return json_({ success: false, error: "Malformed request body" });
  }
  const action = params.action || body.action || null;

  let lock = null;
  if (WRITE_ACTIONS.indexOf(action) > -1) {
    lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) {
      return json_({ success: false, error: "Planilha ocupada, tente novamente." });
    }
  }

  try {

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getRecordSheet_(ss);

    // ── REGISTROS ──────────────────────────────────────────
    // Ler exige senha válida. A página é pública e o endereço do backend está
    // no código-fonte dela, então sem isso qualquer um com o link leria a
    // jornada de trabalho de todo mundo.
    if (action === "getRecords") {
      if (!authUser_(readUsers_(ss), body)) {
        return json_({ success: false, authRequired: true, error: "Sign in required to read the records." });
      }
      return json_({ success: true, records: readRecords_(sheet) });
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
    // Devolve só o que a tela precisa. Os hashes não saem daqui.
    if (action === "getUsers") {
      const users = readUsers_(ss).map(function (u) {
        return {
          key: u.key,
          name: u.name,
          admin: u.admin,
          hasPassword: !!u.pwdHash,
          hasWord: !!u.wordHash
        };
      });
      return json_({ success: true, users: users });
    }

    if (action === "saveUsers") {
      const incoming = body.users;
      if (!Array.isArray(incoming) || !incoming.length) {
        return json_({ success: false, error: "Empty user list refused" });
      }
      // preserva as senhas de quem continua na lista
      const current = readUsers_(ss);
      const byKey = {};
      current.forEach(function (u) { byKey[u.key] = u; });
      const merged = incoming.map(function (u) {
        const key = String(u.key).trim();
        const old = byKey[key] || {};
        return {
          key: key,
          name: String(u.name).trim(),
          admin: u.admin === true,
          pwdHash: old.pwdHash || "",
          wordHash: old.wordHash || ""
        };
      });
      writeUsers_(ss, merged);
      return json_({ success: true });
    }

    // ── AUTENTICAÇÃO ───────────────────────────────────────
    if (action === "login") {
      const user = findUser_(ss, body.user);
      if (!user) return json_({ success: false, error: "This collaborator no longer exists." });
      if (!user.pwdHash) {
        // primeiro acesso: a tela pede para criar a senha
        return json_({ success: false, needsPassword: true, error: "No password set yet." });
      }
      if (String(body.pwdHash || "") !== user.pwdHash) {
        return json_({ success: false, error: "Incorrect password." });
      }
      // Os registros vão junto: a sessão acabou de ser autenticada, então isso
      // poupa uma ida ao servidor logo depois — e cada ida custa ~1-2 segundos.
      return json_({
        success: true,
        user: { key: user.key, name: user.name, admin: user.admin, hasWord: !!user.wordHash },
        records: readRecords_(sheet)
      });
    }

    // Troca de senha. Precisa de uma prova, conferida aqui no servidor:
    // a senha atual, a palavra de segurança, ou a senha de um admin.
    if (action === "setPassword") {
      const users = readUsers_(ss);
      const user = pickUser_(users, body.user);
      if (!user) return json_({ success: false, error: "This collaborator no longer exists." });
      const newHash = String(body.pwdHash || "");
      if (!newHash) return json_({ success: false, error: "Missing new password." });

      const allowed = authorize_(users, user, body);
      if (allowed !== true) return json_({ success: false, error: allowed });

      user.pwdHash = newHash;
      if (body.wordHash && !user.wordHash) user.wordHash = String(body.wordHash);
      writeUsers_(ss, users);
      return json_({ success: true });
    }

    if (action === "setSecurityWord") {
      const users = readUsers_(ss);
      const user = pickUser_(users, body.user);
      if (!user) return json_({ success: false, error: "This collaborator no longer exists." });
      const newWord = String(body.wordHash || "");
      if (!newWord) return json_({ success: false, error: "Missing security word." });

      const allowed = authorize_(users, user, body);
      if (allowed !== true) return json_({ success: false, error: allowed });

      user.wordHash = newWord;
      writeUsers_(ss, users);
      return json_({ success: true });
    }

    // Admin zera a senha de alguém — a pessoa cria uma nova no próximo acesso.
    if (action === "resetPassword") {
      const users = readUsers_(ss);
      const user = pickUser_(users, body.user);
      if (!user) return json_({ success: false, error: "This collaborator no longer exists." });
      if (!isAdminProof_(users, body)) {
        return json_({ success: false, error: "Admin password required." });
      }
      user.pwdHash = "";
      writeUsers_(ss, users);
      return json_({ success: true });
    }

    // getSettings/saveSettings foram removidos de propósito. Eram a senha única
    // do admin; uma aba com o código antigo em cache liria "nenhuma senha" e
    // voltaria a aceitar a senha padrão. Sem a ação, aquela aba falha e usa o
    // próprio armazenamento local, sem abrir brecha.
    return json_({ success: false, error: "Unknown action" });

  } catch (err) {
    return json_({ success: false, error: err.toString() });
  } finally {
    if (lock) lock.releaseLock();
  }
}

// ═══════════════════════════════════════════════════════════
// AUTORIZAÇÃO
// ═══════════════════════════════════════════════════════════
// Devolve true, ou a mensagem do motivo da recusa.
function authorize_(users, user, body) {
  // primeiro acesso: ninguém definiu senha ainda
  if (!user.pwdHash) return true;
  if (body.currentPwdHash && String(body.currentPwdHash) === user.pwdHash) return true;
  if (body.wordHash && user.wordHash && String(body.wordHash) === user.wordHash) return true;
  if (isAdminProof_(users, body)) return true;
  return "Current password, security word or admin password required.";
}

// Quem é a pessoa da requisição, conferindo a senha aqui no servidor.
// Devolve o usuário, ou null se a senha não bate.
function authUser_(users, body) {
  const user = pickUser_(users, body.user);
  if (!user || !user.pwdHash) return null;
  return String(body.pwdHash || "") === user.pwdHash ? user : null;
}

function isAdminProof_(users, body) {
  if (!body.adminKey || !body.adminPwdHash) return false;
  const admin = pickUser_(users, body.adminKey);
  if (!admin || !admin.admin || !admin.pwdHash) return false;
  return String(body.adminPwdHash) === admin.pwdHash;
}

function pickUser_(users, key) {
  const wanted = String(key || "").trim();
  if (!wanted) return null;
  for (let i = 0; i < users.length; i++) {
    if (users[i].key === wanted) return users[i];
  }
  return null;
}

function findUser_(ss, key) {
  return pickUser_(readUsers_(ss), key);
}

// ═══════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════
function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function readRecords_(sheet) {
  const data = sheet.getDataRange().getValues();
  const headerRow = data[0];
  return data.slice(1)
    .filter(function (row) { return String(row[0]).trim() !== ""; })
    .map(function (row) {
      const obj = {};
      headerRow.forEach(function (h, i) { obj[h] = row[i]; });
      return obj;
    });
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
  if (!sh) return migrateLegacySecret_(ss, DEFAULT_USERS.map(cloneUser_));

  const rows = sh.getDataRange().getValues();
  const header = rows.length ? rows[0].map(function (h) { return String(h).trim(); }) : [];
  const col = {};
  USER_COLUMNS.forEach(function (name) { col[name] = header.indexOf(name); });

  const users = rows.slice(1)
    .filter(function (row) { return String(row[0]).trim() !== ""; })
    .map(function (row) {
      const pick = function (name, fallbackIndex) {
        const i = col[name] > -1 ? col[name] : fallbackIndex;
        return i > -1 && i < row.length ? row[i] : "";
      };
      return {
        key: String(pick("key", 0)).trim(),
        name: String(pick("name", 1) || pick("key", 0)).trim(),
        admin: TRUTHY.indexOf(String(pick("admin", 2)).trim().toLowerCase()) > -1,
        pwdHash: String(pick("pwdHash", 3) || "").trim(),
        wordHash: String(pick("wordHash", 4) || "").trim()
      };
    });

  if (!users.length) return migrateLegacySecret_(ss, DEFAULT_USERS.map(cloneUser_));
  return migrateLegacySecret_(ss, users);
}

function cloneUser_(u) {
  return { key: u.key, name: u.name, admin: u.admin, pwdHash: u.pwdHash || "", wordHash: u.wordHash || "" };
}

// A senha do admin morava na aba Settings, quando era uma só para todos.
// Na primeira leitura ela passa para o primeiro admin sem senha própria, e a
// aba Settings é esvaziada — senão, com dois admins, o segundo herdaria a mesma
// senha na leitura seguinte.
function migrateLegacySecret_(ss, users) {
  const legacy = readSettings_(ss);
  if (!legacy.pwdHash) return users;

  let target = null;
  for (let i = 0; i < users.length; i++) {
    if (users[i].admin && !users[i].pwdHash) { target = users[i]; break; }
  }
  if (target) {
    target.pwdHash = legacy.pwdHash;
    if (!target.wordHash && legacy.wordHash) target.wordHash = legacy.wordHash;
    writeUsers_(ss, users);
  }
  clearLegacySecret_(ss);
  return users;
}

function clearLegacySecret_(ss) {
  const sh = ss.getSheetByName("Settings");
  if (!sh) return;
  sh.clear();
  sh.appendRow(["key", "value"]);
  sh.appendRow(["migrated", "os hashes agora ficam na aba Users, um por pessoa"]);
}

function writeUsers_(ss, users) {
  let sh = ss.getSheetByName("Users");
  if (!sh) sh = ss.insertSheet("Users");
  sh.clear();
  sh.appendRow(USER_COLUMNS);
  users.forEach(function (u) {
    sh.appendRow([
      asText_(u.key),
      asText_(u.name),
      u.admin === true,
      asText_(u.pwdHash || ""),
      asText_(u.wordHash || "")
    ]);
  });
  sh.setFrozenRows(1);
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

// Roda esta função no editor (▶ Run) para descobrir a planilha ligada ao script.
function qualPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Nome: " + ss.getName());
  Logger.log("Link: " + ss.getUrl());
}
