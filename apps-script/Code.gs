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

// Pedidos de correção de ponto. Ficam aqui e não em Records, para um pedido
// pendente não entrar em nenhum cálculo de horas antes de ser aprovado. Ao
// aprovar, é este backend que grava a batida em Records — e a linha do pedido
// permanece como histórico de quem pediu, por quê, e quem aprovou.
const REQUEST_COLUMNS = ["id", "user", "name", "action", "targetId", "type", "date",
                         "time", "reason", "doc", "status", "createdAt", "reviewedBy", "reviewedAt"];
const PUNCH_TYPES_BE = ["Clock In", "Lunch Out", "Lunch In", "Clock Out"];
const OCCURRENCE_TYPES_BE = ["Medical Leave", "Holiday", "Day Off", "Vacation"];

// Uma ocorrência de dia inteiro é neutra no saldo e é registrada direto. Uma com
// janela de horário credita horas trabalhadas, então passa por aprovação — é da
// mesma natureza de uma correção de batida.
function parseWindowBE_(time) {
  const m = String(time || "").match(/^(\d{1,2}:\d{2}(?::\d{2})?)\s*-\s*(\d{1,2}:\d{2}(?::\d{2})?)$/);
  if (!m) return null;
  const secs = function (t) {
    const p = String(t).split(":");
    return (parseInt(p[0], 10) || 0) * 3600 + (parseInt(p[1], 10) || 0) * 60 + (parseInt(p[2], 10) || 0);
  };
  return secs(m[2]) > secs(m[1]) ? { from: m[1], to: m[2] } : null;
}

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
                       "setPassword", "setSecurityWord", "resetPassword",
                       "createRequest", "reviewRequest"];

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
    // Opening the Records sheet costs a read, and half the actions never touch
    // it — getUsers, called on every page load, among them. Open it on demand.
    let recordSheet = null;
    const sheet = function () {
      if (!recordSheet) recordSheet = getRecordSheet_(ss);
      return recordSheet;
    };

    // ── REGISTROS ──────────────────────────────────────────
    // Ler exige senha válida. A página é pública e o endereço do backend está
    // no código-fonte dela, então sem isso qualquer um com o link leria a
    // jornada de trabalho de todo mundo.
    if (action === "getRecords") {
      if (!authUser_(readUsers_(ss), body)) {
        return json_({ success: false, authRequired: true, error: "Sign in required to read the records." });
      }
      return json_({ success: true, records: readRecords_(sheet()) });
    }

    if (action === "addRecord") {
      const rec = body.record || {};
      if (!rec.id) return json_({ success: false, error: "Record without id" });
      sheet().appendRow(recordToRow_(rec));
      return json_({ success: true });
    }

    if (action === "updateRecord") {
      const rec = body.record || {};
      const rowIndex = findRowById_(sheet(), rec.id);
      if (rowIndex === -1) return json_({ success: false, error: "Record not found" });
      sheet().getRange(rowIndex, 1, 1, RECORD_COLUMNS.length).setValues([recordToRow_(rec)]);
      return json_({ success: true });
    }

    if (action === "deleteRecord") {
      const rowIndex = findRowById_(sheet(), body.id);
      if (rowIndex === -1) return json_({ success: false, error: "Record not found" });
      sheet().deleteRow(rowIndex);
      return json_({ success: true });
    }

    // ── PEDIDOS DE CORREÇÃO ────────────────────────────────
    if (action === "getRequests") {
      const me = authUser_(readUsers_(ss), body);
      if (!me) return json_({ success: false, authRequired: true, error: "Sign in required." });
      const all = readRequests_(ss);
      // cada um vê os seus; o admin vê todos, porque é quem aprova
      return json_({ success: true, requests: me.admin ? all : all.filter(r => r.user === me.key) });
    }

    if (action === "createRequest") {
      const me = authUser_(readUsers_(ss), body);
      if (!me) return json_({ success: false, authRequired: true, error: "Sign in required." });

      const req = body.request || {};
      // ninguém pede correção no nome de outra pessoa
      if (String(req.user || "") !== me.key) {
        return json_({ success: false, error: "You can only request corrections for yourself." });
      }
      if (!req.id || !req.date || !req.time) {
        return json_({ success: false, error: "Date and time are required." });
      }
      if (!String(req.reason || "").trim()) {
        return json_({ success: false, error: "A reason is required." });
      }

      const isOccurrence = String(req.action || "") === "occurrence";
      let targetId = "";

      if (isOccurrence) {
        if (OCCURRENCE_TYPES_BE.indexOf(String(req.type)) === -1) {
          return json_({ success: false, error: "Invalid occurrence type." });
        }
        // só a de janela passa por aqui: a de dia inteiro é neutra e vai direto
        if (!parseWindowBE_(req.time)) {
          return json_({ success: false, error: "An occurrence request needs a valid time window." });
        }
      } else {
        if (PUNCH_TYPES_BE.indexOf(String(req.type)) === -1) {
          return json_({ success: false, error: "Invalid punch type." });
        }
        targetId = String(req.targetId || "");
        if (targetId) {
          // só faz sentido alterar uma batida que existe e é da própria pessoa
          const target = findRecordById_(sheet(), targetId);
          if (!target) return json_({ success: false, error: "The punch to change no longer exists." });
          if (String(target.user) !== me.key) {
            return json_({ success: false, error: "That punch belongs to someone else." });
          }
        }
      }

      appendRequest_(ss, {
        id: String(req.id), user: me.key, name: me.name,
        action: isOccurrence ? "occurrence" : (targetId ? "edit" : "add"),
        targetId: targetId,
        type: String(req.type), date: String(req.date), time: String(req.time),
        reason: String(req.reason).trim(), doc: String(req.doc || "").trim(),
        status: "pending",
        createdAt: String(req.createdAt || ""), reviewedBy: "", reviewedAt: ""
      });
      return json_({ success: true });
    }

    if (action === "reviewRequest") {
      const users = readUsers_(ss);
      if (!isAdminProof_(users, body)) {
        return json_({ success: false, error: "Admin password required." });
      }
      const decision = String(body.decision || "");
      if (decision !== "approve" && decision !== "reject") {
        return json_({ success: false, error: "Decision must be approve or reject." });
      }

      const requests = readRequests_(ss);
      const req = pickRequest_(requests, body.id);
      if (!req) return json_({ success: false, error: "Request not found." });
      if (req.status !== "pending") {
        return json_({ success: false, error: "This request was already " + req.status + "." });
      }

      if (decision === "approve") {
        if (req.action === "edit") {
          const rowIndex = findRowById_(sheet(), req.targetId);
          if (rowIndex === -1) {
            return json_({ success: false, error: "The punch to change no longer exists." });
          }
          const current = findRecordById_(sheet(), req.targetId);
          sheet().getRange(rowIndex, 1, 1, RECORD_COLUMNS.length).setValues([recordToRow_({
            id: current.id, user: current.user, name: current.name,
            type: req.type, date: req.date, time: req.time,
            // A justificativa tem de ficar na própria batida, igual acontece
            // quando o pedido é de adição. Antes ela só ficava na aba Requests,
            // então uma alteração aprovada aparecia sem explicação em todo lugar
            // onde a batida é mostrada — enquanto uma adição aparecia com ela.
            notes: withReason_(current.notes, req.reason),
            timestamp: current.timestamp
          })]);
        } else if (req.action === "occurrence") {
          // Igual ao que uma ocorrência registrada direto grava: o link do
          // documento e a justificativa, na mesma ordem.
          sheet().appendRow(recordToRow_({
            id: req.id, user: req.user, name: req.name,
            type: req.type, date: req.date, time: req.time,
            notes: [req.doc, req.reason].filter(function (v) { return !!v; }).join(" "),
            timestamp: String(body.reviewedAt || req.createdAt || "")
          }));
        } else {
          sheet().appendRow(recordToRow_({
            id: req.id, user: req.user, name: req.name,
            type: req.type, date: req.date, time: req.time,
            notes: withReason_("", req.reason),
            timestamp: String(body.reviewedAt || req.createdAt || "")
          }));
        }
      }

      req.status = decision === "approve" ? "approved" : "rejected";
      req.reviewedBy = String(body.adminKey || "");
      req.reviewedAt = String(body.reviewedAt || "");
      writeRequests_(ss, requests);
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
      const requests = readRequests_(ss);
      return json_({
        success: true,
        user: { key: user.key, name: user.name, admin: user.admin, hasWord: !!user.wordHash },
        records: readRecords_(sheet()),
        requests: user.admin ? requests : requests.filter(r => r.user === user.key)
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

// Anexa a justificativa da correção ao que a batida já tinha — um link do Drive
// ou uma observação não podem ser perdidos para dar lugar a ela.
function withReason_(existing, reason) {
  const note = "Correction — " + String(reason == null ? "" : reason).trim();
  const before = String(existing == null ? "" : existing).trim();
  return before ? before + " · " + note : note;
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

// ── PEDIDOS ────────────────────────────────────────────────
function getRequestSheet_(ss) {
  let sh = ss.getSheetByName("Requests");
  if (!sh) {
    sh = ss.insertSheet("Requests");
    sh.appendRow(REQUEST_COLUMNS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readRequests_(ss) {
  const sh = getRequestSheet_(ss);
  const rows = sh.getDataRange().getValues();
  const header = rows.length ? rows[0].map(h => String(h).trim()) : [];
  return rows.slice(1)
    .filter(function (row) { return String(row[0]).trim() !== ""; })
    .map(function (row) {
      const obj = {};
      REQUEST_COLUMNS.forEach(function (name) {
        const i = header.indexOf(name);
        obj[name] = i > -1 && i < row.length ? String(row[i] || "").trim() : "";
      });
      return obj;
    });
}

function requestToRow_(req) {
  return REQUEST_COLUMNS.map(function (name) { return asText_(req[name] || ""); });
}

function appendRequest_(ss, req) {
  getRequestSheet_(ss).appendRow(requestToRow_(req));
}

function writeRequests_(ss, requests) {
  const sh = getRequestSheet_(ss);
  sh.clear();
  sh.appendRow(REQUEST_COLUMNS);
  requests.forEach(function (r) { sh.appendRow(requestToRow_(r)); });
  sh.setFrozenRows(1);
}

function pickRequest_(requests, id) {
  const wanted = String(id || "").trim();
  if (!wanted) return null;
  for (let i = 0; i < requests.length; i++) {
    if (requests[i].id === wanted) return requests[i];
  }
  return null;
}

function findRecordById_(sheet, id) {
  const wanted = String(id || "").trim();
  if (!wanted) return null;
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === wanted) {
      const obj = {};
      header.forEach(function (h, j) { obj[String(h).trim()] = data[i][j]; });
      return obj;
    }
  }
  return null;
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
  // This runs once in the life of the spreadsheet, but it was reading the
  // Settings sheet on every single request that looked at a user — a wasted
  // read on every login. A script property answers in microseconds.
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty("legacyMigrated") === "1") return users;

  const legacy = readSettings_(ss);
  if (!legacy.pwdHash) {
    props.setProperty("legacyMigrated", "1");
    return users;
  }

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
  props.setProperty("legacyMigrated", "1");
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
