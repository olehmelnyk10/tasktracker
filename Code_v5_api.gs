/**
 * TaskTracker v5 — Рамедас Україна
 * ─────────────────────────────────
 * Виправлення v5:
 *  1. Telegram: детальне логування, виправлена прив'язка
 *  2. Статус "На роботі" замість "Вільний" скрізь
 *  3. Права: Виконано/Скасовано — тільки автор задачі
 *  4. Ролі: Співробітник (новий) = Працівник
 *  5. getWorkloadForDate доступний ВСІМ (не тільки менеджерам)
 *  6. getUserDayPlan доступний ВСІМ
 *  7. Telegram sendMessage з повним логуванням відповіді
 */

const TZ       = "GMT+3";
const ROLE_MGR = ["Адмін", "Дирекція", "Керівник"];
// Всі ролі системи
const ALL_ROLES = ["Адмін", "Дирекція", "Керівник", "Співробітник"];
// Статус "На роботі" замість старого "Вільний"
const STATUS_AT_WORK = "На роботі";

// ─── ВХІДНА ТОЧКА ────────────────────────────────
// ─── CORS ────────────────────────────────────────────────────
function _cors(out) {
  return out
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader("Access-Control-Allow-Origin", "*")
    .setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS")
    .setHeader("Access-Control-Allow-Headers", "Content-Type");
}
function _ok(data) {
  return _cors(ContentService.createTextOutput(JSON.stringify(data)));
}
function _err(msg) {
  return _cors(ContentService.createTextOutput(JSON.stringify({ ok: false, err: String(msg) })));
}

// ─── GET: ?fn=loginUser&p={"login":"...","password":"..."} ────
function doGet(e) {
  try {
    const p = e && e.parameter ? e.parameter : {};
    const fn = p.fn || "";
    let params = {};
    try { params = p.p ? JSON.parse(p.p) : {}; } catch(_) {}
    if (!fn) return _ok({ ok: true, app: "TaskTracker", version: "5.0" });
    return _route(fn, params);
  } catch(err) { return _err("doGet: " + err.message); }
}

// ─── РОУТЕР ──────────────────────────────────────────────────
function _route(fn, p) {
  switch (fn) {
    case "loginUser":          return _ok(loginUser(p.login, p.password));
    case "getAppData":         return _ok(getAppData(p.caller));
    case "getRawUsers":        return _ok({ ok: true, rows: getRawUsers() });
    case "createTask":         return _ok(createTask(p.data, p.caller));
    case "updateTaskStatus":   return _ok(updateTaskStatus(p.taskId, p.status, p.caller, p.revComment||""));
    case "editTask":           return _ok(editTask(p.taskId, p.data, p.caller));
    case "deleteTask":         return _ok(deleteTask(p.taskId, p.caller));
    case "addComment":         return _ok(addComment(p.taskId, p.text, p.caller));
    case "getComments":        return _ok({ ok: true, comments: getComments(p.taskId) });
    case "uploadTaskFile":     return _ok(uploadTaskFile(p.taskId, p.fileName, p.base64, p.mimeType, p.caller));
    case "addDriveLinkToTask": return _ok(addDriveLinkToTask(p.taskId, p.url, p.name, p.caller));
    case "getTaskFiles":       return _ok(getTaskFiles(p.taskId));
    case "deleteTaskFile":     return _ok(deleteTaskFile(p.fileId, p.caller));
    case "getAnalytics":       return _ok(getAnalytics(p.caller, p.dept||"", p.period||"month"));
    case "setWorkload":        return _ok(setWorkload(p.login, p.date, p.startTime, p.endTime, p.status, p.note, p.caller));
    case "getWorkloadForDate": return _ok(getWorkloadForDate(p.date, p.caller));
    case "getUserDayPlan":     return _ok(getUserDayPlan(p.login, p.date, p.caller));
    case "saveDayPlan":        return _ok(saveDayPlan(p.login, p.date, p.slots, p.caller));
    case "getInfoRecords":     return _ok(getInfoRecords(p.dept||"", p.caller));
    case "saveInfoRecord":     return _ok(saveInfoRecord(p.rec, p.caller));
    case "deleteInfoRecord":   return _ok(deleteInfoRecord(p.id, p.caller));
    case "uploadInfoFile":     return _ok(uploadInfoFile(p.infoId, p.fileName, p.base64, p.mimeType, p.caller));
    case "getInfoFiles":       return _ok(getInfoFiles(p.infoId));
    case "deleteInfoFile":     return _ok(deleteInfoFile(p.fileId, p.caller));
    case "addUser":            return _ok(addUser(p.data, p.caller));
    case "updateUser":         return _ok(updateUser(p.oldLogin, p.data, p.caller));
    case "deleteUser":         return _ok(deleteUser(p.login, p.caller));
    case "generateTgCode":     return _ok(generateTgCode(p.caller));
    case "getTgBotName":       return _ok(getTgBotName());
    case "getAnnouncements":   return _ok(getAnnouncements());
    case "saveAnnouncement":   return _ok(saveAnnouncement(p.data, p.caller));
    case "deleteAnnouncement": return _ok(deleteAnnouncement(p.id, p.caller));
    default:                   return _err("Unknown function: " + fn);
  }
}

function getAppUrl() {
  try { return ScriptApp.getService().getUrl(); } catch(_) { return ""; }
}

// ─── ДОПОМІЖНІ ───────────────────────────────────
function ss()     { return SpreadsheetApp.getActiveSpreadsheet(); }
function sh(name) { return ss().getSheetByName(name); }

function getOrCreate(name, headers) {
  let s = sh(name);
  if (!s) {
    s = ss().insertSheet(name);
    if (headers && headers.length) s.appendRow(headers);
  }
  return s;
}

function cfg(key) {
  try {
    const rows = sh("Settings").getDataRange().getValues();
    for (const r of rows) if (String(r[0]).trim() === key) return String(r[1]).trim();
    return "";
  } catch(_) { return ""; }
}

function nowStr() {
  return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd HH:mm:ss");
}

function uid() {
  return "T-" + Utilities.getUuid().substring(0, 8).toUpperCase();
}

function parseExecs(raw) {
  if (!raw || raw === "[]" || raw === "null" || raw === "undefined") return [];
  try {
    const a = JSON.parse(raw);
    if (Array.isArray(a)) return a.map(s => String(s).toLowerCase().trim()).filter(Boolean);
  } catch(_) {}
  return String(raw).split(",").map(s => s.trim().toLowerCase()).filter(Boolean);
}

function audit(actor, action, target, detail) {
  try {
    getOrCreate("AuditLog", ["timestamp","actor","action","targetId","detail"])
      .appendRow([nowStr(), actor||"", action||"", target||"", String(detail||"").substring(0,400)]);
  } catch(_) {}
}

function getUser(login) {
  if (!login) return null;
  const cl = String(login).toLowerCase().trim();
  try {
    const rows = sh("Users").getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      if (String(rows[i][0]).toLowerCase().trim() === cl)
        return {
          login: cl,
          name:  String(rows[i][1]||""),
          dept:  String(rows[i][2]||""),
          role:  String(rows[i][3]||""),
          tgId:  String(rows[i][5]||""),
          email: String(rows[i][7]||""),   // окреме поле email (col 8)
          row:   i + 1
        };
    }
  } catch(_) {}
  return null;
}

function checkRole(login, allowed) {
  const u = getUser(login);
  if (!u) throw new Error("Не авторизовано");
  if (!allowed.includes(u.role)) throw new Error("Недостатньо прав: " + u.role);
  return u;
}

// Перевірка чи є у ролі права менеджера
function isMgrRole(role) {
  return ROLE_MGR.includes(role);
}

// ─── АВТОРИЗАЦІЯ ─────────────────────────────────
function loginUser(login, password) {
  try {
    if (!login || !password) return { ok: false, err: "Заповніть усі поля" };
    const cl   = String(login).toLowerCase().trim();
    const cp   = String(password).trim();
    const rows = sh("Users").getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      if (String(rows[i][0]).toLowerCase().trim() !== cl) continue;
      if (String(rows[i][4]||"").trim() !== cp) return { ok: false, err: "Невірний пароль" };
      audit(cl, "LOGIN", "", "");
      return {
        ok: true,
        user: {
          login: cl,
          name:  String(rows[i][1]||""),
          dept:  String(rows[i][2]||""),
          role:  String(rows[i][3]||""),
          tg:    !!(rows[i][5])
        }
      };
    }
    return { ok: false, err: "Користувача не знайдено" };
  } catch(e) { return { ok: false, err: "Помилка: " + e.message }; }
}

// ─── ГОЛОВНИЙ ENDPOINT ───────────────────────────
function getAppData(callerLogin) {
  try {
    const t0      = Date.now();
    const myLogin = String(callerLogin||"").toLowerCase().trim();
    const caller  = getUser(myLogin);
    if (!caller) return { ok: false, err: "Не авторизовано" };

    const taskRows = sh("Tasks").getDataRange().getValues();
    const userRows = sh("Users").getDataRange().getValues();

    const uMap = {};
    const userList = [];
    for (let i = 1; i < userRows.length; i++) {
      if (!userRows[i][0]) continue;
      const ul = String(userRows[i][0]).toLowerCase().trim();
      uMap[ul] = {
        name: String(userRows[i][1]||""),
        dept: String(userRows[i][2]||""),
        role: String(userRows[i][3]||"")
      };
      userList.push({
        login: ul,
        name:  String(userRows[i][1]||""),
        dept:  String(userRows[i][2]||""),
        role:  String(userRows[i][3]||"")
      });
    }

    const tasks = [];
    for (let i = 1; i < taskRows.length; i++) {
      const r = taskRows[i];
      if (!r[0]) continue;
      if (String(r[0]) === "id" || String(r[1]) === "title") continue;

      const execs  = parseExecs(String(r[6]||""));
      const coExec = String(r[7]||"").toLowerCase().trim();
      const author = String(r[8]||"").toLowerCase().trim();
      const dept   = String(r[5]||"");

      const isExec = execs.includes(myLogin) || coExec === myLogin;
      const isAuth = author === myLogin;
      const isDept = dept === caller.dept;

      // Видимість задачі по ролі
      let visible = false;
      if (caller.role === "Адмін")     visible = true;
      else if (caller.role === "Дирекція") visible = true;
      else if (caller.role === "Керівник") visible = isDept || isExec || isAuth;
      else visible = isExec || isAuth; // Співробітник

      if (!visible) continue;

      tasks.push({
        id:          String(r[0]),
        title:       String(r[1]||""),
        desc:        String(r[2]||""),
        type:        String(r[3]||"Операційна"),
        status:      String(r[4]||"Нова"),
        dept:        dept,
        executors:   execs,
        execNames:   execs.map(l => (uMap[l]||{}).name || l).join(", "),
        coExec:      coExec,
        coExecName:  coExec ? ((uMap[coExec]||{}).name || coExec) : "",
        author:      author,
        authorName:  (uMap[author]||{}).name || author,
        deadline:    r[9]  ? String(r[9]).replace("T"," ").substring(0,16) : "",
        startedAt:   r[10] ? String(r[10]).substring(0,16) : "",
        closedAt:    r[11] ? String(r[11]).substring(0,16) : "",
        meetLink:    String(r[12]||""),
        priority:    String(r[13]||"Середній"),
        revComment:  String(r[14]||""),
        isExec:      isExec,
        isAuthor:    isAuth
      });
    }

    Logger.log("getAppData: " + tasks.length + " tasks, " + (Date.now()-t0) + "ms");
    return { ok: true, tasks, users: userList };
  } catch(e) {
    Logger.log("getAppData error: " + e.message);
    return { ok: false, err: e.message };
  }
}

function getRawUsers() {
  try { return sh("Users").getDataRange().getValues(); } catch(_) { return []; }
}

// ─── ЗАДАЧІ ──────────────────────────────────────
function createTask(data, authorLogin) {
  try {
    checkRole(authorLogin, ROLE_MGR);
    if (!data.title || !String(data.title).trim()) return { ok: false, err: "Введіть назву" };

    let execs = [];
    if (Array.isArray(data.executors) && data.executors.length)
      execs = data.executors.map(l => String(l).toLowerCase().trim()).filter(Boolean);
    else if (data.executor)
      execs = [String(data.executor).toLowerCase().trim()];
    if (!execs.length) return { ok: false, err: "Вкажіть виконавця" };

    const coExec   = data.coExec ? String(data.coExec).toLowerCase().trim() : "";
    const taskId   = uid();
    const deadline = data.deadline || "";
    let meetLink   = "";

    if (String(data.type||"").toLowerCase() === "нарада" && deadline) {
      try {
        const cal    = CalendarApp.getDefaultCalendar();
        const start  = new Date(deadline);
        const end    = new Date(start.getTime() + 3600000);
        const guests = [...execs, coExec].filter(Boolean).map(l => l.includes("@") ? l : l+"@gmail.com").join(",");
        const ev     = cal.createEvent("["+taskId+"] "+data.title, start, end, { description: data.desc||"", guests, sendInvites: true });
        meetLink     = "https://meet.google.com/lookup/" + ev.getId().replace("@google.com","");
      } catch(e) {
        Logger.log("Meet creation failed: " + e.message);
        meetLink = "https://meet.google.com/new";
      }
    }

    sh("Tasks").appendRow([
      taskId,
      String(data.title).trim(),
      String(data.desc||"").trim(),
      String(data.type||"Операційна"),
      "Нова",
      String(data.dept||""),
      JSON.stringify(execs),
      coExec,
      authorLogin.toLowerCase().trim(),
      deadline,
      "", // startedAt
      "", // closedAt
      meetLink,
      String(data.priority||"Середній"),
      ""  // revComment
    ]);

    audit(authorLogin, "CREATE", taskId, data.title);

    const caller = getUser(authorLogin);
    const callerName = caller ? caller.name : authorLogin;
    const all = [...new Set([...execs, ...(coExec ? [coExec] : [])])];
    _notifyMany(all,
      "🎯 Нова задача: " + data.title,
      "Постановник: " + callerName +
      "\nПріоритет: " + (data.priority||"Середній") +
      "\nДедлайн: " + (deadline||"—") +
      "\nВідділ: " + (data.dept||"—")
    );

    return { ok: true, taskId };
  } catch(e) { return { ok: false, err: e.message }; }
}

function updateTaskStatus(taskId, newStatus, callerLogin, revComment) {
  try {
    const VALID = ["Нова","В роботі","На перевірці","На доопрацювання","Виконано","Скасовано"];
    if (!VALID.includes(newStatus)) return { ok: false, err: "Невірний статус" };

    const caller = getUser(callerLogin);
    if (!caller) return { ok: false, err: "Не авторизовано" };

    const sheet = sh("Tasks");
    const rows  = sheet.getDataRange().getValues();
    const now   = nowStr();

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) !== String(taskId)) continue;

      const execs  = parseExecs(String(rows[i][6]||""));
      const coExec = String(rows[i][7]||"").toLowerCase().trim();
      const author = String(rows[i][8]||"").toLowerCase().trim();
      const title  = String(rows[i][1]||"");
      const me     = callerLogin.toLowerCase().trim();
      const isMgr  = isMgrRole(caller.role);
      const isExec = execs.includes(me) || coExec === me;
      const isAuth = author === me;

      if (!isMgr && !isExec && !isAuth) return { ok: false, err: "Немає прав" };

      // ВАЖЛИВО: Виконано/Скасовано — тільки автор задачі
      if (["Виконано", "Скасовано"].includes(newStatus) && !isAuth && caller.role !== "Адмін") {
        return { ok: false, err: "Підтверджувати виконання може лише автор задачі" };
      }

      // Оновлення статусу
      sheet.getRange(i+1, 5).setValue(newStatus);
      if (newStatus === "В роботі" && !rows[i][10]) sheet.getRange(i+1, 11).setValue(now);
      if (["Виконано","Скасовано"].includes(newStatus)) sheet.getRange(i+1, 12).setValue(now);
      if (newStatus === "На доопрацювання" && revComment)
        sheet.getRange(i+1, 15).setValue("["+now+"] "+caller.name+": "+revComment);

      audit(callerLogin, "STATUS", taskId, "→"+newStatus);

      // Сповіщення
      const SI = {"В роботі":"▶️","На перевірці":"📋","Виконано":"✅","На доопрацювання":"🔄","Скасовано":"🚫"};
      const emoji = SI[newStatus] || "•";
      const baseMsg = [
        "Задача: " + title,
        "Статус: " + newStatus,
        "Хто змінив: " + caller.name,
        "Дата: " + now.substring(0,16)
      ].join("\n");

      // Виконавцям
      const execList = [...new Set([...execs, ...(coExec ? [coExec] : [])])].filter(l => l !== me);
      if (execList.length) _notifyMany(execList, emoji + " " + title, baseMsg);

      // Автору
      if (author && author !== me) {
        _notifyOne(author, emoji + " Оновлення: " + title, [
          "Задача яку ви поставили: " + title,
          "Статус: " + newStatus,
          "Виконавець: " + caller.name,
          "Дата: " + now.substring(0,16)
        ].join("\n"));
      }

      // На доопрацювання — окремий коментар
      if (newStatus === "На доопрацювання" && revComment) {
        _notifyMany([...execs, ...(coExec ? [coExec] : [])],
          "🔄 Потрібне доопрацювання: " + title,
          ["Задача: " + title, "Коментар:\n" + revComment, "Від: " + caller.name, "Дата: " + now.substring(0,16)].join("\n")
        );
      }

      // Виконано — подяка
      if (newStatus === "Виконано") {
        _notifyMany([...execs, ...(coExec ? [coExec] : [])],
          "✅ Задачу прийнято: " + title,
          ["Задача прийнята: " + title, "Прийняв: " + caller.name, "Дата: " + now.substring(0,16)].join("\n")
        );
      }

      return { ok: true };
    }
    return { ok: false, err: "Задачу не знайдено" };
  } catch(e) { return { ok: false, err: e.message }; }
}

function editTask(taskId, data, callerLogin) {
  try {
    const caller = checkRole(callerLogin, ROLE_MGR);
    const sheet  = sh("Tasks");
    const rows   = sheet.getDataRange().getValues();

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) !== String(taskId)) continue;
      if (caller.role === "Керівник" && rows[i][5] !== caller.dept)
        return { ok: false, err: "Задача не вашого відділу" };

      let execs = rows[i][6];
      if (Array.isArray(data.executors) && data.executors.length)
        execs = JSON.stringify(data.executors.map(l => String(l).toLowerCase().trim()).filter(Boolean));

      sheet.getRange(i+1,  2).setValue(data.title    || rows[i][1]);
      sheet.getRange(i+1,  3).setValue(data.desc     !== undefined ? data.desc : rows[i][2]);
      sheet.getRange(i+1,  4).setValue(data.type     || rows[i][3]);
      sheet.getRange(i+1,  5).setValue(data.status   || rows[i][4]);
      sheet.getRange(i+1,  6).setValue(data.dept     || rows[i][5]);
      sheet.getRange(i+1,  7).setValue(execs);
      sheet.getRange(i+1,  8).setValue(data.coExec   !== undefined ? data.coExec : rows[i][7]);
      sheet.getRange(i+1, 10).setValue(data.deadline || rows[i][9]);
      sheet.getRange(i+1, 14).setValue(data.priority || rows[i][13]);

      audit(callerLogin, "EDIT", taskId, String(data.title||"").substring(0,80));
      return { ok: true };
    }
    return { ok: false, err: "Не знайдено" };
  } catch(e) { return { ok: false, err: e.message }; }
}

function deleteTask(taskId, callerLogin) {
  try {
    const caller = getUser(callerLogin);
    if (!caller) return { ok: false, err: "Не авторизовано" };
    const sheet = sh("Tasks");
    const rows  = sheet.getDataRange().getValues();

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) !== String(taskId)) continue;
      const author  = String(rows[i][8]||"").toLowerCase().trim();
      const isAdmin = ["Адмін","Дирекція"].includes(caller.role);
      const isAuthor = author === callerLogin.toLowerCase().trim() && isMgrRole(caller.role);
      if (!isAdmin && !isAuthor) return { ok: false, err: "Немає прав на видалення" };
      sheet.deleteRow(i+1);
      audit(callerLogin, "DELETE", taskId, "");
      return { ok: true };
    }
    return { ok: false, err: "Не знайдено" };
  } catch(e) { return { ok: false, err: e.message }; }
}

// ─── КОМЕНТАРІ ───────────────────────────────────
function addComment(taskId, text, callerLogin) {
  try {
    if (!taskId || !text || !callerLogin) return { ok: false, err: "Невалідні дані" };
    const u = getUser(callerLogin);
    if (!u) return { ok: false, err: "Не авторизовано" };

    getOrCreate("Comments", ["id","taskId","author","authorName","text","createdAt"])
      .appendRow([
        Utilities.getUuid().substring(0,8).toUpperCase(),
        taskId,
        callerLogin.toLowerCase().trim(),
        u.name,
        String(text).substring(0,1000),
        nowStr()
      ]);
    audit(callerLogin, "COMMENT", taskId, String(text).substring(0,60));

    try {
      const rows = sh("Tasks").getDataRange().getValues();
      for (const r of rows) {
        if (String(r[0]) !== String(taskId)) continue;
        const all = [...new Set([
          ...parseExecs(String(r[6]||"")),
          String(r[7]||""),
          String(r[8]||"")
        ].filter(Boolean))].filter(l => l !== callerLogin.toLowerCase().trim());
        _notifyMany(all, "💬 Коментар до ["+taskId+"]", u.name + ": " + String(text).substring(0,200));
        break;
      }
    } catch(_) {}
    return { ok: true };
  } catch(e) { return { ok: false, err: e.message }; }
}

function getComments(taskId) {
  try {
    const rows = getOrCreate("Comments", ["id","taskId","author","authorName","text","createdAt"])
      .getDataRange().getValues();
    return rows
      .filter((r, i) => i > 0 && String(r[1]) === String(taskId))
      .map(r => ({ id: r[0], author: r[2], name: r[3], text: r[4], time: String(r[5]||"").substring(0,16) }));
  } catch(_) { return []; }
}

// ─── ФАЙЛИ ───────────────────────────────────────
function uploadTaskFile(taskId, fileName, base64Data, mimeType, callerLogin) {
  try {
    if (!taskId || !fileName || !base64Data) return { ok: false, err: "Невалідні дані" };
    const u = getUser(callerLogin);
    if (!u) return { ok: false, err: "Не авторизовано" };

    const taskRow  = _getTaskRow(taskId);
    const dept     = taskRow ? String(taskRow[5]||"Загальне") : "Загальне";
    const fname    = "TaskTracker_Files_" + dept.replace(/[\/\\:*?"<>|]/g,"_").substring(0,40);
    const fl       = DriveApp.getFoldersByName(fname);
    const folder   = fl.hasNext() ? fl.next() : DriveApp.createFolder(fname);
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const bytes = Utilities.base64Decode(base64Data);
    const blob  = Utilities.newBlob(bytes, mimeType||"application/octet-stream", fileName);
    const file  = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    getOrCreate("TaskFiles", ["id","taskId","fileName","fileUrl","uploadedBy","uploadedAt"])
      .appendRow([
        Utilities.getUuid().substring(0,8).toUpperCase(),
        taskId, fileName, file.getUrl(),
        callerLogin.toLowerCase().trim(), nowStr()
      ]);
    audit(callerLogin, "FILE_UPLOAD", taskId, fileName);
    return { ok: true, url: file.getUrl(), name: fileName };
  } catch(e) {
    Logger.log("uploadTaskFile: " + e.message);
    return { ok: false, err: e.message };
  }
}

function addDriveLinkToTask(taskId, linkUrl, linkName, callerLogin) {
  try {
    if (!taskId || !linkUrl) return { ok: false, err: "Вкажіть посилання" };
    const u = getUser(callerLogin);
    if (!u) return { ok: false, err: "Не авторизовано" };
    getOrCreate("TaskFiles", ["id","taskId","fileName","fileUrl","uploadedBy","uploadedAt"])
      .appendRow([
        Utilities.getUuid().substring(0,8).toUpperCase(),
        taskId,
        "🔗 " + (linkName || linkUrl.substring(0,50)),
        linkUrl,
        callerLogin.toLowerCase().trim(),
        nowStr()
      ]);
    audit(callerLogin, "LINK_ADD", taskId, linkUrl);
    return { ok: true };
  } catch(e) { return { ok: false, err: e.message }; }
}

function getTaskFiles(taskId) {
  try {
    const rows = getOrCreate("TaskFiles", ["id","taskId","fileName","fileUrl","uploadedBy","uploadedAt"])
      .getDataRange().getValues();
    const files = [];
    for (let i = 1; i < rows.length; i++) {
      if (!rows[i][0] || String(rows[i][1]) !== String(taskId)) continue;
      files.push({
        id:         String(rows[i][0]),
        name:       String(rows[i][2]||""),
        url:        String(rows[i][3]||""),
        uploadedBy: String(rows[i][4]||""),
        uploadedAt: String(rows[i][5]||"").substring(0,16)
      });
    }
    return { ok: true, files };
  } catch(e) { return { ok: false, files: [] }; }
}

function deleteTaskFile(fileId, callerLogin) {
  try {
    const sheet = getOrCreate("TaskFiles", ["id","taskId","fileName","fileUrl","uploadedBy","uploadedAt"]);
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) !== String(fileId)) continue;
      const uploader = String(rows[i][4]||"").toLowerCase().trim();
      const caller   = getUser(callerLogin);
      if (!caller) return { ok: false, err: "Не авторизовано" };
      if (uploader !== callerLogin.toLowerCase().trim() && !isMgrRole(caller.role))
        return { ok: false, err: "Немає прав" };
      sheet.deleteRow(i+1);
      return { ok: true };
    }
    return { ok: false, err: "Файл не знайдено" };
  } catch(e) { return { ok: false, err: e.message }; }
}

function _getTaskRow(taskId) {
  try {
    const rows = sh("Tasks").getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) if (String(rows[i][0]) === String(taskId)) return rows[i];
    return null;
  } catch(_) { return null; }
}

// ─── АНАЛІТИКА ───────────────────────────────────
function getAnalytics(callerLogin, deptFilter, period) {
  try {
    const caller   = checkRole(callerLogin, ROLE_MGR);
    const rows     = sh("Tasks").getDataRange().getValues();
    const userRows = sh("Users").getDataRange().getValues();
    const now      = new Date();

    const uMap = {};
    for (let i = 1; i < userRows.length; i++) {
      if (!userRows[i][0]) continue;
      uMap[String(userRows[i][0]).toLowerCase().trim()] = String(userRows[i][1]||"");
    }

    let dateFrom = new Date(0);
    if (period === "week")    dateFrom = new Date(now.getTime() - 7*86400000);
    if (period === "month")   dateFrom = new Date(now.getTime() - 30*86400000);
    if (period === "quarter") dateFrom = new Date(now.getTime() - 90*86400000);

    let total=0, done=0, inWork=0, overdue=0, revision=0, onTime=0;
    const byExec = {}, byDept = {}, taskList = [];

    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      if (String(r[0]) === "id" || String(r[1]) === "title") continue;

      const dept   = String(r[5]||"");
      const status = String(r[4]||"");
      const dl     = r[9]  ? new Date(r[9])  : null;
      const closed = r[11] ? new Date(r[11]) : null;
      const execs  = parseExecs(String(r[6]||""));
      const author = String(r[8]||"").toLowerCase().trim();

      if (caller.role === "Керівник" && dept !== caller.dept) continue;
      if (deptFilter && dept !== deptFilter) continue;

      const refDate = closed || dl;
      if (refDate && refDate < dateFrom) continue;

      total++;
      if (status === "Виконано")         done++;
      if (status === "В роботі")         inWork++;
      if (status === "На доопрацювання") revision++;
      if (dl && dl < now && !["Виконано","Скасовано"].includes(status)) overdue++;
      if (status === "Виконано" && dl && closed && closed <= dl) onTime++;

      execs.forEach(l => {
        const name = uMap[l] || l;
        byExec[name] = (byExec[name]||0) + 1;
      });

      if (!byDept[dept]) byDept[dept] = { total:0, done:0, inWork:0, overdue:0 };
      byDept[dept].total++;
      if (status === "Виконано") byDept[dept].done++;
      if (status === "В роботі") byDept[dept].inWork++;
      if (dl && dl < now && !["Виконано","Скасовано"].includes(status)) byDept[dept].overdue++;

      taskList.push({
        id:         String(r[0]),
        title:      String(r[1]||""),
        status,
        dept,
        priority:   String(r[13]||"Середній"),
        executor:   execs.map(l => uMap[l]||l).join(", "),
        authorName: uMap[author] || author,
        deadline:   dl     ? Utilities.formatDate(dl,     TZ, "dd.MM.yyyy HH:mm") : "—",
        closedAt:   closed ? Utilities.formatDate(closed, TZ, "dd.MM.yyyy HH:mm") : "—"
      });
    }

    const kpi = done > 0 ? Math.round((onTime/done)*100) : 0;
    return { ok: true, stats: { total, done, inWork, overdue, revision, kpi }, byExec, byDept, taskList };
  } catch(e) { return { ok: false, err: e.message }; }
}

// ─── ГРАФІК (WorkloadLog) ────────────────────────
// "На роботі" замість "Вільний" — нормалізуємо при збереженні
function _normalizeWlStatus(status) {
  if (!status) return STATUS_AT_WORK;
  if (status === "Вільний") return STATUS_AT_WORK;
  return status;
}

function setWorkload(login, date, startTime, endTime, status, note, callerLogin) {
  try {
    const cl = callerLogin.toLowerCase().trim();
    const lg = login.toLowerCase().trim();
    if (lg !== cl) checkRole(callerLogin, ROLE_MGR);

    const normStatus = _normalizeWlStatus(status);
    const sheet = getOrCreate("WorkloadLog", ["login","date","startTime","endTime","status","note","updatedAt"]);
    const rows  = sheet.getDataRange().getValues();
    const now   = nowStr();

    for (let i = 1; i < rows.length; i++) {
      const rl = String(rows[i][0]||"").toLowerCase().trim();
      const rd = rows[i][1] instanceof Date
        ? Utilities.formatDate(rows[i][1], TZ, "yyyy-MM-dd")
        : String(rows[i][1]||"").substring(0,10);
      if (rl === lg && rd === date) {
        sheet.getRange(i+1,3,1,5).setValues([[startTime, endTime, normStatus, note||"", now]]);
        return { ok: true };
      }
    }
    sheet.appendRow([lg, date, startTime, endTime, normStatus, note||"", now]);
    return { ok: true };
  } catch(e) { return { ok: false, err: e.message }; }
}

// getWorkloadForDate — доступний ВСІМ (не тільки менеджерам)
function getWorkloadForDate(targetDate, callerLogin) {
  try {
    const caller   = getUser(callerLogin);
    if (!caller) return { ok: false, err: "Не авторизовано" };

    const users    = sh("Users").getDataRange().getValues();
    const wlRows   = sh("WorkloadLog") ? sh("WorkloadLog").getDataRange().getValues() : [];
    const taskRows = sh("Tasks").getDataRange().getValues();

    const activeByUser = {};
    for (let i = 1; i < taskRows.length; i++) {
      if (!taskRows[i][0] || ["Виконано","Скасовано"].includes(taskRows[i][4])) continue;
      parseExecs(String(taskRows[i][6]||"")).forEach(l => {
        activeByUser[l] = (activeByUser[l]||0) + 1;
      });
    }

    const wlMap = {};
    for (let i = 1; i < wlRows.length; i++) {
      if (!wlRows[i][0]) continue;
      const rd = wlRows[i][1] instanceof Date
        ? Utilities.formatDate(wlRows[i][1], TZ, "yyyy-MM-dd")
        : String(wlRows[i][1]||"").substring(0,10);
      if (rd !== targetDate) continue;
      const l = String(wlRows[i][0]).toLowerCase().trim();
      wlMap[l] = {
        startTime: String(wlRows[i][2]||"09:00"),
        endTime:   String(wlRows[i][3]||"18:00"),
        status:    _normalizeWlStatus(String(wlRows[i][4]||"")),
        note:      String(wlRows[i][5]||"")
      };
    }

    const result = [];
    for (let i = 1; i < users.length; i++) {
      if (!users[i][0]) continue;
      const l = String(users[i][0]).toLowerCase().trim();
      // Керівник бачить тільки свій відділ; всі інші бачать всіх
      if (caller.role === "Керівник" && String(users[i][2]) !== caller.dept) continue;
      const wl = wlMap[l] || { startTime:"09:00", endTime:"18:00", status: STATUS_AT_WORK, note:"" };
      result.push({
        login:       l,
        name:        String(users[i][1]||""),
        dept:        String(users[i][2]||""),
        role:        String(users[i][3]||""),
        activeTasks: activeByUser[l]||0,
        startTime:   wl.startTime,
        endTime:     wl.endTime,
        status:      wl.status,
        note:        wl.note
      });
    }

    result.sort((a, b) => {
      const ord = {"Зайнятий":0,"Відрядження":1,"Лікарняний":2,"Відпустка":3,[STATUS_AT_WORK]:4};
      const sa  = ord[a.status] ?? 5;
      const sb  = ord[b.status] ?? 5;
      if (sa !== sb) return sa - sb;
      return b.activeTasks - a.activeTasks;
    });

    return { ok: true, data: result };
  } catch(e) { return { ok: false, err: e.message }; }
}

// getUserDayPlan — доступний ВСІМ
function getUserDayPlan(targetLogin, targetDate, callerLogin) {
  try {
    const caller = getUser(callerLogin);
    if (!caller) return { ok: false, err: "Не авторизовано" };

    const login = String(targetLogin).toLowerCase().trim();
    const store = PropertiesService.getScriptProperties();
    const raw   = store.getProperty("ttp_" + login + "_" + targetDate);
    const slots = raw ? JSON.parse(raw) : [];

    const wlRows = sh("WorkloadLog").getDataRange().getValues();
    let wl = { startTime:"09:00", endTime:"18:00", status: STATUS_AT_WORK, note:"" };
    for (let i = 1; i < wlRows.length; i++) {
      if (!wlRows[i][0]) continue;
      const rl = String(wlRows[i][0]).toLowerCase().trim();
      const rd = wlRows[i][1] instanceof Date
        ? Utilities.formatDate(wlRows[i][1], TZ, "yyyy-MM-dd")
        : String(wlRows[i][1]||"").substring(0,10);
      if (rl === login && rd === targetDate) {
        wl = {
          startTime: String(wlRows[i][2]||"09:00"),
          endTime:   String(wlRows[i][3]||"18:00"),
          status:    _normalizeWlStatus(String(wlRows[i][4]||"")),
          note:      String(wlRows[i][5]||"")
        };
        break;
      }
    }

    const u = getUser(login);
    return { ok: true, login, name: u ? u.name : login, slots, workload: wl };
  } catch(e) { return { ok: false, err: e.message }; }
}

function saveDayPlan(login, date, slots, callerLogin) {
  try {
    const caller = getUser(callerLogin);
    if (!caller) return { ok: false, err: "Не авторизовано" };
    const cl = callerLogin.toLowerCase().trim();
    const lg = login.toLowerCase().trim();
    if (lg !== cl && !isMgrRole(caller.role)) return { ok: false, err: "Немає прав" };
    PropertiesService.getScriptProperties()
      .setProperty("ttp_" + lg + "_" + date, JSON.stringify(slots||[]));
    return { ok: true };
  } catch(e) { return { ok: false, err: e.message }; }
}

// ─── БАЗА ЗНАНЬ ──────────────────────────────────
function getInfoRecords(dept, callerLogin) {
  try {
    const u = getUser(callerLogin);
    if (!u) return { ok: false, err: "Не авторизовано" };
    const rows = getOrCreate("InfoRecords",
      ["id","dept","heading","category","content","author","createdAt","updatedAt"])
      .getDataRange().getValues();
    const result = [];
    for (let i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      if (dept && rows[i][1] !== dept) continue;
      result.push({
        id:        rows[i][0],
        dept:      rows[i][1],
        heading:   rows[i][2],
        category:  rows[i][3],
        content:   rows[i][4],
        author:    rows[i][5],
        createdAt: String(rows[i][6]||"")
      });
    }
    return { ok: true, records: result };
  } catch(e) { return { ok: false, records: [] }; }
}

function saveInfoRecord(rec, callerLogin) {
  try {
    checkRole(callerLogin, ROLE_MGR);
    const u   = getUser(callerLogin);
    const s   = getOrCreate("InfoRecords",
      ["id","dept","heading","category","content","author","createdAt","updatedAt"]);
    const now = Utilities.formatDate(new Date(), TZ, "dd.MM.yyyy");

    if (rec.id) {
      const rows = s.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] !== rec.id) continue;
        s.getRange(i+1,2,1,7).setValues([[
          rec.dept, rec.heading,
          rec.category||"📋 Загальне",
          rec.content,
          u ? u.name : callerLogin,
          rows[i][6], now
        ]]);
        return { ok: true };
      }
    }
    const id = Utilities.getUuid().substring(0,10).toUpperCase();
    s.appendRow([id, rec.dept, rec.heading, rec.category||"📋 Загальне", rec.content,
      u ? u.name : callerLogin, now, ""]);
    return { ok: true, id };
  } catch(e) { return { ok: false, err: e.message }; }
}

function deleteInfoRecord(id, callerLogin) {
  try {
    checkRole(callerLogin, ROLE_MGR);
    const s    = getOrCreate("InfoRecords",
      ["id","dept","heading","category","content","author","createdAt","updatedAt"]);
    const rows = s.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) { s.deleteRow(i+1); return { ok: true }; }
    }
    return { ok: false };
  } catch(e) { return { ok: false, err: e.message }; }
}

function uploadInfoFile(infoId, fileName, base64Data, mimeType, callerLogin) {
  try {
    checkRole(callerLogin, ROLE_MGR);
    const bytes  = Utilities.base64Decode(base64Data);
    const blob   = Utilities.newBlob(bytes, mimeType||"application/octet-stream", fileName);
    const fl     = DriveApp.getFoldersByName("TaskTracker_InfoFiles");
    const folder = fl.hasNext() ? fl.next() : DriveApp.createFolder("TaskTracker_InfoFiles");
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fid    = Utilities.getUuid().substring(0,8).toUpperCase();
    getOrCreate("InfoFiles", ["id","infoId","fileName","fileUrl","uploadedBy","uploadedAt"])
      .appendRow([fid, infoId, fileName, file.getUrl(), callerLogin.toLowerCase().trim(), nowStr()]);
    return { ok: true, id: fid, url: file.getUrl(), name: fileName };
  } catch(e) { return { ok: false, err: e.message }; }
}

function getInfoFiles(infoId) {
  try {
    const rows = getOrCreate("InfoFiles",
      ["id","infoId","fileName","fileUrl","uploadedBy","uploadedAt"])
      .getDataRange().getValues();
    const files = [];
    for (let i = 1; i < rows.length; i++) {
      if (!rows[i][0] || String(rows[i][1]) !== String(infoId)) continue;
      files.push({ id: rows[i][0], name: rows[i][2], url: rows[i][3], uploadedBy: rows[i][4] });
    }
    return { ok: true, files };
  } catch(_) { return { ok: false, files: [] }; }
}

function deleteInfoFile(fileId, callerLogin) {
  try {
    checkRole(callerLogin, ROLE_MGR);
    const s    = getOrCreate("InfoFiles",
      ["id","infoId","fileName","fileUrl","uploadedBy","uploadedAt"]);
    const rows = s.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(fileId)) { s.deleteRow(i+1); return { ok: true }; }
    }
    return { ok: false };
  } catch(e) { return { ok: false, err: e.message }; }
}

// ─── ПЕРСОНАЛ ────────────────────────────────────
function addUser(data, callerLogin) {
  try {
    checkRole(callerLogin, ["Адмін"]);
    const nl = String(data.login||"").toLowerCase().trim();
    if (!nl) return { ok: false, err: "Вкажіть логін" };
    if (getUser(nl)) return { ok: false, err: "Логін вже існує: " + nl };
    const email = String(data.email||"").trim();
    sh("Users").appendRow([nl, data.name, data.dept, data.role, data.password, "", "", email]);
    audit(callerLogin, "USER_ADD", nl, data.role);
    return { ok: true };
  } catch(e) { return { ok: false, err: e.message }; }
}

function updateUser(oldLogin, data, callerLogin) {
  try {
    checkRole(callerLogin, ["Адмін"]);
    const sheet = sh("Users");
    const rows  = sheet.getDataRange().getValues();
    const ol    = oldLogin.toLowerCase().trim();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).toLowerCase().trim() !== ol) continue;
      const nl = String(data.login||ol).toLowerCase().trim();
      sheet.getRange(i+1,1).setValue(nl);
      sheet.getRange(i+1,2).setValue(data.name  || rows[i][1]);
      sheet.getRange(i+1,3).setValue(data.dept  || rows[i][2]);
      sheet.getRange(i+1,4).setValue(data.role  || rows[i][3]);
      if (data.password) sheet.getRange(i+1,5).setValue(data.password);
      // Оновлюємо email якщо передано
      if (data.email !== undefined) sheet.getRange(i+1,8).setValue(String(data.email||"").trim());
      audit(callerLogin, "USER_UPDATE", ol, data.role);
      return { ok: true };
    }
    return { ok: false, err: "Не знайдено" };
  } catch(e) { return { ok: false, err: e.message }; }
}

function deleteUser(login, callerLogin) {
  try {
    checkRole(callerLogin, ["Адмін"]);
    const sheet = sh("Users");
    const rows  = sheet.getDataRange().getValues();
    const cl    = login.toLowerCase().trim();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).toLowerCase().trim() !== cl) continue;
      sheet.deleteRow(i+1);
      audit(callerLogin, "USER_DELETE", cl, "");
      return { ok: true };
    }
    return { ok: false };
  } catch(e) { return { ok: false }; }
}

// ─── TELEGRAM ────────────────────────────────────
// ВИПРАВЛЕНО: детальне логування, перевірка токену та відповіді

function generateTgCode(callerLogin) {
  try {
    const token = cfg("TG_TOKEN");
    if (!token) {
      audit(callerLogin, "TG_CODE_FAIL", "", "TG_TOKEN не налаштовано в Settings");
      return { ok: false, err: "Telegram не налаштовано. Додайте TG_TOKEN у Settings." };
    }

    const code = Math.random().toString(36).substring(2,8).toUpperCase();
    const s    = getOrCreate("TgCodes", ["code","login","createdAt"]);
    const rows = s.getDataRange().getValues();

    // Видаляємо старі коди цього користувача
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][1]).toLowerCase().trim() === callerLogin.toLowerCase().trim())
        s.deleteRow(i+1);
    }
    s.appendRow([code, callerLogin.toLowerCase().trim(), nowStr()]);
    audit(callerLogin, "TG_CODE_GEN", "", code);

    // Відправляємо код собі в Telegram якщо вже прив'язаний
    const u = getUser(callerLogin);
    if (u && u.tgId) {
      const msg = "🔑 Ваш новий код прив'язки TaskTracker:\n\n<code>CODE " + code + "</code>\n\n⏱ Дійсний 10 хвилин";
      const result = _tgSendWithLog(token, u.tgId, msg);
      Logger.log("TG code sent to " + callerLogin + ": " + JSON.stringify(result));
    }

    return { ok: true, code };
  } catch(e) {
    Logger.log("generateTgCode error: " + e.message);
    return { ok: false, err: e.message };
  }
}

function getTgBotName() {
  try {
    const token = cfg("TG_TOKEN");
    if (!token) return { ok: false, err: "TG_TOKEN не налаштовано" };
    const response = UrlFetchApp.fetch(
      "https://api.telegram.org/bot" + token + "/getMe",
      { muteHttpExceptions: true }
    );
    const r = JSON.parse(response.getContentText());
    Logger.log("getMe response: " + JSON.stringify(r));
    if (r.ok) return { ok: true, name: r.result.username };
    audit("system", "TG_GETME_FAIL", "", JSON.stringify(r));
    return { ok: false, err: r.description || "Невірний токен" };
  } catch(e) {
    Logger.log("getTgBotName error: " + e.message);
    return { ok: false, err: e.message };
  }
}

function doPost(e) {
  try {
    let body = {};
    try { body = JSON.parse(e.postData.contents); } catch(_) {}

    // Telegram update
    if (body.message || body.callback_query || body.update_id) {
      return _handleTelegram(body);
    }

    // API POST call: { fn: "...", p: {...} }
    const fn = body.fn || "";
    const p  = body.p  || {};
    if (!fn) return _err("fn required");
    return _route(fn, p);
  } catch(err) {
    Logger.log("doPost error: " + err.message);
    return ContentService.createTextOutput("ok");
  }
}

function _handleTelegram(update) {
  try {
    if (update.callback_query) _tgAnswerCb(update.callback_query.id);
    const msg = update.message || (update.callback_query && update.callback_query.message);
    if (!msg) return ContentService.createTextOutput("ok");
    const chatId = String(msg.chat.id);
    const text   = update.callback_query ? update.callback_query.data : (msg.text||"").trim();
    const token  = cfg("TG_TOKEN");
    const u      = _tgGetUser(chatId);
    if (!u) {
      if (text === "/start")
        _tgSend(token, chatId, "👋 <b>TaskTracker — Рамедас Україна</b>\n\nДля прив'язки:\n1. Відкрийте TaskTracker\n2. Налаштування → «Отримати код»\n3. Надішліть сюди: <code>CODE XXXXXX</code>");
      else if (text.toUpperCase().startsWith("CODE "))
        _tgBind(token, chatId, text.substring(5).trim().toUpperCase());
      else
        _tgSend(token, chatId, "Надішліть /start або CODE XXXXXX для прив'язки.");
      return ContentService.createTextOutput("ok");
    }
    if      (text === "/start" || text === "/menu" || text === "main") _tgMenu(token, chatId, u);
    else if (text === "/tasks" || text === "my_tasks") _tgMyTasks(token, chatId, u);
    else if (text.startsWith("task_"))   _tgTaskDetail(token, chatId, u, text.replace("task_",""));
    else if (text.startsWith("status_")) {
      const parts = text.replace("status_","").split("|");
      const res   = updateTaskStatus(parts[0], parts[1], u.login, "");
      _tgSend(token, chatId, res.ok ? "✅ "+parts[1] : "❌ "+(res.err||"Помилка"));
    } else _tgMenu(token, chatId, u);
    return ContentService.createTextOutput("ok");
  } catch(err) {
    Logger.log("_handleTelegram: " + err.message);
    return ContentService.createTextOutput("ok");
  }
}


function registerWebhook() {
  const token = cfg("TG_TOKEN");
  const url   = getAppUrl();
  if (!token) { Logger.log("❌ TG_TOKEN не знайдено в Settings"); return; }
  if (!url)   { Logger.log("❌ Спочатку зробіть Deploy"); return; }
  const r = UrlFetchApp.fetch(
    "https://api.telegram.org/bot" + token + "/setWebhook?url=" + encodeURIComponent(url)
  );
  Logger.log("Webhook: " + r.getContentText());
}

// ─── TG ДОПОМІЖНІ ────────────────────────────────
function _tgGetUser(chatId) {
  try {
    const rows = sh("Users").getDataRange().getValues();
    for (let i = 1; i < rows.length; i++)
      if (String(rows[i][5]) === String(chatId))
        return {
          login: String(rows[i][0]).toLowerCase().trim(),
          name:  String(rows[i][1]),
          role:  String(rows[i][3])
        };
    return null;
  } catch(_) { return null; }
}

function _tgBind(token, chatId, code) {
  try {
    const s    = getOrCreate("TgCodes", ["code","login","createdAt"]);
    const rows = s.getDataRange().getValues();

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim().toUpperCase() !== code) continue;

      const login = String(rows[i][1]).toLowerCase().trim();

      // Перевіряємо термін дії (10 хвилин)
      const created    = new Date(rows[i][2]);
      const ageMinutes = (new Date() - created) / 60000;
      if (ageMinutes > 10) {
        s.deleteRow(i+1);
        _tgSend(token, chatId, "⏳ Код прострочено. Згенеруйте новий у Налаштуваннях.");
        return;
      }

      // Зберігаємо chatId
      const uSh  = sh("Users");
      const uR   = uSh.getDataRange().getValues();
      let bound  = false;
      for (let j = 1; j < uR.length; j++) {
        if (String(uR[j][0]).toLowerCase().trim() === login) {
          uSh.getRange(j+1, 6).setValue(chatId);
          bound = true;
          break;
        }
      }
      if (!bound) {
        _tgSend(token, chatId, "❌ Користувача не знайдено.");
        return;
      }

      s.deleteRow(i+1);
      const u = getUser(login);
      audit(login, "TG_BIND", "", "chatId:" + chatId);

      _tgSend(token, chatId,
        "✅ <b>Telegram прив'язано!</b>\n\n" +
        "👤 " + (u ? u.name : login) + "\n" +
        "🏢 " + (u ? u.dept : "") + "\n" +
        "🎭 " + (u ? u.role : "") + "\n\n" +
        "Тепер ви будете отримувати сповіщення про задачі.\n\nНадішліть /menu"
      );
      return;
    }
    _tgSend(token, chatId, "❌ Невірний код. Перевірте і спробуйте ще раз.");
  } catch(err) {
    Logger.log("_tgBind error: " + err.message);
    _tgSend(token, chatId, "❌ Помилка прив'язки: " + err.message);
  }
}

function _tgMenu(token, chatId, u) {
  _tgSend(token, chatId,
    "🏠 <b>TaskTracker</b>\n\n👤 " + u.name + " · " + u.role,
    { inline_keyboard: [
      [{ text: "📋 Мої задачі", callback_data: "my_tasks" }],
      [{ text: "ℹ️ Допомога",   callback_data: "/start"   }]
    ]}
  );
}

function _tgMyTasks(token, chatId, u) {
  const rows  = sh("Tasks").getDataRange().getValues();
  const tasks = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0] || ["Виконано","Скасовано"].includes(rows[i][4])) continue;
    const execs = parseExecs(String(rows[i][6]||""));
    if (!execs.includes(u.login) && String(rows[i][7]||"").toLowerCase() !== u.login) continue;
    tasks.push({
      id:     rows[i][0],
      title:  String(rows[i][1]),
      status: String(rows[i][4]),
      dl:     rows[i][9] ? String(rows[i][9]).substring(0,10) : "—"
    });
  }
  if (!tasks.length) { _tgSend(token, chatId, "✅ Активних задач немає!"); return; }

  const SI = {"Нова":"🆕","В роботі":"▶","На перевірці":"📁","На доопрацювання":"🔴"};
  let text  = "📋 <b>Ваші задачі (" + tasks.length + "):</b>\n\n";
  tasks.slice(0,10).forEach(t => {
    text += (SI[t.status]||"•") + " " + t.title + "\n  📅 " + t.dl + " · " + t.id + "\n\n";
  });
  const btns = tasks.slice(0,5).map(t => [{
    text: t.title.substring(0,35),
    callback_data: "task_" + t.id
  }]);
  btns.push([{ text: "🏠 Меню", callback_data: "main" }]);
  _tgSend(token, chatId, text, { inline_keyboard: btns });
}

function _tgTaskDetail(token, chatId, u, taskId) {
  const rows = sh("Tasks").getDataRange().getValues();
  for (const r of rows) {
    if (String(r[0]) !== taskId) continue;
    const status = String(r[4]);
    const execs  = parseExecs(String(r[6]||""));
    const isExec = execs.includes(u.login) || String(r[7]||"").toLowerCase() === u.login;
    const SI     = {"Нова":"🆕","В роботі":"▶","На перевірці":"📁","На доопрацювання":"🔴","Виконано":"✅","Скасовано":"🚫"};
    let text     = (SI[status]||"•") + " <b>" + r[1] + "</b>\n" +
                   "<code>" + r[0] + "</code>\n" +
                   "📅 " + (r[9] ? String(r[9]).substring(0,16) : "—") + " · " + r[13];
    if (r[14]) text += "\n💬 " + r[14];
    const btns = [];
    if (isExec) {
      if (status === "Нова")
        btns.push([{ text:"▶ Взяти в роботу", callback_data:"status_"+taskId+"|В роботі" }]);
      if (status === "В роботі")
        btns.push([{ text:"📁 На перевірку", callback_data:"status_"+taskId+"|На перевірці" }]);
      if (status === "На доопрацювання")
        btns.push([{ text:"🔄 Після виправлення", callback_data:"status_"+taskId+"|На перевірці" }]);
    }
    btns.push([{ text:"◀ Назад", callback_data:"my_tasks" }]);
    _tgSend(token, chatId, text, { inline_keyboard: btns });
    return;
  }
  _tgSend(token, chatId, "❌ Задачу не знайдено");
}

// ВИПРАВЛЕНО: _tgSend з детальним логуванням
function _tgSend(token, chatId, text, keyboard) {
  if (!token || !chatId) {
    Logger.log("_tgSend: missing token or chatId");
    return null;
  }
  return _tgSendWithLog(token, chatId, text, keyboard);
}

function _tgSendWithLog(token, chatId, text, keyboard) {
  try {
    const payload = {
      chat_id:              String(chatId),
      text:                 String(text).substring(0, 4096),
      parse_mode:           "HTML",
      disable_web_page_preview: true
    };
    if (keyboard) payload.reply_markup = JSON.stringify(keyboard);

    const response = UrlFetchApp.fetch(
      "https://api.telegram.org/bot" + token + "/sendMessage",
      {
        method:      "post",
        contentType: "application/json",
        payload:     JSON.stringify(payload),
        muteHttpExceptions: true
      }
    );

    const responseText = response.getContentText();
    const result       = JSON.parse(responseText);

    if (!result.ok) {
      Logger.log("TG sendMessage FAILED: " + responseText);
      audit("system", "TG_SEND_FAIL", String(chatId), result.description || responseText);
    }

    return result;
  } catch(e) {
    Logger.log("_tgSendWithLog error: " + e.message);
    audit("system", "TG_SEND_ERROR", String(chatId), e.message);
    return null;
  }
}

function _tgAnswerCb(id) {
  try {
    UrlFetchApp.fetch(
      "https://api.telegram.org/bot" + cfg("TG_TOKEN") + "/answerCallbackQuery",
      {
        method:      "post",
        contentType: "application/json",
        payload:     JSON.stringify({ callback_query_id: id }),
        muteHttpExceptions: true
      }
    );
  } catch(_) {}
}

// ─── СПОВІЩЕННЯ ──────────────────────────────────
function _notifyOne(login, subject, body) {
  if (!login) return;
  const u = getUser(String(login).toLowerCase().trim());
  if (!u) return;

  // Email — беремо з окремого поля email (col 8) якщо є, інакше з login
  try {
    const emailAddr = _getUserEmail(u);
    if (emailAddr) {
      MailApp.sendEmail({ to: emailAddr, subject, htmlBody: _emailHtml(subject, body, u.name) });
      Logger.log("Email sent to: " + emailAddr + " (" + u.name + ")");
    } else {
      Logger.log("No email for user: " + u.login);
    }
  } catch(e) {
    Logger.log("Email failed to " + login + ": " + e.message);
  }

  // Telegram
  if (u.tgId) {
    const token = cfg("TG_TOKEN");
    if (token) {
      _tgSendWithLog(token, u.tgId, "<b>" + subject + "</b>\n\n" + body);
    }
  }
}

// Отримати email користувача: спочатку поле email (col 8), потім login якщо це @
function _getUserEmail(u) {
  // Читаємо поле email з таблиці (col 8, індекс 7)
  try {
    const rows = sh("Users").getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      if (String(rows[i][0]).toLowerCase().trim() !== u.login) continue;
      // Col 8 (index 7) — окреме поле email
      const dedicatedEmail = String(rows[i][7]||"").trim();
      if (dedicatedEmail && dedicatedEmail.includes("@")) return dedicatedEmail;
      // Fallback: login якщо містить @
      if (u.login.includes("@")) return u.login;
      return null;
    }
  } catch(_) {}
  if (u.login.includes("@")) return u.login;
  return null;
}

function _notifyMany(logins, subject, body) {
  if (!Array.isArray(logins)) logins = [logins];
  [...new Set(logins.filter(Boolean))].forEach(l => _notifyOne(l, subject, body));
}

function _emailHtml(subject, body, recipientName) {
  const url = getAppUrl();
  return '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f1f5f9;font-family:system-ui,sans-serif;">' +
    '<table width="100%" cellpadding="0" cellspacing="0"><tr><td align="center" style="padding:32px 16px;">' +
    '<table width="500" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:16px;overflow:hidden;border:1px solid #e2e8f0;box-shadow:0 4px 24px rgba(0,0,0,.08);">' +
    '<tr><td style="background:linear-gradient(135deg,#1a237e,#283593);padding:24px 28px;">' +
    '<div style="font-size:18px;font-weight:900;color:#fff;">🎯 TaskTracker</div>' +
    '<div style="font-size:11px;color:rgba(255,255,255,.75);margin-top:2px;letter-spacing:.15em;text-transform:uppercase;">Рамедас Україна</div>' +
    '</td></tr><tr><td style="padding:24px 28px;">' +
    (recipientName ? '<p style="margin:0 0 8px;font-size:12px;color:#94a3b8;">Привіт, <b>' + recipientName + '</b>!</p>' : '') +
    '<h2 style="margin:0 0 12px;font-size:16px;color:#0f172a;font-weight:800;">' + subject + '</h2>' +
    '<div style="font-size:13px;color:#475569;line-height:1.6;white-space:pre-line;background:#f8fafc;border-radius:10px;padding:14px;border-left:3px solid #4f8ef7;">' + body + '</div>' +
    (url ? '<div style="margin-top:20px;"><a href="' + url + '" style="display:inline-block;background:linear-gradient(135deg,#4f8ef7,#7c6ff7);color:#fff;font-weight:700;font-size:13px;padding:11px 24px;border-radius:10px;text-decoration:none;">Відкрити TaskTracker →</a></div>' : '') +
    '</td></tr><tr><td style="background:#f8fafc;padding:12px 28px;font-size:11px;color:#94a3b8;text-align:center;border-top:1px solid #e2e8f0;">Це автоматичне повідомлення · TaskTracker · Рамедас Україна</td></tr>' +
    '</table></td></tr></table></body></html>';
}

// ─── ТРИГЕРИ ─────────────────────────────────────
function morningDigest() {
  try {
    const tasks  = sh("Tasks").getDataRange().getValues();
    const users  = sh("Users").getDataRange().getValues();
    const token  = cfg("TG_TOKEN");
    const byUser = {};

    for (let i = 1; i < tasks.length; i++) {
      if (!tasks[i][0] || ["Виконано","Скасовано"].includes(tasks[i][4])) continue;
      parseExecs(String(tasks[i][6]||"")).forEach(l => {
        if (!byUser[l]) byUser[l] = [];
        byUser[l].push({
          id:     tasks[i][0],
          title:  String(tasks[i][1]),
          status: String(tasks[i][4]),
          dl:     tasks[i][9] ? String(tasks[i][9]).substring(0,10) : "—"
        });
      });
    }

    for (let i = 1; i < users.length; i++) {
      if (!users[i][0]) continue;
      const l     = String(users[i][0]).toLowerCase().trim();
      const tgId  = String(users[i][5]||"");
      const items = byUser[l];
      if (!items || !items.length) continue;

      const SI  = {"Нова":"🆕","В роботі":"▶","На перевірці":"📁","На доопрацювання":"🔴"};
      let msg   = "☀️ Доброго ранку, " + users[i][1] + "!\n\nВаші задачі:\n\n";
      items.forEach(t => { msg += (SI[t.status]||"•") + " " + t.title + "\n  📅 " + t.dl + "\n\n"; });

      if (tgId && token) _tgSendWithLog(token, tgId, msg);

      try {
        const email = l.includes("@") ? l : l + "@gmail.com";
        MailApp.sendEmail({
          to: email,
          subject: "☀️ Ваші задачі на сьогодні — TaskTracker",
          htmlBody: _emailHtml(
            "Ваші задачі на сьогодні",
            items.map(t => t.title + " | " + t.status + " | до " + t.dl).join("\n"),
            users[i][1]
          )
        });
      } catch(_) {}
    }
  } catch(e) { Logger.log("morningDigest: " + e.message); }
}

function deadlineReminders() {
  try {
    const tasks = sh("Tasks").getDataRange().getValues();
    const now   = new Date();
    const in24h = new Date(now.getTime() + 86400000);
    const in25h = new Date(now.getTime() + 90000000);

    for (let i = 1; i < tasks.length; i++) {
      if (!tasks[i][0] || ["Виконано","Скасовано","На перевірці"].includes(tasks[i][4])) continue;
      const dl = tasks[i][9] ? new Date(tasks[i][9]) : null;
      if (!dl || dl < in24h || dl > in25h) continue;
      const execs  = parseExecs(String(tasks[i][6]||""));
      const author = String(tasks[i][8]||"");
      const dlStr  = Utilities.formatDate(dl, TZ, "dd.MM.yyyy HH:mm");
      const all    = [...new Set([...execs, author].filter(Boolean))];
      _notifyMany(all,
        "⚠️ Дедлайн завтра: " + tasks[i][1],
        "[" + tasks[i][0] + "] до " + dlStr + "\nСтатус: " + tasks[i][4]
      );
    }
  } catch(e) { Logger.log("deadlineReminders: " + e.message); }
}

function weeklyBackup() {
  try {
    const file   = DriveApp.getFileById(ss().getId());
    const date   = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd");
    const fl     = DriveApp.getFoldersByName("TaskTracker_Backups");
    const folder = fl.hasNext() ? fl.next() : DriveApp.createFolder("TaskTracker_Backups");
    folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    const copy   = file.makeCopy("Backup_TaskTracker_" + date, folder);
    copy.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    const cutoff = new Date(new Date().getTime() - 180*86400000);
    const files  = folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (f.getId() !== copy.getId() && f.getDateCreated() < cutoff) f.setTrashed(true);
    }
    Logger.log("✅ Backup: " + copy.getUrl());
  } catch(e) { Logger.log("weeklyBackup: " + e.message); }
}

// ─── ОГОЛОШЕННЯ ──────────────────────────────────────────────────
// Доступ: читати — всі; створювати/редагувати/видаляти — тільки Адмін

function getAnnouncements() {
  try {
    const rows = getOrCreate("Announcements",
      ["id","title","body","category","author","createdAt","pinned"])
      .getDataRange().getValues();
    const result = [];
    for (let i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      result.push({
        id:        String(rows[i][0]),
        title:     String(rows[i][1]||""),
        body:      String(rows[i][2]||""),
        category:  String(rows[i][3]||"Загальне"),
        author:    String(rows[i][4]||""),
        createdAt: String(rows[i][5]||""),
        pinned:    rows[i][6] === true || String(rows[i][6]) === "TRUE"
      });
    }
    // Закріплені спочатку, потім за датою (нові зверху)
    result.sort((a, b) => {
      if (a.pinned && !b.pinned) return -1;
      if (!a.pinned && b.pinned) return 1;
      return b.createdAt.localeCompare(a.createdAt);
    });
    return { ok: true, items: result };
  } catch(e) { return { ok: false, items: [], err: e.message }; }
}

function saveAnnouncement(data, callerLogin) {
  try {
    checkRole(callerLogin, ["Адмін"]); // тільки Адмін
    const u   = getUser(callerLogin);
    const s   = getOrCreate("Announcements",
      ["id","title","body","category","author","createdAt","pinned"]);
    const now = nowStr();

    if (data.id) {
      // Редагування
      const rows = s.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) !== String(data.id)) continue;
        s.getRange(i+1,2,1,6).setValues([[
          String(data.title||"").trim(),
          String(data.body||"").trim(),
          String(data.category||"Загальне"),
          u ? u.name : callerLogin,
          rows[i][5], // зберігаємо оригінальну дату
          data.pinned === true ? true : false
        ]]);
        audit(callerLogin, "ANN_EDIT", data.id, data.title);
        return { ok: true };
      }
    }

    // Створення
    const id = "A-" + Utilities.getUuid().substring(0,8).toUpperCase();
    s.appendRow([
      id,
      String(data.title||"").trim(),
      String(data.body||"").trim(),
      String(data.category||"Загальне"),
      u ? u.name : callerLogin,
      now,
      data.pinned === true ? true : false
    ]);
    audit(callerLogin, "ANN_CREATE", id, data.title);

    // Надсилаємо сповіщення ВСІМ співробітникам
    if (data.notify !== false) {
      _notifyAllUsers(
        "📢 " + String(data.title||""),
        String(data.body||""),
        callerLogin
      );
    }

    return { ok: true, id };
  } catch(e) { return { ok: false, err: e.message }; }
}

function deleteAnnouncement(id, callerLogin) {
  try {
    checkRole(callerLogin, ["Адмін"]);
    const s    = getOrCreate("Announcements",
      ["id","title","body","category","author","createdAt","pinned"]);
    const rows = s.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        s.deleteRow(i+1);
        audit(callerLogin, "ANN_DELETE", id, "");
        return { ok: true };
      }
    }
    return { ok: false, err: "Не знайдено" };
  } catch(e) { return { ok: false, err: e.message }; }
}

// Надіслати оголошення ВСІМ користувачам
function _notifyAllUsers(subject, body, fromLogin) {
  try {
    const users  = sh("Users").getDataRange().getValues();
    const token  = cfg("TG_TOKEN");
    const sender = getUser(fromLogin);
    const senderName = sender ? sender.name : fromLogin;

    for (let i = 1; i < users.length; i++) {
      if (!users[i][0]) continue;
      const login = String(users[i][0]).toLowerCase().trim();
      const name  = String(users[i][1]||"");
      const tgId  = String(users[i][5]||"");
      // Email зі спеціального поля (col 8) або login якщо є @
      const emailField = String(users[i][7]||"").trim();
      const emailAddr  = (emailField && emailField.includes("@")) ? emailField :
                         (login.includes("@") ? login : null);

      const fullBody = body + "\n\n— " + senderName;

      if (emailAddr) {
        try {
          MailApp.sendEmail({
            to: emailAddr,
            subject,
            htmlBody: _emailHtml(subject, fullBody, name)
          });
          Logger.log("Ann email sent to: " + emailAddr);
        } catch(e) {
          Logger.log("Ann email failed to " + emailAddr + ": " + e.message);
        }
      }

      if (tgId && token) {
        _tgSendWithLog(token, tgId, "<b>" + subject + "</b>\n\n" + fullBody);
      }
    }
  } catch(e) { Logger.log("_notifyAllUsers error: " + e.message); }
}

// ─── ІНІЦІАЛІЗАЦІЯ ───────────────────────────────
function initSheets() {
  // Users — додано поле email (col 8)
  getOrCreate("Users",        ["login","name","dept","role","password","tgId","note","email"]);
  getOrCreate("Tasks",        ["id","title","desc","type","status","dept","executors","coExec","author","deadline","startedAt","closedAt","meetLink","priority","revComment"]);
  getOrCreate("Comments",     ["id","taskId","author","authorName","text","createdAt"]);
  getOrCreate("WorkloadLog",  ["login","date","startTime","endTime","status","note","updatedAt"]);
  getOrCreate("InfoRecords",  ["id","dept","heading","category","content","author","createdAt","updatedAt"]);
  getOrCreate("InfoFiles",    ["id","infoId","fileName","fileUrl","uploadedBy","uploadedAt"]);
  getOrCreate("Settings",     ["key","value"]);
  getOrCreate("TgCodes",      ["code","login","createdAt"]);
  getOrCreate("AuditLog",     ["timestamp","actor","action","targetId","detail"]);
  getOrCreate("TaskFiles",    ["id","taskId","fileName","fileUrl","uploadedBy","uploadedAt"]);
  // Оголошення
  getOrCreate("Announcements",["id","title","body","category","author","createdAt","pinned"]);
  Logger.log("✅ Всі листи ініціалізовано (v5)");
}

function setupTriggers() {
  // Видаляємо старі тригери
  ScriptApp.getProjectTriggers().forEach(t => {
    if (["morningDigest","deadlineReminders","weeklyBackup"].includes(t.getHandlerFunction()))
      ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("morningDigest")
    .timeBased().atHour(9).everyDays(1).inTimezone("Europe/Kyiv").create();
  ScriptApp.newTrigger("deadlineReminders")
    .timeBased().atHour(9).everyDays(1).inTimezone("Europe/Kyiv").create();
  ScriptApp.newTrigger("weeklyBackup")
    .timeBased().atHour(3).onWeekDay(ScriptApp.WeekDay.MONDAY).inTimezone("Europe/Kyiv").create();
  Logger.log("✅ Тригери встановлено");
}

// ─── ТЕСТ TELEGRAM ───────────────────────────────
// Запустіть вручну для перевірки: testTelegram()
function testTelegram() {
  const token = cfg("TG_TOKEN");
  Logger.log("TG_TOKEN: " + (token ? "є (" + token.substring(0,10) + "...)" : "ВІДСУТНІЙ"));

  if (!token) {
    Logger.log("❌ Додайте TG_TOKEN в лист Settings (колонка A = 'TG_TOKEN', B = ваш токен)");
    return;
  }

  // Перевірка бота
  const meResult = JSON.parse(
    UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/getMe", {muteHttpExceptions:true})
      .getContentText()
  );
  Logger.log("getMe: " + JSON.stringify(meResult));

  if (!meResult.ok) {
    Logger.log("❌ Токен невірний: " + (meResult.description || "помилка"));
    return;
  }

  Logger.log("✅ Бот: @" + meResult.result.username + " (" + meResult.result.first_name + ")");

  // Webhook
  const webhookInfo = JSON.parse(
    UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/getWebhookInfo", {muteHttpExceptions:true})
      .getContentText()
  );
  Logger.log("Webhook: " + JSON.stringify(webhookInfo.result));
}
