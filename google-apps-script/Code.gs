const SHEETS = {
  schedule: "排班資料",
  members: "人員名冊",
  leaves: "請假資料",
  settings: "系統設定"
};

function doGet(e) {
  try {
    if ((e.parameter.action || "getState") !== "getState") {
    return json_({ ok: false, error: "不支援的 action" }, e.parameter.callback);
    }
    return json_({ ok: true, state: loadState_() }, e.parameter.callback);
  } catch (error) {
    return json_({ ok: false, error: error.message }, e.parameter.callback);
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const payload = JSON.parse(e.postData.contents || "{}");
    if (payload.action !== "saveState" || !payload.state) {
      return json_({ ok: false, error: "請提供 action=saveState 與 state" });
    }
    saveState_(payload.state);
    return json_({ ok: true, savedAt: new Date().toISOString() });
  } catch (error) {
    return json_({ ok: false, error: error.message });
  } finally {
    lock.releaseLock();
  }
}

function json_(value, callback) {
  const text = JSON.stringify(value);
  if (callback && /^[A-Za-z_$][\w$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + "(" + text + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name, headers) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sheet;
}

function replaceData_(sheet, headers, rows) {
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.setFrozenRows(1);
}

function loadState_() {
  const membersSheet = sheet_(SHEETS.members, ["id", "姓名", "職稱", "啟用", "本月指定值日天數", "備註"]);
  const membersRows = values_(membersSheet);
  const members = membersRows.slice(1).filter(row => row[0] || row[1]).map(row => ({
    id: String(row[0] || "m" + Math.random().toString(36).slice(2)),
    name: String(row[1] || ""),
    rank: String(row[2] || ""),
    active: row[3] !== false && String(row[3]).toLowerCase() !== "false",
    targetDutyDays: Number(row[4]) || 0,
    notes: String(row[5] || "")
  }));

  const scheduleRows = values_(sheet_(SHEETS.schedule, ["日期", "週次", "人員ID", "人員", "星期", "班別", "支援", "備註"]));
  const duties = {};
  scheduleRows.slice(1).forEach(row => {
    const date = dateString_(row[0]);
    if (!date) return;
    let memberId = "";
    if (row[3]) {
      const member = members.find(item => item.name === String(row[3]));
      memberId = member ? member.id : "";
    }
    if (!memberId) memberId = String(row[2] || "");
    if (memberId) duties[date] = memberId;
  });

  const leaveRows = values_(sheet_(SHEETS.leaves, ["日期", "人員ID", "人員", "假別", "備註"]));
  const leaves = {};
  leaveRows.slice(1).forEach(row => {
    const date = dateString_(row[0]);
    if (!date || !row[1] && !row[2]) return;
    let memberId = "";
    if (row[2]) {
      const member = members.find(item => item.name === String(row[2]));
      memberId = member ? member.id : "";
    }
    if (!memberId) memberId = String(row[1] || "");
    if (!memberId) return;
    if (!leaves[date]) leaves[date] = [];
    leaves[date].push({ memberId: memberId, type: String(row[3] || "特別行程"), notes: String(row[4] || "") });
  });

  const settingsRows = values_(sheet_(SHEETS.settings, ["設定", "值"]));
  const settings = {};
  settingsRows.slice(1).forEach(row => { if (row[0]) settings[String(row[0])] = row[1]; });
  return {
    unitName: String(settings.unitName || "艦隊分署 第五海巡隊"),
    currentYear: Number(settings.currentYear) || new Date().getFullYear(),
    currentMonth: Number(settings.currentMonth) || new Date().getMonth() + 1,
    members: members,
    leaves: leaves,
    duties: duties,
    customHolidays: parseJson_(settings.customHolidays, {}),
    officialCalendarUpdatedAt: String(settings.officialCalendarUpdatedAt || ""),
    monthlyRotationLeaveQuota: Number(settings.monthlyRotationLeaveQuota) || 0,
    trainingDates: parseJson_(settings.trainingDates, {}),
    viewMode: String(settings.viewMode || "table")
  };
}

function saveState_(state) {
  const members = Array.isArray(state.members) ? state.members : [];
  const membersHeaders = ["id", "姓名", "職稱", "啟用", "本月指定值日天數", "備註"];
  const memberRows = members.map(member => [member.id || "", member.name || "", member.rank || "", !!member.active, Number(member.targetDutyDays) || 0, member.notes || ""]);
  replaceData_(sheet_(SHEETS.members, membersHeaders), membersHeaders, memberRows);

  const memberMap = {};
  members.forEach(member => { memberMap[member.id] = member; });
  const scheduleHeaders = ["日期", "週次", "人員ID", "人員", "星期", "班別", "支援", "備註"];
  const duties = state.duties || {};
  const scheduleRows = Object.keys(duties).sort().map(date => {
    const member = memberMap[duties[date]] || {};
    const day = Number(String(date).slice(-2));
    return [date, "第" + Math.ceil(day / 7) + "週", duties[date], member.name || "", weekday_(date), "值日", "", ""];
  });
  replaceData_(sheet_(SHEETS.schedule, scheduleHeaders), scheduleHeaders, scheduleRows);

  const leaveHeaders = ["日期", "人員ID", "人員", "假別", "備註"];
  const leaveRows = [];
  Object.keys(state.leaves || {}).sort().forEach(date => (state.leaves[date] || []).forEach(leave => {
    const member = memberMap[leave.memberId] || {};
    leaveRows.push([date, leave.memberId || "", member.name || "", leave.type || "特別行程", leave.notes || ""]);
  }));
  replaceData_(sheet_(SHEETS.leaves, leaveHeaders), leaveHeaders, leaveRows);

  const settingHeaders = ["設定", "值"];
  const settingRows = [
    ["unitName", state.unitName || ""],
    ["currentYear", state.currentYear || ""],
    ["currentMonth", state.currentMonth || ""],
    ["customHolidays", JSON.stringify(state.customHolidays || {})],
    ["officialCalendarUpdatedAt", state.officialCalendarUpdatedAt || ""],
    ["monthlyRotationLeaveQuota", Number(state.monthlyRotationLeaveQuota) || 0],
    ["trainingDates", JSON.stringify(state.trainingDates || {})],
    ["viewMode", state.viewMode || "table"]
  ];
  replaceData_(sheet_(SHEETS.settings, settingHeaders), settingHeaders, settingRows);
}

function values_(sheet) {
  return sheet.getLastRow() ? sheet.getDataRange().getValues() : [];
}

function parseJson_(value, fallback) {
  try { return value ? JSON.parse(value) : fallback; } catch (error) { return fallback; }
}

function dateString_(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(value).slice(0, 10);
}

function weekday_(date) {
  const days = ["日", "一", "二", "三", "四", "五", "六"];
  return days[new Date(date + "T00:00:00").getDay()];
}
