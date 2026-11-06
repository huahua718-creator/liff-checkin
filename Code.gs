// ==========================================
// 🚀 核心加速：GAS CacheService 記憶體快取 (快取 20 分鐘)
// ==========================================
const CACHE_TTL_SECONDS = 1200; // 快取存活時間：20 分鐘

function getAppCache() {
  return CacheService.getScriptCache();
}

// 1. 教師主名單快取
function getCachedTeachers() {
  const cache = getAppCache();
  const cached = cache.get("CACHE_TEACHER_LIST");
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {}
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teacherSheet = ss.getSheetByName("教師主名單");
  if (!teacherSheet) return [];
  const data = teacherSheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[0] ? row[0].toString().trim() : "";
    if (!name) continue;
    let idCode = row[1] !== undefined && row[1] !== null ? row[1].toString().trim() : "";
    if (idCode.length > 0 && idCode.length < 3) idCode = idCode.padStart(3, '0');
    list.push({
      rowIndex: i + 1,
      name: name,
      idCode: idCode,
      lineUserId: row[2] ? row[2].toString().trim() : "",
      status: row[3] ? row[3].toString().trim() : "",
      group: row[5] ? row[5].toString().trim() : ""
    });
  }

  try {
    const jsonStr = JSON.stringify(list);
    if (jsonStr.length < 100000) {
      cache.put("CACHE_TEACHER_LIST", jsonStr, CACHE_TTL_SECONDS);
    }
  } catch (e) {
    Logger.log("教師快取寫入失敗：" + e);
  }
  return list;
}

function clearTeacherCache() {
  try {
    getAppCache().remove("CACHE_TEACHER_LIST");
  } catch (e) {}
}

// 2. 會議地點與時間設定快取
function getCachedMeetings() {
  const cache = getAppCache();
  const cached = cache.get("CACHE_MEETING_SETTINGS");
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {}
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("會議地點與時間設定");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const meetings = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && row[0] !== 0) continue;
    meetings.push({
      rowIndex: i + 1,
      name: row[0] ? row[0].toString().trim() : "",
      dateRule: row[1] instanceof Date ? Utilities.formatDate(row[1], "Asia/Taipei", "yyyy/MM/dd") : (row[1] ? row[1].toString().trim() : ""),
      startTime: row[2] instanceof Date ? Utilities.formatDate(row[2], "Asia/Taipei", "HH:mm") : (row[2] ? row[2].toString().trim() : ""),
      endTime: row[3] instanceof Date ? Utilities.formatDate(row[3], "Asia/Taipei", "HH:mm") : (row[3] ? row[3].toString().trim() : ""),
      passcode: row[4] ? row[4].toString().trim() : "",
      status: row[5] ? row[5].toString().trim() : "停用",
      target: row[6] ? row[6].toString().trim() : "全體",
      meetingAdmins: row[7] ? row[7].toString().trim() : ""
    });
  }

  try {
    const jsonStr = JSON.stringify(meetings);
    if (jsonStr.length < 100000) {
      cache.put("CACHE_MEETING_SETTINGS", jsonStr, CACHE_TTL_SECONDS);
    }
  } catch (e) {
    Logger.log("會議快取寫入失敗：" + e);
  }
  return meetings;
}

function clearMeetingCache() {
  try {
    getAppCache().remove("CACHE_MEETING_SETTINGS");
  } catch (e) {}
}

// 3. 管理員名單快取
function getCachedSystemAdmins() {
  const cache = getAppCache();
  const cached = cache.get("CACHE_SYSTEM_ADMINS");
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {}
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adminSheet = ss.getSheetByName("管理員名單");
  if (!adminSheet) return [];
  const adminData = adminSheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < adminData.length; i++) {
    if (adminData[i][1]) {
      list.push(adminData[i][1].toString().trim());
    }
  }
  try {
    cache.put("CACHE_SYSTEM_ADMINS", JSON.stringify(list), CACHE_TTL_SECONDS);
  } catch (e) {}
  return list;
}

function clearAdminCache() {
  try {
    getAppCache().remove("CACHE_SYSTEM_ADMINS");
  } catch (e) {}
}

// 4. 系統參數快取
function getCachedSystemSettings() {
  const cache = getAppCache();
  const cached = cache.get("CACHE_SYSTEM_SETTINGS");
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {}
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("系統參數") || ss.getSheetByName("系統參數與聯絡設定");
  const settings = {};
  if (sheet) {
    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0]) {
        const key = data[i][0].toString().trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
        settings[key] = data[i][1] ? data[i][1].toString().trim() : "";
      }
    }
    if (!settings["ADMIN_CONTACT_NAME"] && sheet.getRange("B1").getValue()) {
      settings["ADMIN_CONTACT_NAME"] = sheet.getRange("B1").getValue().toString().trim();
    }
    if (!settings["ADMIN_CONTACT_PHONE"] && sheet.getRange("B2").getValue()) {
      settings["ADMIN_CONTACT_PHONE"] = sheet.getRange("B2").getValue().toString().trim();
    }
    if (!settings["LINE_CHANNEL_ACCESS_TOKEN"] && sheet.getRange("B3").getValue()) {
      settings["LINE_CHANNEL_ACCESS_TOKEN"] = sheet.getRange("B3").getValue().toString().trim();
    }
  }
  try {
    cache.put("CACHE_SYSTEM_SETTINGS", JSON.stringify(settings), CACHE_TTL_SECONDS);
  } catch (e) {}
  return settings;
}

function clearSettingsCache() {
  try {
    getAppCache().remove("CACHE_SYSTEM_SETTINGS");
  } catch (e) {}
}

// 5. 清除全部快取 (提供給試算表管理員選單點擊使用)
function clearAllSystemCaches() {
  clearTeacherCache();
  clearMeetingCache();
  clearAdminCache();
  clearSettingsCache();
  SpreadsheetApp.getActiveSpreadsheet().toast("系統快取已全部清除，下次連線將重新載入試算表最新資料！", "🔄 快取清除完成", 5);
}

// ==========================================
// API 主路由 ( doPost ) - 相容 LIFF 與 LINE Webhook
// ==========================================
function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);

    // 1. 判斷是否為 LINE Bot 傳入的訊息事件 (Webhook)
    if (req.events && Array.isArray(req.events)) {
      handleLineWebhook(req.events);
      return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
    }

    // 2. 來自 GitHub Pages / LIFF 前端的 API 請求
    const action = req.action;
    let result;

    switch (action) {
      case "getInitialData":
        result = getInitialData(req.myLineUserId);
        break;
      case "bindUser":
        result = bindUser(req.myLineUserId, req.name, req.idCode);
        break;
      case "processQrCheckIn":
        result = processQrCheckIn(req.myLineUserId, req.scannedCode);
        break;
      case "getAdminDashboardData":
        result = getAdminDashboardData(req.targetMeetingName, req.myLineUserId);
        break;
      case "adminManualCheckIn":
        result = adminManualCheckIn(req.name, req.myLineUserId, req.targetMeetingName);
        break;
      case "sendReminders":
        result = sendReminders(req.targetIds, req.meetingName);
        break;
      default:
        result = { success: false, message: "無效的操作指令" };
      case "getMeetingSettingsData":
        result = getMeetingSettingsData(req.myLineUserId);
        break;
      case "saveMeetingSetting":
        result = saveMeetingSetting(req.myLineUserId, req.meetingData);
        break;
      case "toggleMeetingStatus":
        result = toggleMeetingStatus(req.myLineUserId, req.rowIndex, req.newStatus);
        break;
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: "後端執行異常: " + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// 1. 初始化身分與會議驗證 (getInitialData)
// ==========================================
function getInitialData(myLineUserId) {
  let isBound = false;
  let userName = "";
  let userGroup = "";
  let userTitle = "";

  if (myLineUserId) {
    const teachers = getCachedTeachers();
    const cleanMyId = myLineUserId.toString().trim();

    for (let i = 0; i < teachers.length; i++) {
      if (teachers[i].lineUserId && teachers[i].lineUserId === cleanMyId) {
        isBound = true;
        userName = teachers[i].name;
        userGroup = teachers[i].group;
        
        const rawTitle = teachers[i].group;
        if (rawTitle) {
          userTitle = rawTitle.split(/[,/、\s]+/)[0];
        }
        break;
      }
    }
  }

  if (!isBound) {
    return {
      isBound: false,
      adminContact: getSystemSetting("ADMIN_CONTACT_NAME"),
      adminPhone: getSystemSetting("ADMIN_CONTACT_PHONE")
    };
  }

  // 判斷權限：最高系統管理員 OR 任一啟用會議之專屬管理員
  const isSysAdmin = isSystemAdmin(myLineUserId);
  const manageableMeetings = getManageableMeetingsForUser(userName, myLineUserId);
  const isAdmin = isSysAdmin || manageableMeetings.length > 0;

  // 1. 取得該教師當前時間段內「所有」符合的會議 (使用快取)
  const activeMeetings = getActiveMeetingsForUser(userName, userGroup);

  // 2. 查詢該教師今日「已簽到」的會議項目 (倒序高效率檢索)
  const todayStr = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd");
  const signedSet = new Set();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recordSheet = ss.getSheetByName("簽到紀錄");
  if (recordSheet) {
    const lastRow = recordSheet.getLastRow();
    if (lastRow > 1) {
      // 僅讀取 A~E 欄 (日期、會議名、姓名、LINE ID、狀態)
      const records = recordSheet.getRange(2, 1, lastRow - 1, 5).getValues();
      for (let i = records.length - 1; i >= 0; i--) {
        if (!records[i][0]) continue;
        const rTime = Utilities.formatDate(new Date(records[i][0]), "Asia/Taipei", "yyyy/MM/dd");
        if (rTime !== todayStr && (records.length - i) > 150) break; // 歷史超過 150 筆且非今日則提早中斷

        const rMeeting = records[i][1] ? records[i][1].toString().trim() : "";
        const rLineId = records[i][3] ? records[i][3].toString().trim() : "";
        const rStatus = records[i][4] ? records[i][4].toString().trim() : "";

        if (rTime === todayStr && rLineId === myLineUserId && (rStatus === "成功" || rStatus === "管理員代簽")) {
          signedSet.add(rMeeting);
        }
      }
    }
  }

  // 3. 為每場會議標記是否已簽到 (isSigned)
  const meetingsWithStatus = activeMeetings.map(m => {
    return {
      name: m.name,
      passcode: m.passcode,
      target: m.target,
      isSigned: signedSet.has(m.name)
    };
  });

  return {
    isBound: true,
    userName: userName,
    userTitle: userTitle,
    isAdmin: isAdmin,
    isSysAdmin: isSysAdmin,
    activeMeetings: meetingsWithStatus,
    activeMeeting: meetingsWithStatus.length > 0 ? meetingsWithStatus[0] : null
  };
}

// ==========================================
// 2. 身分綁定 (bindUser) - 3 位數識別碼
// ==========================================
function bindUser(myLineUserId, name, idCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teacherSheet = ss.getSheetByName("教師主名單");
  if (!teacherSheet) return { success: false, message: "找不到「教師主名單」試算表！" };

  const cleanInputIdCode = idCode.trim().padStart(3, '0');
  const cleanInputName = name.trim();

  const data = teacherSheet.getDataRange().getValues();
  let targetRowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    const rowName = data[i][0] ? data[i][0].toString().trim() : "";
    let rowIdCode = data[i][1] !== undefined && data[i][1] !== null ? data[i][1].toString().trim() : "";
    const rowLineId = data[i][2] ? data[i][2].toString().trim() : "";

    if (rowIdCode.length > 0 && rowIdCode.length < 3) {
      rowIdCode = rowIdCode.padStart(3, '0');
    }

    if (rowName === cleanInputName && rowIdCode === cleanInputIdCode) {
      if (rowLineId !== "") {
        if (rowLineId === myLineUserId) {
          return { success: false, message: "您已完成綁定，無需重複操作。" };
        } else {
          return { success: false, message: "⚠️ 此身分已被其他 LINE 帳號綁定，若有疑慮請聯絡管理員。" };
        }
      }
      targetRowIndex = i + 1;
      break;
    }
  }

  if (targetRowIndex === -1) {
    return { success: false, message: "⚠️ 查無此人或身分識別碼錯誤！請確認輸入是否正確。" };
  }

  const nowStr = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm:ss");
  teacherSheet.getRange(targetRowIndex, 3).setValue(myLineUserId);
  teacherSheet.getRange(targetRowIndex, 4).setValue("已綁定");
  teacherSheet.getRange(targetRowIndex, 5).setValue(nowStr);

  clearTeacherCache(); // 清除快取，使新綁定即時生效

  return { success: true, message: "🎉 身分綁定成功！" };
}

// ==========================================
// 3. QR Code 比對與簽到 (processQrCheckIn)
// ==========================================
function processQrCheckIn(myLineUserId, scannedCode) {
  const teacherInfo = getTeacherInfoByLineId(myLineUserId);
  if (!teacherInfo || !teacherInfo.name) {
    return { success: false, message: "請先完成身分綁定！" };
  }

  const teacherName = teacherInfo.name;
  const teacherGroup = teacherInfo.group;

  const activeMeetings = getActiveMeetingsForUser(teacherName, teacherGroup);
  if (!activeMeetings || activeMeetings.length === 0) {
    return { success: false, message: "您目前無需要參加的開會場次，或已超過簽到時間。" };
  }

  const targetMeeting = activeMeetings.find(m => m.passcode && m.passcode.trim() === scannedCode.trim());

  if (!targetMeeting) {
    return { success: false, message: "⚠️ 簽到失敗：條碼不符！您掃描的不是您當前需參加會議的官方 QR Code。" };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recordSheet = ss.getSheetByName("簽到紀錄");
  const todayStr = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd");
  if (recordSheet) {
    const lastRow = recordSheet.getLastRow();
    if (lastRow > 1) {
      const records = recordSheet.getRange(2, 1, lastRow - 1, 5).getValues();
      for (let i = records.length - 1; i >= 0; i--) {
        if (!records[i][0]) continue;
        const rTime = Utilities.formatDate(new Date(records[i][0]), "Asia/Taipei", "yyyy/MM/dd");
        if (rTime !== todayStr && (records.length - i) > 100) break;

        const rMeeting = records[i][1];
        const rLineId = records[i][3];
        const rStatus = records[i][4];

        if (rTime === todayStr && rMeeting === targetMeeting.name && rLineId === myLineUserId) {
          if (rStatus === "成功" || rStatus === "管理員代簽") {
            return { success: false, message: `您剛才已完成【${targetMeeting.name}】的簽到，請勿重複簽到！` };
          }
        }
      }
    }
  }

  const nowStr = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm:ss");
  if (recordSheet) {
    recordSheet.appendRow([
      nowStr, targetMeeting.name, teacherName, myLineUserId, 
      "成功", "QR碼驗證通過", "-", "-", "掃描簽到"
    ]);
  }

  return { success: true, message: `🎉 【${targetMeeting.name}】簽到成功！` };
}

// ==========================================
// 4. 管理員看板與相關邏輯 (支援分權管理)
// ==========================================
function getAdminDashboardData(targetMeetingName, myLineUserId) {
  // 1. 取得請求者姓名
  const teacherInfo = getTeacherInfoByLineId(myLineUserId);
  const userName = teacherInfo ? teacherInfo.name : "";

  // 2. 取得該使用者「有權限管理」的所有會議清單
  const manageableMeetings = getManageableMeetingsForUser(userName, myLineUserId);
  if (manageableMeetings.length === 0) {
    return { success: false, message: "⚠️ 您無權限開啟管理員看板或當前無可管理的會議。" };
  }

  // 3. 確定當前選擇的會議
  let meetingInfo = null;
  if (targetMeetingName) {
    meetingInfo = manageableMeetings.find(m => m.name === targetMeetingName);
  }
  if (!meetingInfo) {
    meetingInfo = manageableMeetings[0];
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recordSheet = ss.getSheetByName("簽到紀錄");
  const todayStr = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd");
  const checkedSet = new Set();
  if (recordSheet) {
    const records = recordSheet.getDataRange().getValues();
    for (let i = 1; i < records.length; i++) {
      if (!records[i][0]) continue;
      const rTime = Utilities.formatDate(new Date(records[i][0]), "Asia/Taipei", "yyyy/MM/dd");
      if (rTime === todayStr && records[i][1] === meetingInfo.name && (records[i][4] === "成功" || records[i][4] === "管理員代簽")) {
        checkedSet.add(records[i][2].toString().trim());
      }
    }
  }

  const uncheckedList = [];
  let totalCount = 0;
  const teachers = getCachedTeachers();
  for (let i = 0; i < teachers.length; i++) {
    const tName = teachers[i].name;
    if (!tName) continue;
    const tGroup = teachers[i].group;

    if (isTeacherMatchMeetingTarget(tName, tGroup, meetingInfo.target)) {
      totalCount++;
      if (!checkedSet.has(tName)) {
        uncheckedList.push({
          name: tName,
          lineUserId: teachers[i].lineUserId,
          isBound: teachers[i].status === "已綁定"
        });
      }
    }
  }

  return {
    success: true,
    data: {
      meetingName: meetingInfo.name,
      totalCount: totalCount,
      checkedCount: checkedSet.size,
      uncheckedList: uncheckedList,
      allMeetings: manageableMeetings.map(m => m.name) // 僅傳回該使用者可管理的會議選單
    }
  };
}

function adminManualCheckIn(name, myLineUserId, targetMeetingName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recordSheet = ss.getSheetByName("簽到紀錄");
  
  const teacherInfo = getTeacherInfoByLineId(myLineUserId);
  const adminName = teacherInfo ? teacherInfo.name : "管理員";

  const meetingInfo = getMeetingByName(targetMeetingName) || getLatestEnabledMeeting();

  const nowStr = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm:ss");
  recordSheet.appendRow([nowStr, meetingInfo.name, name, "-", "管理員代簽", "-", "-", "-", `代簽 (${adminName})`]);
  return { success: true, message: `已成功為 ${name} 老師手動代簽【${meetingInfo.name}】！` };
}

function sendReminders(targetIds, meetingName) {
  const token = getSystemSetting("LINE_CHANNEL_ACCESS_TOKEN");
  if (!token) return { success: false, message: "未設定 LINE_CHANNEL_ACCESS_TOKEN！" };
  const validIds = targetIds.filter(id => id && id.trim() !== "");
  if (validIds.length === 0) return { success: false, message: "未簽到同仁皆尚未綁定 LINE 帳號。" };

  let successCount = 0;
  validIds.forEach(id => {
    const payload = {
      to: id,
      messages: [{
        type: "text",
        text: `🔔 【會議簽到提醒】\n您好：「${meetingName}」目前開放簽到中，請點開選單並【掃描現場 QR Code】完成簽到。`
      }]
    };
    try {
      const res = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
        method: "post",
        headers: { "Content-Type": "application/json", "Authorization": "Bearer " + token },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      if (res.getResponseCode() === 200) successCount++;
    } catch (e) {}
  });
  return { success: true, message: `已成功發送提醒給 ${successCount} 位同仁！` };
}

// ==========================================
// 5. 權限檢查與輔助函式
// ==========================================
// 判斷是否為「最高系統管理員」 (使用快取)
function isSystemAdmin(myLineUserId) {
  if (!myLineUserId) return false;
  const adminIds = getCachedSystemAdmins();
  return adminIds.includes(myLineUserId.toString().trim());
}

// 取得使用者可存取/管理的所有會議 (使用快取)
function getManageableMeetingsForUser(userName, myLineUserId) {
  const allMeetings = getCachedMeetings();
  const sysAdmin = isSystemAdmin(myLineUserId);
  const manageableList = [];

  for (let i = 0; i < allMeetings.length; i++) {
    const m = allMeetings[i];
    if (m.status !== "啟用" || !m.name) continue;

    // 權限檢查：若為最高系統管理員 OR 人名包含在 H 欄中
    if (sysAdmin || isNameInMeetingAdmins(userName, m.meetingAdmins)) {
      manageableList.push({
        name: m.name,
        passcode: m.passcode,
        target: m.target
      });
    }
  }
  return manageableList;
}

function isNameInMeetingAdmins(userName, meetingAdminsStr) {
  if (!userName || !meetingAdminsStr) return false;
  const adminNames = meetingAdminsStr.split(/[,，\s]+/).map(s => s.trim());
  return adminNames.includes(userName.trim());
}

function getTeacherInfoByLineId(myLineUserId) {
  if (!myLineUserId) return null;
  const teachers = getCachedTeachers();
  const cleanMyId = myLineUserId.toString().trim();
  for (let i = 0; i < teachers.length; i++) {
    if (teachers[i].lineUserId && teachers[i].lineUserId === cleanMyId) {
      return {
        name: teachers[i].name,
        group: teachers[i].group
      };
    }
  }
  return null;
}

function getActiveMeetingsForUser(teacherName, teacherGroup) {
  const allMeetings = getCachedMeetings();
  const now = new Date();
  const currentMinutes = now.getHours() * 60 + now.getMinutes();

  const activeMeetings = [];

  for (let i = 0; i < allMeetings.length; i++) {
    const m = allMeetings[i];
    if (m.status !== "啟用") continue;
    if (isMeetingOnDate(m.dateRule, now)) {
      const startMin = parseTimeToMinutes(m.startTime);
      const endMin = parseTimeToMinutes(m.endTime);
      if (startMin !== -1 && endMin !== -1 && currentMinutes >= startMin && currentMinutes <= endMin) {
        if (isTeacherMatchMeetingTarget(teacherName, teacherGroup, m.target)) {
          activeMeetings.push({
            name: m.name,
            passcode: m.passcode,
            target: m.target
          });
        }
      }
    }
  }
  return activeMeetings;
}

function getMeetingByName(name) {
  if (!name) return null;
  const allMeetings = getCachedMeetings();
  for (let i = 0; i < allMeetings.length; i++) {
    const m = allMeetings[i];
    if (m.name === name.trim() && m.status === "啟用") {
      return {
        name: m.name,
        passcode: m.passcode,
        target: m.target
      };
    }
  }
  return null;
}

function getLatestEnabledMeeting() {
  const allMeetings = getCachedMeetings();
  for (let i = allMeetings.length - 1; i >= 0; i--) {
    const m = allMeetings[i];
    if (m.status === "啟用") {
      return {
        name: m.name,
        passcode: m.passcode,
        target: m.target
      };
    }
  }
  return null;
}

function getAllEnabledMeetingNames() {
  const allMeetings = getCachedMeetings();
  const list = [];
  for (let i = 0; i < allMeetings.length; i++) {
    const m = allMeetings[i];
    if (m.status === "啟用" && m.name) list.push(m.name);
  }
  return list;
}

function isTeacherMatchMeetingTarget(teacherName, teacherGroupsStr, meetingTargetStr) {
  if (!meetingTargetStr || meetingTargetStr === "全體" || meetingTargetStr === "全體教師") return true;
  const targets = meetingTargetStr.split(/[,，\s]+/).map(s => s.trim());
  const teacherGroups = teacherGroupsStr.split(/[,，\s]+/).map(s => s.trim());
  if (targets.includes(teacherName)) return true;
  for (let group of teacherGroups) {
    if (group && targets.includes(group)) return true;
  }
  return false;
}

function isMeetingOnDate(dayRuleVal, targetDate) {
  if (!dayRuleVal) return false;
  const dayMap = ["週日", "週一", "週二", "週三", "週四", "週五", "週六"];
  const currentDayStr = dayMap[targetDate.getDay()];
  const dateStr1 = Utilities.formatDate(targetDate, "Asia/Taipei", "yyyy/MM/dd");
  let ruleStr = dayRuleVal instanceof Date ? Utilities.formatDate(dayRuleVal, "Asia/Taipei", "yyyy/MM/dd") : dayRuleVal.toString().trim();
  return ruleStr.includes(currentDayStr) || ruleStr.includes(dateStr1);
}

function parseTimeToMinutes(timeVal) {
  if (timeVal instanceof Date) return timeVal.getHours() * 60 + timeVal.getMinutes();
  if (typeof timeVal === "string" && timeVal.includes(":")) {
    const parts = timeVal.split(":");
    return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
  }
  return -1;
}

// ==========================================
// 6. 系統參數與預熱功能
// ==========================================
function getSystemSetting(keyName) {
  if (!keyName) return "";
  const searchKey = keyName.trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
  const settings = getCachedSystemSettings();
  return settings[searchKey] || "";
}

function smartPing() {
  const allMeetings = getCachedMeetings();
  const now = new Date();
  const currentMinutes = now.getHours() * 60 + now.getMinutes();
  let isAnyMeetingActive = false;

  for (let i = 0; i < allMeetings.length; i++) {
    const m = allMeetings[i];
    if (m.status !== "啟用") continue;

    if (isMeetingOnDate(m.dateRule, now)) {
      const startMin = parseTimeToMinutes(m.startTime);
      const endMin = parseTimeToMinutes(m.endTime);
      if (startMin !== -1 && endMin !== -1 && currentMinutes >= (startMin - 10) && currentMinutes <= endMin) {
        isAnyMeetingActive = true;
        break;
      }
    }
  }

  if (isAnyMeetingActive) {
    getCachedTeachers(); // 預熱快取
    Logger.log("🔥 偵測到開放中之會議，已完成預熱：" + Utilities.formatDate(now, "Asia/Taipei", "yyyy/MM/dd HH:mm:ss"));
  }
}

// ==========================================
// 7. LINE Bot Webhook 與管理員文字選單 (點擊即下令)
// ==========================================
function handleLineWebhook(events) {
  events.forEach(event => {
    if (event.type === "message" && event.message.type === "text") {
      const userText = event.message.text.trim();
      const lowerText = userText.toLowerCase();
      const userId = event.source.userId;
      const replyToken = event.replyToken;

      const teacherInfo = getTeacherInfoByLineId(userId);
      const userName = teacherInfo ? teacherInfo.name : "";
      const sysAdmin = isSystemAdmin(userId);
      const manageableMeetings = getManageableMeetingsForUser(userName, userId);
      const isAdmin = sysAdmin || manageableMeetings.length > 0;

      // 1. 若非管理員：提供一般同仁友善引導
      if (!isAdmin) {
        handleGeneralUserMessage(userId, replyToken, userText, teacherInfo);
        return;
      }

      // 2. 管理員指令處理 (點擊或輸入文字)
      if (lowerText === "qrcode" || lowerText === "qr" || userText.includes("條碼") || userText.includes("qr碼")) {
        processAdminQrCodeRequest(userId, replyToken);
      } else if (userText.includes("概況") || userText.includes("狀況") || userText.includes("統計") || lowerText === "status") {
        processAdminStatusRequest(userId, replyToken);
      } else if (userText.includes("未到") || userText.includes("未簽") || userText.includes("缺席")) {
        processAdminUncheckedListRequest(userId, replyToken);
      } else if (userText.includes("快取") || lowerText.includes("cache") || userText.includes("重整")) {
        processAdminClearCacheRequest(userId, replyToken);
      } else {
        // 預設或輸入「選單、指令、管理、help」：傳送管理員點擊選單
        sendAdminMenu(userId, replyToken, userName);
      }
    }
  });
}

// 產生管理員專屬的 Quick Reply 橫向快捷按鈕 (點擊直接送出指令)
function getAdminQuickReplyItems() {
  const liffUrl = "https://liff.line.me/" + (getSystemSetting("LIFF_ID") || "2011810210-ESVfHIBL");
  return [
    {
      type: "action",
      action: {
        type: "message",
        label: "📷 索取條碼 (QR)",
        text: "QR"
      }
    },
    {
      type: "action",
      action: {
        type: "message",
        label: "📊 簽到概況",
        text: "簽到概況"
      }
    },
    {
      type: "action",
      action: {
        type: "message",
        label: "👥 未簽到名單",
        text: "未簽到名單"
      }
    },
    {
      type: "action",
      action: {
        type: "message",
        label: "🔄 清除快取",
        text: "清除快取"
      }
    },
    {
      type: "action",
      action: {
        type: "uri",
        label: "📱 開啟管理看板",
        uri: liffUrl
      }
    },
    {
      type: "action",
      action: {
        type: "message",
        label: "❓ 顯示選單",
        text: "選單"
      }
    }
  ];
}

// 傳送管理員文字指令選單
function sendAdminMenu(userId, replyToken, userName) {
  const teacherInfo = getTeacherInfoByLineId(userId);
  const name = userName || (teacherInfo ? teacherInfo.name : "管理員");
  const sysAdmin = isSystemAdmin(userId);
  const roleName = sysAdmin ? "最高系統管理員" : "會議管理員";

  const menuText = 
`👮‍♂️ 【管理員文字指令選單】
您好，${name} (${roleName})！

您可直接「點擊下方橫向快捷按鈕」秒速下令，或輸入指令文字：

📷 【QR】：產生今日會議現場簽到條碼
📊 【簽到概況】：查看今日會議應到/實到統計
👥 【未簽到名單】：列出尚未完成簽到的同仁
🔄 【清除快取】：手動清除試算表快取
📱 【開啟管理看板】：進入手機簽到與管理看板

👇 請點擊鍵盤上方快捷按鈕即可直接執行！`;

  replyLineMessage(replyToken, [{
    type: "text",
    text: menuText
  }], getAdminQuickReplyItems());
}

// 一般同仁訊息引導
function handleGeneralUserMessage(userId, replyToken, userText, teacherInfo) {
  const liffUrl = "https://liff.line.me/" + (getSystemSetting("LIFF_ID") || "2011810210-ESVfHIBL");
  if (teacherInfo && teacherInfo.name) {
    replyLineMessage(replyToken, [{
      type: "text",
      text: `👋 您好，${teacherInfo.name} 老師！\n歡迎使用學校會議簽到系統，請點擊下方按鈕開啟簽到頁面完成簽到。`
    }], [
      {
        type: "action",
        action: {
          type: "uri",
          label: "📱 開啟簽到頁面",
          uri: liffUrl
        }
      }
    ]);
  } else {
    replyLineMessage(replyToken, [{
      type: "text",
      text: `👋 您好！您尚未完成教師身分綁定。\n請點擊下方按鈕開啟頁面完成綁定：`
    }], [
      {
        type: "action",
        action: {
          type: "uri",
          label: "🔗 首次身分綁定",
          uri: liffUrl
        }
      }
    ]);
  }
}

// 1. 產生今日簽到條碼
function processAdminQrCodeRequest(userId, replyToken) {
  const teacherInfo = getTeacherInfoByLineId(userId);
  const userName = teacherInfo ? teacherInfo.name : "";
  const sysAdmin = isSystemAdmin(userId);

  const todayMeetings = getTodayEnabledMeetings();
  const allowedMeetings = todayMeetings.filter(m => sysAdmin || isNameInMeetingAdmins(userName, m.meetingAdmin));

  if (allowedMeetings.length === 0) {
    replyLineMessage(replyToken, [{
      type: "text",
      text: "ℹ️ 您目前無管理權限，或今日無可管理的啟用中會議。"
    }], getAdminQuickReplyItems());
    return;
  }

  const messages = [];
  allowedMeetings.forEach(meeting => {
    if (!meeting.passcode) {
      messages.push({
        type: "text",
        text: `📌 【${meeting.name}】未設定現場驗證碼。`
      });
    } else {
      const qrImageUrl = `https://api.qrserver.com/v1/create-qr-code/?size=500x500&data=${encodeURIComponent(meeting.passcode)}`;
      messages.push({
        type: "text",
        text: `📷 【${meeting.name}】現場簽到 QR Code\n驗證碼：${meeting.passcode}`
      });
      messages.push({
        type: "image",
        originalContentUrl: qrImageUrl,
        previewImageUrl: qrImageUrl
      });
    }
  });

  replyLineMessage(replyToken, messages.slice(0, 5), getAdminQuickReplyItems());
}

// 2. 查詢今日簽到概況
function processAdminStatusRequest(userId, replyToken) {
  const teacherInfo = getTeacherInfoByLineId(userId);
  const userName = teacherInfo ? teacherInfo.name : "";
  const sysAdmin = isSystemAdmin(userId);

  const todayMeetings = getTodayEnabledMeetings();
  const allowedMeetings = todayMeetings.filter(m => sysAdmin || isNameInMeetingAdmins(userName, m.meetingAdmin));

  if (allowedMeetings.length === 0) {
    replyLineMessage(replyToken, [{
      type: "text",
      text: "ℹ️ 今日查無您可管理的啟用中會議。"
    }], getAdminQuickReplyItems());
    return;
  }

  const todayStr = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recordSheet = ss.getSheetByName("簽到紀錄");

  const checkedMap = new Map();
  if (recordSheet) {
    const lastRow = recordSheet.getLastRow();
    if (lastRow > 1) {
      const records = recordSheet.getRange(2, 1, lastRow - 1, 5).getValues();
      for (let i = records.length - 1; i >= 0; i--) {
        if (!records[i][0]) continue;
        const rTime = Utilities.formatDate(new Date(records[i][0]), "Asia/Taipei", "yyyy/MM/dd");
        if (rTime !== todayStr && (records.length - i) > 150) break;

        const rMeeting = records[i][1] ? records[i][1].toString().trim() : "";
        const rName = records[i][2] ? records[i][2].toString().trim() : "";
        const rStatus = records[i][4] ? records[i][4].toString().trim() : "";

        if (rTime === todayStr && (rStatus === "成功" || rStatus === "管理員代簽")) {
          if (!checkedMap.has(rMeeting)) checkedMap.set(rMeeting, new Set());
          checkedMap.get(rMeeting).add(rName);
        }
      }
    }
  }

  const teachers = getCachedTeachers();
  let text = `📊 【今日會議簽到概況】\n📅 日期：${todayStr}\n`;

  allowedMeetings.forEach((m, idx) => {
    let totalCount = 0;
    for (let i = 0; i < teachers.length; i++) {
      if (teachers[i].name && isTeacherMatchMeetingTarget(teachers[i].name, teachers[i].group, m.target || "全體")) {
        totalCount++;
      }
    }

    const checkedSet = checkedMap.get(m.name) || new Set();
    const checkedCount = checkedSet.size;
    const uncheckedCount = Math.max(0, totalCount - checkedCount);
    const rate = totalCount > 0 ? ((checkedCount / totalCount) * 100).toFixed(1) + "%" : "0%";

    text += `\n${idx + 1}. 📍 【${m.name}】`;
    text += `\n   👥 應到：${totalCount} 人 | ✅ 實到：${checkedCount} 人`;
    text += `\n   ⏳ 未到：${uncheckedCount} 人 | 📈 出席率：${rate}`;
  });

  text += `\n\n👉 點擊下方【👥 未簽到名單】可查看缺席名冊。`;

  replyLineMessage(replyToken, [{
    type: "text",
    text: text
  }], getAdminQuickReplyItems());
}

// 3. 查詢今日未簽到人員名單
function processAdminUncheckedListRequest(userId, replyToken) {
  const teacherInfo = getTeacherInfoByLineId(userId);
  const userName = teacherInfo ? teacherInfo.name : "";
  const sysAdmin = isSystemAdmin(userId);

  const todayMeetings = getTodayEnabledMeetings();
  const allowedMeetings = todayMeetings.filter(m => sysAdmin || isNameInMeetingAdmins(userName, m.meetingAdmin));

  if (allowedMeetings.length === 0) {
    replyLineMessage(replyToken, [{
      type: "text",
      text: "ℹ️ 今日查無您可管理的啟用中會議。"
    }], getAdminQuickReplyItems());
    return;
  }

  const todayStr = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recordSheet = ss.getSheetByName("簽到紀錄");

  const checkedMap = new Map();
  if (recordSheet) {
    const lastRow = recordSheet.getLastRow();
    if (lastRow > 1) {
      const records = recordSheet.getRange(2, 1, lastRow - 1, 5).getValues();
      for (let i = records.length - 1; i >= 0; i--) {
        if (!records[i][0]) continue;
        const rTime = Utilities.formatDate(new Date(records[i][0]), "Asia/Taipei", "yyyy/MM/dd");
        if (rTime !== todayStr && (records.length - i) > 150) break;

        const rMeeting = records[i][1] ? records[i][1].toString().trim() : "";
        const rName = records[i][2] ? records[i][2].toString().trim() : "";
        const rStatus = records[i][4] ? records[i][4].toString().trim() : "";

        if (rTime === todayStr && (rStatus === "成功" || rStatus === "管理員代簽")) {
          if (!checkedMap.has(rMeeting)) checkedMap.set(rMeeting, new Set());
          checkedMap.get(rMeeting).add(rName);
        }
      }
    }
  }

  const teachers = getCachedTeachers();
  let text = `👥 【今日未簽到名單】\n📅 日期：${todayStr}\n`;

  allowedMeetings.forEach(m => {
    const checkedSet = checkedMap.get(m.name) || new Set();
    const uncheckeds = [];

    for (let i = 0; i < teachers.length; i++) {
      const tName = teachers[i].name;
      if (tName && isTeacherMatchMeetingTarget(tName, teachers[i].group, m.target || "全體")) {
        if (!checkedSet.has(tName)) {
          const boundTag = teachers[i].status === "已綁定" ? "已綁定" : "未綁定";
          uncheckeds.push(`${tName} (${boundTag})`);
        }
      }
    }

    text += `\n📌 【${m.name}】(未到 ${uncheckeds.length} 人)`;
    if (uncheckeds.length === 0) {
      text += `\n   🎉 全員皆已完成簽到！`;
    } else {
      const displayList = uncheckeds.slice(0, 25);
      text += `\n   ` + displayList.join("、");
      if (uncheckeds.length > 25) {
        text += `\n   ...等共 ${uncheckeds.length} 人`;
      }
    }
  });

  text += `\n\n💡 提示：若需一鍵 LINE 催簽提醒或代簽，請點擊【📱 開啟管理看板】。`;

  replyLineMessage(replyToken, [{
    type: "text",
    text: text
  }], getAdminQuickReplyItems());
}

// 4. 清除系統快取
function processAdminClearCacheRequest(userId, replyToken) {
  clearAllSystemCaches();
  replyLineMessage(replyToken, [{
    type: "text",
    text: "🔄 【系統快取已清除完成】\n已成功重整所有試算表快取！下次連線將直接讀取試算表之最新資料。"
  }], getAdminQuickReplyItems());
}

function getTodayEnabledMeetings() {
  const allMeetings = getCachedMeetings();
  const now = new Date();
  const meetings = [];

  for (let i = 0; i < allMeetings.length; i++) {
    const m = allMeetings[i];
    if (m.status !== "啟用") continue;

    if (isMeetingOnDate(m.dateRule, now)) {
      meetings.push({
        name: m.name,
        passcode: m.passcode,
        target: m.target || "全體",
        meetingAdmin: m.meetingAdmins // 會議管理員
      });
    }
  }
  return meetings;
}

function replyLineMessage(replyToken, messages, quickReplyItems) {
  const token = getSystemSetting("LINE_CHANNEL_ACCESS_TOKEN");
  if (!token) return;

  if (quickReplyItems && quickReplyItems.length > 0 && messages.length > 0) {
    const lastMsg = messages[messages.length - 1];
    lastMsg.quickReply = {
      items: quickReplyItems
    };
  }

  const payload = { replyToken: replyToken, messages: messages };
  try {
    UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply", {
      method: "post",
      headers: { "Content-Type": "application/json", "Authorization": "Bearer " + token },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log("LINE 回覆失敗: " + e);
  }
}

// ==========================================
// 8. 試算表選單與擴充工具
// ==========================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('會議管理功能')
    .addItem('🏷️ 批次新增教師標籤', 'showTagModal')
    .addItem('📷 產生會議驗證碼 QR Code', 'showQRCodeModal')
    .addItem('📊 產生會議出席統計報表', 'showReportModal')
    .addSeparator()
    .addItem('🔄 清除系統快取 (重載設定)', 'clearAllSystemCaches')
    .addToUi();
}

function showTagModal() {
  const html = HtmlService.createTemplateFromFile('TagDialog').evaluate().setWidth(480).setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, '批次設定教師職務/群組標籤');
}

function getTeacherListForTagging() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teacherSheet = ss.getSheetByName("教師主名單");
  if (!teacherSheet) return [];
  const data = teacherSheet.getDataRange().getValues();
  const teachers = [];
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0] ? data[i][0].toString().trim() : "";
    if (name) {
      teachers.push({
        rowIndex: i + 1,
        name: name,
        currentTag: data[i][5] ? data[i][5].toString().trim() : ""
      });
    }
  }
  return teachers;
}

function addTagsToTeachers(selectedRowIndexes, newTag) {
  try {
    if (!newTag || newTag.trim() === "") return { success: false, message: "標籤主題不可為空！" };
    if (!selectedRowIndexes || selectedRowIndexes.length === 0) return { success: false, message: "請至少勾選一位教師！" };

    const cleanTag = newTag.trim();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const teacherSheet = ss.getSheetByName("教師主名單");

    let updatedCount = 0;
    selectedRowIndexes.forEach(rowIndex => {
      const cell = teacherSheet.getRange(rowIndex, 6);
      const currentVal = cell.getValue() ? cell.getValue().toString().trim() : "";

      if (!currentVal) {
        cell.setValue(cleanTag);
        updatedCount++;
      } else {
        const tags = currentVal.split(/[,，]\s*/).map(t => t.trim());
        if (!tags.includes(cleanTag)) {
          tags.push(cleanTag);
          cell.setValue(tags.join(", "));
          updatedCount++;
        }
      }
    });

    clearTeacherCache(); // 清除快取

    return { success: true, message: `🎉 成功為 ${updatedCount} 位教師新增標籤【${cleanTag}】！` };
  } catch (e) {
    return { success: false, message: "寫入失敗: " + e.toString() };
  }
}

function showQRCodeModal() {
  const template = HtmlService.createTemplateFromFile('QRCodeDialog');
  template.meetings = getEnabledMeetingsForQRCode();
  const html = template.evaluate().setWidth(480).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '📷 現場驗證碼 QR Code');
}

function showReportModal() {
  const template = HtmlService.createTemplateFromFile('ReportDialog');
  template.meetings = getAllMeetingNames();
  const html = template.evaluate().setWidth(420).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, '📊 產生會議出席統計報表');
}

function getEnabledMeetingsForQRCode() {
  const allMeetings = getCachedMeetings();
  const meetings = [];
  for (let i = 0; i < allMeetings.length; i++) {
    const m = allMeetings[i];
    if (m.name && m.status === "啟用") {
      meetings.push({
        name: m.name,
        passcode: m.passcode,
        target: m.target
      });
    }
  }
  return meetings;
}

function getAllMeetingNames() {
  const allMeetings = getCachedMeetings();
  const names = [];
  for (let i = 0; i < allMeetings.length; i++) {
    if (allMeetings[i].name) names.push(allMeetings[i].name);
  }
  return names;
}

function generateAttendanceReport(meetingName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const meetingInfo = getMeetingByName(meetingName) || getLatestEnabledMeeting();
    if (!meetingInfo) return { success: false, message: "找不到指定的會議設定資料！" };

    const recordSheet = ss.getSheetByName("簽到紀錄");

    const checkInMap = new Map();
    if (recordSheet) {
      const records = recordSheet.getDataRange().getValues();
      for (let i = 1; i < records.length; i++) {
        const rMeeting = records[i][1] ? records[i][1].toString().trim() : "";
        const rStatus = records[i][4] ? records[i][4].toString().trim() : "";
        if (rMeeting === meetingName && (rStatus === "成功" || rStatus === "管理員代簽")) {
          const rName = records[i][2] ? records[i][2].toString().trim() : "";
          const rTime = records[i][0] ? Utilities.formatDate(new Date(records[i][0]), "Asia/Taipei", "yyyy/MM/dd HH:mm:ss") : "";
          const rMethod = records[i][8] ? records[i][8].toString().trim() : (records[i][5] || "簽到");
          checkInMap.set(rName, { time: rTime, status: "已出席", method: rMethod });
        }
      }
    }

    const teachers = getCachedTeachers();
    const reportRows = [];
    let attendedCount = 0;
    let totalTarget = 0;

    for (let i = 0; i < teachers.length; i++) {
      const tName = teachers[i].name;
      if (!tName) continue;
      const tGroup = teachers[i].group;

      if (isTeacherMatchMeetingTarget(tName, tGroup, meetingInfo.target)) {
        totalTarget++;
        if (checkInMap.has(tName)) {
          attendedCount++;
          const info = checkInMap.get(tName);
          reportRows.push([tName, tGroup || "-", "已出席", info.time, info.method]);
        } else {
          reportRows.push([tName, tGroup || "-", "未出席", "-", "-"]);
        }
      }
    }

    const absentCount = totalTarget - attendedCount;
    const rate = totalTarget > 0 ? ((attendedCount / totalTarget) * 100).toFixed(1) + "%" : "0%";

    const cleanSheetName = `報表_${meetingName.substring(0, 15)}`;
    let reportSheet = ss.getSheetByName(cleanSheetName);
    if (reportSheet) {
      reportSheet.clear();
    } else {
      reportSheet = ss.insertSheet(cleanSheetName);
    }

    reportSheet.getRange("A1").setValue(`📊 【${meetingName}】會議出席統計報表`).setFontSize(14).setFontWeight("bold");
    reportSheet.getRange("A2:E2").setValues([["應到人數", "實到人數", "未到人數", "出席率", "報表建立時間"]])
      .setFontWeight("bold").setBackground("#e2e8f0").setHorizontalAlignment("center");
    
    const nowStr = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm");
    reportSheet.getRange("A3:E3").setValues([[totalTarget, attendedCount, absentCount, rate, nowStr]])
      .setHorizontalAlignment("center");

    reportSheet.getRange("A5:E5").setValues([["教師姓名", "職務/標籤", "出席狀態", "簽到時間", "簽到方式"]])
      .setFontWeight("bold").setBackground("#1DB446").setFontColor("#ffffff").setHorizontalAlignment("center");

    if (reportRows.length > 0) {
      reportSheet.getRange(6, 1, reportRows.length, 5).setValues(reportRows);
      for (let r = 0; r < reportRows.length; r++) {
        const statusCell = reportSheet.getRange(6 + r, 3);
        if (reportRows[r][2] === "已出席") {
          statusCell.setBackground("#d4edda").setFontColor("#155724").setFontWeight("bold").setHorizontalAlignment("center");
        } else {
          statusCell.setBackground("#f8d7da").setFontColor("#721c24").setFontWeight("bold").setHorizontalAlignment("center");
        }
      }
    }

    reportSheet.autoResizeColumns(1, 5);
    ss.setActiveSheet(reportSheet);

    return { success: true, message: `🎉 已成功產生【${cleanSheetName}】統計頁籤！` };

  } catch (e) {
    return { success: false, message: "產生報表失敗：" + e.toString() };
  }
}
// ==========================================
// 🆕 最高系統管理員專屬：會議地點與時間設定 API
// ==========================================
function getMeetingSettingsData(myLineUserId) {
  if (!isSystemAdmin(myLineUserId)) {
    return { success: false, message: "⚠️ 權限不足，僅限最高系統管理員操作！" };
  }

  // 取得快取的最新會議設定
  const meetings = getCachedMeetings();
  return { success: true, meetings: meetings };
}

function saveMeetingSetting(myLineUserId, meetingData) {
  if (!isSystemAdmin(myLineUserId)) {
    return { success: false, message: "⚠️ 權限不足，僅限最高系統管理員操作！" };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("會議地點與時間設定");
  if (!sheet) return { success: false, message: "找不到「會議地點與時間設定」頁籤！" };

  const rowIndex = meetingData.rowIndex;
  const rowValues = [
    meetingData.name,
    meetingData.dateRule,
    meetingData.startTime,
    meetingData.endTime,
    meetingData.passcode,
    meetingData.status,
    meetingData.target || "全體",
    meetingData.meetingAdmins || ""
  ];

  if (rowIndex && rowIndex > 1) {
    // 編輯現有會議
    sheet.getRange(rowIndex, 1, 1, 8).setValues([rowValues]);
    clearMeetingCache(); // 清除快取
    return { success: true, message: `🎉 成功更新【${meetingData.name}】會議設定！` };
  } else {
    // 新增會議
    sheet.appendRow(rowValues);
    clearMeetingCache(); // 清除快取
    return { success: true, message: `🎉 成功新增【${meetingData.name}】會議設定！` };
  }
}

function toggleMeetingStatus(myLineUserId, rowIndex, newStatus) {
  if (!isSystemAdmin(myLineUserId)) {
    return { success: false, message: "⚠️ 權限不足，僅限最高系統管理員操作！" };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("會議地點與時間設定");
  if (!sheet) return { success: false, message: "找不到「會議地點與時間設定」頁籤！" };

  sheet.getRange(rowIndex, 6).setValue(newStatus); // 修改 F 欄啟用狀態
  clearMeetingCache(); // 清除快取
  return { success: true, message: `已將該會議切換為【${newStatus}】！` };
}