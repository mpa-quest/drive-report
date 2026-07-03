// =====================
// 設定
// =====================
const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");

// =====================
// GETリクエスト（スタッフ照合・会社一覧・運行記録照会）
// =====================
function doGet(e) {
  const action = e.parameter.action;

  if (action === "getStaff")           { return getStaff(e.parameter.lineUserId); }
  if (action === "getCompanies")       { return getCompanies(); }
  if (action === "getStaffList")       { return getStaffList(); }
  if (action === "getRecords")         { return getRecords(e.parameter); }
  if (action === "getRecordsByStaff")  { return getRecordsByStaff(e.parameter); }
  if (action === "getMyRecords")       { return getMyRecords(e.parameter); }
  if (action === "verifyCustomer")     { return verifyCustomer(e.parameter); }
  if (action === "getCustomerRecords") { return getCustomerRecords(e.parameter); }
  if (action === "getFormOptions")     { return getFormOptions(); }

  return jsonResponse({ error: "unknown action" });
}

// =====================
// POSTリクエスト（運行記録の保存・スタッフ登録）
// =====================
function doPost(e) {
  const payload = JSON.parse(e.postData.contents);
  const action = payload.action;

  if (action === "submitReport")       { return submitReport(payload); }
  if (action === "registerStaff")      { return registerStaff(payload); }
  if (action === "uploadExpenseFiles") { return uploadExpenseFiles(payload); }

  return jsonResponse({ error: "unknown action" });
}

// =====================
// スタッフ新規登録
// =====================
function registerStaff(payload) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("スタッフマスタ");
  const data  = sheet.getDataRange().getValues();

  // 重複チェック
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === payload.lineUserId) {
      return jsonResponse({ success: false, reason: "already registered" });
    }
  }

  // スタッフIDを自動付番（E + 最終行 を4桁ゼロ埋め）
  const lastRow = sheet.getLastRow();
  const staffId = "E" + String(lastRow).padStart(4, "0");

  sheet.appendRow([payload.lineUserId, staffId, payload.staffName]);

  return jsonResponse({ success: true, staffId });
}

// =====================
// スタッフ照合
// =====================
function getStaff(lineUserId) {
  const ss         = SpreadsheetApp.openById(SPREADSHEET_ID);
  const staffSheet = ss.getSheetByName("スタッフマスタ");
  const staffData  = staffSheet.getDataRange().getValues();

  for (let i = 1; i < staffData.length; i++) {
    if (staffData[i][0] !== lineUserId) continue;

    const staffId      = staffData[i][1];
    const staffName    = staffData[i][2];
    const companyIdStr = staffData[i][3] ? String(staffData[i][3]) : "";
    const companyIds   = companyIdStr
      ? companyIdStr.split(",").map(id => id.trim()).filter(id => id)
      : [];

    const companySheet = ss.getSheetByName("会社マスタ");
    const companyData  = companySheet.getDataRange().getValues();
    const companyMap   = {};
    for (let j = 1; j < companyData.length; j++) {
      companyMap[String(companyData[j][0]).trim()] = companyData[j][1];
    }

    const companies = companyIds.map(id => ({ id, name: companyMap[id] || id }));
    return jsonResponse({ found: true, staffId, staffName, companyIds: companies });
  }

  return jsonResponse({ found: false });
}

// =====================
// 会社マスタ一覧を返す（records.html の会社プルダウン用）
// =====================
function getCompanies() {
  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const data = ss.getSheetByName("会社マスタ").getDataRange().getValues();
  const companies = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    companies.push({
      id:   String(data[i][0]).trim(),
      name: String(data[i][1]).trim(),
    });
  }
  return jsonResponse({ companies });
}

// =====================
// スタッフマスタ一覧を返す（records.html スタッフ検索用）
// =====================
function getStaffList() {
  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const data = ss.getSheetByName("スタッフマスタ").getDataRange().getValues();
  const staff = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][1]) continue;
    staff.push({
      staffId:   String(data[i][1]).trim(),
      staffName: String(data[i][2]).trim(),
    });
  }
  return jsonResponse({ staff });
}

// =====================
// 自分の運行記録を返す（myrecords.html 用・会社またぎ全件）
// =====================
function getMyRecords(params) {
  const lineUserId = String(params.lineUserId || "").trim();
  const year       = parseInt(params.year);
  const month      = parseInt(params.month);

  if (!lineUserId || isNaN(year) || isNaN(month)) {
    return jsonResponse({ error: "invalid params", records: [] });
  }

  const ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
  const staffData = ss.getSheetByName("スタッフマスタ").getDataRange().getValues();
  let staffId = null, staffName = null;
  for (let i = 1; i < staffData.length; i++) {
    if (String(staffData[i][0]).trim() === lineUserId) {
      staffId   = String(staffData[i][1]).trim();
      staffName = String(staffData[i][2]).trim();
      break;
    }
  }
  if (!staffId) return jsonResponse({ error: "staff not found", records: [] });

  const sheet       = ss.getSheetByName("運行記録");
  const cm          = getColumnMap(sheet);
  const values      = sheet.getDataRange().getValues();
  const displayVals = sheet.getDataRange().getDisplayValues();
  const records     = [];

  for (let i = 1; i < values.length; i++) {
    const row     = values[i];
    const dispRow = displayVals[i];

    if (String(row[cm["スタッフID"]]).trim() !== staffId) continue;

    const dateVal = row[cm["日付"]];
    if (!dateVal) continue;
    const date = new Date(dateVal);
    if (date.getFullYear() !== year || date.getMonth() + 1 !== month) continue;

    records.push(buildRecord(dispRow, cm));
  }

  records.sort((a, b) => {
    if (a.date !== b.date) return a.date < b.date ? -1 : 1;
    return (a.departureTime || "") < (b.departureTime || "") ? -1 : 1;
  });

  return jsonResponse({ staffId, staffName, records });
}

// =====================
// 指定スタッフ・年月の運行記録を全列返す（records.html スタッフ検索用）
// =====================
function getRecordsByStaff(params) {
  const staffId = String(params.staffId || "").trim();
  const year    = parseInt(params.year);
  const month   = parseInt(params.month);

  if (!staffId || isNaN(year) || isNaN(month)) {
    return jsonResponse({ error: "invalid params", records: [] });
  }

  const ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet       = ss.getSheetByName("運行記録");
  const cm          = getColumnMap(sheet);
  const values      = sheet.getDataRange().getValues();
  const displayVals = sheet.getDataRange().getDisplayValues();
  const records     = [];

  for (let i = 1; i < values.length; i++) {
    const row     = values[i];
    const dispRow = displayVals[i];

    if (String(row[cm["スタッフID"]]).trim() !== staffId) continue;

    const dateVal = row[cm["日付"]];
    if (!dateVal) continue;
    const date = new Date(dateVal);
    if (date.getFullYear() !== year || date.getMonth() + 1 !== month) continue;

    records.push(buildRecord(dispRow, cm));
  }

  records.sort((a, b) => {
    if (a.date !== b.date) return a.date < b.date ? -1 : 1;
    return (a.departureTime || "") < (b.departureTime || "") ? -1 : 1;
  });

  return jsonResponse({ records });
}

// =====================
// 指定会社・年月の運行記録を全列返す（records.html 用）
// =====================
function getRecords(params) {
  const companyId = String(params.companyId || "").trim();
  const year      = parseInt(params.year);
  const month     = parseInt(params.month);

  if (!companyId || isNaN(year) || isNaN(month)) {
    return jsonResponse({ error: "invalid params", records: [] });
  }

  const ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet       = ss.getSheetByName("運行記録");
  const cm          = getColumnMap(sheet);
  const values      = sheet.getDataRange().getValues();
  const displayVals = sheet.getDataRange().getDisplayValues();
  const records     = [];

  for (let i = 1; i < values.length; i++) {
    const row     = values[i];
    const dispRow = displayVals[i];

    const dateVal = row[cm["日付"]];
    if (!dateVal) continue;
    const date = new Date(dateVal);
    if (date.getFullYear() !== year || date.getMonth() + 1 !== month) continue;

    // 会社ID正規化（"C001_会社名" 形式にも対応）
    const rawCompany   = String(row[cm["会社"]] || "").trim();
    const rowCompanyId = rawCompany.includes("_") ? rawCompany.split("_")[0] : rawCompany;
    if (rowCompanyId !== companyId) continue;

    records.push(buildRecord(dispRow, cm));
  }

  records.sort((a, b) => {
    if (a.date !== b.date) return a.date < b.date ? -1 : 1;
    return (a.departureTime || "") < (b.departureTime || "") ? -1 : 1;
  });

  return jsonResponse({ records });
}

// =====================
// 運行記録の保存
// =====================
function submitReport(payload) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("運行記録");

  const now       = new Date();
  const timestamp = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");

  // 会社名を先に取得して ID_名前 形式を作成
  const companySheet = ss.getSheetByName("会社マスタ");
  const companyData  = companySheet.getDataRange().getValues();
  let companyLabel   = payload.company;
  for (let i = 1; i < companyData.length; i++) {
    if (String(companyData[i][0]).trim() === String(payload.company).trim()) {
      companyLabel = payload.company + "_" + companyData[i][1];
      break;
    }
  }
  payload.companyLabel = companyLabel;

  // ヘッダー行から列マップを取得して動的に行を構成
  const colMap  = getColumnMap(sheet);
  const lastCol = Math.max(...Object.values(colMap)) + 1;
  const row     = new Array(lastCol).fill("");

  function setCol(name, value) {
    if (colMap[name] !== undefined) row[colMap[name]] = value;
  }

  setCol("タイムスタンプ", timestamp);
  setCol("会社",           companyLabel);
  setCol("日付",           payload.date);
  setCol("スタッフID",     payload.staffId);
  setCol("運行者",         payload.staffName);
  setCol("車両",           payload.vehicle);
  setCol("出発",           payload.departure);
  setCol("行き先",         payload.destination);
  setCol("到着",           payload.arrival);
  setCol("出発時間", payload.departureTime || "");
  setCol("到着時間", payload.arrivalTime   || "");
  setCol("到着日",         payload.arrivalDate || payload.date);
  setCol("稼働時間",       payload.workingHours);
  setCol("利用者",         payload.passenger);
  setCol("利用目的",       payload.purpose);
  setCol("メーター走行前", payload.meterBefore);
  setCol("メーター走行後", payload.meterAfter);
  setCol("走行距離",       payload.distance);
  setCol("電車通勤",       payload.trainCommute);
  setCol("会社クレカ払い", payload.companyCardPayment ? "1" : "");

  // 立替費用を転記（1報告につき1行・項目ごとに金額とチェックを横並び）
  const expenseItems = ["ガソリン代", "燃料代", "パーキング代"];
  (payload.expenses || []).forEach(function(exp) {
    if (!exp.item || !expenseItems.includes(exp.item)) return;
    setCol(exp.item,              exp.amount   || "");
    setCol(exp.item + "_チェック", exp.noCharge ? "1" : "");
  });

  setCol("備考",           payload.memo);

  sheet.appendRow(row);

  // 時間系の列を文字列形式に設定（時刻シリアル値への自動変換を防ぐ）
  const lastRow2    = sheet.getLastRow();
  const colMap2     = getColumnMap(sheet);
  const depTimeCol  = colMap2["出発時間"]  + 1;
  const arrTimeCol  = colMap2["到着時間"]  + 1;
  const workHrsCol  = colMap2["稼働時間"]  + 1;

  [[depTimeCol,  payload.departureTime || ""],
   [arrTimeCol,  payload.arrivalTime   || ""],
   [workHrsCol,  payload.workingHours  || ""],
  ].forEach(function([col, val]) {
    sheet.getRange(lastRow2, col).setNumberFormat("@");
    sheet.getRange(lastRow2, col).setValue(val);
  });

  // LINEに通知
  notifyLineGroup(payload);

  return jsonResponse({ success: true });
}

// =====================
// LINEグループ通知
// =====================
function notifyLineGroup(payload) {
  const props   = PropertiesService.getScriptProperties();
  const token   = props.getProperty("LINE_CHANNEL_ACCESS_TOKEN");
  const groupId = props.getProperty("LINE_GROUP_ID");

  if (!token || !groupId) return;

  const text =
    "🚗 運行報告が届きました\n" +
    "─────────────────\n" +
    "🏢 会社：" + (payload.companyLabel  || payload.company || "―") + "\n" +
    "📅 日付：" + (payload.date          || "―") + "\n" +
    "👤 運行者：" + (payload.staffName   || "―") + "（" + (payload.staffId || "―") + "）\n" +
    "🚐 車両：" + (payload.vehicle       || "―") + "\n" +
    "─────────────────\n" +
    "📍 出発：" + (payload.departure     || "―") + "\n" +
    "📍 行き先：" + (payload.destination || "―") + "\n" +
    "📍 到着：" + (payload.arrival       || "―") + "\n" +
    "─────────────────\n" +
    "🕐 出発時間：" + (payload.departureTime || "―") + "\n" +
    "🕐 到着時間：" + (payload.arrivalTime   || "―") +
    (payload.arrivalDate && payload.arrivalDate !== payload.date
      ? "（" + payload.arrivalDate + "）"
      : "") + "\n" +
    "⏱ 稼働時間：" + (payload.workingHours  || "―") + "\n" +
    "─────────────────\n" +
    "👥 利用者：" + (payload.passenger   || "―") + "\n" +
    "🎯 利用目的：" + (payload.purpose   || "―") + "\n" +
    "─────────────────\n" +
    "📊 メーター前：" + (payload.meterBefore || "―") + "\n" +
    "📊 メーター後：" + (payload.meterAfter  || "―") + "\n" +
    "📏 走行距離：" + (payload.distance   || "―") + "\n" +
    "🚃 電車通勤：" + (payload.trainCommute !== "" ? payload.trainCommute + "円" : "―") + "\n" +
    (payload.companyCardPayment ? "💳 会社クレカ払い：あり\n" : "") +
    "📝 備考：" + (payload.memo          || "―");

  const options = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + token,
    },
    payload: JSON.stringify({
      to: groupId,
      messages: [{ type: "text", text: text }],
    }),
  };

  try {
    UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", options);
  } catch(err) {
    Logger.log("LINE通知エラー: " + err.toString());
  }
}

// =====================
// colMapを使ってrecordオブジェクトを生成するヘルパー
// ※ ヘッダー名はスプレッドシートの1行目と完全一致させること
// =====================
function buildRecord(dispRow, cm) {
  function g(name) { return cm[name] !== undefined ? (dispRow[cm[name]] || "") : ""; }
  return {
    timestamp:     g("タイムスタンプ"),
    company:       g("会社"),
    date:          g("日付"),
    staffId:       g("スタッフID"),
    staffName:     g("運行者"),
    vehicle:       g("車両"),
    departure:     g("出発"),
    destination:   g("行き先"),
    arrival:       g("到着"),
    departureTime: g("出発時間"),
    arrivalTime:   g("到着時間"),
    arrivalDate:   g("到着日"),
    workingHours:  g("稼働時間"),
    passenger:     g("利用者"),
    purpose:       g("利用目的"),
    meterBefore:   g("メーター走行前"),
    meterAfter:    g("メーター走行後"),
    distance:      g("走行距離"),
    trainCommute:  g("電車通勤"),
    companyCardPayment: g("会社クレカ払い"),
    gasoline:      g("ガソリン代"),
    gasolineNC:    g("ガソリン代_チェック"),
    fuel:          g("燃料代"),
    fuelNC:        g("燃料代_チェック"),
    parking:       g("パーキング代"),
    parkingNC:     g("パーキング代_チェック"),
    memo:          g("備考"),
  };
}

// =====================
// ヘッダー行から列名→インデックスのマップを動的生成
// =====================
function getColumnMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach(function(h, i) {
    if (h) map[String(h).trim()] = i;
  });
  return map;
}

// =====================
// JSONレスポンス生成
// =====================
function jsonResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// =====================
// お客様向け：パスワード照合＋運行記録取得を一体化（customer.html 用）
// パスワードなしでは絶対にデータを返さない
// =====================
function getCustomerRecords(params) {
  const companyId = String(params.companyId || "").trim();
  const password  = String(params.password  || "").trim();
  const year      = parseInt(params.year);
  const month     = parseInt(params.month);

  if (!companyId || !password || isNaN(year) || isNaN(month)) {
    return jsonResponse({ error: "invalid params", records: [] });
  }

  const ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
  const companyData = ss.getSheetByName("会社マスタ").getDataRange().getValues();

  // ① パスワード照合（G列 = index 6）
  let companyName = null;
  for (let i = 1; i < companyData.length; i++) {
    if (String(companyData[i][0]).trim() !== companyId) continue;
    const stored = String(companyData[i][6] || "").trim();
    if (!stored || stored !== password) {
      return jsonResponse({ error: "unauthorized", records: [] });
    }
    companyName = String(companyData[i][1]).trim();
    break;
  }
  if (!companyName) {
    return jsonResponse({ error: "unauthorized", records: [] });
  }

  // ② 運行記録を取得
  const sheet       = ss.getSheetByName("運行記録");
  const cm          = getColumnMap(sheet);
  const values      = sheet.getDataRange().getValues();
  const displayVals = sheet.getDataRange().getDisplayValues();
  const records     = [];

  for (let i = 1; i < values.length; i++) {
    const row     = values[i];
    const dispRow = displayVals[i];

    const dateVal = row[cm["日付"]];
    if (!dateVal) continue;
    const date = new Date(dateVal);
    if (date.getFullYear() !== year || date.getMonth() + 1 !== month) continue;

    const rawCompany   = String(row[cm["会社"]] || "").trim();
    const rowCompanyId = rawCompany.includes("_") ? rawCompany.split("_")[0] : rawCompany;
    if (rowCompanyId !== companyId) continue;

    records.push(buildRecord(dispRow, cm));
  }

  records.sort((a, b) => {
    if (a.date !== b.date) return a.date < b.date ? -1 : 1;
    return (a.departureTime || "") < (b.departureTime || "") ? -1 : 1;
  });

  return jsonResponse({ companyName, records });
}

// =====================
// お客様パスワード照合のみ（customer.html ログイン用）
// =====================
function verifyCustomer(params) {
  const companyId = String(params.companyId || "").trim();
  const password  = String(params.password  || "").trim();

  if (!companyId || !password) {
    return jsonResponse({ verified: false, reason: "invalid params" });
  }

  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const data = ss.getSheetByName("会社マスタ").getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== companyId) continue;

    // G列（index 6）がパスワード
    const stored = String(data[i][6] || "").trim();
    if (!stored)          return jsonResponse({ verified: false, reason: "no password set" });
    if (stored === password) return jsonResponse({ verified: true, companyName: String(data[i][1]).trim() });
    return jsonResponse({ verified: false, reason: "wrong password" });
  }

  return jsonResponse({ verified: false, reason: "company not found" });
}

// =====================
// フォーム選択肢マスタ取得（種別は日本語：車両・出発・到着・利用目的）
// =====================
function getFormOptions() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("選択肢マスタ");
  if (!sheet) {
    return jsonResponse({
      vehicle:   ["アルファード", "レクサスLM", "レクサスLS"],
      departure: ["社長自宅", "会社", "ドライバー自宅", "駐車場"],
      arrival:   ["社長自宅", "会社", "ドライバー自宅", "駐車場"],
      purpose:   ["送迎", "会食"],
    });
  }
  const categoryMap = { "車両": "vehicle", "出発": "departure", "到着": "arrival", "利用目的": "purpose" };
  const rows    = sheet.getDataRange().getValues();
  const options = { vehicle: [], departure: [], arrival: [], purpose: [] };
  for (let i = 1; i < rows.length; i++) {
    const categoryJa = String(rows[i][0]).trim();
    const value      = String(rows[i][1]).trim();
    if (!value || value === "その他") continue;
    const key = categoryMap[categoryJa];
    if (key) options[key].push(value);
  }
  return jsonResponse(options);
}

// =====================
// 立替費用ファイルをDriveに保存（submitReport内から呼ぶ・画像なし版）
// =====================
function saveExpenseFiles_(payload) {
  Logger.log("立替費用メタデータ記録：" + (payload.expenses || []).length + " 件");
}

// =====================
// 立替費用画像のアップロード（フロントから別リクエストで呼ばれる）
// =====================
function uploadExpenseFiles(payload) {
  const props  = PropertiesService.getScriptProperties();
  const rootId = props.getProperty("EXPENSE_FOLDER_ID");
  if (!rootId) {
    Logger.log("EXPENSE_FOLDER_ID が未設定のため立替費用ファイルをスキップ");
    return jsonResponse({ success: true, skipped: true });
  }
  const rootFolder   = DriveApp.getFolderById(rootId);
  const companyLabel = payload.companyName || payload.company || "不明";
  const dateParts    = (payload.date || "").split("-");
  const targetLabel  = dateParts.length >= 2
    ? dateParts[0] + "年" + String(dateParts[1]).padStart(2, "0") + "月" : "不明";
  const dateStr8 = (payload.date || "").replace(/-/g, "");
  let savedCount = 0;
  (payload.expenses || []).forEach(function(exp) {
    if (!exp.images || exp.images.length === 0) return;
    const companyFolder = getOrCreateSubFolder_(rootFolder,    companyLabel);
    const monthFolder   = getOrCreateSubFolder_(companyFolder, targetLabel);
    const itemFolder    = getOrCreateSubFolder_(monthFolder,   exp.item);
    exp.images.forEach(function(img, idx) {
      const suffix   = exp.images.length > 1 ? "_" + (idx + 1) : "";
      const fileName = dateStr8 + "_" + exp.item + "_" + (exp.amount || "0") + suffix + ".jpg";
      const blob     = Utilities.newBlob(Utilities.base64Decode(img.data), "image/jpeg", fileName);
      itemFolder.createFile(blob);
      Logger.log("立替費用ファイル保存：" + fileName);
      savedCount++;
    });
  });
  return jsonResponse({ success: true, savedCount });
}

function getOrCreateSubFolder_(parent, name) {
  const existing = parent.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();
  return parent.createFolder(name);
}
