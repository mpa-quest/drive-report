// =====================
// 月初処理メイン
// =====================
function runMonthlyProcess() {
  // 年・月選択ダイアログを表示
  const ui = SpreadsheetApp.getUi();

  const yearResult = ui.prompt(
    "請求月の選択（1/2）",
    "請求年を入力してください（例：2026）",
    ui.ButtonSet.OK_CANCEL
  );
  if (yearResult.getSelectedButton() !== ui.Button.OK) {
    ui.alert("処理をキャンセルしました。");
    return;
  }
  const inputYear = parseInt(yearResult.getResponseText().trim());
  if (isNaN(inputYear) || inputYear < 2000 || inputYear > 2100) {
    ui.alert("年の入力が正しくありません。処理を中止します。");
    return;
  }

  const monthResult = ui.prompt(
    "請求月の選択（2/2）",
    "請求月を入力してください（例：6）",
    ui.ButtonSet.OK_CANCEL
  );
  if (monthResult.getSelectedButton() !== ui.Button.OK) {
    ui.alert("処理をキャンセルしました。");
    return;
  }
  const inputMonth = parseInt(monthResult.getResponseText().trim());
  if (isNaN(inputMonth) || inputMonth < 1 || inputMonth > 12) {
    ui.alert("月の入力が正しくありません（1〜12で入力してください）。処理を中止します。");
    return;
  }

  // 確認ダイアログ
  const confirmLabel = inputYear + "年" + String(inputMonth).padStart(2, "0") + "月";
  const confirm = ui.alert(
    "実行確認",
    confirmLabel + "分の月初処理を実行します。よろしいですか？",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) {
    ui.alert("処理をキャンセルしました。");
    return;
  }

  Logger.log("=== 月初処理開始：" + confirmLabel + " ===");

  try {

    const props = PropertiesService.getScriptProperties();
    const SPREADSHEET_ID    = props.getProperty("SPREADSHEET_ID");
    const TEMPLATE_ID       = props.getProperty("TEMPLATE_SPREADSHEET_ID");
    const STAFF_FOLDER_ID   = props.getProperty("STAFF_FOLDER_ID");
    const COMPANY_FOLDER_ID = props.getProperty("COMPANY_FOLDER_ID");
    const KANRI_FOLDER_ID   = props.getProperty("KANRI_FOLDER_ID");

    const db = SpreadsheetApp.openById(SPREADSHEET_ID);

    const targetYear  = inputYear;
    const targetMonth = inputMonth;
    const targetLabel = targetYear + "年" + String(targetMonth).padStart(2, "0") + "月";

    // 振込期日：翌月20日
    const paymentYear  = targetMonth === 12 ? targetYear + 1 : targetYear;
    const paymentMonth = targetMonth === 12 ? 1 : targetMonth + 1;
    const paymentDate  = paymentYear + "年" + String(paymentMonth).padStart(2, "0") + "月20日";

    Logger.log("振込期日：" + paymentDate);

    // 月フォルダを取得または作成（既存ファイルがあればゴミ箱へ）
    Logger.log("--- フォルダ準備開始 ---");
    const staffMonthFolder   = getOrCreateFolder(STAFF_FOLDER_ID,   targetLabel);
    const companyMonthFolder = getOrCreateFolder(COMPANY_FOLDER_ID, targetLabel);
    const kanriMonthFolder   = getOrCreateFolder(KANRI_FOLDER_ID,   targetLabel);
    clearFolderFiles_(staffMonthFolder,   "スタッフ請求書");
    clearFolderFiles_(companyMonthFolder, "お客様請求書");
    clearFolderFiles_(kanriMonthFolder,   "運行管理表");
    Logger.log("フォルダ準備完了");

    // 前月の運行記録を取得
    Logger.log("--- 運行記録の取得開始 ---");
    const records = getMonthlyRecords(db, targetYear, targetMonth);
    Logger.log("運行記録取得完了：" + records.length + " 件");

    if (records.length === 0) {
      Logger.log("⚠️ 該当月の運行記録が0件のため処理を中止します");
      ui.alert("該当月の運行記録が0件です。処理を中止します。");
      return;
    }

    // スタッフマスタ・会社マスタを取得
    Logger.log("--- マスタデータの取得開始 ---");
    const staffMap   = getStaffMap(db);
    const companyMap = getCompanyMap(db);
    Logger.log("スタッフマスタ：" + Object.keys(staffMap).length + " 件");
    Logger.log("会社マスタ：" + Object.keys(companyMap).length + " 件");

    // ① ドライバー請求書を生成
    Logger.log("--- ① ドライバー請求書の生成開始 ---");
    SpreadsheetApp.getActiveSpreadsheet().toast("① ドライバー請求書を生成中...", "🚗 月初処理", 5);
    generateStaffInvoices(records, staffMap, targetLabel, paymentDate, TEMPLATE_ID, staffMonthFolder);
    Logger.log("① ドライバー請求書の生成完了");

    // ② 運行管理表・お客様請求書を生成
    Logger.log("--- ② 運行管理表・お客様請求書の生成開始 ---");
    SpreadsheetApp.getActiveSpreadsheet().toast("② 運行管理表・お客様請求書を生成中...", "🚗 月初処理", 5);
    generateCompanyDocs(records, companyMap, targetLabel, paymentDate, TEMPLATE_ID, kanriMonthFolder, companyMonthFolder);
    Logger.log("② 運行管理表・お客様請求書の生成完了");

    // ③ Gmailの下書きを作成（お客様請求書 ＆ 運行管理表 を両方添付）
    Logger.log("--- ③ Gmail下書きの作成開始 ---");
    SpreadsheetApp.getActiveSpreadsheet().toast("③ Gmail下書きを作成中...", "🚗 月初処理", 5);
    createGmailDrafts(companyMap, companyMonthFolder, kanriMonthFolder, targetLabel);
    Logger.log("③ Gmail下書きの作成完了");

    Logger.log("=== 月初処理完了：" + targetLabel + " ===");
    ui.alert("✅ 月初処理が完了しました。\n\n対象月：" + targetLabel + "\n\n① ドライバー請求書\n② 運行管理表・お客様請求書\n③ Gmail下書き\n\nすべての処理が正常に完了しました。");

  } catch(e) {
    Logger.log("❌ 月初処理エラー：" + e.toString());
    notifyErrorLine_("月初処理（一括）", e);
    ui.alert("❌ エラーが発生しました。\n\n" + e.message + "\n\nLINEグループにエラー通知を送信しました。");
  }
}

// =====================
// 前月の運行記録を取得
// =====================
function getMonthlyRecords(db, year, month) {
  const sheet = db.getSheetByName("運行記録");
  const data  = sheet.getDataRange().getValues();
  const cm    = getColumnMapFromSheet(sheet);
  const records = [];

  for (let i = 1; i < data.length; i++) {
    const row     = data[i];
    const dateStr = row[cm["日付"]];
    if (!dateStr) continue;
    const date = new Date(dateStr);
    if (date.getFullYear() === year && date.getMonth() + 1 === month) {
      // 会社IDを正規化（"C001_ABC株式会社" 形式にも対応）
      const rawCompany = String(row[cm["会社"]] || "").trim();
      const normalizedCompany = rawCompany.includes("_") ? rawCompany.split("_")[0] : rawCompany;

      records.push({
        timestamp:     row[cm["タイムスタンプ"]],
        company:       normalizedCompany,
        date:          dateStr,
        staffId:       row[cm["スタッフID"]],
        staffName:     row[cm["運行者"]],
        vehicle:       row[cm["車両"]],
        departure:     row[cm["出発"]],
        destination:   row[cm["行き先"]],
        arrival:       row[cm["到着"]],
        departureTime: row[cm["出発時間"]],
        arrivalTime:   row[cm["到着時間"]],
        arrivalDate:   row[cm["到着日"]]   || "",
        workingHours:  row[cm["稼働時間"]],
        passenger:     row[cm["利用者"]],
        purpose:       row[cm["利用目的"]],
        meterBefore:   row[cm["メーター走行前"]],
        meterAfter:    row[cm["メーター走行後"]],
        distance:      row[cm["走行距離"]],
        trainCommute:       row[cm["電車通勤"]],
        memo:               row[cm["備考"]],
        companyCardPayment: ["1", 1, true].includes(row[cm["会社クレカ払い"]]),
        gasoline:           row[cm["ガソリン代"]],
        gasolineNC:         ["1", 1, true].includes(row[cm["ガソリン代_チェック"]]),
        fuel:               row[cm["燃料代"]],
        fuelNC:             ["1", 1, true].includes(row[cm["燃料代_チェック"]]),
        parking:            row[cm["パーキング代"]],
        parkingNC:          ["1", 1, true].includes(row[cm["パーキング代_チェック"]]),
      });
    }
  }
  return records;
}

// =====================
// スタッフマスタをマップで取得
// =====================
function getStaffMap(db) {
  const sheet = db.getSheetByName("スタッフマスタ");
  const data  = sheet.getDataRange().getValues();
  const map   = {};
  for (let i = 1; i < data.length; i++) {
    const staffId = data[i][1];
    map[staffId] = {
      lineUserId:  data[i][0],
      staffId:     data[i][1],
      staffName:   data[i][2],
      companyIds:  data[i][3] ? String(data[i][3]).split(",").map(s => s.trim()) : [],
      email:       data[i][4],
      regNumber:   data[i][5],
      zip:         data[i][6],
      address:     data[i][7],
      phone:       data[i][8],
      bankName:    data[i][9],
      branchName:  data[i][10],
      accountNum:  data[i][11],
    };
  }
  return map;
}

// =====================
// 会社マスタをマップで取得
// =====================
function getCompanyMap(db) {
  const sheet = db.getSheetByName("会社マスタ");
  const data  = sheet.getDataRange().getValues();
  const map   = {};
  for (let i = 1; i < data.length; i++) {
    const companyId = data[i][0];
    map[companyId] = {
      companyId:   data[i][0],
      companyName: data[i][1],
      email:       data[i][2],
      contactName: data[i][3],
    };
  }
  return map;
}

// =====================
// ① ドライバー請求書を生成
// =====================
function generateStaffInvoices(records, staffMap, targetLabel, paymentDate, templateId, folder) {
  // 発行日（マクロ実行日）・請求番号用月キー（YYYYMM形式）
  const now            = new Date();
  const issueDate      = Utilities.formatDate(now, "Asia/Tokyo", "yyyy年MM月dd日");
  const invoiceMonthKey = targetLabel.replace("年", "").replace("月", ""); // 例: "202506"

  const grouped = {};
  records.forEach(function(r) {
    if (!grouped[r.staffId]) grouped[r.staffId] = [];
    grouped[r.staffId].push(r);
  });

  const staffIds = Object.keys(grouped);
  Logger.log("対象スタッフ数：" + staffIds.length + " 名");

  staffIds.forEach(function(staffId) {
    const staff = staffMap[staffId];
    if (!staff) {
      Logger.log("⚠️ スタッフマスタに未登録のスタッフIDをスキップ：" + staffId);
      return;
    }

    const rows = grouped[staffId];
    const fileName = "【請求書】" + staffId + "_" + staff.staffName + "_" + targetLabel;
    Logger.log("  生成開始：" + fileName + "（" + rows.length + " 件）");

    // テンプレートをコピー
    const copy = DriveApp.getFileById(templateId).makeCopy(fileName, folder);
    const ss   = SpreadsheetApp.openById(copy.getId());
    const sheet = ss.getSheetByName("スタッフ請求書");

    // 変数を置換
    replaceInSheet(sheet, "{{スタッフ名}}",  staff.staffName  || "");
    replaceInSheet(sheet, "{{登録番号}}",    staff.regNumber  || "");
    replaceInSheet(sheet, "{{郵便番号}}",    staff.zip        || "");
    replaceInSheet(sheet, "{{住所}}",        staff.address    || "");
    replaceInSheet(sheet, "{{電話番号}}",    staff.phone      || "");
    replaceInSheet(sheet, "{{銀行名}}",      staff.bankName   || "");
    replaceInSheet(sheet, "{{支店名}}",      staff.branchName || "");
    replaceInSheet(sheet, "{{口座番号}}",    staff.accountNum || "");
    replaceInSheet(sheet, "{{発行日}}",      issueDate);
    replaceInSheet(sheet, "{{請求番号}}",    staffId + "_" + invoiceMonthKey + "_00");

    // 振込期日（C16固定）
    sheet.getRange("C16").setValue(paymentDate);

    // 明細行を書き込み
    writeDetailRows(sheet, rows, "staff");

    // スタッフ請求書シートのみPDF出力
    savePdfSingleSheet(ss, copy, fileName, folder, "スタッフ請求書");
    Logger.log("  生成完了：" + fileName);
  });
}

// =====================
// ② 運行管理表・お客様請求書を生成
// =====================
function generateCompanyDocs(records, companyMap, targetLabel, paymentDate, templateId, kanriFolder, companyFolder) {
  // 発行日（マクロ実行日）・請求番号用月キー（YYYYMM形式）
  const now             = new Date();
  const issueDate       = Utilities.formatDate(now, "Asia/Tokyo", "yyyy年MM月dd日");
  const invoiceMonthKey = targetLabel.replace("年", "").replace("月", ""); // 例: "202506"

  const grouped = {};
  records.forEach(function(r) {
    const cid = r.company;
    if (!grouped[cid]) grouped[cid] = [];
    grouped[cid].push(r);
  });

  const companyIds = Object.keys(grouped);
  Logger.log("対象会社数：" + companyIds.length + " 社");

  companyIds.forEach(function(companyId) {
    const company = companyMap[companyId];
    if (!company) {
      Logger.log("⚠️ 会社マスタに未登録の会社IDをスキップ：" + companyId);
      return;
    }
    // 日付・出発時間の昇順でソート
    const rows = grouped[companyId].slice().sort(function(a, b) {
      const dateA = a.date ? new Date(a.date).getTime() : 0;
      const dateB = b.date ? new Date(b.date).getTime() : 0;
      if (dateA !== dateB) return dateA - dateB;
      const tA = String(a.departureTime || "").trim();
      const tB = String(b.departureTime || "").trim();
      return tA < tB ? -1 : tA > tB ? 1 : 0;
    });
    Logger.log("  処理開始：" + companyId + "_" + company.companyName + "（" + rows.length + " 件）");

    // 運行管理表
    const kanriName = "【運行管理表】" + companyId + "_" + company.companyName + "_" + targetLabel;
    Logger.log("    運行管理表 生成中：" + kanriName);
    const kanriCopy  = DriveApp.getFileById(templateId).makeCopy(kanriName, kanriFolder);
    const kanriSS    = SpreadsheetApp.openById(kanriCopy.getId());
    const kanriSheet = kanriSS.getSheetByName("運行管理表");

    // A1セルの {{会社名}} を置換（「{{会社名}}様　運行管理表」という形式）
    const a1Val = kanriSheet.getRange("A1").getValue();
    kanriSheet.getRange("A1").setValue(String(a1Val).replace("{{会社名}}", company.companyName));

    // 月を書き込み（D1セル）
    const monthNum = parseInt(targetLabel.replace(/\d+年(\d+)月/, "$1"));
    kanriSheet.getRange("D1").setValue(monthNum + "月");

    writeDetailRows(kanriSheet, rows, "kanri");
    savePdfSingleSheet(kanriSS, kanriCopy, kanriName, kanriFolder, "運行管理表");
    Logger.log("    運行管理表 生成完了：" + kanriName);

    // お客様請求書
    const invoiceName = "【請求書】" + companyId + "_" + company.companyName + "_" + targetLabel;
    Logger.log("    お客様請求書 生成中：" + invoiceName);
    const invoiceCopy  = DriveApp.getFileById(templateId).makeCopy(invoiceName, companyFolder);
    const invoiceSS    = SpreadsheetApp.openById(invoiceCopy.getId());
    const invoiceSheet = invoiceSS.getSheetByName("お客様請求書");

    replaceInSheet(invoiceSheet, "{{会社名}}",   company.companyName);
    replaceInSheet(invoiceSheet, "{{発行日}}",   issueDate);
    replaceInSheet(invoiceSheet, "{{請求番号}}", companyId + "_" + invoiceMonthKey + "_00");

    // 振込期日（C16固定）
    invoiceSheet.getRange("C16").setValue(paymentDate);

    writeDetailRows(invoiceSheet, rows, "company");
    savePdfSingleSheet(invoiceSS, invoiceCopy, invoiceName, companyFolder, "お客様請求書");
    Logger.log("    お客様請求書 生成完了：" + invoiceName);
  });
}

// =====================
// ③ Gmailの下書きを作成（お客様請求書＆運行管理表を2点添付）
// =====================
function createGmailDrafts(companyMap, companyFolder, kanriFolder, targetLabel) {
  // 差出人メールアドレス（スクリプトプロパティ FROM_EMAIL）
  const props     = PropertiesService.getScriptProperties();
  const fromEmail = props.getProperty("FROM_EMAIL") || "";
  // ① 会社フォルダ（請求書）内のファイルをマッピング
  const companyFiles = companyFolder.getFiles();
  const companyFileMap = {};
  while (companyFiles.hasNext()) {
    const file = companyFiles.next();
    companyFileMap[file.getName()] = file;
  }

  // ② 運行管理表フォルダ内のファイルをマッピング
  const kanriFiles = kanriFolder.getFiles();
  const kanriFileMap = {};
  while (kanriFiles.hasNext()) {
    const file = kanriFiles.next();
    kanriFileMap[file.getName()] = file;
  }

  let draftCount = 0;
  let skipCount  = 0;

  Object.keys(companyMap).forEach(function(companyId) {
    const company = companyMap[companyId];
    if (!company.email) {
      Logger.log("  メールアドレス未設定のためスキップ：" + companyId + "_" + company.companyName);
      skipCount++;
      return;
    }

    // それぞれのファイル名を作成
    const invoiceName = "【請求書】" + companyId + "_" + company.companyName + "_" + targetLabel + ".pdf";
    const kanriName   = "【運行管理表】" + companyId + "_" + company.companyName + "_" + targetLabel + ".pdf";

    const invoiceFile = companyFileMap[invoiceName];
    const kanriFile   = kanriFileMap[kanriName];

    // 請求書ファイルがなければスキップ
    if (!invoiceFile) {
      Logger.log("  ⚠️ 請求書PDFが見つからないためスキップ：" + invoiceName);
      skipCount++;
      return;
    }

    Logger.log("  下書き作成中：" + company.email + "（" + company.companyName + "）");
    if (!kanriFile) {
      Logger.log("  ⚠️ 運行管理表PDFが見つからないため請求書のみ添付：" + kanriName);
    }

    const subject = "【ご請求】" + targetLabel + "分　株式会社エグゼクティブサポート";
    
    const body =
      (company.contactName || company.companyName) + " 様\n\n" +
      "いつもお世話になっております。\n" +
      "株式会社エグゼクティブサポートでございます。\n\n" +
      targetLabel + "分のご請求書および運行管理表をお送りいたします。\n" +
      "ご確認のほど、よろしくお願いいたします。\n\n" +
      "──────────────────\n" +
      "株式会社エグゼクティブサポート\n" +
      "──────────────────";

    // 添付ファイルの配列を準備
    const attachments = [invoiceFile.getAs(MimeType.PDF)];
    
    // 運行管理表ファイルが見つかった場合のみ、添付配列に追加
    if (kanriFile) {
      attachments.push(kanriFile.getAs(MimeType.PDF));
    }

    GmailApp.createDraft(
      company.email,
      subject,
      body,
      {
        attachments: attachments,
        ...(fromEmail ? { from: fromEmail } : {}),
      }
    );
    Logger.log("  下書き作成完了：" + company.companyName + "（添付 " + attachments.length + " 件）");
    draftCount++;
  });

  Logger.log("Gmail下書き作成：成功 " + draftCount + " 件 ／ スキップ " + skipCount + " 件");
}

// =====================
// シート内の文字列を置換
// =====================
function replaceInSheet(sheet, placeholder, value) {
  const range  = sheet.getDataRange();
  const values = range.getValues();
  let changed  = false;
  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      if (typeof values[i][j] === "string" && values[i][j].includes(placeholder)) {
        values[i][j] = values[i][j].replace(placeholder, value);
        changed = true;
      }
    }
  }
  if (changed) range.setValues(values);
}

// =====================
// 明細行を書き込み（不具合改修済）
// =====================
function writeDetailRows(sheet, rows, type) {
  // 「日付」ヘッダーの次の行を明細開始行として探す（全列スキャン・完全一致）
  const data = sheet.getDataRange().getValues();
  let startRow = -1;
  for (let i = 0; i < data.length; i++) {
    for (let j = 0; j < data[i].length; j++) {
      if (String(data[i][j]).trim() === "日付") {
        startRow = i + 2;
        break;
      }
    }
    if (startRow !== -1) break;
  }
  if (startRow === -1) {
    Logger.log("⚠️ writeDetailRows：「日付」ヘッダーが見つかりません（type=" + type + "）");
    return;
  }
  Logger.log("    明細書き込み開始：type=" + type + "、開始行=" + startRow + "、件数=" + rows.length);

  // 運行記録シートの「見た目の文字列」をそのまま取得するためにDisplayValuesを取得する
  const props = PropertiesService.getScriptProperties();
  const db = SpreadsheetApp.openById(props.getProperty("SPREADSHEET_ID"));
  const recordSheet = db.getSheetByName("運行記録");
  const displayValues = recordSheet.getDataRange().getDisplayValues();

  rows.forEach(function(r, idx) {
    const row = startRow + idx;
    
    if (type === "kanri") {
      let dispDepartureTime = r.departureTime;
      let dispArrivalTime   = r.arrivalTime;
      
      // r.dateを「YYYY/MM/DD」形式の文字列に変換して型ミスマッチを防ぐ
      let formattedTargetDate = "";
      if (r.date instanceof Date) {
        formattedTargetDate = Utilities.formatDate(r.date, Session.getScriptTimeZone(), "yyyy/MM/dd");
      } else if (r.date) {
        const d = new Date(r.date);
        if (!isNaN(d.getTime())) {
          formattedTargetDate = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy/MM/dd");
        } else {
          formattedTargetDate = String(r.date);
        }
      }

      const targetTimestamp = r.timestamp ? String(r.timestamp).trim() : "";

      // 元のシートの行を特定し「見かけ上の時間・稼働時間」を取得する
      const dcm = getColumnMapFromSheet(recordSheet); // 動的列マップ
      let matched = false;
      let workingSerial = 0;
      for (let k = 1; k < displayValues.length; k++) {
        const rowTimestamp = displayValues[k][dcm["タイムスタンプ"]] ? String(displayValues[k][dcm["タイムスタンプ"]]).trim() : "";
        const rowDateStr   = displayValues[k][dcm["日付"]]           ? String(displayValues[k][dcm["日付"]]).trim()           : "";
        const rowStaffId   = displayValues[k][dcm["スタッフID"]]     ? String(displayValues[k][dcm["スタッフID"]]).trim()     : "";

        const normalizedRowDate    = rowDateStr.replace(/-/g, "/");
        const normalizedTargetDate = formattedTargetDate.replace(/-/g, "/");

        if ((targetTimestamp !== "" && rowTimestamp === targetTimestamp) ||
            (normalizedRowDate === normalizedTargetDate && rowStaffId === String(r.staffId).trim())) {
          dispDepartureTime = displayValues[k][dcm["出発時間"]];
          dispArrivalTime   = displayValues[k][dcm["到着時間"]];

          // 稼働時間をdisplayValuesから取得（24時間超えも正しく取れる）
          const whDisplay = String(displayValues[k][dcm["稼働時間"]] || "").trim();
          let workingMinutes = 0;
          // "X時間Y分" 形式
          const jpH = whDisplay.match(/(\d+)時間/);
          const jpM = whDisplay.match(/(\d+)分/);
          if (jpH || jpM) {
            workingMinutes = (jpH ? parseInt(jpH[1]) : 0) * 60 + (jpM ? parseInt(jpM[1]) : 0);
          } else {
            // "H:MM" / "HH:MM" / "H:MM:SS" 形式（24時間超えも対応）
            const hm = whDisplay.match(/^(\d+):(\d{2})/);
            if (hm) workingMinutes = parseInt(hm[1]) * 60 + parseInt(hm[2]);
          }
          workingSerial = workingMinutes / (24 * 60);

          matched = true;
          break;
        }
      }
      if (!matched) {
        Logger.log("    ⚠️ 元レコードが特定できませんでした：staffId=" + r.staffId + "、date=" + formattedTargetDate);
      }

      // 日付・到着日をJST文字列に変換
      const tz = Session.getScriptTimeZone();
      function toDateStr(val) {
        if (!val) return "";
        if (val instanceof Date) return Utilities.formatDate(val, tz, "yyyy/MM/dd");
        const d = new Date(val);
        return isNaN(d.getTime()) ? String(val) : Utilities.formatDate(d, tz, "yyyy/MM/dd");
      }
      const displayDate        = toDateStr(r.date);
      const displayArrivalDate = toDateStr(r.arrivalDate);

      // 立替費用：会社クレカ払いチェックありはスタッフ請求書に反映しない→運行管理表では実費記載
      // お客様請求不要チェックありは会社請求書・運行管理表に反映しない→0表示
      const gasolineVal = r.gasolineNC ? 0 : (parseFloat(r.gasoline) || 0);
      const fuelVal     = r.fuelNC     ? 0 : (parseFloat(r.fuel)     || 0);
      const parkingVal  = r.parkingNC  ? 0 : (parseFloat(r.parking)  || 0);

      // データを書き込み
      // 列順: 日付,運行者,車両,出発,行き先,到着,出発時間,到着日,到着時間,稼働時間,利用者,利用目的,メーター前,メーター後,走行距離,電車通勤,ガソリン代,燃料代,パーキング代,備考
      sheet.getRange(row, 1, 1, 20).setValues([[
        displayDate, r.staffName, r.vehicle,
        r.departure, r.destination, r.arrival,
        dispDepartureTime, displayArrivalDate, dispArrivalTime, workingSerial,
        r.passenger, r.purpose,
        r.meterBefore, r.meterAfter, r.distance,
        r.trainCommute,
        gasolineVal || "", fuelVal || "", parkingVal || "",
        r.memo,
      ]]);

      // 稼働時間（J列 = 10列目）の表示形式を「[h]:mm」に設定
      sheet.getRange(row, 10).setNumberFormat("[h]:mm");
      
    } else {
      // お客様請求書・ドライバー請求書用
      // ⚠️ 金額計算ロジック未実装のため現時点では書き込みをスキップ
      // ヒアリング結果確定後に実装予定（単価・超過料金・消費税計算など）
      Logger.log("    [SKIP] 請求書明細書き込みスキップ（金額計算ロジック未実装）：" + type);
    }
  });

  // 運行管理表（kanri）の場合のみ、合計行（37行目固定）にSUM式を設定する
  // ※稼働時間合計はテンプレートに既存式があるためGASからは書き込まない
  if (type === "kanri" && rows.length > 0) {
    const lastDataRow = startRow + rows.length - 1;
    const targetTotalRow = 37;

    Logger.log("    合計行を設定：" + targetTotalRow + "行目（データ " + startRow + "〜" + lastDataRow + "行）");

    // 電車通勤合計（P列=16）・ガソリン代合計（Q列=17）・燃料代合計（R列=18）・パーキング代合計（S列=19）
    sheet.getRange(targetTotalRow, 16).setFormula("=SUM(P" + startRow + ":P" + lastDataRow + ")");
    sheet.getRange(targetTotalRow, 17).setFormula("=SUM(Q" + startRow + ":Q" + lastDataRow + ")");
    sheet.getRange(targetTotalRow, 18).setFormula("=SUM(R" + startRow + ":R" + lastDataRow + ")");
    sheet.getRange(targetTotalRow, 19).setFormula("=SUM(S" + startRow + ":S" + lastDataRow + ")");
  }

  Logger.log("    明細書き込み完了：type=" + type);
}

// =====================
// 指定シートのみPDFに変換して保存
// =====================
function savePdfSingleSheet(ss, copyFile, fileName, folder, sheetName) {
  Logger.log("      PDF出力開始：" + fileName + "（シート：" + sheetName + "）");
  SpreadsheetApp.flush();

  const sheet = ss.getSheetByName(sheetName);
  const sheetId = sheet.getSheetId();

  // シートごとに印刷設定を切り替え
  const isLandscape = sheetName === "運行管理表";

  const url =
    "https://docs.google.com/spreadsheets/d/" + ss.getId() +
    "/export?format=pdf" +
    "&gid=" + sheetId +
    "&size=A4" +
    "&portrait=" + (!isLandscape) +
    "&fitw=true" +
    "&gridlines=false" +
    "&sheetnames=false" +
    "&printtitle=false" +
    "&pagenumbers=false" +
    "&fzr=false" +
    "&top_margin=0.25" +
    "&bottom_margin=0.25" +
    "&left_margin=0.25" +
    "&right_margin=0.25";

  const token = ScriptApp.getOAuthToken();

  const blob = UrlFetchApp.fetch(url, {
    headers: {
      "Authorization": "Bearer " + token
    }
  }).getBlob().setName(fileName + ".pdf");

  folder.createFile(blob);
  Logger.log("      PDF出力完了・Driveへ保存：" + fileName + ".pdf");

  copyFile.setTrashed(true);
  Logger.log("      コピー元スプレッドシートを削除：" + fileName);
}

// =====================
// ヘッダー行から列名→インデックスのマップを動的生成
// =====================
function getColumnMapFromSheet(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach(function(h, i) {
    if (h) map[String(h).trim()] = i;
  });
  return map;
}

// =====================
// スプレッドシートを開いたときにカスタムメニューを追加
// =====================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🚗 運行管理")
    .addItem("月初処理を実行（一括）", "runMonthlyProcess")
    .addSeparator()
    .addItem("① ドライバー請求書を生成", "runStep1")
    .addItem("② 運行管理表・お客様請求書を生成", "runStep2")
    .addItem("③ Gmail下書きを作成", "runStep3")
    .addSeparator()
    .addItem("🗑️ 月フォルダ内ファイルを一括削除", "clearMonthFolder")
    .addToUi();
}


// =====================
// 年月入力ダイアログ（共通）
// =====================
function promptYearMonth_() {
  const ui = SpreadsheetApp.getUi();

  const yearResult = ui.prompt(
    "請求月の選択（1/2）",
    "請求年を入力してください（例：2026）",
    ui.ButtonSet.OK_CANCEL
  );
  if (yearResult.getSelectedButton() !== ui.Button.OK) return null;
  const inputYear = parseInt(yearResult.getResponseText().trim());
  if (isNaN(inputYear) || inputYear < 2000 || inputYear > 2100) {
    ui.alert("年の入力が正しくありません。処理を中止します。");
    return null;
  }

  const monthResult = ui.prompt(
    "請求月の選択（2/2）",
    "請求月を入力してください（例：6）",
    ui.ButtonSet.OK_CANCEL
  );
  if (monthResult.getSelectedButton() !== ui.Button.OK) return null;
  const inputMonth = parseInt(monthResult.getResponseText().trim());
  if (isNaN(inputMonth) || inputMonth < 1 || inputMonth > 12) {
    ui.alert("月の入力が正しくありません（1〜12で入力してください）。処理を中止します。");
    return null;
  }

  const confirmLabel = inputYear + "年" + String(inputMonth).padStart(2, "0") + "月";
  const confirm = ui.alert(
    "実行確認",
    confirmLabel + "分の処理を実行します。よろしいですか？",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) {
    ui.alert("処理をキャンセルしました。");
    return null;
  }

  return { year: inputYear, month: inputMonth };
}

// =====================
// ① ドライバー請求書のみ生成
// =====================
function runStep1() {
  const ym = promptYearMonth_();
  if (!ym) return;

  const props = PropertiesService.getScriptProperties();
  const db    = SpreadsheetApp.openById(props.getProperty("SPREADSHEET_ID"));

  const targetLabel = ym.year + "年" + String(ym.month).padStart(2, "0") + "月";
  const paymentYear  = ym.month === 12 ? ym.year + 1 : ym.year;
  const paymentMonth = ym.month === 12 ? 1 : ym.month + 1;
  const paymentDate  = paymentYear + "年" + String(paymentMonth).padStart(2, "0") + "月20日";

  Logger.log("=== ① ドライバー請求書生成開始：" + targetLabel + " ===");
  SpreadsheetApp.getActiveSpreadsheet().toast("① ドライバー請求書を生成中...", "🚗 月初処理", 5);

  try {
    const records  = getMonthlyRecords(db, ym.year, ym.month);
    const staffMap = getStaffMap(db);
    Logger.log("運行記録：" + records.length + " 件　スタッフ：" + Object.keys(staffMap).length + " 名");

    const staffFolder = getOrCreateFolder(props.getProperty("STAFF_FOLDER_ID"), targetLabel);
    clearFolderFiles_(staffFolder, "スタッフ請求書");
    generateStaffInvoices(records, staffMap, targetLabel, paymentDate, props.getProperty("TEMPLATE_SPREADSHEET_ID"), staffFolder);

    Logger.log("=== ① ドライバー請求書生成完了：" + targetLabel + " ===");
    SpreadsheetApp.getUi().alert("✅ 処理が完了しました。\n\n① ドライバー請求書の生成が完了しました。\n対象月：" + targetLabel);
  } catch(e) {
    Logger.log("❌ ① ドライバー請求書生成エラー：" + e.toString());
    notifyErrorLine_("① ドライバー請求書生成", e);
    SpreadsheetApp.getUi().alert("❌ エラーが発生しました。\n\n" + e.message + "\n\nLINEグループにエラー通知を送信しました。");
  }
}

// =====================
// ② 運行管理表・お客様請求書のみ生成
// =====================
function runStep2() {
  const ym = promptYearMonth_();
  if (!ym) return;

  const props = PropertiesService.getScriptProperties();
  const db    = SpreadsheetApp.openById(props.getProperty("SPREADSHEET_ID"));

  const targetLabel = ym.year + "年" + String(ym.month).padStart(2, "0") + "月";
  const paymentYear  = ym.month === 12 ? ym.year + 1 : ym.year;
  const paymentMonth = ym.month === 12 ? 1 : ym.month + 1;
  const paymentDate  = paymentYear + "年" + String(paymentMonth).padStart(2, "0") + "月20日";

  Logger.log("=== ② 運行管理表・お客様請求書生成開始：" + targetLabel + " ===");
  SpreadsheetApp.getActiveSpreadsheet().toast("② 運行管理表・お客様請求書を生成中...", "🚗 月初処理", 5);

  try {
    const records    = getMonthlyRecords(db, ym.year, ym.month);
    const companyMap = getCompanyMap(db);
    Logger.log("運行記録：" + records.length + " 件　会社：" + Object.keys(companyMap).length + " 社");

    const kanriFolder   = getOrCreateFolder(props.getProperty("KANRI_FOLDER_ID"),   targetLabel);
    const companyFolder = getOrCreateFolder(props.getProperty("COMPANY_FOLDER_ID"), targetLabel);
    clearFolderFiles_(kanriFolder,   "運行管理表");
    clearFolderFiles_(companyFolder, "お客様請求書");
    generateCompanyDocs(records, companyMap, targetLabel, paymentDate, props.getProperty("TEMPLATE_SPREADSHEET_ID"), kanriFolder, companyFolder);

    Logger.log("=== ② 運行管理表・お客様請求書生成完了：" + targetLabel + " ===");
    SpreadsheetApp.getUi().alert("✅ 処理が完了しました。\n\n② 運行管理表・お客様請求書の生成が完了しました。\n対象月：" + targetLabel);
  } catch(e) {
    Logger.log("❌ ② 運行管理表・お客様請求書生成エラー：" + e.toString());
    notifyErrorLine_("② 運行管理表・お客様請求書生成", e);
    SpreadsheetApp.getUi().alert("❌ エラーが発生しました。\n\n" + e.message + "\n\nLINEグループにエラー通知を送信しました。");
  }
}

// =====================
// ③ Gmail下書きのみ作成
// =====================
function runStep3() {
  const ym = promptYearMonth_();
  if (!ym) return;

  const props = PropertiesService.getScriptProperties();
  const db    = SpreadsheetApp.openById(props.getProperty("SPREADSHEET_ID"));

  const targetLabel   = ym.year + "年" + String(ym.month).padStart(2, "0") + "月";
  const companyMap    = getCompanyMap(db);
  const companyFolder = getOrCreateFolder(props.getProperty("COMPANY_FOLDER_ID"), targetLabel);
  const kanriFolder   = getOrCreateFolder(props.getProperty("KANRI_FOLDER_ID"),   targetLabel);

  Logger.log("=== ③ Gmail下書き作成開始：" + targetLabel + " ===");
  SpreadsheetApp.getActiveSpreadsheet().toast("③ Gmail下書きを作成中...", "🚗 月初処理", 5);

  try {
    createGmailDrafts(companyMap, companyFolder, kanriFolder, targetLabel);
    Logger.log("=== ③ Gmail下書き作成完了：" + targetLabel + " ===");
    SpreadsheetApp.getUi().alert("✅ 処理が完了しました。\n\n③ Gmail下書きの作成が完了しました。\n対象月：" + targetLabel);
  } catch(e) {
    Logger.log("❌ ③ Gmail下書き作成エラー：" + e.toString());
    notifyErrorLine_("③ Gmail下書き作成", e);
    SpreadsheetApp.getUi().alert("❌ エラーが発生しました。\n\n" + e.message + "\n\nLINEグループにエラー通知を送信しました。");
  }
}
// =====================
// フォルダを取得または作成
// =====================
function getOrCreateFolder(parentId, name) {
  const parent   = DriveApp.getFolderById(parentId);
  const existing = parent.getFoldersByName(name);
  if (existing.hasNext()) {
    const folder = existing.next();
    Logger.log("  既存フォルダを使用：" + name);
    return folder;
  }
  const folder = parent.createFolder(name);
  Logger.log("  新規フォルダを作成：" + name);
  return folder;
}

// =====================
// エラー通知（LINEグループへ）
// =====================
function notifyErrorLine_(stepName, error) {
  const props   = PropertiesService.getScriptProperties();
  const token   = props.getProperty("LINE_CHANNEL_ACCESS_TOKEN");
  const groupId = props.getProperty("LINE_GROUP_ID");
  if (!token || !groupId) {
    Logger.log("LINE通知スキップ（トークンまたはグループID未設定）");
    return;
  }
  const text = "🚨 月初処理でエラーが発生しました\n─────────────────\n処理：" + stepName + "\n内容：" + error.message + "\n─────────────────\nGASの実行ログを確認してください。\n\nこの内容をコピーして、システム管理者にお問い合わせください。";
  try {
    UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + token,
      },
      payload: JSON.stringify({
        to: groupId,
        messages: [{ type: "text", text: text }],
      }),
    });
  } catch(e) {
    Logger.log("エラー通知のLINE送信に失敗: " + e.toString());
  }
}

// =====================
// フォルダ内の既存ファイルをゴミ箱へ移動（共通ヘルパー）
// =====================
function clearFolderFiles_(folder, label) {
  const files = folder.getFiles();
  let count = 0;
  while (files.hasNext()) {
    files.next().setTrashed(true);
    count++;
  }
  if (count > 0) {
    Logger.log("  既存ファイルをゴミ箱へ移動：" + label + "（" + count + " 件）");
  }
}

// =====================
// 月フォルダ内ファイルを手動で一括削除（メニューから実行）
// =====================
function clearMonthFolder() {
  const ui = SpreadsheetApp.getUi();

  const yearResult = ui.prompt("フォルダ初期化（1/2）", "対象年を入力してください（例：2026）", ui.ButtonSet.OK_CANCEL);
  if (yearResult.getSelectedButton() !== ui.Button.OK) return;
  const inputYear = parseInt(yearResult.getResponseText().trim());
  if (isNaN(inputYear) || inputYear < 2000 || inputYear > 2100) {
    ui.alert("年の入力が正しくありません。処理を中止します。"); return;
  }

  const monthResult = ui.prompt("フォルダ初期化（2/2）", "対象月を入力してください（例：6）", ui.ButtonSet.OK_CANCEL);
  if (monthResult.getSelectedButton() !== ui.Button.OK) return;
  const inputMonth = parseInt(monthResult.getResponseText().trim());
  if (isNaN(inputMonth) || inputMonth < 1 || inputMonth > 12) {
    ui.alert("月の入力が正しくありません（1〜12で入力してください）。処理を中止します。"); return;
  }

  const targetLabel = inputYear + "年" + String(inputMonth).padStart(2, "0") + "月";
  const confirm = ui.alert("削除確認",
    "【" + targetLabel + "】フォルダ内のファイルをすべてゴミ箱に移動します。\n\n対象：\n・スタッフ請求書フォルダ\n・お客様請求書フォルダ\n・運行管理表フォルダ\n\nよろしいですか？",
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) { ui.alert("処理をキャンセルしました。"); return; }

  const props = PropertiesService.getScriptProperties();
  const folderDefs = [
    { id: props.getProperty("STAFF_FOLDER_ID"),   label: "スタッフ請求書" },
    { id: props.getProperty("COMPANY_FOLDER_ID"), label: "お客様請求書" },
    { id: props.getProperty("KANRI_FOLDER_ID"),   label: "運行管理表" },
  ];

  let totalDeleted = 0;
  folderDefs.forEach(function(def) {
    const parent = DriveApp.getFolderById(def.id);
    const subFolders = parent.getFoldersByName(targetLabel);
    if (!subFolders.hasNext()) { Logger.log("  フォルダなし（スキップ）：" + def.label); return; }
    const subFolder = subFolders.next();
    const files = subFolder.getFiles();
    let count = 0;
    while (files.hasNext()) { files.next().setTrashed(true); count++; }
    totalDeleted += count;
    Logger.log("  削除完了：" + def.label + "（" + count + " 件）");
  });

  ui.alert("✅ 削除が完了しました。\n\n対象月：" + targetLabel + "\n合計 " + totalDeleted + " 件のファイルをゴミ箱に移動しました。");
}
