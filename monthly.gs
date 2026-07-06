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
    generateCompanyDocs(records, companyMap, targetLabel, targetYear, targetMonth, TEMPLATE_ID, kanriMonthFolder, companyMonthFolder);
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
// =====================
// 請求パターン名（マスタの「請求パターン」列に入れる値）を1箇所で管理
// 新しい例外パターンを実装する際は、ここに追記して各マスタ列に対応する値を入れる
//   ・会社・スタッフで同じ内容の割増ロジックは、同じパターン名を使い回す
// =====================
const BILLING_PATTERN = {
  CONSOLIDATED: "一本化",   // 会社のみ：運行代・交通費・立替費用・時間超過分を1行「貴社運行管理請負費用」にまとめる
  HOLIDAY_A:    "休日加算A", // 会社・スタッフ共通：水・土・日の記録日数×1万円（税抜）を別行加算
};
const COMPANY_BILLING_PATTERNS = [BILLING_PATTERN.CONSOLIDATED, BILLING_PATTERN.HOLIDAY_A];
const STAFF_BILLING_PATTERNS   = [BILLING_PATTERN.HOLIDAY_A];

function getStaffMap(db) {
  const sheet = db.getSheetByName("スタッフマスタ");
  const data  = sheet.getDataRange().getValues();
  const cm    = getColumnMapFromSheet(sheet); // ヘッダー名→列インデックス（0-indexed）
  const map   = {};

  // ヘッダー名の揺れに対応するためのフォールバック付き取得ヘルパー
  function col(row, names, fallbackIdx) {
    for (let n = 0; n < names.length; n++) {
      if (cm[names[n]] !== undefined) return row[cm[names[n]]];
    }
    return fallbackIdx !== undefined ? row[fallbackIdx] : undefined;
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const staffId = col(row, ["スタッフID"], 1);
    if (!staffId) continue;

    // 会社マスタと同様、「例外」列に値があるスタッフは、給与ロジック未確定のためとりあえずスキップ
    // ただし「請求パターン」列に実装済みパターン（STAFF_BILLING_PATTERNS）が入っている場合はスキップせず特殊ロジックを適用
    const exceptionNote  = col(row, ["例外"]);
    const billingPattern = String(col(row, ["請求パターン"]) || "").trim();
    const isImplementedStaffPattern = STAFF_BILLING_PATTERNS.indexOf(billingPattern) !== -1;
    if (exceptionNote && !isImplementedStaffPattern) {
      Logger.log("  ⚠️ 例外ルールありのためスキップ（要個別対応）：" + staffId + "　内容：" + exceptionNote);
      continue;
    }

    const companyIdsRaw = col(row, ["担当会社"], 3);

    map[staffId] = {
      lineUserId:  col(row, ["LINEユーザーID"], 0),
      staffId:     staffId,
      staffName:   col(row, ["名前"], 2),
      companyIds:  companyIdsRaw ? String(companyIdsRaw).split(",").map(s => s.trim()) : [],
      email:       col(row, ["メール"], 4),
      regNumber:   col(row, ["登録番号"], 5),
      zip:         col(row, ["郵便番号"], 6),
      address:     col(row, ["住所"], 7),
      phone:       col(row, ["電話番号"], 8),
      bankName:    col(row, ["銀行名"], 9),
      branchName:  col(row, ["支店名"], 10),
      accountNum:  col(row, ["口座番号"], 11),
      basicPay:    Number(col(row, ["基本給"])) || 0,
      payUnit:     col(row, ["単位"]),        // "1時間" 等の時間単位、または "月額固定"
      overRate:    Number(col(row, ["超過単価"])) || 0,
      billingPattern: billingPattern,
    };
  }
  return map;
}

// =====================
// スタッフの月給を計算
//   ・単位が「月額固定」→ 実稼働時間に関係なく基本給そのまま支給（超過なし）
//   ・単位が「◯時間」→ その時間までは基本給、超えた分は超過単価×超過時間を加算
//     （超過単価が空欄の場合は会社マスタと同様、単価（基本給÷基本時間）×実稼働時間の実額制）
// ※ まだ請求書明細への書き込みには接続していません
// =====================
function calculateStaffFee_(staff, totalWorkedHours) {
  if (!staff) return null;

  const unit = String(staff.payUnit || "").trim();
  const basicPay = staff.basicPay || 0;

  // 月額固定 → 実稼働時間に関係なく基本給そのまま（超過なし）
  if (unit === "月額固定") {
    return basicPay;
  }

  const basicHours = parseHoursToNumber_(unit);
  if (!basicHours) {
    Logger.log("  ⚠️ 単位から基本時間を判定できないため給与計算不可：" + staff.staffId);
    return null;
  }

  // 超過単価が空欄 → 単価×実稼働時間のみ（超過という概念自体がない）
  if (!staff.overRate) {
    const unitPrice = basicPay / basicHours;
    return unitPrice * totalWorkedHours;
  }

  // 超過単価が設定されている場合：基本時間を超えた分だけ超過単価を加算
  const overHours = Math.max(0, totalWorkedHours - basicHours);
  return basicPay + overHours * staff.overRate;
}

// =====================
// 会社マスタをマップで取得
// =====================
function getCompanyMap(db) {
  const sheet = db.getSheetByName("会社マスタ");
  const data  = sheet.getDataRange().getValues();
  const cm    = getColumnMapFromSheet(sheet); // ヘッダー名→列インデックス（0-indexed）
  const map   = {};

  // ヘッダー名の揺れに対応するためのフォールバック付き取得ヘルパー
  function col(row, names, fallbackIdx) {
    for (let n = 0; n < names.length; n++) {
      if (cm[names[n]] !== undefined) return row[cm[names[n]]];
    }
    return fallbackIdx !== undefined ? row[fallbackIdx] : undefined;
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const companyId = col(row, ["会社ID"], 0);
    if (!companyId) continue;

    // ①「例外」列に値がある会社は、料金ロジック未確定のためお客様請求書の生成は見送るが、
    //    運行管理表（社内向け）は通常通り作成するため、マップからは除外せずフラグとして保持する
    const exceptionNote = col(row, ["例外"]);
    if (exceptionNote) {
      Logger.log("  ⚠️ 例外ルールあり（要個別対応・請求書は生成しません）：" + companyId + "　内容：" + exceptionNote);
    }

    // 「請求パターン」列：例外の中でも実装済みの特殊料金パターンを示す
    //   BILLING_PATTERN.CONSOLIDATED　→ 運行代・交通費・立替費用・時間超過分を1行「貴社運行管理請負費用」にまとめる
    //   BILLING_PATTERN.HOLIDAY_A     → 水・土・日に運行があった日数×1万円（税抜）を「休日運行費」として別行加算
    const billingPattern = String(col(row, ["請求パターン"]) || "").trim();

    map[companyId] = {
      companyId:   companyId,
      companyName: col(row, ["会社名"], 1),
      email:       col(row, ["アドレス", "メールアドレス", "メール"], 2),
      contactName: col(row, ["担当者名"], 3),
      closingDay:  col(row, ["締日"]),
      category:    col(row, ["分類"]),      // "専属" or "スポット"
      basicFee:    Number(col(row, ["基本料金"])) || 0,
      basicHours:  col(row, ["基本時間"]),
      overRate:    Number(col(row, ["超過単価"])) || 0,
      exception:   exceptionNote || "",     // 値があり、かつ請求パターンが未実装のものは請求書生成をスキップ
      billingPattern: billingPattern,
    };
  }
  return map;
}

// =====================
// ③ 基本料金が発生するかどうかを判定
//   ・専属：当月の運行実績が0件でも基本料金は発生する
//   ・スポット：当月1日でも運行があれば、基本時間未達でも基本料金は発生する
// =====================
function shouldApplyBasicFee_(company, companyRecordsThisMonth) {
  if (!company) return false;
  const hasAnyRecord = !!(companyRecordsThisMonth && companyRecordsThisMonth.length > 0);

  if (company.category === "専属") {
    return true;
  }
  if (company.category === "スポット") {
    return hasAnyRecord;
  }
  Logger.log("  ⚠️ 分類（専属／スポット）が未設定のため基本料金の判定不可：" + company.companyId);
  return false;
}

// =====================
// 会社ごとの月額料金を計算
//   ・超過単価が空欄の場合：超過という区分をせず、単価（基本料金÷基本時間）×実際の稼働時間の実額制
//   ・超過単価が設定されている場合：基本時間までは基本料金、超えた時間分は超過単価×超過時間を加算
//     （③のルールにより、専属は当月実績0件でも基本料金は発生、スポットは1日でも運行があれば基本時間未達でも基本料金発生）
// ※ まだ請求書明細への書き込みには接続していません（金額の書式・消費税計算が未確定のため）
// =====================
function calculateCompanyFee_(company, totalWorkedHours) {
  if (!company) return null;

  const basicFee   = company.basicFee || 0;
  const basicHours = parseHoursToNumber_(company.basicHours);

  // 超過単価が空欄 → 単価×実稼働時間のみ（超過という概念自体がない）
  if (!company.overRate) {
    if (!basicHours) {
      Logger.log("  ⚠️ 基本時間が未設定のため単価計算不可：" + company.companyId);
      return null;
    }
    const unitPrice = basicFee / basicHours;
    return unitPrice * totalWorkedHours;
  }

  // 超過単価が設定されている場合：基本時間を超えた分だけ超過単価を加算
  if (!basicHours) {
    Logger.log("  ⚠️ 基本時間が未設定のため超過計算不可（基本料金のみ）：" + company.companyId);
    return basicFee;
  }
  const overHours = Math.max(0, totalWorkedHours - basicHours);
  return basicFee + overHours * company.overRate;
}

// 対象レコード群の中から「水・土・日」に運行/出勤があったユニークな日数をカウントする
// ※ 同じ日に複数レコードがあっても1日としてカウント（二重計上しない）
function countHolidayDates_(rows) {
  const seen = {};
  (rows || []).forEach(function(r) {
    if (!r.date) return;
    const d = (r.date instanceof Date) ? r.date : new Date(r.date);
    if (isNaN(d.getTime())) return;
    const dow = d.getDay(); // 0:日 3:水 6:土
    if (dow === 0 || dow === 3 || dow === 6) {
      const key = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy/MM/dd");
      seen[key] = true;
    }
  });
  return Object.keys(seen).length;
}

// "苗字　名前"（全角スペース区切り）から苗字だけを抽出する。全角スペースが無ければ全体をそのまま返す
function extractSurname_(fullName) {
  if (!fullName) return fullName;
  const parts = String(fullName).split("　");
  return parts[0];
}

// "9時間" "180時間" のような表記から数値（時間）を取り出す
function parseHoursToNumber_(hoursText) {
  if (hoursText === undefined || hoursText === null || hoursText === "") return null;
  if (typeof hoursText === "number") return hoursText;
  const m = String(hoursText).match(/(\d+(\.\d+)?)/);
  return m ? parseFloat(m[1]) : null;
}

// =====================
// 請求書明細（基本料金・超過時間・立替費用）を組み立てるための共通ヘルパー
// =====================

// 運行記録シートの表示値から、該当レコード1件分の実稼働時間（分）を取得
// ※ writeDetailRows（kanri）内のマッチングロジックと同等
function getMatchedWorkingMinutes_(r, dcm, displayValues) {
  const tz = Session.getScriptTimeZone();
  let formattedTargetDate = "";
  if (r.date instanceof Date) {
    formattedTargetDate = Utilities.formatDate(r.date, tz, "yyyy/MM/dd");
  } else if (r.date) {
    const d = new Date(r.date);
    formattedTargetDate = isNaN(d.getTime()) ? String(r.date) : Utilities.formatDate(d, tz, "yyyy/MM/dd");
  }
  const targetTimestamp = r.timestamp ? String(r.timestamp).trim() : "";

  for (let k = 1; k < displayValues.length; k++) {
    const rowTimestamp = displayValues[k][dcm["タイムスタンプ"]] ? String(displayValues[k][dcm["タイムスタンプ"]]).trim() : "";
    const rowDateStr   = displayValues[k][dcm["日付"]]           ? String(displayValues[k][dcm["日付"]]).trim()           : "";
    const rowStaffId   = displayValues[k][dcm["スタッフID"]]     ? String(displayValues[k][dcm["スタッフID"]]).trim()     : "";

    const normalizedRowDate    = rowDateStr.replace(/-/g, "/");
    const normalizedTargetDate = formattedTargetDate.replace(/-/g, "/");

    if ((targetTimestamp !== "" && rowTimestamp === targetTimestamp) ||
        (normalizedRowDate === normalizedTargetDate && rowStaffId === String(r.staffId).trim())) {
      const whDisplay = String(displayValues[k][dcm["稼働時間"]] || "").trim();
      let workingMinutes = 0;
      const jpH = whDisplay.match(/(\d+)時間/);
      const jpM = whDisplay.match(/(\d+)分/);
      if (jpH || jpM) {
        workingMinutes = (jpH ? parseInt(jpH[1]) : 0) * 60 + (jpM ? parseInt(jpM[1]) : 0);
      } else {
        const hm = whDisplay.match(/^(\d+):(\d{2})/);
        if (hm) workingMinutes = parseInt(hm[1]) * 60 + parseInt(hm[2]);
      }
      return workingMinutes;
    }
  }
  return 0;
}

// 明細行（rows）の合計稼働時間を時間（小数）で算出
function getTotalWorkedHours_(rows, dcm, displayValues) {
  let totalMinutes = 0;
  rows.forEach(function(r) {
    totalMinutes += getMatchedWorkingMinutes_(r, dcm, displayValues);
  });
  return totalMinutes / 60;
}

// 超過時間の端数処理：15分単位で切り上げ
function ceilToQuarterHour_(hours) {
  return Math.ceil(hours * 4) / 4;
}

// 税込金額 → 税抜金額。立替費用（フォーム入力＝税込）用
// 正しい順序：消費税＝税込金額÷11（切り捨て）を先に算出 → 税抜金額＝税込金額－消費税
function toTaxExcluded_(taxIncludedAmount) {
  const tax = Math.floor(taxIncludedAmount / 11);
  return taxIncludedAmount - tax;
}

// toTaxExcluded_ と同じ計算式で消費税額だけを取り出す（合計金額が税込入力の原本と1円もズレないようにするため）
// ※ 集計時にfloor(税抜合計×10%)で消費税を再計算すると、変換時の消費税と丸めの関係で1円ズレることがあるため、
//    税込→税抜に変換した項目は、変換時に確定した消費税額をそのまま合計に使う
function taxPortionOfIncluded_(taxIncludedAmount) {
  return Math.floor(taxIncludedAmount / 11);
}

// "2026年07月" のような対象月ラベルから、その月の1日を "yyyy/MM/dd" 文字列で返す
function getFirstDayOfMonthStr_(targetLabel) {
  const m = String(targetLabel || "").match(/(\d+)年(\d+)月/);
  if (!m) return "";
  const y = parseInt(m[1]);
  const mo = String(parseInt(m[2])).padStart(2, "0");
  return y + "/" + mo + "/01";
}

// 会社向け請求書の明細行（運行代・交通費・立替費用・時間超過分）を組み立てる
// ※ 会社マスタの基本料金・超過単価は税抜で入力されている前提
// ※ 表示順：運行代 → 交通費 → 立替費用（ガソリン代・燃料代・パーキング代） → 時間超過分
function buildCompanyInvoiceLineItems_(company, rows, dcm, displayValues) {
  const items = [];
  if (!company) return items;

  const applyBasicFee = shouldApplyBasicFee_(company, rows);
  const basicHours    = parseHoursToNumber_(company.basicHours);
  const totalHours    = getTotalWorkedHours_(rows, dcm, displayValues);
  const hasOverRate   = !!company.overRate;
  const pattern       = String(company.billingPattern || "").trim();
  const isConsolidated = pattern === BILLING_PATTERN.CONSOLIDATED; // 会社：1行にまとめる
  const isHolidayAdd   = pattern === BILLING_PATTERN.HOLIDAY_A;    // 会社：休日運行費を別行加算

  // 統合パターンの場合は明細を一旦別配列に貯めて、最後に合算した1行に差し替える
  const baseItems = [];

  // ① 運行代（基本料金 or 超過単価空欄時の実額制）
  if (!hasOverRate) {
    // 超過単価が空欄 → 単価（基本料金÷基本時間）×実稼働時間の実額制（超過という概念なし）
    if (applyBasicFee && basicHours) {
      baseItems.push({ date: "", content: "運行管理費", qty: totalHours, unit: "時間", unitPrice: company.basicFee / basicHours });
    }
  } else if (applyBasicFee) {
    baseItems.push({ date: "", content: "運行管理費", qty: 1, unit: "式", unitPrice: company.basicFee });
  }

  // ② 交通費（電車通勤）：お客様への請求不要チェックの概念はないため常に計上
  const trainTotal = rows.reduce(function(sum, r) { return sum + (parseFloat(r.trainCommute) || 0); }, 0);
  if (trainTotal > 0) {
    baseItems.push({ date: "", content: "交通費", qty: 1, unit: "式", unitPrice: toTaxExcluded_(trainTotal), taxAmountOverride: taxPortionOfIncluded_(trainTotal) });
  }

  // ③ 立替費用（ガソリン代・燃料代・パーキング代）：お客様への請求不要チェックは除外
  const expenseTotals = { "ガソリン代": 0, "燃料代": 0, "パーキング代": 0 };
  rows.forEach(function(r) {
    if (!r.gasolineNC) expenseTotals["ガソリン代"]   += parseFloat(r.gasoline) || 0;
    if (!r.fuelNC)     expenseTotals["燃料代"]       += parseFloat(r.fuel)     || 0;
    if (!r.parkingNC)  expenseTotals["パーキング代"] += parseFloat(r.parking)  || 0;
  });
  ["ガソリン代", "燃料代", "パーキング代"].forEach(function(label) {
    const taxIncluded = expenseTotals[label];
    if (taxIncluded > 0) {
      baseItems.push({ date: "", content: label, qty: 1, unit: "式", unitPrice: toTaxExcluded_(taxIncluded), taxAmountOverride: taxPortionOfIncluded_(taxIncluded) });
    }
  });

  // ④ 時間超過分（超過単価が設定されている場合のみ）
  if (hasOverRate && basicHours) {
    const overHours = ceilToQuarterHour_(Math.max(0, totalHours - basicHours));
    if (overHours > 0) {
      baseItems.push({ date: "", content: "時間超過分" + overHours + "時間", qty: overHours, unit: "時間", unitPrice: company.overRate });
    }
  }

  if (isConsolidated) {
    // 一本化パターン：運行代・交通費・立替費用・時間超過分をすべて合算して1行にまとめる
    // ※ 各明細の消費税額（taxAmountOverride、無ければ税抜額×10%を切り捨て）も合算し、
    //    合算後の1行にそのまま引き継ぐ（集計時の再計算による1円ズレを防ぐため）
    let total = 0;
    let totalTaxOverride = 0;
    baseItems.forEach(function(it) {
      const amount = (it.qty || 0) * (it.unitPrice || 0);
      total += amount;
      totalTaxOverride += (it.taxAmountOverride !== undefined) ? it.taxAmountOverride : Math.floor(amount * 0.1);
    });
    if (total > 0 || applyBasicFee) {
      items.push({ date: "", content: "貴社運行管理請負費用", qty: 1, unit: "式", unitPrice: total, taxAmountOverride: totalTaxOverride });
    }
  } else {
    baseItems.forEach(function(it) { items.push(it); });
  }

  // 休日加算Aパターン：水・土・日に運行があった日数×1万円（税抜）を「休日運行費」として別行加算
  if (isHolidayAdd) {
    const holidayDays = countHolidayDates_(rows);
    if (holidayDays > 0) {
      items.push({ date: "", content: "休日運行費", qty: holidayDays, unit: "日", unitPrice: 10000 });
    }
  }

  // お客様請求書は自社（インボイス登録済み）発行のため、常に課税対象（10%）
  items.forEach(function(it) { it.taxRate = 0.1; });

  return items;
}

// スタッフ向け請求書の明細行（運行代・交通費・立替費用・時間超過分）を組み立てる
// ※ スタッフマスタの基本給・超過単価も会社マスタと同様、税抜で入力されている前提
// ※ 表示順：運行代 → 交通費 → 立替費用（ガソリン代・燃料代・パーキング代） → 時間超過分
function buildStaffInvoiceLineItems_(staff, rows, dcm, displayValues) {
  const items = [];
  if (!staff) return items;

  const unit       = String(staff.payUnit || "").trim();
  const totalHours = getTotalWorkedHours_(rows, dcm, displayValues);
  let basicHours   = null;
  let hasOverRate  = false;

  // 登録番号（インボイス発行事業者番号）がある場合のみ課税対象（10%）。
  // 登録番号が空欄＝免税事業者扱いのため、税率は0%にする（消費税0円、小計＝合計になる）
  const isTaxable = !!staff.regNumber;

  // ① 運行代
  if (unit === "月額固定") {
    // 実稼働時間に関係なく基本給そのまま（超過なし）
    items.push({ date: "", content: "運行代", qty: 1, unit: "式", unitPrice: staff.basicPay });
  } else {
    basicHours = parseHoursToNumber_(unit);
    hasOverRate = !!staff.overRate;
    if (basicHours) {
      if (!hasOverRate) {
        items.push({ date: "", content: "運行代", qty: totalHours, unit: "時間", unitPrice: staff.basicPay / basicHours });
      } else {
        items.push({ date: "", content: "運行代", qty: 1, unit: "式", unitPrice: staff.basicPay });
      }
    } else {
      Logger.log("  ⚠️ 単位から基本時間を判定できないため給与明細を計算できません：" + staff.staffId);
    }
  }

  // ② 交通費（電車通勤）：会社クレカ払いの場合は対象外
  // フォーム入力は税込のまま。登録番号があれば税抜に変換、なければ税込金額をそのまま使う
  const trainTotal = rows.reduce(function(sum, r) {
    return sum + (r.companyCardPayment ? 0 : (parseFloat(r.trainCommute) || 0));
  }, 0);
  if (trainTotal > 0) {
    items.push({ date: "", content: "交通費", qty: 1, unit: "式", unitPrice: isTaxable ? toTaxExcluded_(trainTotal) : trainTotal, taxAmountOverride: isTaxable ? taxPortionOfIncluded_(trainTotal) : undefined });
  }

  // ③ 立替費用（ガソリン代・燃料代・パーキング代）：会社クレカ払いの場合は対象外
  // フォーム入力は税込のまま。登録番号があれば税抜に変換、なければ税込金額をそのまま使う
  const expenseTotals = { "ガソリン代": 0, "燃料代": 0, "パーキング代": 0 };
  rows.forEach(function(r) {
    if (r.companyCardPayment) return;
    expenseTotals["ガソリン代"]   += parseFloat(r.gasoline) || 0;
    expenseTotals["燃料代"]       += parseFloat(r.fuel)     || 0;
    expenseTotals["パーキング代"] += parseFloat(r.parking)  || 0;
  });
  ["ガソリン代", "燃料代", "パーキング代"].forEach(function(label) {
    const taxIncluded = expenseTotals[label];
    if (taxIncluded > 0) {
      items.push({ date: "", content: label, qty: 1, unit: "式", unitPrice: isTaxable ? toTaxExcluded_(taxIncluded) : taxIncluded, taxAmountOverride: isTaxable ? taxPortionOfIncluded_(taxIncluded) : undefined });
    }
  });

  // ④ 時間超過分（時間単位かつ超過単価が設定されている場合のみ）
  if (unit !== "月額固定" && hasOverRate && basicHours) {
    const overHours = ceilToQuarterHour_(Math.max(0, totalHours - basicHours));
    if (overHours > 0) {
      items.push({ date: "", content: "時間超過分" + overHours + "時間", qty: overHours, unit: "時間", unitPrice: staff.overRate });
    }
  }

  // ⑤ 休日加算Aパターン：水・土・日に出勤した日数×1万円（税抜）を「休日手当」として別行加算
  if (String(staff.billingPattern || "").trim() === BILLING_PATTERN.HOLIDAY_A) {
    const holidayDays = countHolidayDates_(rows);
    if (holidayDays > 0) {
      items.push({ date: "", content: "休日手当", qty: holidayDays, unit: "日", unitPrice: 10000 });
    }
  }

  items.forEach(function(it) { it.taxRate = isTaxable ? 0.1 : 0; });

  return items;
}

// 組み立てた明細を請求書テンプレートの明細エリア（B/D/L/M/N/P列）へ書き込む
// Q列（金額）は "=IF(L<>"",L*N,"")" 形式の数式を都度設定する（テンプレート側で未設定の行があるため）
// dateStr：明細行すべてに共通で入れる日付（対象月の1日）
function writeInvoiceLineItems_(sheet, items, startRow, dateStr) {
  items.forEach(function(item, idx) {
    const row = startRow + idx;
    sheet.getRange(row, 2).setValue(dateStr || "");                    // B: 日付（対象月の1日）
    sheet.getRange(row, 4).setValue(item.content);                     // D: 内容
    sheet.getRange(row, 12).setValue(item.qty);                        // L: 数量
    sheet.getRange(row, 13).setValue(item.unit);                       // M: 単位
    sheet.getRange(row, 14).setValue(Math.round(item.unitPrice));      // N: 単価（税抜）
    sheet.getRange(row, 16).setValue(item.taxRate !== undefined ? item.taxRate : 0.1); // P: 税率（登録番号が無いスタッフは空欄＝非課税）
    sheet.getRange(row, 17).setFormula("=IF(L" + row + "<>\"\",L" + row + "*N" + row + ",\"\")"); // Q: 金額（税抜）
  });
}

// 小計・消費税・合計をテンプレートの数式（SUM/SUMIF）に頼らず、GAS側で直接計算して値として書き込む
// D38：10%対象の消費税／F38：10%対象の金額（税抜）
// D39：8%対象の消費税／F39：8%対象の金額（税抜）
// Q36：小計／Q37：消費税／Q38：合計
// B11：ご請求金額（税込）＝合計と同じ値を直接書き込む
// startRow：明細の書き込み開始行（税込原本とのズレ調整で、最後の課税明細のQ列を直接上書きするために必要）
function writeInvoiceSummary_(sheet, items, startRow) {
  const amounts = items.map(function(item) { return (item.qty || 0) * (Math.round(item.unitPrice) || 0); });

  let net10 = 0, net8 = 0, netOther = 0;
  items.forEach(function(item, i) {
    const amount = amounts[i];
    if (item.taxRate === 0.1) {
      net10 += amount;
    } else if (item.taxRate === 0.08) {
      net8 += amount;
    } else {
      netOther += amount; // 登録番号なしスタッフ等、非課税扱い（税率未設定）
    }
  });

  // インボイス制度のルールに則り、消費税の端数処理は「税率ごとに請求書1枚につき1回」だけ行う
  const tax10 = Math.floor(net10 * 0.10);
  const tax8  = Math.floor(net8  * 0.08);
  const taxTotal = tax10 + tax8;

  // 税込で入力された項目の原本金額（税込→税抜変換時に確定した消費税＝taxAmountOverrideを使う）と、
  // 上記の「税率ごとに1回だけ」ルールで出した合計との差額を求める
  // ※ 通常は税込⇔税抜の変換誤差でごくわずか（±1円程度）に収まる想定
  let idealGrandTotal = 0;
  items.forEach(function(item, i) {
    const amount = amounts[i];
    if (item.taxAmountOverride !== undefined) {
      idealGrandTotal += amount + item.taxAmountOverride;
    } else if (item.taxRate === 0.1 || item.taxRate === 0.08) {
      idealGrandTotal += amount + Math.floor(amount * item.taxRate);
    } else {
      idealGrandTotal += amount;
    }
  });

  let subtotal   = net10 + net8 + netOther;
  let grandTotal = subtotal + taxTotal;
  const diff = idealGrandTotal - grandTotal;

  // ズレがある場合、最後の課税対象明細1行のQ列（金額）だけを直接上書きして吸収する
  // （L×Nの数式ではなく、その行だけ確定値を書き込む。他の行・単価は一切変更しない）
  if (diff !== 0 && startRow !== undefined) {
    let targetIdx = -1;
    for (let i = items.length - 1; i >= 0; i--) {
      if (items[i].taxRate === 0.1 || items[i].taxRate === 0.08) { targetIdx = i; break; }
    }
    if (targetIdx !== -1) {
      sheet.getRange(startRow + targetIdx, 17).setValue(amounts[targetIdx] + diff); // Q列を確定値で上書き
      subtotal   += diff;
      grandTotal += diff;
    }
  }

  sheet.getRange("F38").setValue(net10);
  sheet.getRange("F39").setValue(net8);
  sheet.getRange("D38").setValue(tax10);
  sheet.getRange("D39").setValue(tax8);
  sheet.getRange("Q36").setValue(subtotal);   // 小計
  sheet.getRange("Q37").setValue(taxTotal);   // 消費税
  sheet.getRange("Q38").setValue(grandTotal); // 合計
  sheet.getRange("B11").setValue(grandTotal); // ご請求金額（税込）
}

// =====================
// 会社マスタの「締日」（支払日パターン："20日" or "月末"）に応じて支払期日を算出
//   ・"月末" → 翌月末日
//   ・それ以外（"20日"含む・未設定含む） → 翌月20日（デフォルト）
// =====================
function calculatePaymentDate_(targetYear, targetMonth, closingDayType) {
  const paymentYear  = targetMonth === 12 ? targetYear + 1 : targetYear;
  const paymentMonth = targetMonth === 12 ? 1 : targetMonth + 1;

  const type = String(closingDayType || "").trim();
  if (type === "月末") {
    // new Date(year, month, 0) は「month月の0日目」＝「month-1月の末日」を返す仕様を利用し、
    // paymentMonthの月末日を取得する
    const lastDay = new Date(paymentYear, paymentMonth, 0).getDate();
    return paymentYear + "年" + String(paymentMonth).padStart(2, "0") + "月" + lastDay + "日";
  }
  return paymentYear + "年" + String(paymentMonth).padStart(2, "0") + "月20日";
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

    // 明細行を書き込み（小計・消費税・合計もこの中でGAS側が直接計算して書き込む）
    writeDetailRows(sheet, rows, "staff", staff, targetLabel);

    // スタッフ請求書シートのみPDF出力
    savePdfSingleSheet(ss, copy, fileName, folder, "スタッフ請求書");
    Logger.log("  生成完了：" + fileName);
  });
}

// =====================
// ② 運行管理表・お客様請求書を生成
// =====================
function generateCompanyDocs(records, companyMap, targetLabel, targetYear, targetMonth, templateId, kanriFolder, companyFolder) {
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

  // ③「専属」会社は当月の運行実績が0件でも基本料金が発生するため、
  //    運行記録がなくても請求対象に含める（明細0件の請求書を生成）
  Object.keys(companyMap).forEach(function(companyId) {
    const company = companyMap[companyId];
    if (company.category === "専属" && !grouped[companyId]) {
      grouped[companyId] = [];
      Logger.log("  専属契約かつ当月運行実績0件のため、明細0件で請求対象に追加：" + companyId + "_" + company.companyName);
    }
  });

  const companyIds = Object.keys(grouped);
  Logger.log("対象会社数：" + companyIds.length + " 社");

  companyIds.forEach(function(companyId) {
    const company = companyMap[companyId];
    if (!company) {
      Logger.log("⚠️ 会社マスタに未登録の会社IDをスキップ：" + companyId);
      return;
    }

    // ③ 基本料金の発生有無を判定（金額の書き込み自体は未実装のためログ出力のみ）
    const applyBasicFee = shouldApplyBasicFee_(company, grouped[companyId]);
    Logger.log("  基本料金判定：" + companyId + "_" + company.companyName + "（分類：" + (company.category || "未設定") + "）→ " + (applyBasicFee ? "発生する" : "発生しない"));
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

    // お客様請求書（①「例外」ありでも、請求パターンが実装済み（COMPANY_BILLING_PATTERNS）なら通常通り生成する）
    const isImplementedCompanyPattern = COMPANY_BILLING_PATTERNS.indexOf(company.billingPattern) !== -1;
    if (company.exception && !isImplementedCompanyPattern) {
      Logger.log("    ⚠️ 例外ルールありのためお客様請求書はスキップ：" + companyId + "_" + company.companyName + "（" + company.exception + "）");
      return;
    }

    const invoiceName = "【請求書】" + companyId + "_" + company.companyName + "_" + targetLabel;
    Logger.log("    お客様請求書 生成中：" + invoiceName);
    const invoiceCopy  = DriveApp.getFileById(templateId).makeCopy(invoiceName, companyFolder);
    const invoiceSS    = SpreadsheetApp.openById(invoiceCopy.getId());
    const invoiceSheet = invoiceSS.getSheetByName("お客様請求書");

    replaceInSheet(invoiceSheet, "{{会社名}}",   company.companyName);
    replaceInSheet(invoiceSheet, "{{発行日}}",   issueDate);
    replaceInSheet(invoiceSheet, "{{請求番号}}", companyId + "_" + invoiceMonthKey + "_00");

    // 振込期日（会社マスタの「締日」＝支払日パターンに応じて算出：「20日」or「月末」）
    const companyPaymentDate = calculatePaymentDate_(targetYear, targetMonth, company.closingDay);
    invoiceSheet.getRange("C16").setValue(companyPaymentDate);

    // 明細行を書き込み（小計・消費税・合計もこの中でGAS側が直接計算して書き込む）
    writeDetailRows(invoiceSheet, rows, "company", company, targetLabel);
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

    // 請求書・運行管理表のどちらもなければスキップ（例外ありの会社は請求書が無いこともある）
    if (!invoiceFile && !kanriFile) {
      Logger.log("  ⚠️ 請求書・運行管理表とも見つからないためスキップ：" + companyId + "_" + company.companyName);
      skipCount++;
      return;
    }

    Logger.log("  下書き作成中：" + company.email + "（" + company.companyName + "）");
    if (!invoiceFile) {
      Logger.log("  ⚠️ 請求書PDFが見つからないため運行管理表のみ添付：" + invoiceName);
    }
    if (!kanriFile) {
      Logger.log("  ⚠️ 運行管理表PDFが見つからないため請求書のみ添付：" + kanriName);
    }

    const subject = "【ご請求】" + targetLabel + "分　" + company.companyName + " 御中";

    const body =
      (company.contactName || company.companyName) + " 様\n\n" +
      "いつもお世話になっております。\n" +
      "株式会社エグゼクティブサポートでございます。\n\n" +
      targetLabel + "分のご請求書および運行管理表をお送りいたします。\n" +
      "ご確認のほど、よろしくお願いいたします。\n\n" +
      "──────────────────\n" +
      "株式会社エグゼクティブサポート\n" +
      "──────────────────";

    // 添付ファイルの配列を準備（存在するものだけ添付）
    const attachments = [];
    if (invoiceFile) attachments.push(invoiceFile.getAs(MimeType.PDF));
    if (kanriFile)   attachments.push(kanriFile.getAs(MimeType.PDF));

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
function writeDetailRows(sheet, rows, type, masterRecord, targetLabel) {
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

  // ヘッダー名→列インデックスマップをループ外で1回だけ取得（パフォーマンス対策）
  // ※ startRow-1 がヘッダー行（テンプレートが複数行タイトルを持つ場合でも正しく取得できる）
  function buildColMap(targetSheet, headerRow) {
    const headers = targetSheet.getRange(headerRow, 1, 1, targetSheet.getLastColumn()).getValues()[0];
    const map = {};
    headers.forEach(function(h, i) { if (h) map[String(h).trim()] = i + 1; }); // 1-indexed
    return map;
  }
  const kanriCm = (type === "kanri") ? buildColMap(sheet, startRow - 1) : {};
  const dcm     = getColumnMapFromSheet(recordSheet); // 運行記録シートは1行目がヘッダー

  // お客様請求書・スタッフ請求書：日ごとの明細ではなく、運行代・時間超過分・立替費用の集計行を書き込む
  if (type === "company" || type === "staff") {
    const items = (type === "company")
      ? buildCompanyInvoiceLineItems_(masterRecord, rows, dcm, displayValues)
      : buildStaffInvoiceLineItems_(masterRecord, rows, dcm, displayValues);
    const dateStr = getFirstDayOfMonthStr_(targetLabel); // すべて対象月の1日
    writeInvoiceLineItems_(sheet, items, startRow, dateStr);
    writeInvoiceSummary_(sheet, items, startRow); // 小計・消費税・合計はテンプレート数式を使わずGAS側で直接計算
    Logger.log("    明細書き込み完了：type=" + type + "、" + items.length + " 件");
    return;
  }

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

      // 元のシートの行を特定し「見かけ上の時間・稼働時間」を取得する（dcmはループ外で取得済み）
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

      // 立替費用：請求不要チェックありは0表示
      const gasolineVal = r.gasolineNC ? 0 : (parseFloat(r.gasoline) || 0);
      const fuelVal     = r.fuelNC     ? 0 : (parseFloat(r.fuel)     || 0);
      const parkingVal  = r.parkingNC  ? 0 : (parseFloat(r.parking)  || 0);

      // テンプレートのヘッダー名で動的に列マップを取得して書き込む（列順変更に対応）
      // ヘッダー名で列位置を動的取得しつつ、書き込みは行単位でまとめて実行（高速化）
      function getCol(name) { return kanriCm[name] !== undefined ? kanriCm[name] : -1; } // buildColMapは既に1-indexed

      // 書き込む列・値のペアを配列で構築
      const writes = [
        ["日付",         displayDate],
        ["運行者",       extractSurname_(r.staffName)],
        ["車両",         r.vehicle],
        ["出発",         r.departure],
        ["行き先",       r.destination],
        ["到着",         r.arrival],
        ["出発時間",     dispDepartureTime],
        ["到着日",       displayArrivalDate],
        ["到着時間",     dispArrivalTime],
        ["稼働時間",     workingSerial],
        ["利用者",       r.passenger],
        ["利用目的",     r.purpose],
        ["メーター走行前", r.meterBefore],
        ["メーター走行後", r.meterAfter],
        ["走行距離",     r.distance],
        ["電車通勤",     r.trainCommute],
        ["ガソリン代",   gasolineVal || ""],
        ["燃料代",       fuelVal     || ""],
        ["パーキング代", parkingVal  || ""],
        ["備考",         r.memo],
      ];

      // 列番号が存在するものだけセルに書き込む
      writes.forEach(function(pair) {
        const colIdx = getCol(pair[0]);
        if (colIdx === -1) return;
        sheet.getRange(row, colIdx).setValue(pair[1]);
      });

      // 稼働時間の表示形式を[h]:mmに設定
      const whColIdx = getCol("稼働時間");
      if (whColIdx !== -1) sheet.getRange(row, whColIdx).setNumberFormat("[h]:mm");
    }
  });

  // 運行管理表（kanri）の場合のみ、合計行にSUM式を設定する
  // ※稼働時間合計はテンプレートに既存式があるためGASからは書き込まない
  if (type === "kanri" && rows.length > 0) {
    const lastDataRow = startRow + rows.length - 1;
    const targetTotalRow = 37; // 合計値行（36行目がラベル、37行目が数値）

    Logger.log("    合計行を設定：" + targetTotalRow + "行目（データ " + startRow + "〜" + lastDataRow + "行）");

    // ヘッダー名で列番号を動的取得してSUM式を設定（kanriCmを再利用・ヘッダー行は startRow-1）
    const sumTargets = ["電車通勤", "ガソリン代", "燃料代", "パーキング代"];
    sumTargets.forEach(function(colName) {
      if (kanriCm[colName] === undefined) return;
      const colIdx    = kanriCm[colName]; // buildColMapは既に1-indexed
      const colLetter = columnToLetter_(colIdx);
      sheet.getRange(targetTotalRow, colIdx)
           .setFormula("=SUM(" + colLetter + startRow + ":" + colLetter + lastDataRow + ")");
    });
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
// 列番号（1-indexed）をA1形式のアルファベットに変換するヘルパー
function columnToLetter_(col) {
  let letter = "";
  while (col > 0) {
    const mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

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
    generateCompanyDocs(records, companyMap, targetLabel, ym.year, ym.month, props.getProperty("TEMPLATE_SPREADSHEET_ID"), kanriFolder, companyFolder);

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
