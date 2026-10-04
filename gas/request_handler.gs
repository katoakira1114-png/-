// =============================================================================
// 川畑水産（観賞魚） 生体リクエスト・採集要望受付 GASスクリプト
// =============================================================================

// スクリプトプロパティから取得（未設定時の初期フォールバック値）
var SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty("REQUEST_SPREADSHEET_ID") || "";
var LINE_CHANNEL_ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty("LINE_CHANNEL_ACCESS_TOKEN") || "4z5y4BBsTTTEZgfmcaUF7cbc1T1k6qePWr/xKn3Yuk0E1pi4wxY3uPtSar9xW3F8FxIDpqTpma7lfn3HXhQ19LyTaDkNU6s779LkSLtPCUCZNb3nKg/tDSJXLNuADTfjDGoD8SQMw/CYDBdUAgkq3gdB04t89/1O/w1cDnyilFU=";
var ADMIN_LINE_USER_ID = PropertiesService.getScriptProperties().getProperty("ADMIN_LINE_USER_ID") || "";

var SHEET_NAME = "採集リクエスト一覧";

// 権限認証用ダミー関数（初回実行時に許可を与えるため）
function authorize() {
  SpreadsheetApp.getActiveSpreadsheet();
  DriveApp.getRootFolder();
  UrlFetchApp.fetch("https://www.google.com", { muteHttpExceptions: true });
}

/**
 * スプレッドシートを取得またはバインドする
 */
function getTargetSpreadsheet() {
  if (SPREADSHEET_ID && SPREADSHEET_ID !== "") {
    try {
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    } catch (e) {
      Logger.log("openById failed, fallback to active spreadsheet: " + e.toString());
    }
  }
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    Logger.log("getActiveSpreadsheet failed: " + e.toString());
  }
  return null;
}

/**
 * 採集リクエストシートの初期化およびヘッダー検証
 */
function ensureRequestSheet(ss) {
  var sheet = ss.getSheetByName(SHEET_NAME);
  var headers = [
    "受付日時",
    "リクエストID",
    "お名前 / 店舗名",
    "電話番号",
    "メールアドレス",
    "LINE_UID",
    "お届け先都道府県",
    "希望生体一覧(明細)",
    "品目数",
    "希望納期",
    "合計目安予算",
    "注意事項同意",
    "備考・飼育環境",
    "対応ステータス",
    "管理者メモ"
  ];

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground("#005bac")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
    sheet.setFrozenRows(1);

    // 列幅の調整
    sheet.setColumnWidth(1, 150); // 受付日時
    sheet.setColumnWidth(2, 130); // リクエストID
    sheet.setColumnWidth(3, 160); // お名前
    sheet.setColumnWidth(4, 130); // 電話番号
    sheet.setColumnWidth(5, 180); // メール
    sheet.setColumnWidth(6, 140); // LINE_UID
    sheet.setColumnWidth(7, 120); // 都道府県
    sheet.setColumnWidth(8, 300); // 希望生体一覧(明細)
    sheet.setColumnWidth(9, 80);  // 品目数
    sheet.setColumnWidth(10, 130); // 納期
    sheet.setColumnWidth(11, 120); // 予算
    sheet.setColumnWidth(12, 90);  // 同意
    sheet.setColumnWidth(13, 250); // 備考
    sheet.setColumnWidth(14, 110); // ステータス
    sheet.setColumnWidth(15, 200); // メモ
  } else {
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers])
        .setBackground("#005bac")
        .setFontColor("#ffffff")
        .setFontWeight("bold")
        .setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
    }
  }

  return sheet;
}

/**
 * GETリクエストハンドラ（疎通確認 & 顧客情報検索API）
 */
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? String(e.parameter.action).trim() : "";

  // 顧客マスタ検索API（電話番号 / LINE UID / 名前）
  if (action === "lookupCustomer" || action === "searchCustomer") {
    var queryPhone = (e.parameter.phone || "").replace(/[^0-9]/g, "");
    if (queryPhone.length === 10 && !queryPhone.startsWith("0")) queryPhone = "0" + queryPhone;
    var queryUid = (e.parameter.userId || "").trim().toLowerCase();
    var queryName = (e.parameter.name || "").trim().toLowerCase();

    var ss = getTargetSpreadsheet();
    var foundCustomer = null;
    if (ss) {
      var cSheet = ss.getSheetByName("顧客マスタ");
      if (cSheet) {
        var cData = cSheet.getDataRange().getValues();
        if (cData.length > 1) {
          var h = cData[0];
          var lUid = -1, lName = -1, lParamPhone = -1, lAddr = -1;
          for (var i = 0; i < h.length; i++) {
            var title = String(h[i]).trim();
            if (title.indexOf("LINE") !== -1 || title.indexOf("UID") !== -1) lUid = i;
            else if (title.indexOf("店舗") !== -1 || title.indexOf("氏名") !== -1) lName = i;
            else if (title.indexOf("電話") !== -1 || title.indexOf("TEL") !== -1) lParamPhone = i;
            else if (title.indexOf("住所") !== -1) lAddr = i;
          }
          if (lUid === -1) lUid = 1;
          if (lName === -1) lName = 2;
          if (lParamPhone === -1) lParamPhone = 3;

          for (var r = 1; r < cData.length; r++) {
            var rowUid = String(cData[r][lUid] || "").trim().toLowerCase();
            var rowPhone = String(cData[r][lParamPhone] || "").replace(/[^0-9]/g, "");
            if (rowPhone.length === 10 && !rowPhone.startsWith("0")) rowPhone = "0" + rowPhone;
            var rowName = String(cData[r][lName] || "").trim().toLowerCase();

            var isMatched = false;
            if (queryPhone && rowPhone && queryPhone === rowPhone) isMatched = true;
            else if (queryUid && rowUid && queryUid === rowUid) isMatched = true;
            else if (queryName && rowName && queryName === rowName) isMatched = true;

            if (isMatched) {
              var fullAddress = (lAddr !== -1) ? String(cData[r][lAddr] || "").trim() : "";
              var prefMatch = fullAddress.match(/(東京都|北海道|(?:京都|大阪)府|.{2,3}県)/);
              var extractedPref = prefMatch ? prefMatch[0] : "";

              foundCustomer = {
                found: true,
                name: String(cData[r][lName] || "").trim(),
                phone: String(cData[r][lParamPhone] || "").trim(),
                address: fullAddress,
                prefecture: extractedPref,
                lineUserId: String(cData[r][lUid] || "").trim()
              };
              break;
            }
          }
        }
      }
    }

    return ContentService.createTextOutput(JSON.stringify(
      foundCustomer || { found: false }
    )).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: "ok",
    service: "生体リクエスト・採集要望受付API",
    timestamp: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * POSTリクエストハンドラ（フォーム受付）
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (lockErr) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "回線が混み合っています。少し時間をおいて再度お試しください。"
    })).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var rawText = (e && e.postData && e.postData.contents) ? e.postData.contents : "{}";
    var data = JSON.parse(rawText);

    var ss = getTargetSpreadsheet();
    if (!ss) {
      throw new Error("対象のスプレッドシートが見つかりません。REQUEST_SPREADSHEET_ID を設定してください。");
    }

    var sheet = ensureRequestSheet(ss);

    var now = new Date();
    var dateStr = Utilities.formatDate(now, "JST", "yyyy/MM/dd HH:mm:ss");
    var reqId = "REQ-" + Utilities.formatDate(now, "JST", "yyyyMMdd-HHmmss");

    var items = data.items || [];
    if (items.length === 0 && data.species) {
      items = [{
        category: data.category || "その他",
        species: data.species || "",
        quantity: data.quantitySize || "1",
        size: "",
        budget: data.budget || ""
      }];
    }

    var itemsSummaryLines = [];
    var totalBudgetSummary = data.totalBudget || "";
    for (var itIdx = 0; itIdx < items.length; itIdx++) {
      var itm = items[itIdx];
      var itmLine = "・[" + (itm.category || "その他") + "] " + (itm.species || "未指定") + " × " + (itm.quantity || "1");
      if (itm.size) itmLine += " (" + itm.size + ")";
      if (itm.budget) itmLine += " [予算: " + itm.budget + "]";
      itemsSummaryLines.push(itmLine);
    }
    var itemsSummaryText = itemsSummaryLines.join("\n");

    var rowValues = [
      dateStr,
      reqId,
      String(data.name || "").trim(),
      String(data.phone || "").trim(),
      String(data.email || "").trim(),
      String(data.lineUserId || "").trim(),
      String(data.prefecture || "").trim(),
      itemsSummaryText,
      items.length,
      String(data.deadline || "").trim(),
      String(totalBudgetSummary || "").trim(),
      data.agreedToTerms ? "同意済" : "未同意",
      String(data.notes || "").trim(),
      "新規受付",
      ""
    ];

    sheet.appendRow(rowValues);

    // LINEプッシュ通知（管理者宛 & お客様宛）
    sendRequestLineNotifications({
      reqId: reqId,
      dateStr: dateStr,
      data: data,
      items: items,
      itemsSummaryText: itemsSummaryText,
      totalBudget: totalBudgetSummary
    });

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      requestId: reqId,
      message: "採集リクエストを正常に受け付けました。"
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log("doPost Error: " + err.toString());
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.message || err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/**
 * LINE Messaging API (Push Message) 送信
 */
function pushLineMessage(toUserId, textMessage) {
  if (!toUserId || !LINE_CHANNEL_ACCESS_TOKEN) return false;
  try {
    var url = "https://api.line.me/v2/bot/message/push";
    var payload = {
      to: toUserId,
      messages: [{ type: "text", text: textMessage }]
    };
    var options = {
      method: "post",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + LINE_CHANNEL_ACCESS_TOKEN
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    var res = UrlFetchApp.fetch(url, options);
    return res.getResponseCode() === 200;
  } catch (e) {
    Logger.log("pushLineMessage exception: " + e.toString());
    return false;
  }
}

/**
 * リクエスト受付時のLINE通知（管理者およびお客様）
 */
function sendRequestLineNotifications(info) {
  var d = info.data;
  var items = info.items || [];
  var itemsSummary = info.itemsSummaryText || "";

  // 1. 管理者（店主）宛通知
  var adminUid = ADMIN_LINE_USER_ID;
  if (adminUid) {
    var adminText = "🎣【新規 生体リクエスト・採集要望受付】\n" +
      "-------------------\n" +
      "受付番号: " + info.reqId + "\n" +
      "日時: " + info.dateStr + "\n" +
      "お名前: " + (d.name || "未入力") + " 様\n" +
      "TEL: " + (d.phone || "未入力") + "\n" +
      "お届け先: " + (d.prefecture || "未入力") + "\n" +
      "-------------------\n" +
      "【ご希望生体一覧 (" + items.length + "品目)】\n" +
      itemsSummary + "\n" +
      "-------------------\n" +
      "⏳希望納期: " + (d.deadline || "指定なし") + "\n" +
      (info.totalBudget ? "💰目安予算: " + info.totalBudget + "\n" : "") +
      "💡飼育環境・備考:\n" + (d.notes ? d.notes : "なし") + "\n\n" +
      "※スプレッドシートの「採集リクエスト一覧」をご確認ください。";

    pushLineMessage(adminUid, adminText);
  }

  // 2. お客様宛の確認通知（LINE UIDが取得できている場合）
  if (d.lineUserId) {
    var userText = "【採集リクエストを承りました】\n" +
      (d.name || "お客様") + " 様\n\n" +
      "川畑水産への採集リクエストありがとうございます。\n" +
      "下記の内容で受け付けいたしました。\n\n" +
      "🔖受付番号: " + info.reqId + "\n\n" +
      "【ご希望生体 (" + items.length + "品目)】\n" +
      itemsSummary + "\n\n" +
      "⏳希望納期: " + (d.deadline || "時期不問") + "\n\n" +
      "当店の採集員が海況・天候を確認し、採集後の状態確認・トリートメント完了次第、改めてご連絡いたします。\n" +
      "※自然環境下での採集のため、海況によりお時間をいただく場合がございます。";

    pushLineMessage(d.lineUserId, userText);
  }
}
