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
  // コンテナバインド（スプレッドシートの拡張機能から作成した場合）
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
    "希望カテゴリ",
    "希望種・生体名",
    "数量・サイズ感",
    "希望納期",
    "目安予算",
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
    sheet.setColumnWidth(8, 130); // カテゴリ
    sheet.setColumnWidth(9, 200); // 生体名
    sheet.setColumnWidth(10, 140); // 数量サイズ
    sheet.setColumnWidth(11, 130); // 納期
    sheet.setColumnWidth(12, 110); // 予算
    sheet.setColumnWidth(13, 90);  // 同意
    sheet.setColumnWidth(14, 250); // 備考
    sheet.setColumnWidth(15, 110); // ステータス
    sheet.setColumnWidth(16, 200); // メモ
  } else {
    // 既存シートのヘッダーが未設定の場合
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
 * GETリクエストハンドラ（疎通確認用）
 */
function doGet(e) {
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

    var rowValues = [
      dateStr,
      reqId,
      String(data.name || "").trim(),
      String(data.phone || "").trim(),
      String(data.email || "").trim(),
      String(data.lineUserId || "").trim(),
      String(data.prefecture || "").trim(),
      String(data.category || "").trim(),
      String(data.species || "").trim(),
      String(data.quantitySize || "").trim(),
      String(data.deadline || "").trim(),
      String(data.budget || "").trim(),
      data.agreedToTerms ? "同意済" : "未同意",
      String(data.notes || "").trim(),
      "新規受付", // 初期対応ステータス
      "" // 管理者メモ
    ];

    sheet.appendRow(rowValues);

    // LINEプッシュ通知（管理者宛 & お客様宛）
    sendRequestLineNotifications({
      reqId: reqId,
      dateStr: dateStr,
      data: data
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
      "🐟希望カテゴリ: " + (d.category || "未選択") + "\n" +
      "📝希望生体名: " + (d.species || "未入力") + "\n" +
      "📦数量/サイズ: " + (d.quantitySize || "未入力") + "\n" +
      "⏳希望納期: " + (d.deadline || "指定なし") + "\n" +
      "💰目安予算: " + (d.budget || "指定なし") + "\n" +
      "-------------------\n" +
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
      "🔖受付番号: " + info.reqId + "\n" +
      "🐟希望生体: " + (d.species || "") + " (" + (d.category || "") + ")\n" +
      "📦数量・サイズ: " + (d.quantitySize || "") + "\n" +
      "⏳希望納期: " + (d.deadline || "") + "\n\n" +
      "当店の採集員が天候や海況を確認の上、採集可否や目処が立ち次第、改めてご連絡いたします。\n" +
      "※天候や海の状況によってはお時間をいただく場合がございます。";

    pushLineMessage(d.lineUserId, userText);
  }
}
