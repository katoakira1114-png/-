var SPREADSHEET_ID = "1ZiNel8TGJ-e7JMolkRagVOSnr1UczvQzpTa-F-s4p3o";
var PDF_FOLDER_NAME = "川畑水産_納品書PDF";
var EDIT_FOLDER_NAME = "川畑水産_納品書(編集用)";
var BACKUP_FOLDER_NAME = "川畑水産_スプレッドシート自動バックアップ";

// LINEアクセストークンの解決（スクリプトプロパティから取得・ハードコード禁止）
function getLineChannelAccessToken() {
  var token = PropertiesService.getScriptProperties().getProperty("LINE_CHANNEL_ACCESS_TOKEN") ||
              PropertiesService.getScriptProperties().getProperty("LINE_ACCESS_TOKEN");
  if (!token || token.trim() === "") {
    console.error("【セキュリティ警告】LINE_CHANNEL_ACCESS_TOKENがスクリプトプロパティに未設定です。コード内平文保存を防止するため、GASのプロジェクト設定でスクリプトプロパティを設定してください。");
    return "";
  }
  return token.trim();
}

// 管理者（店主）のLINEユーザーIDの解決（スクリプトプロパティ最優先、ハードコードIDがあれば補完）
function getAdminLineUserId() {
  var uid = PropertiesService.getScriptProperties().getProperty("ADMIN_LINE_USER_ID");
  return (uid && uid.trim() !== "") ? uid.trim() : "";
}

// ピッキング管理用PINコード（スクリプトプロパティADMIN_PIN優先、デフォルト7210）
function getAdminPin() {
  var pin = PropertiesService.getScriptProperties().getProperty("ADMIN_PIN");
  return (pin && pin.trim() !== "") ? pin.trim() : "7210";
}

// サイトに表示する正規カテゴリ定義（表示名・対応シート名候補）
var CATEGORY_DEFS = [
  { label: "魚", sheetNames: ["魚"] },
  { label: "シャコガイ", sheetNames: ["シャコガイ", "シャコ貝"] },
  { label: "貝", sheetNames: ["貝"] },
  { label: "サンゴ", sheetNames: ["サンゴ"] },
  { label: "イソギンチャク", sheetNames: ["イソギンチャク"] },
  { label: "擬岩", sheetNames: ["擬岩"] },
  { label: "その他", sheetNames: ["その他"] }
];

function authorize() {
  SpreadsheetApp.getActiveSpreadsheet();
  DriveApp.getRootFolder();
  UrlFetchApp.fetch("https://www.google.com", { muteHttpExceptions: true });
}

var FISH_IMAGE_FOLDER_NAME = "川畑水産_観賞魚画像";

// アップロード画像をGoogle Driveの観賞魚フォルダへ保存し、直リンクURLを発行
function saveImageToDrive(base64Data, fileName) {
  var folder;
  var folders = DriveApp.getFoldersByName(FISH_IMAGE_FOLDER_NAME);
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(FISH_IMAGE_FOLDER_NAME);
  }

  var pureBase64 = base64Data;
  var mimeType = "image/jpeg";
  if (base64Data.indexOf("data:") === 0) {
    var parts = base64Data.split(",");
    var mimeMatch = parts[0].match(/:(.*?);/);
    if (mimeMatch) mimeType = mimeMatch[1];
    pureBase64 = parts[1];
  }

  var decodedBytes = Utilities.base64Decode(pureBase64);
  var blob = Utilities.newBlob(decodedBytes, mimeType, fileName || ("fish_" + new Date().getTime() + ".jpg"));
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var fileId = file.getId();
  var directUrl = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w1000";
  return {
    fileId: fileId,
    url: directUrl,
    viewUrl: file.getUrl()
  };
}

// カテゴリに対応するシートを取得するヘルパー
function getSheetForCategory(ss, catLabel) {
  for (var i = 0; i < CATEGORY_DEFS.length; i++) {
    if (CATEGORY_DEFS[i].label === catLabel) {
      for (var s = 0; s < CATEGORY_DEFS[i].sheetNames.length; s++) {
        var sh = ss.getSheetByName(CATEGORY_DEFS[i].sheetNames[s]);
        if (sh) return sh;
      }
    }
  }
  return ss.getSheetByName(catLabel);
}

// 顧客登録申請シートと顧客マスタの照合を行い、登録済み顧客を「承認済み」に自動更新する
function syncApplicationStatus(ss) {
  var appSheet = ss.getSheetByName("顧客登録申請");
  var cSheet = ss.getSheetByName("顧客マスタ");
  if (!appSheet || !cSheet) return;

  var appData = appSheet.getDataRange().getValues();
  if (appData.length < 2) return;
  var cData = cSheet.getDataRange().getValues();
  if (cData.length < 2) return;

  // 顧客マスタのヘッダー特定
  var cHeaders = cData[0];
  var cColName = -1, cColPhone = -1;
  for (var ch = 0; ch < cHeaders.length; ch++) {
    var title = String(cHeaders[ch]).trim();
    if (title.indexOf("店舗") !== -1 || title.indexOf("氏名") !== -1) cColName = ch;
    else if (title.indexOf("電話") !== -1 || title.indexOf("TEL") !== -1) cColPhone = ch;
  }
  if (cColName === -1) cColName = 2; // デフォルト C列
  if (cColPhone === -1) cColPhone = 3; // デフォルト D列

  // 顧客登録申請のヘッダー特定
  var appHeaders = appData[0];
  var aColName = -1, aColPhone = -1, aColStatus = -1;
  for (var ah = 0; ah < appHeaders.length; ah++) {
    var atitle = String(appHeaders[ah]).trim();
    if (atitle.indexOf("店舗") !== -1 || atitle.indexOf("氏名") !== -1) aColName = ah;
    else if (atitle.indexOf("電話") !== -1 || atitle.indexOf("TEL") !== -1) aColPhone = ah;
    else if (atitle.indexOf("ステータス") !== -1) aColStatus = ah;
  }
  if (aColName === -1) aColName = 2;
  if (aColPhone === -1) aColPhone = 3;
  if (aColStatus === -1) aColStatus = 5;

  // 顧客マスタの登録済みリストをマップ化 (正規化電話番号 + 正規化店舗名)
  var registeredMap = {};
  for (var cr = 1; cr < cData.length; cr++) {
    var cPhone = String(cData[cr][cColPhone] || "").replace(/[^0-9]/g, "");
    if (cPhone.length === 10 && !cPhone.startsWith("0")) cPhone = "0" + cPhone;
    var cName = String(cData[cr][cColName] || "").trim().toLowerCase();
    if (cPhone) {
      registeredMap[cPhone] = true;
      if (cName) registeredMap[cPhone + "_" + cName] = true;
    }
  }

  // 申請シートの「承認待ち」行をチェックして更新
  for (var ar = 1; ar < appData.length; ar++) {
    var curStatus = String(appData[ar][aColStatus] || "").trim();
    if (curStatus === "承認待ち" || curStatus === "") {
      var aPhone = String(appData[ar][aColPhone] || "").replace(/[^0-9]/g, "");
      if (aPhone.length === 10 && !aPhone.startsWith("0")) aPhone = "0" + aPhone;
      var aName = String(appData[ar][aColName] || "").trim().toLowerCase();

      // 電話番号と店舗名が一致する情報が登録されていれば自動で「承認済み」
      var isMatched = false;
      if (aPhone && registeredMap[aPhone + "_" + aName]) {
        isMatched = true;
      } else if (aPhone && registeredMap[aPhone]) {
        isMatched = true;
      }

      if (isMatched) {
        appSheet.getRange(ar + 1, aColStatus + 1).setValue("承認済み")
          .setBackground("#e8f8f5").setFontColor("#117a65").setFontWeight("bold");
      }
    }
  }
}

// 顧客マスタの列構成を整理・移行（B:LINE_UID, C:店舗名, D:電話番号, E:住所, F〜L:各カテゴリ許可, M〜O:商品許可A,B,C）
function ensureCustomerMasterColumns(ss) {
  var cSheet = ss.getSheetByName("顧客マスタ");
  if (!cSheet) return;

  var lastCol = cSheet.getLastColumn();
  if (lastCol < 1) return;
  var headers = cSheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // 必要な列名定義
  var targetCatHeaders = ["魚", "シャコガイ", "貝", "サンゴ", "イソギンチャク", "擬岩", "その他"];
  var targetProdHeaders = ["商品許可A", "商品許可B", "商品許可C"];

  // 既存の列インデックスを調査
  var colMap = {};
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).trim();
    colMap[h] = i + 1;
  }

  // 旧仕様「カテゴリ許可」「商品許可」の単一列が存在し、新チェックボックス列が未作成の場合は移行
  var hasOldCat = ("カテゴリ許可" in colMap);
  var hasOldProd = ("商品許可" in colMap || "許可商品" in colMap);
  var hasNewCat = ("魚" in colMap && "擬岩" in colMap);
  var hasNewProd = ("商品許可A" in colMap);

  // カテゴリチェックボックス列の追加
  for (var c = 0; c < targetCatHeaders.length; c++) {
    var catName = targetCatHeaders[c];
    if (!(catName in colMap)) {
      var newCol = cSheet.getLastColumn() + 1;
      cSheet.getRange(1, newCol).setValue(catName)
        .setBackground("#005bac").setFontColor("#ffffff").setFontWeight("bold");
      cSheet.setColumnWidth(newCol, 75);
      var maxRows = Math.max(cSheet.getMaxRows() - 1, 1);
      cSheet.getRange(2, newCol, maxRows, 1).insertCheckboxes();
      colMap[catName] = newCol;

      // 既存顧客について旧「カテゴリ許可」から値を引き継ぎ
      if (hasOldCat) {
        var oldCatCol = colMap["カテゴリ許可"];
        var numRows = cSheet.getLastRow() - 1;
        if (numRows > 0) {
          var oldVals = cSheet.getRange(2, oldCatCol, numRows, 1).getValues();
          for (var r = 0; r < numRows; r++) {
            var v = String(oldVals[r][0] || "").trim();
            // 空欄・全ての場合は全許可
            if (!v || v === "全て" || v === "全カテゴリ" || v.indexOf(catName) !== -1 || (catName === "シャコガイ" && v.indexOf("シャコ貝") !== -1)) {
              cSheet.getRange(r + 2, newCol).setValue(true);
            }
          }
        }
      }
    }
  }

  // 商品グループ許可チェックボックス列の追加 (A, B, C)
  for (var p = 0; p < targetProdHeaders.length; p++) {
    var prodName = targetProdHeaders[p];
    var groupLetter = prodName.replace("商品許可", "");
    if (!(prodName in colMap)) {
      var newCol = cSheet.getLastColumn() + 1;
      cSheet.getRange(1, newCol).setValue(prodName)
        .setBackground("#005bac").setFontColor("#ffffff").setFontWeight("bold");
      cSheet.setColumnWidth(newCol, 90);
      var maxRows = Math.max(cSheet.getMaxRows() - 1, 1);
      cSheet.getRange(2, newCol, maxRows, 1).insertCheckboxes();
      colMap[prodName] = newCol;

      // 旧「商品許可」から引き継ぎ
      if (hasOldProd) {
        var oldProdCol = colMap["商品許可"] || colMap["許可商品"];
        var numRows = cSheet.getLastRow() - 1;
        if (numRows > 0) {
          var oldVals = cSheet.getRange(2, oldProdCol, numRows, 1).getValues();
          for (var r = 0; r < numRows; r++) {
            var v = String(oldVals[r][0] || "").toUpperCase().trim();
            if (v.indexOf(groupLetter) !== -1) {
              cSheet.getRange(r + 2, newCol).setValue(true);
            }
          }
        }
      }
    }
  }

  // 旧単一列のデータバリデーションクリアまたは非表示（不要な古い入力規則によるエラー回避）
  if (hasOldCat) {
    try { cSheet.getRange(2, colMap["カテゴリ許可"], Math.max(cSheet.getMaxRows() - 1, 1), 1).clearDataValidations(); } catch(e) {}
  }
  if (hasOldProd) {
    try {
      var opCol = colMap["商品許可"] || colMap["許可商品"];
      cSheet.getRange(2, opCol, Math.max(cSheet.getMaxRows() - 1, 1), 1).clearDataValidations();
    } catch(e) {}
  }
}

// 「擬岩」シートの列構成を他シート（魚シート等）に合わせて完全に統一
// 統一列順: [商品ID, 商品名, サイズ, 単価, vol, 在庫, 画像, 売り切れ日時]
function ensureGiganiSheet(ss) {
  var sheet = ss.getSheetByName("擬岩");
  var standardHeaders = ["商品ID", "商品名", "サイズ", "単価", "vol", "在庫", "画像", "売り切れ日時"];

  if (!sheet) {
    sheet = ss.insertSheet("擬岩");
    sheet.getRange(1, 1, 1, standardHeaders.length).setValues([standardHeaders])
      .setBackground("#005bac")
      .setFontColor("#ffffff")
      .setFontWeight("bold");
    sheet.setFrozenRows(1);
    
    // 1. 通常商品（全顧客OK）: ID="擬岩_1"
    // 2. A許可限定商品: ID="A_001"
    var testData = [
      ["擬岩_1", "擬岩", "Mサイズ", 3000, 500, 50, "", ""],
      ["A_001", "【限定】特選擬岩", "Lサイズ", 5000, 800, 20, ""]
    ];
    sheet.getRange(2, 1, testData.length, standardHeaders.length).setValues(testData);
  } else {
    // 既存シートのヘッダーを検査し、商品IDが先頭でない（旧順序）場合は並び替え・統一
    var curData = sheet.getDataRange().getValues();
    if (curData.length >= 1) {
      var curHeader = curData[0];
      var firstTitle = String(curHeader[0] || "").trim();
      if (firstTitle !== "商品ID" && firstTitle !== "ID" && firstTitle !== "id") {
        // 列インデックス特定
        var cId = -1, cName = -1, cSize = -1, cPrice = -1, cVol = -1, cStock = -1, cImg = -1, cSold = -1;
        for (var c = 0; c < curHeader.length; c++) {
          var h = String(curHeader[c]).trim();
          if (h.indexOf("ID") !== -1 || h.indexOf("id") !== -1) cId = c;
          else if (h.indexOf("名") !== -1 || h.indexOf("商品") !== -1) cName = c;
          else if (h.indexOf("サイズ") !== -1 || h.indexOf("規格") !== -1) cSize = c;
          else if (h.indexOf("価格") !== -1 || h.indexOf("単価") !== -1) cPrice = c;
          else if (h.indexOf("vol") !== -1 || h.indexOf("容量") !== -1) cVol = c;
          else if (h.indexOf("在庫") !== -1) cStock = c;
          else if (h.indexOf("画像") !== -1 || h.indexOf("img") !== -1) cImg = c;
          else if (h.indexOf("売り切れ日時") !== -1) cSold = c;
        }

        var newRows = [];
        newRows.push(standardHeaders);
        for (var r = 1; r < curData.length; r++) {
          var row = curData[r];
          var pId = (cId !== -1) ? row[cId] : "擬岩_" + r;
          var pName = (cName !== -1) ? row[cName] : "";
          if (!pName) continue;
          newRows.push([
            pId,
            pName,
            (cSize !== -1) ? row[cSize] : "",
            (cPrice !== -1) ? row[cPrice] : "",
            (cVol !== -1) ? row[cVol] : "",
            (cStock !== -1) ? row[cStock] : "",
            (cImg !== -1) ? row[cImg] : "",
            (cSold !== -1) ? row[cSold] : ""
          ]);
        }
        sheet.clearContents();
        sheet.getRange(1, 1, newRows.length, standardHeaders.length).setValues(newRows);
        sheet.getRange(1, 1, 1, standardHeaders.length)
          .setBackground("#005bac").setFontColor("#ffffff").setFontWeight("bold");
        sheet.setFrozenRows(1);
      }
    }

    // A_001テストデータが存在するか確認、なければ追記
    var updatedData = sheet.getDataRange().getValues();
    var hasA = false;
    for (var r = 1; r < updatedData.length; r++) {
      if (String(updatedData[r][0] || "").indexOf("A_") === 0 || String(updatedData[r][1] || "").indexOf("【限定】") !== -1) {
        hasA = true;
        break;
      }
    }
    if (!hasA) {
      sheet.appendRow(["A_001", "【限定】特選擬岩", "Lサイズ", 5000, 800, 20, "", ""]);
    }
  }
}

function doGet(e) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var reqUserId = (e && e.parameter && e.parameter.userId) ? String(e.parameter.userId).trim() : "";
    var reqPhone = (e && e.parameter && e.parameter.phone) ? String(e.parameter.phone).trim() : "";
    var action = (e && e.parameter && e.parameter.action) ? String(e.parameter.action).trim() : "";
    var customerInfo = null;

    // メンテナンス処理の定期実行（毎リクエストではなく、キャッシュ等で1時間以上経過時または明示的sync時のみ実行）
    var cache = CacheService.getScriptCache();
    var lastSync = cache.get("LAST_SCHEMA_SYNC");
    var forceSync = (e && e.parameter && (e.parameter.sync === "1" || e.parameter.maintenance === "1"));
    if (forceSync || !lastSync) {
      try {
        ensureCustomerMasterColumns(ss);
        ensureGiganiSheet(ss);
        syncApplicationStatus(ss);
        cache.put("LAST_SCHEMA_SYNC", "done", 3600); // 1時間キャッシュ
      } catch (syncErr) {
        console.warn("Schema maintenance warning: " + syncErr);
      }
    }

    // 【診断API】顧客マスタと申請シートの登録状況確認用
    if (action === "debugCustomerMaster") {
      var dbgCSheet = ss.getSheetByName("顧客マスタ");
      var dbgCData = dbgCSheet ? dbgCSheet.getDataRange().getValues() : [];
      var dbgAppSheet = ss.getSheetByName("顧客登録申請");
      var dbgAppData = dbgAppSheet ? dbgAppSheet.getDataRange().getValues() : [];
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        headers: dbgCData.length > 0 ? dbgCData[0] : [],
        customers: dbgCData.slice(1).map(function(r) {
          return {
            uid: String(r[1] || ""),
            name: String(r[2] || ""),
            phone: String(r[3] || ""),
            rate: r[0],
            rawRow: r
          };
        }),
        applications: dbgAppData.slice(1).map(function(r) {
          return {
            date: r[0],
            uid: String(r[1] || ""),
            name: String(r[2] || ""),
            phone: String(r[3] || ""),
            status: String(r[5] || "")
          };
        })
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 【ピッキングリスト取得API】出荷ピッキング画面用（現場専用ツールとしてダイレクト表示）
    if (action === "getPickingList") {

      var detailSheet = ss.getSheetByName("注文明細");
      if (!detailSheet) {
        return ContentService.createTextOutput(JSON.stringify({
          status: "success",
          orders: [],
          totalSummary: []
        })).setMimeType(ContentService.MimeType.JSON);
      }

      var dData = detailSheet.getDataRange().getValues();
      var orderMap = {};
      var totalItemMap = {};
      var showAll = (e && e.parameter && e.parameter.showAll === "1");

      for (var r = 1; r < dData.length; r++) {
        var row = dData[r];
        var isShipped = row[10] === true || String(row[10]).toUpperCase() === "TRUE";

        if (!isShipped || showAll) {
          var dateVal = row[0];
          var dateStr = (dateVal instanceof Date) ? Utilities.formatDate(dateVal, "JST", "yyyy/MM/dd HH:mm") : String(dateVal || "");
          var orderId = String(row[1] || "").trim();
          var custName = String(row[2] || "お客様").trim();
          var phone = String(row[3] || "").trim();
          var category = String(row[4] || "").trim();
          var prodName = String(row[5] || "").trim();
          var qty = Number(row[6]) || 0;
          var price = row[7];
          var subtotal = row[8];
          var pdfUrl = String(row[9] || "").trim();
          var ngDates = String(row[13] || "").trim();
          var notes = String(row[14] || "").trim();

          if (!orderId) continue;

          if (!orderMap[orderId]) {
            orderMap[orderId] = {
              orderId: orderId,
              orderDate: dateStr,
              customerName: custName,
              phone: phone,
              pdfUrl: pdfUrl,
              ngDates: ngDates,
              notes: notes,
              isShipped: isShipped,
              items: [],
              rowIndices: []
            };
          }

          if (!isShipped) {
            orderMap[orderId].isShipped = false;
          }

          orderMap[orderId].items.push({
            rowNumber: r + 1,
            category: category,
            name: prodName,
            qty: qty,
            price: price,
            isShipped: isShipped
          });
          orderMap[orderId].rowIndices.push(r + 1);

          if (!isShipped && prodName) {
            var catKey = category || "その他";
            if (!totalItemMap[catKey]) totalItemMap[catKey] = {};
            if (!totalItemMap[catKey][prodName]) {
              totalItemMap[catKey][prodName] = { name: prodName, category: catKey, totalQty: 0, orderCount: 0 };
            }
            totalItemMap[catKey][prodName].totalQty += qty;
            totalItemMap[catKey][prodName].orderCount += 1;
          }
        }
      }

      var ordersList = Object.keys(orderMap).map(function(k) { return orderMap[k]; });
      ordersList.sort(function(a, b) { return a.orderId.localeCompare(b.orderId); });

      var totalList = [];
      Object.keys(totalItemMap).forEach(function(cat) {
        var itemsInCat = Object.keys(totalItemMap[cat]).map(function(pName) {
          return totalItemMap[cat][pName];
        });
        itemsInCat.sort(function(a, b) { return b.totalQty - a.totalQty; });
        totalList.push({
          category: cat,
          items: itemsInCat
        });
      });

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        orders: ordersList,
        totalSummary: totalList
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 【在庫管理一覧取得API】社内在庫管理・商品登録画面用
    if (action === "getStockList") {
      var allItems = [];
      var catList = ["魚", "シャコガイ", "貝", "サンゴ", "イソギンチャク", "擬岩", "その他"];

      for (var ci = 0; ci < CATEGORY_DEFS.length; ci++) {
        var cDef = CATEGORY_DEFS[ci];
        var sh = getSheetForCategory(ss, cDef.label);
        if (!sh) continue;

        var sData = sh.getDataRange().getValues();
        if (sData.length < 2) continue;

        var sHead = sData[0];
        var colI = -1, colN = -1, colSz = -1, colP = -1, colV = -1, colS = -1, colIm = -1, colSd = -1;
        for (var sc = 0; sc < sHead.length; sc++) {
          var titleStr = String(sHead[sc]).trim();
          if (titleStr.indexOf("ID") !== -1 || titleStr.indexOf("id") !== -1) colI = sc;
          else if (titleStr.indexOf("名") !== -1 || titleStr.indexOf("商品") !== -1) colN = sc;
          else if (titleStr.indexOf("サイズ") !== -1 || titleStr.indexOf("規格") !== -1) colSz = sc;
          else if (titleStr.indexOf("価格") !== -1 || titleStr.indexOf("単価") !== -1) colP = sc;
          else if (titleStr.indexOf("vol") !== -1 || titleStr.indexOf("容量") !== -1) colV = sc;
          else if (titleStr.indexOf("在庫") !== -1) colS = sc;
          else if (titleStr.indexOf("画像") !== -1 || titleStr.indexOf("img") !== -1) colIm = sc;
          else if (titleStr.indexOf("売り切れ日時") !== -1) colSd = sc;
        }

        if (colN === -1) colN = 1;
        if (colP === -1) colP = 3;
        if (colS === -1) colS = 5;

        for (var sr = 1; sr < sData.length; sr++) {
          var sRow = sData[sr];
          var itName = String(sRow[colN] || "").trim();
          if (!itName) continue;

          var itId = (colI !== -1 && sRow[colI]) ? String(sRow[colI]).trim() : cDef.label + "_" + sr;
          var itPrice = sRow[colP];
          var itIsAsk = String(itPrice).trim().toLowerCase() === "ask" || itPrice === "";
          var itNumPrice = itIsAsk ? 0 : (Number(itPrice) || 0);

          allItems.push({
            category: cDef.label,
            rowNumber: sr + 1,
            id: itId,
            name: itName,
            size: (colSz !== -1) ? String(sRow[colSz] || "").trim() : "",
            price: itNumPrice,
            isAsk: itIsAsk,
            vol: (colV !== -1) ? (Number(sRow[colV]) || 0) : 0,
            stock: (colS !== -1) ? (parseInt(sRow[colS], 10) || 0) : 0,
            img: (colIm !== -1) ? String(sRow[colIm] || "").trim() : "",
            soldOutAt: (colSd !== -1 && sRow[colSd]) ? String(sRow[colSd]) : ""
          });
        }
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        categories: catList,
        items: allItems
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 【顧客情報参照API】採集リクエストフォームや注文フォームからの自動補完用
    if (action === "lookupCustomer" || action === "searchCustomer") {
      var queryPhone = (e.parameter.phone || "").replace(/[^0-9]/g, "");
      if (queryPhone.length === 10 && !queryPhone.startsWith("0")) queryPhone = "0" + queryPhone;
      var queryUid = (e.parameter.userId || "").trim().toLowerCase();
      var queryName = (e.parameter.name || "").trim().toLowerCase();

      var foundCustomer = null;
      var cSheetLookup = ss.getSheetByName("顧客マスタ");
      if (cSheetLookup) {
        var cDataLookup = cSheetLookup.getDataRange().getValues();
        if (cDataLookup.length > 1) {
          var hLookup = cDataLookup[0];
          var lUid = -1, lName = -1, lParamPhone = -1, lAddr = -1;
          for (var lh = 0; lh < hLookup.length; lh++) {
            var lTitle = String(hLookup[lh]).trim();
            if (lTitle.indexOf("LINE") !== -1 || lTitle.indexOf("UID") !== -1) lUid = lh;
            else if (lTitle.indexOf("店舗") !== -1 || lTitle.indexOf("氏名") !== -1) lName = lh;
            else if (lTitle.indexOf("電話") !== -1 || lTitle.indexOf("TEL") !== -1) lParamPhone = lh;
            else if (lTitle.indexOf("住所") !== -1) lAddr = lh;
          }
          if (lUid === -1) lUid = 1;
          if (lName === -1) lName = 2;
          if (lParamPhone === -1) lParamPhone = 3;

          for (var lr = 1; lr < cDataLookup.length; lr++) {
            var rowUid = String(cDataLookup[lr][lUid] || "").trim().toLowerCase();
            var rowPhone = String(cDataLookup[lr][lParamPhone] || "").replace(/[^0-9]/g, "");
            if (rowPhone.length === 10 && !rowPhone.startsWith("0")) rowPhone = "0" + rowPhone;
            var rowName = String(cDataLookup[lr][lName] || "").trim().toLowerCase();

            var isMatched = false;
            if (queryPhone && rowPhone && queryPhone === rowPhone) isMatched = true;
            else if (queryUid && rowUid && queryUid === rowUid) isMatched = true;
            else if (queryName && rowName && queryName === rowName) isMatched = true;

            if (isMatched) {
              var fullAddress = (lAddr !== -1) ? String(cDataLookup[lr][lAddr] || "").trim() : "";
              // 都道府県の抽出（個人情報保護のため詳細住所・番地は返却せず都道府県のみ）
              var prefMatch = fullAddress.match(/(東京都|北海道|(?:京都|大阪)府|.{2,3}県)/);
              var extractedPref = prefMatch ? prefMatch[0] : "";

              foundCustomer = {
                found: true,
                name: String(cDataLookup[lr][lName] || "").trim(),
                phone: String(cDataLookup[lr][lParamPhone] || "").trim(),
                address: "", // セキュリティ対策: 詳細住所の無制限露出を防止
                prefecture: extractedPref,
                lineUserId: "" // セキュリティ対策: UID漏洩を防止
              };
              break;
            }
          }
        }
      }

      return ContentService.createTextOutput(JSON.stringify(
        foundCustomer || { found: false }
      )).setMimeType(ContentService.MimeType.JSON);
    }

    // 顧客マスタから顧客情報を特定（動的ヘッダー判定で列順の変更に対応）
    var cSheet = ss.getSheetByName("顧客マスタ");
    if (cSheet) {
      var cData = cSheet.getDataRange().getValues();
      var reqRawPhone = reqPhone.replace(/[^0-9]/g, "");
      if (reqRawPhone.length === 10 && !reqRawPhone.startsWith("0")) reqRawPhone = "0" + reqRawPhone;

      // ヘッダー列インデックスの完全動的特定
      var colUid = -1;
      var colName = -1;
      var colPhone = -1;
      var colRate = -1;
      var colAddr = -1;
      var colCatMap = {}; // { "魚": colIdx, ... }
      var colProdMap = {}; // { "A": colIdx, "B": colIdx, "C": colIdx }

      var cHeaders = cData[0];
      for (var ch = 0; ch < cHeaders.length; ch++) {
        var chTitle = String(cHeaders[ch]).trim();
        var chTitleLower = chTitle.toLowerCase();
        if (chTitleLower.indexOf("line") !== -1 || chTitleLower.indexOf("uid") !== -1 || chTitleLower.indexOf("userid") !== -1 || chTitle.indexOf("ユーザーID") !== -1) {
          colUid = ch;
        } else if (chTitle.indexOf("店舗") !== -1 || chTitle.indexOf("氏名") !== -1 || chTitle.indexOf("顧客名") !== -1) {
          colName = ch;
        } else if (chTitle.indexOf("電話") !== -1 || chTitle.indexOf("TEL") !== -1) {
          colPhone = ch;
        } else if (chTitle.indexOf("掛率") !== -1 || chTitle.indexOf("掛け率") !== -1 || chTitle.indexOf("割引率") !== -1) {
          colRate = ch;
        } else if (chTitle.indexOf("住所") !== -1) {
          colAddr = ch;
        } else if (chTitle === "魚" || chTitle === "シャコガイ" || chTitle === "貝" || chTitle === "サンゴ" || chTitle === "イソギンチャク" || chTitle === "擬岩" || chTitle === "その他") {
          colCatMap[chTitle] = ch;
        } else if (chTitle === "商品許可A" || chTitle === "許可A") {
          colProdMap["A"] = ch;
        } else if (chTitle === "商品許可B" || chTitle === "許可B") {
          colProdMap["B"] = ch;
        } else if (chTitle === "商品許可C" || chTitle === "許可C") {
          colProdMap["C"] = ch;
        }
      }

      // ヘッダーで見つからなかった場合のフォールバック（新配置 B:UID, C:店舗名, D:電話番号）
      if (colUid === -1) colUid = 1; // B列
      if (colName === -1) colName = 2; // C列
      if (colPhone === -1) colPhone = 3; // D列
      if (colRate === -1) colRate = 0; // A列(または旧仕様)

      var cleanReqUserId = reqUserId.replace(/[\s\uFEFF\xA0]/g, "").toLowerCase();

      for (var i = 1; i < cData.length; i++) {
        var dbUserId = String(cData[i][colUid] || "").replace(/[\s\uFEFF\xA0]/g, "").toLowerCase();
        var matchUid = (dbUserId !== "" && cleanReqUserId !== "" && dbUserId === cleanReqUserId);

        // 列ズレ・誤配置対策: もしcolUidで一致しなかった場合、その行の全列からUIDを照合
        if (!matchUid && cleanReqUserId.length >= 8) {
          for (var cIdx = 0; cIdx < cData[i].length; cIdx++) {
            var cellVal = String(cData[i][cIdx] || "").replace(/[\s\uFEFF\xA0]/g, "").toLowerCase();
            if (cellVal === cleanReqUserId) {
              matchUid = true;
              dbUserId = cellVal;
              break;
            }
          }
        }

        if (matchUid) {
          var rateVal = 1.0;
          if (colRate !== -1 && cData[i][colRate] !== undefined) {
            rateVal = parseFloat(String(cData[i][colRate] || "1").trim());
            if (isNaN(rateVal) || rateVal <= 0) rateVal = 1.0;
          }

          // 1. 各カテゴリ許可チェックボックスの判定
          var allowedCategories = [];
          var catKeys = ["魚", "シャコガイ", "貝", "サンゴ", "イソギンチャク", "擬岩", "その他"];
          var hasAnyCategoryCheckbox = false;

          for (var ck = 0; ck < catKeys.length; ck++) {
            var kName = catKeys[ck];
            if (colCatMap[kName] !== undefined) {
              hasAnyCategoryCheckbox = true;
              var isChecked = cData[i][colCatMap[kName]] === true || String(cData[i][colCatMap[kName]]).toUpperCase() === "TRUE";
              if (isChecked) {
                allowedCategories.push(kName);
              }
            }
          }

          // カテゴリ列が未設定、またはチェックボックスが1つもTRUEでない場合は「全カテゴリ許可」とする
          if (!hasAnyCategoryCheckbox || allowedCategories.length === 0) {
            allowedCategories = ["魚", "シャコガイ", "貝", "サンゴ", "イソギンチャク", "擬岩", "その他"];
          }

          // 2. 商品グループ許可チェックボックス（A, B, C）の判定
          var allowedGroups = {
            A: (colProdMap["A"] !== undefined) ? (cData[i][colProdMap["A"]] === true || String(cData[i][colProdMap["A"]]).toUpperCase() === "TRUE") : false,
            B: (colProdMap["B"] !== undefined) ? (cData[i][colProdMap["B"]] === true || String(cData[i][colProdMap["B"]]).toUpperCase() === "TRUE") : false,
            C: (colProdMap["C"] !== undefined) ? (cData[i][colProdMap["C"]] === true || String(cData[i][colProdMap["C"]]).toUpperCase() === "TRUE") : false
          };

          customerInfo = {
            lineUserId: dbUserId,
            name: String(cData[i][colName] || "お客様").trim(),
            discountRate: rateVal,
            phone: String(cData[i][colPhone] || "").trim(),
            address: (colAddr !== -1) ? String(cData[i][colAddr] || "").trim() : "",
            allowedCategories: allowedCategories,
            allowedProductGroups: allowedGroups
          };
          break;
        }
      }
    }

    // 厳格なアクセス制御：未登録ユーザーには商品情報を一切返さない
    if (!customerInfo) {
      return ContentService.createTextOutput(JSON.stringify({
        isRegistered: false,
        customer: null,
        products: [],
        userHistory: { purchaseCounts: {}, lastOrderItems: [] }
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 登録済み顧客のみ商品データを読み込み（カテゴリ許可 & 商品グループ許可で厳格にフィルタリング）
    var products = [];
    var nowTime = new Date().getTime();

    for (var catIdx = 0; catIdx < CATEGORY_DEFS.length; catIdx++) {
      var catDef = CATEGORY_DEFS[catIdx];

      // 【カテゴリ認可チェック】顧客に許可されていないカテゴリはスキップ
      if (customerInfo.allowedCategories.indexOf(catDef.label) === -1) {
        continue;
      }

      var sheet = getSheetForCategory(ss, catDef.label);
      if (!sheet) continue;

      var sheetName = catDef.label; // 画面表示用カテゴリ名
      var data = sheet.getDataRange().getValues();
      if (data.length < 2) continue;

      var header = data[0];
      var colName = -1, colPrice = -1, colVol = -1, colStock = -1, colSize = -1, colImg = -1, colId = -1, colSoldOut = -1;

      for (var c = 0; c < header.length; c++) {
        var h = String(header[c]).trim();
        if (h.indexOf("名") !== -1 || h.indexOf("商品") !== -1) colName = c;
        else if (h.indexOf("価格") !== -1 || h.indexOf("単価") !== -1) colPrice = c;
        else if (h.indexOf("vol") !== -1 || h.indexOf("容量") !== -1) colVol = c;
        else if (h.indexOf("在庫") !== -1) colStock = c;
        else if (h.indexOf("サイズ") !== -1 || h.indexOf("規格") !== -1) colSize = c;
        else if (h.indexOf("画像") !== -1 || h.indexOf("img") !== -1) colImg = c;
        else if (h.indexOf("ID") !== -1 || h.indexOf("id") !== -1) colId = c;
        else if (h.indexOf("売り切れ日時") !== -1) colSoldOut = c;
      }

      if (colName === -1) colName = 1;
      if (colPrice === -1) colPrice = 3;
      if (colVol === -1) colVol = 4;
      if (colStock === -1) colStock = 5;

      for (var r = 1; r < data.length; r++) {
        var row = data[r];
        var pName = String(row[colName] || "").trim();
        if (!pName) continue;

        var pId = (colId !== -1 && row[colId]) ? String(row[colId]).trim() : sheetName + "_" + r;

        // 【商品グループ認可チェック】商品IDの先頭文字が A, B, C の場合は対応グループ許可が必要
        var firstChar = pId.charAt(0).toUpperCase();
        if (firstChar === "A" || firstChar === "B" || firstChar === "C") {
          // 該当グループの許可を持っていない顧客には表示しない
          if (!customerInfo.allowedProductGroups[firstChar]) {
            continue;
          }
        }
        // 頭にアルファベット(A, B, C)が付いていない商品は全顧客に表示OK

        var numStock = Number(row[colStock]) || 0;
        if (numStock <= 0 && colSoldOut !== -1 && row[colSoldOut]) {
          var soldOutTime = new Date(row[colSoldOut]).getTime();
          if (nowTime - soldOutTime > 2 * 60 * 60 * 1000) continue;
        }

        var priceVal = row[colPrice];
        var isAsk = String(priceVal).trim().toLowerCase() === "ask" || priceVal === "";
        var numPrice = isAsk ? 0 : (Number(priceVal) || 0);

        products.push({
          id: pId,
          name: pName,
          category: sheetName,
          price: numPrice,
          isAsk: isAsk,
          vol: Number(row[colVol]) || 0,
          stock: numStock,
          size: colSize !== -1 ? String(row[colSize] || "").trim() : "",
          img: colImg !== -1 ? String(row[colImg] || "").trim() : "",
          isNew: false
        });
      }
    }

    // 顧客の販売履歴集計（よく買う商品・前回の注文商品）
    var userHistory = {
      purchaseCounts: {},
      lastOrderItems: []
    };

    var detailSheet = ss.getSheetByName("注文明細");
    if (detailSheet) {
      var detailData = detailSheet.getDataRange().getValues();
      if (detailData.length > 1) {
        var custPhoneClean = (customerInfo.phone || "").replace(/[^0-9]/g, "");
        var custNameClean = (customerInfo.name || "").trim();

        var lastOrderId = "";
        var lastOrderItemsMap = {};

        // 最新行から過去に遡って集計
        for (var dr = detailData.length - 1; dr >= 1; dr--) {
          var rowCustName = String(detailData[dr][2] || "").trim();
          var rowPhone = String(detailData[dr][3] || "").replace(/[^0-9]/g, "");
          var isMatch = (custPhoneClean !== "" && rowPhone !== "" && custPhoneClean === rowPhone) ||
                        (custNameClean !== "" && rowCustName === custNameClean);

          if (isMatch) {
            var orderIdVal = String(detailData[dr][1] || "").trim();
            var prodName = String(detailData[dr][5] || "").trim();
            var qty = Number(detailData[dr][6]) || 1;

            if (prodName) {
              userHistory.purchaseCounts[prodName] = (userHistory.purchaseCounts[prodName] || 0) + qty;

              // 最新の注文IDに含まれる商品名を記録
              if (!lastOrderId && orderIdVal) {
                lastOrderId = orderIdVal;
              }
              if (lastOrderId && orderIdVal === lastOrderId) {
                lastOrderItemsMap[prodName] = true;
              }
            }
          }
        }
        userHistory.lastOrderItems = Object.keys(lastOrderItemsMap);
      }
    }

    return ContentService.createTextOutput(JSON.stringify({
      isRegistered: true,
      customer: customerInfo,
      products: products,
      userHistory: userHistory
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      isRegistered: false,
      customer: null,
      products: [],
      userHistory: { purchaseCounts: {}, lastOrderItems: [] },
      error: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "混雑しております。時間をおいて再度お試しください。"
    })).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var contents = JSON.parse(e.postData.contents);
    var action = contents.action || "order";

    // 顧客登録の申請処理（スプレッドシートへの申請ログバックアップ記録）
    if (action === "apply" || action === "register") {
      var customerReg = contents.customer || {};
      var ssReg = SpreadsheetApp.openById(SPREADSHEET_ID);
      var appSheet = ssReg.getSheetByName("顧客登録申請");

      if (!appSheet) {
        appSheet = ssReg.insertSheet("顧客登録申請");
        var appHeaders = ["申請日時", "LINE_UID", "店舗名/氏名", "電話番号", "住所", "ステータス"];
        appSheet.getRange(1, 1, 1, appHeaders.length).setValues([appHeaders])
          .setBackground("#005bac").setFontColor("#ffffff").setFontWeight("bold");
        appSheet.setFrozenRows(1);
      }

      var nowAppStr = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");
      appSheet.appendRow([
        nowAppStr,
        customerReg.lineUserId || "",
        customerReg.name || "",
        customerReg.phone || "",
        customerReg.address || "",
        "承認待ち"
      ]);

      // 既に顧客マスタに登録済みであれば即座に「承認済み」へ更新
      syncApplicationStatus(ssReg);

      // 【自動LINE通知】新規の顧客登録申請を管理者(店主)宛てに通知
      sendApplicationLineNotification(customerReg);

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "申請を受け付けました。管理者の承認をお待ちください。"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 生体リクエスト・採集要望の処理
    if (action === "request") {
      return handleSpecimenRequest(contents);
    }

    // 【発送ステータス更新API】ピッキング画面からタップで発送済みに切り替え
    if (action === "updateShipStatus" || action === "toggleShipped") {

      var ssDetail = SpreadsheetApp.openById(SPREADSHEET_ID);
      var dSheet = ssDetail.getSheetByName("注文明細");
      if (!dSheet) throw new Error("注文明細シートが見つかりません。");

      var targetOrderId = contents.orderId;
      var targetRowNumbers = contents.rowNumbers || [];
      var newStatus = (contents.isShipped !== undefined) ? !!contents.isShipped : true;

      if (targetRowNumbers.length > 0) {
        for (var ri = 0; ri < targetRowNumbers.length; ri++) {
          var rNum = targetRowNumbers[ri];
          if (rNum >= 2 && rNum <= dSheet.getLastRow()) {
            dSheet.getRange(rNum, 11).setValue(newStatus);
          }
        }
      } else if (targetOrderId) {
        var dVals = dSheet.getDataRange().getValues();
        for (var rowIdx = 1; rowIdx < dVals.length; rowIdx++) {
          if (String(dVals[rowIdx][1] || "").trim() === String(targetOrderId).trim()) {
            dSheet.getRange(rowIdx + 1, 11).setValue(newStatus);
          }
        }
      }

      SpreadsheetApp.flush();

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "発送ステータスを更新しました",
        isShipped: newStatus
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 【生体追加API】電話追加などでピッキング管理画面から手動で生体・品目を追加
    if (action === "addOrderItem") {
      var targetOrderId = String(contents.orderId || "").trim();
      var itemData = contents.item || {};
      var prodName = String(itemData.name || "").trim();
      var qty = parseInt(itemData.qty, 10) || 1;
      var category = String(itemData.category || "魚").trim();
      var isAsk = !!itemData.isAsk;
      var unitPrice = isAsk ? "ASK" : (Number(itemData.price) || 0);
      var subtotal = isAsk ? "ASK" : (Number(itemData.subtotal) || (unitPrice * qty));

      if (!prodName) throw new Error("商品名を入力してください。");
      if (qty <= 0) throw new Error("匹数(数量)は1以上を指定してください。");

      var ssMaster = SpreadsheetApp.openById(SPREADSHEET_ID);
      var dSheet = ssMaster.getSheetByName("注文明細");
      if (!dSheet) throw new Error("注文明細シートが見つかりません。");

      var nowTime = new Date();
      var dateFormatted = Utilities.formatDate(nowTime, "JST", "yyyy/MM/dd HH:mm:ss");

      var custName = "電話注文客";
      var custPhone = "";
      var pdfUrl = "";
      var editSsUrl = "";
      var ngDatesStr = "";
      var notesStr = "【電話追加】";
      var hyperlinkFormula = "";

      // 既存注文への追加の場合、注文情報を引き継ぐ
      if (targetOrderId && targetOrderId !== "NEW") {
        var dVals = dSheet.getDataRange().getValues();
        for (var rIdx = 1; rIdx < dVals.length; rIdx++) {
          if (String(dVals[rIdx][1] || "").trim() === targetOrderId) {
            custName = String(dVals[rIdx][2] || custName).trim();
            custPhone = String(dVals[rIdx][3] || "").trim();
            pdfUrl = String(dVals[rIdx][9] || "").trim();
            editSsUrl = String(dVals[rIdx][12] || "").trim();
            ngDatesStr = String(dVals[rIdx][13] || "").trim();
            notesStr = String(dVals[rIdx][14] || "").trim();
            if (notesStr && notesStr.indexOf("電話追加") === -1) {
              notesStr += " (電話追加あり)";
            }
            break;
          }
        }
      } else {
        var dateIdStr = Utilities.formatDate(nowTime, "JST", "yyyyMMdd_HHmmss");
        targetOrderId = "TEL-" + dateIdStr;
        if (contents.customerName) custName = String(contents.customerName).trim();
        if (contents.phone) custPhone = String(contents.phone).trim();
      }

      var shippingLineId = PropertiesService.getScriptProperties().getProperty("OFFICIAL_LINE_ID") || "@721pyqfn";
      var shippingMsg = "【発送完了のお知らせ】\n" + custName + " 様\n\nご注文いただいた商品を本日発送いたしました。\n基本的には明日の午前中到着予定となります。\n到着後のご準備をお願いいたします。" + (pdfUrl ? "\n\n📄納品書PDF:\n" + pdfUrl : "");
      var lineUrl = "https://line.me/R/oaMessage/" + shippingLineId + "/?" + encodeURIComponent(shippingMsg);
      hyperlinkFormula = "=HYPERLINK(\"" + lineUrl + "\", \"📤 発送LINEを送る\")";

      // 注文明細に行を追加
      var newRow = [
        dateFormatted,
        targetOrderId,
        custName,
        custPhone,
        category,
        prodName,
        qty,
        unitPrice,
        subtotal,
        pdfUrl,
        false, // 未発送
        hyperlinkFormula,
        editSsUrl,
        ngDatesStr,
        notesStr
      ];

      dSheet.appendRow(newRow);

      // 在庫の減算（シートに該当商品が存在する場合は在庫を減算。在庫0以下・マイナスも許容）
      try {
        var catSheet = getSheetForCategory(ssMaster, category);
        if (catSheet) {
          var cData = catSheet.getDataRange().getValues();
          if (cData.length > 1) {
            var cHead = cData[0];
            var colN = -1, colS = -1;
            for (var c = 0; c < cHead.length; c++) {
              var th = String(cHead[c]).trim();
              if (th.indexOf("名") !== -1 || th.indexOf("商品") !== -1) colN = c;
              else if (th.indexOf("在庫") !== -1) colS = c;
            }
            if (colN !== -1 && colS !== -1) {
              for (var rowI = 1; rowI < cData.length; rowI++) {
                if (String(cData[rowI][colN]).trim() === prodName) {
                  var curStock = parseInt(cData[rowI][colS], 10) || 0;
                  var updatedStock = curStock - qty;
                  catSheet.getRange(rowI + 1, colS + 1).setValue(updatedStock);
                  break;
                }
              }
            }
          }
        }
      } catch (stockErr) {
        console.warn("Stock update warning: " + stockErr);
      }

      SpreadsheetApp.flush();

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "生体「" + prodName + "」(" + qty + "匹) を追加しました",
        orderId: targetOrderId
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 【在庫・商品情報更新API】在庫管理画面からの増減・編集用
    if (action === "updateStock") {
      var category = String(contents.category || "").trim();
      var prodId = String(contents.id || "").trim();
      var rowNum = parseInt(contents.rowNumber, 10) || -1;
      var newStock = parseInt(contents.stock, 10);
      if (isNaN(newStock)) throw new Error("在庫数を指定してください。");

      var ssStock = SpreadsheetApp.openById(SPREADSHEET_ID);
      var targetSheet = getSheetForCategory(ssStock, category);
      if (!targetSheet) throw new Error("カテゴリシート「" + category + "」が見つかりません。");

      var tData = targetSheet.getDataRange().getValues();
      var tHead = tData[0];
      var colI = -1, colN = -1, colSz = -1, colP = -1, colV = -1, colS = -1, colIm = -1, colSd = -1;
      for (var c = 0; c < tHead.length; c++) {
        var hTitle = String(tHead[c]).trim();
        if (hTitle.indexOf("ID") !== -1 || hTitle.indexOf("id") !== -1) colI = c;
        else if (hTitle.indexOf("名") !== -1 || hTitle.indexOf("商品") !== -1) colN = c;
        else if (hTitle.indexOf("サイズ") !== -1 || hTitle.indexOf("規格") !== -1) colSz = c;
        else if (hTitle.indexOf("価格") !== -1 || hTitle.indexOf("単価") !== -1) colP = c;
        else if (hTitle.indexOf("vol") !== -1 || hTitle.indexOf("容量") !== -1) colV = c;
        else if (hTitle.indexOf("在庫") !== -1) colS = c;
        else if (hTitle.indexOf("画像") !== -1 || hTitle.indexOf("img") !== -1) colIm = c;
        else if (hTitle.indexOf("売り切れ日時") !== -1) colSd = c;
      }

      if (colS === -1) throw new Error("在庫列が見つかりません。");

      // 該当行の特定
      var foundRow = -1;
      if (rowNum >= 2 && rowNum <= tData.length) {
        if (colI !== -1 && String(tData[rowNum - 1][colI]).trim() === prodId) {
          foundRow = rowNum;
        } else if (String(tData[rowNum - 1][colN]).trim() === String(contents.name || "").trim()) {
          foundRow = rowNum;
        }
      }
      if (foundRow === -1) {
        for (var r = 1; r < tData.length; r++) {
          if ((colI !== -1 && String(tData[r][colI]).trim() === prodId) ||
              (colN !== -1 && String(tData[r][colN]).trim() === String(contents.name || "").trim())) {
            foundRow = r + 1;
            break;
          }
        }
      }

      if (foundRow === -1) throw new Error("対象の商品が見つかりません。");

      // 在庫数の更新
      targetSheet.getRange(foundRow, colS + 1).setValue(newStock);

      // 売り切れ日時列のメンテナンス
      if (colSd !== -1) {
        if (newStock <= 0) {
          var curSold = targetSheet.getRange(foundRow, colSd + 1).getValue();
          if (!curSold) {
            targetSheet.getRange(foundRow, colSd + 1).setValue(new Date());
          }
        } else {
          targetSheet.getRange(foundRow, colSd + 1).setValue("");
        }
      }

      // 単価、サイズ、vol、画像等の更新指定があれば反映
      if (contents.price !== undefined && colP !== -1) {
        targetSheet.getRange(foundRow, colP + 1).setValue(contents.isAsk ? "ASK" : contents.price);
      }
      if (contents.size !== undefined && colSz !== -1) {
        targetSheet.getRange(foundRow, colSz + 1).setValue(contents.size);
      }
      if (contents.vol !== undefined && colV !== -1) {
        targetSheet.getRange(foundRow, colV + 1).setValue(contents.vol);
      }
      if (contents.img !== undefined && colIm !== -1) {
        targetSheet.getRange(foundRow, colIm + 1).setValue(contents.img);
      }

      SpreadsheetApp.flush();

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "在庫を更新しました",
        newStock: newStock
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 【新規商品登録API】管理画面からの新規登録・画像Drive保存
    if (action === "createProduct") {
      var category = String(contents.category || "魚").trim();
      var prodName = String(contents.name || "").trim();
      var tier = String(contents.tier || "NONE").trim().toUpperCase();
      var size = String(contents.size || "").trim();
      var stock = parseInt(contents.stock, 10);
      if (isNaN(stock)) stock = 1;
      var isAsk = !!contents.isAsk;
      var price = isAsk ? "ASK" : (Number(contents.price) || 0);
      var vol = Number(contents.vol) || 100;

      if (!prodName) throw new Error("商品名を入力してください。");

      var ssProd = SpreadsheetApp.openById(SPREADSHEET_ID);
      var targetSheet = getSheetForCategory(ssProd, category);
      if (!targetSheet) {
        targetSheet = ssProd.insertSheet(category);
        var stdHeaders = ["商品ID", "商品名", "サイズ", "単価", "vol", "在庫", "画像", "売り切れ日時"];
        targetSheet.getRange(1, 1, 1, stdHeaders.length).setValues([stdHeaders])
          .setBackground("#005bac").setFontColor("#ffffff").setFontWeight("bold");
        targetSheet.setFrozenRows(1);
      }

      // 画像のDrive保存（指定がある場合・複数枚対応）
      var imageUrls = [];
      if (contents.images && Array.isArray(contents.images) && contents.images.length > 0) {
        for (var mi = 0; mi < contents.images.length; mi++) {
          var imgObj = contents.images[mi];
          if (imgObj && imgObj.base64) {
            try {
              var s = saveImageToDrive(imgObj.base64, imgObj.fileName || (prodName + "_" + (mi + 1) + ".jpg"));
              imageUrls.push(s.url);
            } catch (err) {
              console.error("Image upload error: " + err);
            }
          }
        }
      } else if (contents.imageBase64) {
        try {
          var saved = saveImageToDrive(contents.imageBase64, contents.imageFileName || (prodName + ".jpg"));
          imageUrls.push(saved.url);
        } catch (imgErr) {
          console.error("Image upload failed: " + imgErr);
        }
      } else if (contents.imageUrl) {
        imageUrls.push(String(contents.imageUrl).trim());
      }
      var imageUrl = imageUrls.join(",");

      // 商品IDの自動生成（頭にABCをつけるかどうかの指定に対応）
      var nowTime = new Date();
      var idDateStr = Utilities.formatDate(nowTime, "JST", "yyyyMMdd_HHmmss");
      var autoId = "";
      if (tier === "A" || tier === "B" || tier === "C") {
        autoId = tier + "_" + category + "_" + idDateStr;
      } else {
        autoId = category + "_" + idDateStr;
      }

      var rowData = [
        autoId,
        prodName,
        size,
        price,
        vol,
        stock,
        imageUrl,
        (stock <= 0) ? Utilities.formatDate(nowTime, "JST", "yyyy/MM/dd HH:mm:ss") : ""
      ];

      targetSheet.appendRow(rowData);
      SpreadsheetApp.flush();

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "商品「" + prodName + "」を登録しました",
        productId: autoId,
        imageUrl: imageUrl,
        imageUrls: imageUrls
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 【画像アップロードAPI】既存商品への画像添付・Drive保存（複数枚対応）
    if (action === "uploadImage") {
      var savedUrls = [];
      if (contents.images && Array.isArray(contents.images) && contents.images.length > 0) {
        for (var ui = 0; ui < contents.images.length; ui++) {
          var uObj = contents.images[ui];
          if (uObj && uObj.base64) {
            var s = saveImageToDrive(uObj.base64, uObj.fileName || ("fish_" + (ui + 1) + ".jpg"));
            savedUrls.push(s.url);
          }
        }
      } else if (contents.imageBase64) {
        var s = saveImageToDrive(contents.imageBase64, contents.imageFileName || "fish.jpg");
        savedUrls.push(s.url);
      } else {
        throw new Error("画像データがありません。");
      }

      var finalUrl = savedUrls.join(",");

      // 商品IDとカテゴリが指定されていればシートにも即時反映
      if (contents.id && contents.category) {
        try {
          var ssUp = SpreadsheetApp.openById(SPREADSHEET_ID);
          var shUp = getSheetForCategory(ssUp, contents.category);
          if (shUp) {
            var upData = shUp.getDataRange().getValues();
            var colId = -1, colImg = -1;
            for (var c = 0; c < upData[0].length; c++) {
              var h = String(upData[0][c]).trim();
              if (h.indexOf("ID") !== -1 || h.indexOf("id") !== -1) colId = c;
              else if (h.indexOf("画像") !== -1 || h.indexOf("img") !== -1) colImg = c;
            }
            if (colId !== -1 && colImg !== -1) {
              for (var r = 1; r < upData.length; r++) {
                if (String(upData[r][colId]).trim() === String(contents.id).trim()) {
                  var curImgVal = String(upData[r][colImg] || "").trim();
                  var newImgVal = "";
                  if (contents.mode === "append" && curImgVal) {
                    newImgVal = curImgVal + "," + finalUrl;
                  } else {
                    newImgVal = finalUrl;
                  }
                  shUp.getRange(r + 1, colImg + 1).setValue(newImgVal);
                  finalUrl = newImgVal;
                  break;
                }
              }
            }
          }
        } catch(e) {
          console.warn("Sheet image update warning: " + e);
        }
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "画像をGoogle Driveに保存しました",
        url: savedUrls[0] || "",
        urls: savedUrls,
        combinedUrl: finalUrl
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // --- ここから注文処理 ---
    var items = contents.items || [];
    var customer = contents.customer || {};
    var shippingFee = Number(contents.shippingFee) || 0;
    var ngDates = contents.ngDates || []; // 受取不可日
    var notes = contents.notes || ""; // 備考

    if (items.length === 0) throw new Error("注文データが空です。");

    var masterSS = SpreadsheetApp.openById(SPREADSHEET_ID);
    var templateSheet = masterSS.getSheetByName("納品書テンプレート");
    if (!templateSheet) {
      createDeliveryNoteTemplate();
      templateSheet = masterSS.getSheetByName("納品書テンプレート");
    }

    var now = new Date();
    var dateStr = Utilities.formatDate(now, "JST", "yyyyMMdd_HHmmss");
    var monthDayStr = Utilities.formatDate(now, "JST", "M月d日");
    var orderId = "ORD-" + dateStr;
    var customerName = customer.name || "お客様";

    var newSSName = monthDayStr + " " + customerName + " 納品書";
    var newSS = SpreadsheetApp.create(newSSName);
    var tempSpreadsheetId = newSS.getId();
    var copiedSheet = templateSheet.copyTo(newSS);
    copiedSheet.setName("納品書");

    var newSheets = newSS.getSheets();
    for (var ns = 0; ns < newSheets.length; ns++) {
      if (newSheets[ns].getName() !== "納品書") newSS.deleteSheet(newSheets[ns]);
    }

    copiedSheet.getRange("E2").setValue(now);
    copiedSheet.getRange("A4").setValue(customerName + " 様");

    var itemRows = [];
    var subtotalAmount = 0;
    var totalVol = 0;

    for (var idx = 0; idx < items.length; idx++) {
      var item = items[idx];
      var qty = parseInt(item.qty, 10);
      if (isNaN(qty) || qty <= 0) continue;
      item.qty = qty;

      var appliedPrice = Number(item.appliedPrice);
      if (isNaN(appliedPrice) || appliedPrice < 0) {
        appliedPrice = 0;
      }
      item.appliedPrice = appliedPrice;

      if (idx < 22) itemRows.push([String(item.name || "").trim(), item.qty, item.isAsk ? "ASK" : item.appliedPrice]);
      if (!item.isAsk) subtotalAmount += item.appliedPrice * item.qty;
      totalVol += (Number(item.vol) || 0) * item.qty;
    }

    if (itemRows.length > 0) copiedSheet.getRange(20, 2, itemRows.length, 3).setValues(itemRows);
    SpreadsheetApp.flush();

    var pdfUrl = generatePdfAndGetUrl(newSS, newSSName);
    var editSsUrl = saveSpreadsheetToFolder(tempSpreadsheetId, EDIT_FOLDER_NAME);

    updateProductStocks(masterSS, items, now);
    recordOrderDetailsAndSetupPicking(masterSS, {
      orderId: orderId,
      date: now,
      customer: customer,
      items: items,
      pdfUrl: pdfUrl,
      editSsUrl: editSsUrl,
      ngDates: ngDates,
      notes: notes
    });
    updateMonthlySales(masterSS, now, subtotalAmount, shippingFee);

    // 【自動LINE通知】管理者(店主)宛およびお客様宛にプッシュ通知を送信
    sendOrderLineNotifications({
      orderId: orderId,
      date: now,
      customer: customer,
      items: items,
      pdfUrl: pdfUrl,
      subtotal: subtotalAmount,
      shippingFee: shippingFee,
      ngDates: ngDates,
      notes: notes
    });

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      invoiceUrl: pdfUrl
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function saveSpreadsheetToFolder(fileId, folderName) {
  var file = DriveApp.getFileById(fileId);
  var folders = DriveApp.getFoldersByName(folderName);
  var folder;
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(folderName);
  }
  file.moveTo(folder);
  return file.getUrl();
}

function generatePdfAndGetUrl(spreadsheet, fileName) {
  var folder;
  var folders = DriveApp.getFoldersByName(PDF_FOLDER_NAME);
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(PDF_FOLDER_NAME);
  }

  var url = "https://docs.google.com/spreadsheets/d/" + spreadsheet.getId() + "/export?exportFormat=pdf&format=pdf&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false&fzr=false";
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, { headers: { "Authorization": "Bearer " + token } });
  var pdfFile = folder.createFile(response.getBlob().setName(fileName + ".pdf"));
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return pdfFile.getUrl();
}

function updateProductStocks(ss, items, nowTime) {
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var sheet = getSheetForCategory(ss, item.category);
    if (!sheet) continue;

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) continue;

    var header = data[0];
    var colName = -1, colStock = -1, colId = -1, colSoldOut = -1;
    for (var c = 0; c < header.length; c++) {
      var h = String(header[c]).trim();
      if (h.indexOf("名") !== -1 || h.indexOf("商品") !== -1) colName = c;
      else if (h.indexOf("在庫") !== -1) colStock = c;
      else if (h.indexOf("ID") !== -1 || h.indexOf("id") !== -1) colId = c;
      else if (h.indexOf("売り切れ日時") !== -1) colSoldOut = c;
    }
    if (colStock === -1) continue;

    if (colSoldOut === -1) {
      colSoldOut = header.length;
      sheet.getRange(1, colSoldOut + 1).setValue("売り切れ日時").setBackground("#ffff00").setFontWeight("bold");
    }

    for (var r = 1; r < data.length; r++) {
      if ((colId !== -1 && String(data[r][colId]).trim() === String(item.id).trim()) ||
          (colName !== -1 && String(data[r][colName]).trim() === String(item.name).trim())) {
        var currentStock = parseInt(data[r][colStock], 10) || 0;
        var newStock = Math.max(0, currentStock - item.qty);
        sheet.getRange(r + 1, colStock + 1).setValue(newStock);

        if (newStock === 0 && currentStock > 0) {
          sheet.getRange(r + 1, colSoldOut + 1).setValue(nowTime);
        }
        break;
      }
    }
  }
}

function recordOrderDetailsAndSetupPicking(ss, data) {
  var detailSheet = ss.getSheetByName("注文明細");

  var headers = [
    "注文日時", "注文番号", "店舗名", "電話番号", "商品カテゴリ", "商品名", "数量", "単価", "小計",
    "納品書PDF", "発送済チェック", "発送用LINEリンク", "編集用データ", "受取不可日", "備考"
  ];

  if (!detailSheet) {
    detailSheet = ss.insertSheet("注文明細");
    detailSheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground("#005bac").setFontColor("#ffffff").setFontWeight("bold");
    detailSheet.setFrozenRows(1);
    detailSheet.setColumnWidth(11, 80);
    detailSheet.setColumnWidth(12, 150);
    detailSheet.setColumnWidth(13, 100);
    detailSheet.setColumnWidth(14, 150);
    detailSheet.setColumnWidth(15, 150);
  } else {
    // 既存シートに受取不可日・備考列がない場合は追加
    var curHeaders = detailSheet.getRange(1, 1, 1, detailSheet.getLastColumn()).getValues()[0];
    if (curHeaders.length < headers.length) {
      detailSheet.getRange(1, curHeaders.length + 1, 1, headers.length - curHeaders.length)
        .setValues([headers.slice(curHeaders.length)])
        .setBackground("#005bac").setFontColor("#ffffff").setFontWeight("bold");
    }
  }

  var dateFormatted = Utilities.formatDate(data.date, "JST", "yyyy/MM/dd HH:mm:ss");
  var shippingMsg = "【発送完了のお知らせ】\n" + (data.customer.name || "お客様") + " 様\n\nご注文いただいた商品を本日発送いたしました。\n基本的には明日の午前中到着予定となります。\n到着後のご準備をお願いいたします。\n\n📄納品書PDF:\n" + data.pdfUrl;
  var shippingLineId = PropertiesService.getScriptProperties().getProperty("OFFICIAL_LINE_ID") || "@721pyqfn";
  var lineUrl = "https://line.me/R/oaMessage/" + shippingLineId + "/?" + encodeURIComponent(shippingMsg);
  var hyperlinkFormula = "=HYPERLINK(\"" + lineUrl + "\", \"📤 発送LINEを送る\")";

  var ngDatesStr = Array.isArray(data.ngDates) ? data.ngDates.join(", ") : (data.ngDates || "");
  var notesStr = data.notes || "";

  var rows = [];
  for (var i = 0; i < data.items.length; i++) {
    var item = data.items[i];
    rows.push([
      dateFormatted,
      data.orderId,
      data.customer.name || "",
      data.customer.phone || "",
      item.category || "",
      item.name,
      item.qty,
      item.isAsk ? "ASK" : item.appliedPrice,
      item.isAsk ? "ASK" : item.appliedPrice * item.qty,
      data.pdfUrl,
      false,
      hyperlinkFormula,
      data.editSsUrl,
      ngDatesStr,
      notesStr
    ]);
  }

  if (rows.length > 0) {
    var startRow = detailSheet.getLastRow() + 1;
    detailSheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
    detailSheet.getRange(startRow, 11, rows.length, 1).insertCheckboxes();
  }

  var pickStore = ss.getSheetByName("店舗別ピッキング");
  if (!pickStore) {
    pickStore = ss.insertSheet("店舗別ピッキング");
    pickStore.getRange("A1").setFormula("=QUERY('注文明細'!A:M, \"select A, C, E, F, G where K = false order by C, A\", 1)");
    pickStore.getRange("A1:E1").setBackground("#d81b60").setFontColor("#ffffff").setFontWeight("bold");
  }

  var pickTotal = ss.getSheetByName("合計ピッキング");
  if (!pickTotal) {
    pickTotal = ss.insertSheet("合計ピッキング");
    pickTotal.getRange("A1").setFormula("=QUERY('注文明細'!A:M, \"select F, sum(G) where K = false group by F label sum(G) '未発送分の必要合計数'\", 1)");
    pickTotal.getRange("A1:B1").setBackground("#27ae60").setFontColor("#ffffff").setFontWeight("bold");
  }
}

function updateMonthlySales(ss, date, subtotal, shippingFee) {
  var sheet = ss.getSheetByName("月別売上集計");
  if (!sheet) {
    sheet = ss.insertSheet("月別売上集計");
    var headers = ["年月", "注文件数", "商品売上合計(税抜)", "送料合計", "総合計金額(税込見込)", "最終更新日時"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground("#005bac").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  var monthKey = Utilities.formatDate(date, "JST", "yyyy年MM月");
  var dateFormatted = Utilities.formatDate(date, "JST", "yyyy/MM/dd HH:mm:ss");
  var data = sheet.getDataRange().getValues();
  var targetRow = -1;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === monthKey) {
      targetRow = i + 1;
      break;
    }
  }

  var totalWithTax = Math.floor((subtotal + shippingFee) * 1.1);

  if (targetRow !== -1) {
    var curCount = Number(sheet.getRange(targetRow, 2).getValue()) || 0;
    var curSubtotal = Number(sheet.getRange(targetRow, 3).getValue()) || 0;
    var curShipping = Number(sheet.getRange(targetRow, 4).getValue()) || 0;
    var curTotalWithTax = Number(sheet.getRange(targetRow, 5).getValue()) || 0;

    sheet.getRange(targetRow, 2).setValue(curCount + 1);
    sheet.getRange(targetRow, 3).setValue(curSubtotal + subtotal);
    sheet.getRange(targetRow, 4).setValue(curShipping + shippingFee);
    sheet.getRange(targetRow, 5).setValue(curTotalWithTax + totalWithTax);
    sheet.getRange(targetRow, 6).setValue(dateFormatted);
  } else {
    sheet.appendRow([monthKey, 1, subtotal, shippingFee, totalWithTax, dateFormatted]);
  }
}

function createDeliveryNoteTemplate() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("納品書テンプレート");
  if (sheet) return;

  sheet = ss.insertSheet("納品書テンプレート");
  sheet.setGridlines(true);

  var columnWidths = [40, 220, 60, 90, 100, 120];
  for (var i = 0; i < columnWidths.length; i++) {
    sheet.setColumnWidth(i + 1, columnWidths[i]);
  }

  sheet.getRange("A2:F2").merge().setValue("納　品　書").setFontSize(18).setFontWeight("bold").setHorizontalAlignment("center");
  var dateRange = sheet.getRange("E2:F2");
  dateRange.setNumberFormat("yyyy年m月d日");

  sheet.getRange("A4:C4").merge().setFontSize(14).setFontWeight("bold");
  sheet.getRange("A5:C5").merge().setValue("毎度ありがとうございます。下記の通り納品申し上げます。").setFontSize(9);

  var sellerInfo = [
    ["川畑水産"],
    ["〒900-0000 沖縄県那覇市港町"],
    ["TEL: 090-XXXX-XXXX"],
    ["登録番号: T1234567890123"]
  ];
  sheet.getRange(4, 5, sellerInfo.length, 1).setValues(sellerInfo).setFontSize(9);

  sheet.getRange("A7:D8").merge().setFormula("=\"合 計 金 額 : ¥ \" & TEXT(E15, \"#,##0\") & \" -\"").setFontSize(14).setFontWeight("bold").setBackground("#f0f4f8").setHorizontalAlignment("center").setVerticalAlignment("middle");

  var detailHeaders = [["No.", "品名・摘要", "数量", "単価", "金額", "備考"]];
  sheet.getRange(19, 1, 1, 6).setValues(detailHeaders).setBackground("#005bac").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");

  var detailRows = [];
  var formulas = [];
  for (var r = 0; r < 22; r++) {
    detailRows.push([r + 1, "", "", "", "", ""]);
    var rowNum = 20 + r;
    formulas.push(["=IF(OR(C" + rowNum + "=\"\", D" + rowNum + "=\"\"), \"\", IF(D" + rowNum + "=\"ASK\", \"ASK\", C" + rowNum + "*D" + rowNum + "))"]);
  }
  sheet.getRange(20, 1, 22, 6).setValues(detailRows);
  sheet.getRange(20, 5, 22, 1).setFormulas(formulas);

  var extraRows = [
    ["", "", "", "小計(税抜)", "=SUM(E20:E41)"],
    ["", "", "", "消費税(10%)", "=FLOOR(E42*0.1)"],
    ["", "", "", "送料", 0],
    ["", "", "", "合計", "=E42+E43+E44"]
  ];
  sheet.getRange(42, 2, 4, 3).setValues([
    ["", "", "小計(税抜)"],
    ["", "", "消費税(10%)"],
    ["", "", "送料"],
    ["", "", "合計"]
  ]).setFontWeight("bold");
  sheet.getRange(42, 5, 4, 1).setFormulas([
    ["=SUM(E20:E41)"],
    ["=FLOOR(E42*0.1)"],
    ["0"],
    ["=E42+E43+E44"]
  ]).setFontWeight("bold");
}

// -----------------------------------------------------------------------------
// 【LINE通知機能】Messaging API (Push Message) による自動送信
// -----------------------------------------------------------------------------
function pushLineMessage(toUserId, textMessage) {
  var token = getLineChannelAccessToken();
  if (!token) {
    console.error("【LINE送信エラー】LINE_CHANNEL_ACCESS_TOKEN が未設定です。スクリプトプロパティをご確認ください。");
    return { success: false, error: "ACCESS_TOKEN_NOT_SET" };
  }

  if (!toUserId || String(toUserId).trim() === "") {
    console.error("【LINE送信エラー】送信先ユーザーID (toUserId) が空です。");
    return { success: false, error: "TO_USER_ID_EMPTY" };
  }

  toUserId = String(toUserId).trim();

  // ユーザーIDの形式バリデーション（LINE UIDは通常 'U' または 'C', 'R' から始まる33文字前後の文字列）
  if (!/^[UCR][0-9a-f]{32}$/i.test(toUserId)) {
    console.warn("【LINE送信警告】送信先ID '" + toUserId + "' はLINEの個別UID（Uから始まる33文字の文字列）ではない可能性があります。公式LINEアカウントの@IDや表示名ではプッシュ送信できません。");
  }

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
        "Authorization": "Bearer " + token
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    var res = UrlFetchApp.fetch(url, options);
    var resCode = res.getResponseCode();
    var resBody = res.getContentText();

    if (resCode !== 200) {
      console.error("【LINE送信失敗】HTTP " + resCode + " to: " + toUserId + " Response: " + resBody);
      return { success: false, statusCode: resCode, response: resBody };
    }

    console.log("【LINE送信成功】to: " + toUserId + " (HTTP " + resCode + ")");
    return { success: true, statusCode: resCode };
  } catch (e) {
    console.error("【LINE送信例外】to: " + toUserId + " Exception: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

// 注文確定時の自動LINE通知（管理者宛 ＆ お客様宛）
function sendOrderLineNotifications(order) {
  try {
    var dateStr = Utilities.formatDate(order.date, "JST", "yyyy/MM/dd HH:mm");
    var custName = (order.customer && order.customer.name) ? order.customer.name : "お客様";
    var custPhone = (order.customer && order.customer.phone) ? order.customer.phone : "";
    var custUid = (order.customer && order.customer.lineUserId) ? order.customer.lineUserId : "";

    // 商品明細テキスト作成
    var itemListText = "";
    for (var i = 0; i < order.items.length; i++) {
      var it = order.items[i];
      var pStr = it.isAsk ? "ASK" : "¥" + Number(it.appliedPrice).toLocaleString();
      itemListText += "・" + it.name + " × " + it.qty + " (" + pStr + ")\n";
    }

    var ngText = "";
    if (Array.isArray(order.ngDates) && order.ngDates.length > 0) {
      ngText = "\n🚫受取不可日: " + order.ngDates.join(", ");
    } else if (order.ngDates) {
      ngText = "\n🚫受取不可日: " + order.ngDates;
    }

    var notesText = order.notes ? "\n📝備考: " + order.notes : "";
    var totalAmountTax = Math.floor(((Number(order.subtotal) || 0) + (Number(order.shippingFee) || 0)) * 1.1);

    // 1. お客様宛の確認メッセージ
    var userMsg = "【ご注文ありがとうございます】\n" +
      custName + " 様\n\n" +
      "下記の内容でご注文を承りました。\n\n" +
      "📅注文日時: " + dateStr + "\n" +
      "🔖注文番号: " + order.orderId + "\n\n" +
      "【ご注文商品】\n" + itemListText +
      "\n商品小計(税抜): ¥" + Number(order.subtotal || 0).toLocaleString() +
      "\n送料: " + (order.shippingFee > 0 ? "¥" + Number(order.shippingFee).toLocaleString() : "別途確認") +
      "\n総合計(税込見込): ¥" + totalAmountTax.toLocaleString() +
      ngText + notesText + "\n\n" +
      "📄納品書PDF:\n" + order.pdfUrl + "\n\n" +
      "商品の発送準備が整い次第、改めてご連絡いたします。";

    if (custUid) {
      pushLineMessage(custUid, userMsg);
    }

    // 2. 店主（管理者）宛の即時注文通知
    var adminTargetUid = getAdminLineUserId();
    if (adminTargetUid) {
      var adminMsg = "🔔【新着注文が入りました！】\n\n" +
        "店舗名: " + custName + " 様\n" +
        "TEL: " + custPhone + "\n" +
        "注文番号: " + order.orderId + "\n" +
        "注文日時: " + dateStr + "\n\n" +
        "【商品内容】\n" + itemListText +
        "\n合計見込: ¥" + totalAmountTax.toLocaleString() +
        ngText + notesText + "\n\n" +
        "📄納品書確認:\n" + order.pdfUrl;

      pushLineMessage(adminTargetUid, adminMsg);
    } else {
      console.warn("【注文通知スキップ】ADMIN_LINE_USER_ID が未設定のため、店主へのLINE注文通知はスキップされました。");
    }

  } catch (err) {
    console.error("sendOrderLineNotifications error: " + err.toString());
  }
}

// 顧客登録申請時の管理者通知
function sendApplicationLineNotification(customerReg) {
  var adminTargetUid = getAdminLineUserId();
  if (!adminTargetUid) {
    console.warn("【申請通知スキップ】ADMIN_LINE_USER_ID が未設定のため、店主へのLINE申請通知はスキップされました。");
    return;
  }

  try {
    var regMsg = "👤【新規の顧客登録申請がありました】\n\n" +
      "店舗名/氏名: " + (customerReg.name || "未入力") + "\n" +
      "電話番号: " + (customerReg.phone || "未入力") + "\n" +
      "LINE_UID: " + (customerReg.lineUserId || "未取得") + "\n" +
      "住所: " + (customerReg.address || "未入力") + "\n\n" +
      "※スプレッドシートの「顧客登録申請」シートを確認の上、顧客マスタへの反映を行ってください。";

    pushLineMessage(adminTargetUid, regMsg);
  } catch (e) {
    Logger.log("sendApplicationLineNotification error: " + e.toString());
  }
}

// -----------------------------------------------------------------------------
// 【スプレッドシート定期バックアップ機能】
// -----------------------------------------------------------------------------
// スプレッドシート全体のコピーを指定フォルダに保存（古いバックアップは自動クリーンアップ）
function backupSpreadsheet() {
  try {
    var originalFile = DriveApp.getFileById(SPREADSHEET_ID);
    var folders = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
    var backupFolder;
    if (folders.hasNext()) {
      backupFolder = folders.next();
    } else {
      backupFolder = DriveApp.createFolder(BACKUP_FOLDER_NAME);
    }

    var now = new Date();
    var timeStamp = Utilities.formatDate(now, "JST", "yyyyMMdd_HHmmss");
    var backupName = "【バックアップ】" + originalFile.getName() + "_" + timeStamp;

    // ファイルの複製
    var copiedFile = originalFile.makeCopy(backupName, backupFolder);
    Logger.log("Backup completed: " + copiedFile.getName());

    // 過去のバックアップファイル整理（直近10件を残して古いものを削除）
    var files = backupFolder.getFiles();
    var fileList = [];
    while (files.hasNext()) {
      var f = files.next();
      fileList.push({ file: f, date: f.getDateCreated().getTime() });
    }

    fileList.sort(function(a, b) { return b.date - a.date; }); // 新しい順
    var MAX_BACKUPS = 10;
    if (fileList.length > MAX_BACKUPS) {
      for (var i = MAX_BACKUPS; i < fileList.length; i++) {
        fileList[i].file.setTrashed(true);
        Logger.log("Old backup trashed: " + fileList[i].file.getName());
      }
    }

    return "バックアップが完了しました: " + backupName;
  } catch (err) {
    Logger.log("backupSpreadsheet error: " + err.toString());
    throw err;
  }
}

// バックアップを週1回（毎週月曜午前3時）定期実行するためのトリガー自動登録関数
function setupWeeklyBackupTrigger() {
  // 既存の同名トリガーを削除して重複を防止
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "backupSpreadsheet") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 毎週月曜日の深夜3時に実行するトリガーを新規登録
  ScriptApp.newTrigger("backupSpreadsheet")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(3)
    .create();

  Logger.log("毎週月曜午前3時のバックアップトリガーを設定しました。");
}

// -----------------------------------------------------------------------------
// 【生体リクエスト・採集要望受付処理】
// -----------------------------------------------------------------------------
function handleSpecimenRequest(data) {
  try {
    var ssTarget;
    // 独立したリクエスト専用スプレッドシートIDがあればそちらを使用、なければメインスプレッドシート
    var reqSsId = PropertiesService.getScriptProperties().getProperty("REQUEST_SPREADSHEET_ID");
    if (reqSsId && reqSsId !== "") {
      try {
        ssTarget = SpreadsheetApp.openById(reqSsId);
      } catch (e) {
        ssTarget = SpreadsheetApp.openById(SPREADSHEET_ID);
      }
    } else {
      ssTarget = SpreadsheetApp.openById(SPREADSHEET_ID);
    }

    var sheetName = "採集リクエスト一覧";
    var sheet = ssTarget.getSheetByName(sheetName);
    var headers = [
      "受付日時", "リクエストID", "お名前 / 店舗名", "電話番号", "メールアドレス",
      "LINE_UID", "お届け先都道府県", "希望生体一覧(明細)", "品目数", "希望納期",
      "合計目安予算", "注意事項同意", "備考・飼育環境", "対応ステータス", "管理者メモ"
    ];

    if (!sheet) {
      sheet = ssTarget.insertSheet(sheetName);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers])
        .setBackground("#005bac").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(8, 300); // 希望生体一覧(明細)
      sheet.setColumnWidth(13, 250); // 備考
    }

    var now = new Date();
    var dateStr = Utilities.formatDate(now, "JST", "yyyy/MM/dd HH:mm:ss");
    var reqId = "REQ-" + Utilities.formatDate(now, "JST", "yyyyMMdd-HHmmss");

    // 複数生体アイテムの整形
    var items = data.items || [];
    if (items.length === 0 && data.species) {
      // 単一指定時の後方互換
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

    // 管理者(店主)宛LINE通知
    var adminTargetUid = getAdminLineUserId();
    var adminLineResult = null;
    if (adminTargetUid) {
      var adminMsg = "🎣【新規 生体リクエスト・採集要望受付】\n" +
        "-------------------\n" +
        "受付番号: " + reqId + "\n" +
        "日時: " + dateStr + "\n" +
        "お名前: " + (data.name || "未入力") + " 様\n" +
        "TEL: " + (data.phone || "未入力") + "\n" +
        "お届け先: " + (data.prefecture || "未入力") + "\n" +
        "-------------------\n" +
        "【ご希望生体一覧 (" + items.length + "品目)】\n" +
        itemsSummaryText + "\n" +
        "-------------------\n" +
        "⏳希望納期: " + (data.deadline || "指定なし") + "\n" +
        (totalBudgetSummary ? "💰目安予算: " + totalBudgetSummary + "\n" : "") +
        "💡飼育環境・備考:\n" + (data.notes ? data.notes : "なし") + "\n\n" +
        "※スプレッドシートの「採集リクエスト一覧」をご確認ください。";

      adminLineResult = pushLineMessage(adminTargetUid, adminMsg);
    } else {
      console.warn("【管理者LINE通知スキップ】スクリプトプロパティ ADMIN_LINE_USER_ID が未設定のため、店主へのLINE通知は送信されませんでした。LINE通知を受け取るには、GASのプロジェクト設定で ADMIN_LINE_USER_ID に店主様の個別UID(Uから始まる文字列)を設定してください。");
    }

    // お客様宛LINE通知（UID取得時）
    var customerLineResult = null;
    if (data.lineUserId) {
      var userMsg = "【採集リクエストを承りました】\n" +
        (data.name || "お客様") + " 様\n\n" +
        "川畑水産への採集リクエストありがとうございます。\n" +
        "下記の内容で受け付けいたしました。\n\n" +
        "🔖受付番号: " + reqId + "\n\n" +
        "【ご希望生体 (" + items.length + "品目)】\n" +
        itemsSummaryText + "\n\n" +
        "⏳希望納期: " + (data.deadline || "時期不問") + "\n\n" +
        "当店の採集員が海況・天候を確認し、採集後の状態確認・トリートメント完了次第、改めてご連絡いたします。\n" +
        "※自然環境下での採集のため、海況によりお時間をいただく場合がございます。";

      customerLineResult = pushLineMessage(data.lineUserId, userMsg);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      requestId: reqId,
      message: "採集リクエストを受け付けました。",
      adminNotified: adminLineResult ? adminLineResult.success : false,
      adminNotifyDetail: adminLineResult || { skipped: !adminTargetUid, reason: "ADMIN_LINE_USER_ID_NOT_CONFIGURED" }
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log("handleSpecimenRequest error: " + err.toString());
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.message || err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// -----------------------------------------------------------------------------
// 【LINE通知 動作確認・診断用テスト関数】
// -----------------------------------------------------------------------------
// GASエディタ上でこの関数を選択して「実行」ボタンを押すと、
// プロパティの設定状況とLINE送信の成否が実行ログに一目で出力されます。
function testSendLineNotification() {
  console.log("=== LINE通知設定 診断開始 ===");

  var token = getLineChannelAccessToken();
  var adminUid = getAdminLineUserId();

  console.log("1. アクセストークン確認: " + (token ? ("設定あり (先頭: " + token.substring(0, 15) + "...)") : "❌ 未設定"));
  console.log("2. 管理者LINE UID確認: " + (adminUid ? ("設定あり (" + adminUid + ")") : "❌ 未設定 (ADMIN_LINE_USER_ID を設定してください)"));

  if (!token) {
    console.error("【診断結果】アクセストークンが取得できませんでした。スクリプトプロパティの LINE_CHANNEL_ACCESS_TOKEN をご確認ください。");
    return "ERROR: LINE_CHANNEL_ACCESS_TOKEN is missing";
  }

  if (!adminUid) {
    console.warn("【診断結果】管理者LINE UIDが設定されていません。GASの「プロジェクトの設定」＞「スクリプト プロパティ」に ADMIN_LINE_USER_ID （Uから始まる文字列）を設定してください。");
    return "WARNING: ADMIN_LINE_USER_ID is missing";
  }

  console.log("3. テストメッセージ送信試行中...");
  var testMsg = "🔔【川畑水産 テスト通知】\nLINE Messaging APIの疎通確認テストです。\n正常に通知が届いています！\n日時: " + Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");

  var result = pushLineMessage(adminUid, testMsg);
  console.log("4. 送信結果: " + JSON.stringify(result));
  console.log("=== LINE通知設定 診断終了 ===");

  return result;
}


