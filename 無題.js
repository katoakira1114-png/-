var SPREADSHEET_ID = "1l5uJihoagHbMAwiCulpt9Zy9PIV6RRDneUYJ6QbiBbU";
var LINE_CHANNEL_ACCESS_TOKEN = "4z5y4BBsTTTEZgfmcaUF7cbc1T1k6qePWr/xKn3Yuk0E1pi4wxY3uPtSar9xW3F8FxIDpqTpma7lfn3HXhQ19LyTaDkNU6s779LkSLtPCUCZNb3nKg/tDSJXLNuADTfjDGoD8SQMw/CYDBdUAgkq3gdB04t89/1O/w1cDnyilFU=";
var PDF_FOLDER_NAME = "川畑水産_納品書PDF";
var EDIT_FOLDER_NAME = "川畑水産_納品書(編集用)";

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
    var customerInfo = null;

    // 「顧客マスタ」の列拡張（チェックボックス化）&「擬岩」シート統一
    ensureCustomerMasterColumns(ss);
    ensureGiganiSheet(ss);

    // 顧客登録申請とマスタの自動承認ステータス同期
    syncApplicationStatus(ss);

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
        if (chTitle.indexOf("LINE") !== -1 || chTitle.indexOf("UID") !== -1 || chTitle.indexOf("ユーザーID") !== -1) {
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

      for (var i = 1; i < cData.length; i++) {
        var dbUserId = String(cData[i][colUid] || "").trim();
        var dbPhone = String(cData[i][colPhone] || "").replace(/[^0-9]/g, "");
        if (dbPhone.length === 10 && !dbPhone.startsWith("0")) dbPhone = "0" + dbPhone;

        var matchUid = (dbUserId !== "" && reqUserId !== "" && dbUserId.toLowerCase() === reqUserId.toLowerCase());
        var matchPhone = (dbPhone !== "" && reqRawPhone !== "" && dbPhone === reqRawPhone);

        if (matchUid || matchPhone) {
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

          // 電話番号一致でLINE ID未登録だった場合は自動補完
          if (matchPhone && dbUserId === "" && reqUserId !== "") {
            cSheet.getRange(i + 1, colUid + 1).setValue(reqUserId);
            customerInfo.lineUserId = reqUserId;
          }
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

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "申請を受け付けました。管理者の承認をお待ちください。"
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
      if (idx < 22) itemRows.push([item.name, item.qty, item.isAsk ? "ASK" : item.appliedPrice]);
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
  var lineUrl = "https://line.me/R/oaMessage/@877qwvyc/?" + encodeURIComponent(shippingMsg);
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
