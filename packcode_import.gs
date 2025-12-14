function onOpen(e) {
  var TARGET_ID = "ここに対象スプレッドシートのIDを入れる"; // 対象スプレッドシートのID
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (ss.getId() === TARGET_ID) {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('パックコードインポートツール')
      .addItem('インポート実行', 'runImportPackcode')
      .addToUi();
  }
}

// 画像をOCRしてスプシに記載する処理
function runImportPackcode() {
  var addedCount = runOcrProcess(); // OCR処理で追加した件数
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("シート1");
  var deletedCount = deleteInvalidFormatRows(sheet, 1); // バリデーションで削除した件数

  // ダイアログを表示
  SpreadsheetApp.getUi().alert(
    "インポート完了\n" +
    "追加件数: " + addedCount + " 行\n" +
    "削除件数: " + deletedCount + " 行"
  );
}

// OCR処理（追加件数を返す）
function runOcrProcess() {
  var FOLDER_ID = "ここに対象フォルダのIDを入れる"; // 対象となる画像フォルダのID
  var folder = DriveApp.getFolderById(FOLDER_ID);
  var files = folder.getFiles();
  var addedCount = 0;

  while (files.hasNext()) {
    var file = files.next();

    var resource = {
      title: "OCR結果",
      mimeType: MimeType.GOOGLE_DOCS
    };
    var docFile = Drive.Files.copy(resource, file.getId());
    var doc = DocumentApp.openById(docFile.id);
    var text = doc.getBody().getText();

    // 改行ごとに分割
    var lines = text.split(/\r?\n/);
    var SHEET_NAME = "シート1"; // シート名は任意に変更してよい
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

    // 行ごとのループ（若干バグあり）
    for (var i = 0; i < lines.length; i++) {
      if (lines[i].trim() !== "") {
        sheet.appendRow([lines[i], new Date(), "未登録"]);
        addedCount++;

        var lastRow = sheet.getLastRow();
        var rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(["未登録", "登録済"], true)
          .setAllowInvalid(false)
          .build();
        sheet.getRange(lastRow, 3).setDataValidation(rule);
      }
    }

    // 処理済みの画像ファイルは削除
    // Drive.Files.remove(file.getId());
    // OCR処理に使ったdoc削除
    Drive.Files.remove(docFile.id);
  }

  return addedCount;
}

// バリデーションチェック
function deleteInvalidFormatRows(sheet, column) {
  var lastRow = sheet.getLastRow();
  // XXXX-XXXX-XXXX形式のみ
  var regex = /^[A-Z]{4}-[A-Z]{4}-[A-Z]{4}$/;
  var deletedCount = 0;

  // チェックに弾かれた数
  for (var i = lastRow; i >= 1; i--) {
    var value = sheet.getRange(i, column).getValue();
    if (!regex.test(value)) {
      sheet.deleteRow(i);
      deletedCount++;
    }
  }
  return deletedCount;
}
