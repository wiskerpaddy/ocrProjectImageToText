function installedOnEditExtract(e) {
  const range = e.range;
  const sheet = range.getSheet();

  // A1（またはB2などチェックボックスのセル）がTRUEになったら実行
  if (range.getA1Notation() === "B2" && e.value === "TRUE") { // 画像に合わせてB2に変更
    const logCell = sheet.getRange("A5"); // ログ出力用セル
    logCell.setValue("⌛ 文字列抽出作成処理中...");

    // GASの logCell.setValue() は、スクリプトの全処理が終わるか、
    // 処理が一時停止（待機）するまでスプレッドシート側の表示が更新されないため、
    // 以下関数を使って強制更新
    SpreadsheetApp.flush();

    try {
      processImagesToSheet();
      logCell.setValue("✅ 文字列抽出処理が完了しました（" + new Date().toLocaleTimeString() + "）");
    } catch (err) {
      // エラー内容をセルに書き出すとデバッグしやすいです
      logCell.setValue("❌ 文字列抽出処理がエラー: " + err.toString());
    } finally {
      range.setValue(false); // チェックを外す
    }
  }
}

/**
 * フォルダ内の全画像をOCRしてスプレッドシートに転記
 */
function processImagesToSheet() {
  const props = PropertiesService.getScriptProperties();
  const INPUT_FOLDER_ID = props.getProperty("INPUT_FOLDER_ID");
  const DONE_FOLDER_ID = props.getProperty("DONE_FOLDER_ID");
  const SPREADSHEET_ID = props.getProperty("SPREADSHEET_ID");
  const SHEET_NAME = props.getProperty("SHEET_NAME");

  if (!INPUT_FOLDER_ID || !DONE_FOLDER_ID) {
    console.warn("プロパティにフォルダIDを設定してください。");
    return;
  }

  const inputFolder = DriveApp.getFolderById(INPUT_FOLDER_ID);
  const doneFolder = DriveApp.getFolderById(DONE_FOLDER_ID);

// --- スプレッドシートの指定処理 ---
  let ss;

  if (SPREADSHEET_ID) {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  } else {
    // ID指定がない場合は、スクリプトが紐付いているシートを開く
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  let sheet = ss.getSheetByName(SHEET_NAME);
  
  // 指定したシート名が存在しない場合の予備処理
  if (!sheet) {
    console.warn(`シート「${SHEET_NAME}」が見つかりません。アクティブなシートを使用します。`);
    sheet = ss.getActiveSheet();
  }
  // ------------------------------
  
  // 1. フォルダ内の画像ファイルを取得（作成日時順にソート）
  let fileList = [];
  const files = inputFolder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType().includes("image")) {
      fileList.push(file);
    }
  }
  fileList.sort((a, b) => a.getDateCreated() - b.getDateCreated());

  if (fileList.length === 0) {
    console.log("処理対象の画像が見つかりませんでした。");
    return;
  }

  let results = [];
  
  // --- Noの初期値計算 ---
  let lastRow = sheet.getLastRow();
  
  // もし最終行が0（空）なら、ヘッダーもないので0からスタート
  // もし最終行が1以上なら、1行目はヘッダーとみなして「最終行 - 1」を現在の件数とする
  let currentNo = lastRow > 0 ? lastRow - 1 : 0;
  // 2. 1枚ずつ順番にOCR処理
  for (let i = 0; i < fileList.length; i++) {
    const file = fileList[i];
    const fileName = file.getName();

    try {
      console.log(`解析開始 (${i + 1}/${fileList.length}): ${fileName}`);
      
      // OCR実行
      const extractedText = extractText(file.getId());

      // [No, 画像名, 読み取り文字列] の形式で配列に追加
      // 改行をスペースに置換して、1つのセルに収まりやすく整形
      const cleanText = extractedText.trim().replace(/\n/g, " ");
      
      results.push([
        ++currentNo, 
        fileName,    
        cleanText    
      ]);

      // 成功したら「処理済みフォルダ」へ移動
      doneFolder.addFile(file);
      inputFolder.removeFile(file);

    } catch (e) {
      console.error(`エラー発生 (${fileName}): ${e.toString()}`);
    }
  }

  // 3. スプレッドシートの最終行に一括追記
  if (results.length > 0) {
    const targetRange = sheet.getRange(sheet.getLastRow() + 1, 1, results.length, 3);
    targetRange.setValues(results);
    console.log(`${results.length} 件の転記が完了しました。`);
  }
}

/**
 * Drive API を使用したOCR処理
 */
function extractText(fileId) {
  const resource = {
    title: "temp_ocr_" + new Date().getTime(),
    mimeType: "application/vnd.google-apps.document"
  };
  
  // Googleドキュメントとしてコピー（OCRを有効化）
  const tempFile = Drive.Files.copy(resource, fileId, { ocr: true, ocrLanguage: "ja" });
  const doc = DocumentApp.openById(tempFile.id);
  const text = doc.getBody().getText();
  
  // 一時ファイルはすぐに削除
  Drive.Files.remove(tempFile.id);
  
  return text;
}