/**
 * スプレッドシートの編集（チェックボックスON）で起動
 */
function installedOnEditExtract(e) {
  const range = e.range;
  const sheet = range.getSheet();

  // B2セルがTRUEになったら実行
  if (range.getA1Notation() === "B2" && e.value === "TRUE") {
    const logCell = sheet.getRange("A5");
    logCell.setValue("⌛ 文字列抽出・ゲームデータ作成中...");

    SpreadsheetApp.flush();

    try {
      processImagesToSheet();
      logCell.setValue("✅ 処理が完了しました（" + new Date().toLocaleTimeString() + "）");
    } catch (err) {
      logCell.setValue("❌ エラー発生: " + err.toString());
    } finally {
      range.setValue(false); // チェックボックスを戻す
    }
  }
}

/**
 * フォルダ内の全画像をOCRしてスプレッドシートに転記し、ゲーム用JSを出力
 */
function processImagesToSheet() {
  const props = PropertiesService.getScriptProperties();
  const INPUT_FOLDER_ID = props.getProperty("INPUT_FOLDER_ID");
  const DONE_FOLDER_ID = props.getProperty("DONE_FOLDER_ID");
  const SPREADSHEET_ID = props.getProperty("SPREADSHEET_ID");
  const SHEET_NAME = props.getProperty("SHEET_NAME");

  if (!INPUT_FOLDER_ID || !DONE_FOLDER_ID) {
    throw new Error("プロパティにフォルダIDを設定してください。");
  }

  const inputFolder = DriveApp.getFolderById(INPUT_FOLDER_ID);
  const doneFolder = DriveApp.getFolderById(DONE_FOLDER_ID);

  let ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME) || ss.getActiveSheet();
  
  // 1. フォルダ内の画像ファイルを取得
  let fileList = [];
  const files = inputFolder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType().includes("image")) {
      fileList.push(file);
    }
  }
  fileList.sort((a, b) => a.getDateCreated() - b.getDateCreated());

  if (fileList.length === 0) return;

  let results = [];
  let allGameWords = []; // ゲーム用データを溜める配列
  let currentNo = sheet.getLastRow() > 0 ? sheet.getLastRow() - 1 : 0;

  // 2. 1枚ずつ順番にOCR処理
  for (let i = 0; i < fileList.length; i++) {
    const file = fileList[i];
    const fileName = file.getName();

    try {
      // OCR実行
      const extractedText = extractText(file.getId());

      // スプレッドシート用：改行をスペースに置換
      const cleanText = extractedText.trim().replace(/\n/g, " ");
      results.push([++currentNo, fileName, cleanText]);

      // ゲーム用データへのパース処理
      const converted = parseTextToGameObjects(extractedText);
      allGameWords = allGameWords.concat(converted);

      // 成功したら「処理済みフォルダ」へ移動
      doneFolder.addFile(file);
      inputFolder.removeFile(file);

    } catch (e) {
      console.error(`エラー発生 (${fileName}): ${e.toString()}`);
    }
  }

  // 3. スプレッドシートへの書き出し
  if (results.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, results.length, 3).setValues(results);
  }

  // 4. ゲーム用 JS ファイルの生成
  if (allGameWords.length > 0) {
    const jsonContent = JSON.stringify(allGameWords, null, 2); // 変数定義を付けない
    const jsonFileName = `words.json`; // 名前を固定（または words_日時.json）
    doneFolder.createFile(jsonFileName, jsonContent, MimeType.PLAIN_TEXT);
  }
}

/**
 * テキストを {text, hint} の形式に変換する
 */
function parseTextToGameObjects(rawText) {
  const lines = rawText.split('\n').map(l => l.trim()).filter(l => l !== "");
  const result = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    // 区切り文字（:や：や-）で分割を試みる
    if (line.includes(':') || line.includes('：') || line.includes('-')) {
      const parts = line.split(/[:：\-]/);
      result.push({
        text: parts[0].trim(),
        hint: parts.slice(1).join(':').trim() || "???"
      });
    } 
    // 区切りがない場合は2行1セットとして処理
    else if (i + 1 < lines.length) {
      result.push({
        text: line,
        hint: lines[i+1].trim()
      });
      i++; 
    }
  }
  return result;
}

/**
 * Drive API を使用したOCR処理
 */
function extractText(fileId) {
  const resource = {
    title: "temp_ocr_" + new Date().getTime(),
    mimeType: "application/vnd.google-apps.document"
  };
  
  const tempFile = Drive.Files.copy(resource, fileId, { ocr: true, ocrLanguage: "ja" });
  const doc = DocumentApp.openById(tempFile.id);
  const text = doc.getBody().getText();
  
  Drive.Files.remove(tempFile.id);
  return text;
}