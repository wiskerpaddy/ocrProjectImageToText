/**
 * インストーラブルトリガー用の関数
 */
function installedOnEditMemcard(e) {
  const range = e.range;
  const sheet = range.getSheet();
  
  // A1（またはB2などチェックボックスのセル）がTRUEになったら実行
  if (range.getA1Notation() === "B3" && e.value === "TRUE") { // 画像に合わせてB2に変更
    const logCell = sheet.getRange("A5"); // ログ出力用セル
    SpreadsheetApp.getActiveSpreadsheet().toast("暗記カード作成処理が開始しました");
    logCell.setValue("⌛ 暗記カード作成処理中（" + new Date().toLocaleTimeString() + "）");
    try {
      processImagesToSheet();
      SpreadsheetApp.getActiveSpreadsheet().toast("暗記カード作成処理が完了しました");
      logCell.setValue("✅ 暗記カード作成完了しました（" + new Date().toLocaleTimeString() + "）");
    } catch (err) {
      // エラー内容をセルに書き出すとデバッグしやすいです
      sheet.getRange("A5").setValue("エラーログ: " + err.toString());
      logCell.setValue("❌ 暗記カード作成エラー: " + err.toString());
    } finally {
      range.setValue(false); // チェックを外す
    }
  }
}
/**
 * 定期実行エントリポイント
 * 2枚の画像を「表・裏」としてペア処理し、成功時のみ移動・CSV保存を行う
 */
function hourlyJob() {
  const props = PropertiesService.getScriptProperties();
  const INPUT_FOLDER_ID = props.getProperty("INPUT_FOLDER_ID");
  const DONE_FOLDER_ID = props.getProperty("DONE_FOLDER_ID");
  const CSV_FOLDER_ID = props.getProperty("CSV_FOLDER_ID");

  if (!INPUT_FOLDER_ID || !DONE_FOLDER_ID || !CSV_FOLDER_ID) {
    console.warn("フォルダIDが未設定です。設定を確認してください。");
    return;
  }

  const inputFolder = DriveApp.getFolderById(INPUT_FOLDER_ID);
  const doneFolder = DriveApp.getFolderById(DONE_FOLDER_ID);
  
  // ファイルを取得して作成日時順にソート（撮影順を担保）
  let fileList = [];
  const files = inputFolder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    // 画像ファイルのみを対象とする
    if (file.getMimeType().includes("image")) {
      fileList.push(file);
    }
  }
  fileList.sort((a, b) => a.getDateCreated() - b.getDateCreated());

  // 2枚ペアに満たない場合は終了
  if (fileList.length < 2) {
    console.log("処理待ちの画像がペア（2枚）に満たないため、待機します。");
    return;
  }

  let newCards = [];

  // 2枚ずつペアにして処理
  for (let i = 0; i < fileList.length - 1; i += 2) {
    const frontFile = fileList[i];
    const backFile = fileList[i + 1];

    try {
      console.log(`処理開始: ${frontFile.getName()} & ${backFile.getName()}`);
      
      // 1. OCR実行（ここでエラーが起きると catch へ飛ぶ）
      const frontText = runOcr(frontFile.getId());
      const backText = runOcr(backFile.getId());

      // 2. テキスト整形
      const front = frontText.split("\n")[0].substring(0, 50); // 1行目をタイトルに
      const back = backText.trim().replace(/\n/g, "<br>");    // 改行をHTMLに

      // 3. 全ての処理が成功したとみなして、リストに追加
      newCards.push([front, back]);

      // 4. 成功した場合のみ、ファイルを移動
      doneFolder.addFile(frontFile);
      inputFolder.removeFile(frontFile);
      doneFolder.addFile(backFile);
      inputFolder.removeFile(backFile);
      
      console.log(`処理成功・移動完了: ${frontFile.getName()}`);

    } catch (e) {
      // エラーが発生した場合はここに飛ぶ
      // 移動処理（removeFile）をスキップするため、画像は INPUT_FOLDER に残ります
      console.error(`エラーのためスキップしました（ファイルは残ります）: ${frontFile.getName()} - ${e.toString()}`);
    }
  }

  // 今回の実行で成功したカードがあればCSVに書き出し
  if (newCards.length > 0) {
    saveToCsv(CSV_FOLDER_ID, newCards);
    console.log(`${newCards.length} 件のペアをCSVに出力しました。`);
  }
}

/**
 * CSVとしてGoogleドライブに保存・追記する
 */
function saveToCsv(folderId, cardDataList) {
  const fileName = "anki_cards_import.csv";
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByName(fileName);
  
  let csvContent = "";
  let file;

  if (files.hasNext()) {
    file = files.next();
    csvContent = file.getBlob().getDataAsString("UTF-8");
    if (csvContent && !csvContent.endsWith("\n")) csvContent += "\r\n";
  }

  const newRows = cardDataList.map(card => {
    const front = card[0].replace(/"/g, '""');
    const back = card[1].replace(/"/g, '""');
    return `"${front}","${back}"`;
  }).join("\r\n");

  csvContent += newRows + "\r\n";

  if (file) {
    file.setContent(csvContent);
  } else {
    folder.createFile(fileName, csvContent, MimeType.PLAIN_TEXT);
  }
}

/**
 * Drive APIを使用したOCR処理
 */
function runOcr(fileId) {
  const resource = {
    title: "ocr-temp",
    mimeType: "application/vnd.google-apps.document"
  };
  // Files:copy が失敗しても呼び出し元（hourlyJob）の catch で捕捉される
  const doc = Drive.Files.copy(resource, fileId, { ocr: true, ocrLanguage: "ja" });
  const docFile = DocumentApp.openById(doc.id);
  const text = docFile.getBody().getText();
  
  // OCR用の一時ドキュメントは必ず削除
  Drive.Files.remove(doc.id);
  return text;
}

/**
 * テキスト分類と整形ロジック（表面をより具体的に修正）
 */
function classifyAndFormat(text) {
  const lines = text.split("\n").filter(x => x.trim().length > 0);
  const isQuestion = text.includes("とは") || text.includes("？");

  if (isQuestion) {
    const front = lines[0]; // 最初の1行を質問文にする
    const back = lines.slice(1).join("<br>");
    return { front, back };
  }

  // 要点抽出の場合も、最初の1行をタイトル的に使う
  const front = lines[0].substring(0, 20) + "..."; // 最初の20文字をタイトルに
  const back = lines.slice(0, 5).join("<br>");
  return { front: front, back: back };
}

function extractQuestion(text) {
  return text.split("\n")[0];
}

function extractAnswer(text) {
  return text.split("\n").slice(1).join("<br>");
}

function summarize(text) {
  const lines = text.split("\n").filter(x => x.trim().length > 0);
  return lines.slice(0, 5).join("<br>");
}