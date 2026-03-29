/**
 * スプレッドシートの編集（チェックボックスON）を検知して起動
 */
function installedOnEditExtract(e) {
  const range = e.range;
  const sheet = range.getSheet();
  
  if (range.getA1Notation() === "B2" && e.value === "TRUE") {
    const writeLog = (msg) => {
      const now = new Date();
      const timeStr = Utilities.formatDate(now, "JST", "HH:mm:ss");
      sheet.appendRow([timeStr, "SYSTEM", msg]); 
      SpreadsheetApp.flush();
    };

    writeLog("🚀 PDF変換処理を開始しました");

    try {
      processImagesToPdf(writeLog);
      writeLog("✅ 全ての処理が正常に完了しました");
    } catch (err) {
      writeLog("❌ エラー発生: " + err.toString());
    } finally {
      range.setValue(false); 
    }
  }
}

/**
 * メインロジック：フォルダ内の画像をPDF化して整理
 */
function processImagesToPdf(log) {
  const props = PropertiesService.getScriptProperties();
  const inputFolder = DriveApp.getFolderById(props.getProperty("INPUT_FOLDER_ID"));
  const doneFolder = DriveApp.getFolderById(props.getProperty("DONE_FOLDER_ID"));

  const files = inputFolder.getFiles();
  let fileList = [];
  while (files.hasNext()) fileList.push(files.next());
  
  if (fileList.length === 0) {
    log("📂 対象ファイルが見つかりませんでした");
    return;
  }

  fileList.sort((a, b) => a.getDateCreated() - b.getDateCreated());

  for (let i = 0; i < fileList.length; i++) {
    const targetFile = fileList[i];
    const fileId = targetFile.getId();
    const fileName = targetFile.getName();

    try {
      log(`[${i + 1}/${fileList.length}] 処理中: ${fileName}`);
      
      // 1. PDF変換
      const pdfName = fileName.replace(/\.[^/.]+$/, "") + ".pdf";
      convertImageToPdf(fileId, pdfName, doneFolder);
      log(`   └ 📄 PDF作成完了`);

      // 2. 元画像の移動（整理）
      doneFolder.addFile(targetFile);
      inputFolder.removeFile(targetFile);
      log(`   └ 🚚 移動完了`);

    } catch (e) {
      log(`   └ ⚠️ 個別エラー (${fileName}): ${e.toString()}`);
    }
  }
}

/**
 * 【切り出し】画像をPDFに変換して指定フォルダに保存
 */
function convertImageToPdf(fileId, pdfName, destinationFolder) {
  const resource = {
    title: "temp_work",
    mimeType: "application/vnd.google-apps.document"
  };
  
  // 直接PDF化できないため、一度ドキュメント形式としてコピー
  const tempDoc = Drive.Files.copy(resource, fileId);
  
  // ドキュメントをPDFのBlobとして取得し、ファイル作成
  const pdfBlob = DriveApp.getFileById(tempDoc.id).getBlob().getAs('application/pdf');
  destinationFolder.createFile(pdfBlob).setName(pdfName);

  // 一時ドキュメントを即座に削除
  Drive.Files.remove(tempDoc.id);
}