function myFunction() {
  // 一覧にしたいフォルダの階層を指定してください。
  // 0の場合、スプレッドシートが存在するフォルダのみのファイル一覧が表示されます。
  const depth = 3;

  const spreadSheet = SpreadsheetApp.getActive();
  const id = spreadSheet.getId();
  const currentFolder = DriveApp.getFileById(id).getParents().next();
  
  // 一つ上の階層のフォルダを取得
  const parentFolders = currentFolder.getParents();
  if (parentFolders.hasNext()) {
    folderRoot = parentFolders.next();
  } else {
    // 親フォルダがない場合は現在のフォルダを使用
    folderRoot = currentFolder;
  }
  
  const sheet = spreadSheet.getSheetByName('シート1');

  const rowInit = 2;
  const rowLast = sheet.getLastRow();
  if (rowLast > 1) {
    sheet.getRange(rowInit, 1, rowLast - rowInit + 1, 100).clear();
  }

  digFolders(folderRoot, sheet, depth, 0)
}

function digFolders(rootFolder, sheet, depth, currentDepth) {
  listFiles(rootFolder, sheet, currentDepth);

  if (depth > 0) {
    const folders = rootFolder.getFolders();
    
    // フォルダを配列に変換してソート
    const folderArray = [];
    while (folders.hasNext()) {
      folderArray.push(folders.next());
    }
    folderArray.sort((a, b) => a.getName().localeCompare(b.getName()));
    
    // ソートされたフォルダを処理
    for (const folder of folderArray) {
      digFolders(folder, sheet, depth - 1, currentDepth + 1);
    }
  }
}

function listFiles(folder, sheet, currentDepth) {
  const files = folder.getFiles();
  
  // ファイルを配列に変換してソート
  const fileArray = [];
  while (files.hasNext()) {
    fileArray.push(files.next());
  }
  fileArray.sort((a, b) => a.getName().localeCompare(b.getName()));
  
  let row = sheet.getLastRow() + 1;

  // ファイルがある場合のみ処理
  if (fileArray.length > 0) {
    // ソートされたファイルを処理
    for (let i = 0; i < fileArray.length; i++) {
      const file = fileArray[i];
      
      // フォルダ名を階層に応じた列に配置（階層0はスキップ、1=A列、2=B列、3=C列）- リンクなし
      // 最初のファイルの行にのみフォルダ名を出力
      if (currentDepth > 0 && i === 0) {
        const folderColumn = currentDepth;
        sheet.getRange(row, folderColumn).setValue(folder.getName());
      }
      
      // ファイル名は常にD列（列4）に配置
      const fileValue = '=HYPERLINK("' + file.getUrl() + '","' + file.getName() + '")';
      sheet.getRange(row, 4).setValue(fileValue);
      
      row = row + 1;
    }
  } else if (currentDepth > 0) {
    // ファイルがない場合でも、階層1以降はフォルダ名だけを出力
    const folderColumn = currentDepth;
    sheet.getRange(row, folderColumn).setValue(folder.getName());
  }
}
