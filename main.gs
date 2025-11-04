function myFunction() {
  // 一覧にしたいフォルダの階層を指定してください。
  // 0の場合、スプレッドシートが存在するフォルダのみのファイル一覧が表示されます。
  const depth = 1;

  const spreadSheet = SpreadsheetApp.getActive();
  const id = spreadSheet.getId();
  folderRoot = DriveApp.getFileById(id).getParents().next();
  const sheet = spreadSheet.getSheetByName('シート1');

  const rowInit = 2;
  const rowLast = sheet.getLastRow();
  if (rowLast > 1) {
    sheet.getRange(rowInit, 1, rowLast - rowInit + 1, 100).clear();
  }

  digFolders(folderRoot, sheet, depth)
}

function digFolders(rootFolder, sheet, depth) {
  listFiles(rootFolder, sheet);

  if (depth > 0) {
    const folders = rootFolder.getFolders();
    while (folders.hasNext()) {
      const folder = folders.next();
      digFolders(folder, sheet, depth - 1)
    }
  }
}

function listFiles(folder, sheet) {
  const files = folder.getFiles();
  let row = sheet.getLastRow() + 1;

  while (files.hasNext()) {
    const file = files.next();
    const folderValue = '=HYPERLINK("' + folder.getUrl() + '","' + folder.getName() + '")';
    sheet.getRange(row, 1).setValue(folderValue);
    const fileValue = '=HYPERLINK("' + file.getUrl() + '","' + file.getName() + '")';
    sheet.getRange(row, 2).setValue(fileValue);
    sheet.getRange(row, 4).setValue(file.getLastUpdated());

    row = row + 1;
  }
}
