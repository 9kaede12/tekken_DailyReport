// GASのトリガーから月一で実行されるぜよ
// 月の初め（1日：深夜0時～1時）に実行
function createSheetAtMonthStart(e) {
  const sheetId = "1pJkWMN1UrB6Z9IUJfV-f-5SsD61FVAsoAKGYc8W6UO8";
  const folderId = "1KrAHJrjhlOv0U6AsyWRzgeAknmjXHes6"

  const originalSheet = SpreadsheetApp.openById(sheetId);

  const fileBase = DriveApp.getFileById(sheetId);
  const nameBase = "鐵建日報データシート"
  const folder = DriveApp.getFolderById(folderId);

  const nowDate = new Date(e.year, e.month - 1, e['day-of-month'], e.hour, e.minute, e.second)
  const fileName = `${nameBase}_${nowDate.getFullYear()}_${nowDate.getMonth() + 1}`
  // 変更の適用
  SpreadsheetApp.flush();
  fileBase.makeCopy(fileName, folder);

  // 元のスプレッドシートからデータを削除
  originalSheet.deleteRows(2, originalSheet.getLastRow() - 1);

  console.log(nowDate)
}
