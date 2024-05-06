const doGet = (e) => {
    const page = (e.parameter.p || "index");
    const htmlOutput = HtmlService.createTemplateFromFile(page);
    //html にタイトルとファビコンを設定.
    return htmlOutput.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width,initial-scale=1')
      .setTitle("鐵建日報集計")
  };

  // 公開URLを取得する関数
  function getScriptUrl() {
    var url = ScriptApp.getService().getUrl();
    return url;
  }

  // スプレッドシートから情報を取得しjsに返す関数
  function getSheetValues(sheetName) {
    try {
      let result = "";
      let sheetId = "";
      if (sheetName == "社員情報") {
        sheetId = "1qVgmwUJkVXu3QOkg793WW1OaWa2MrtIGf7bh-7snZmk";
        const sheet = SpreadsheetApp.openById(sheetId);
        const range = sheet.getDataRange();
        result = range.getValues();

      } else if (sheetName == "日報データ") {

        sheetId = "1pJkWMN1UrB6Z9IUJfV-f-5SsD61FVAsoAKGYc8W6UO8";
        const sheet = SpreadsheetApp.openById(sheetId);
        const range = sheet.getDataRange();
        result = range.getValues();

      } else if (sheetName == "休憩時間") {

        sheetId = "1Sko1LBh-Ar7OGPyCc0MBORbOUkPw2aNZEk7c5f2vfYY";
        const sheet = SpreadsheetApp.openById(sheetId);
        const range = sheet.getDataRange();
        result = range.getValues();
      } else if (sheetName == "勤務形態") {

        sheetId = "1Jj9fQl6gUDY57sdqmhqibz1clIHRWfWIpBNDd2WwP_k";
        const sheet = SpreadsheetApp.openById(sheetId);
        const range = sheet.getDataRange();
        result = range.getValues();
      } else if (sheetName = "伝票番号") {

        sheetId = "1WOWSc-KHUF4VcPtY3mIaNMn30VRpeDovCTTiqJ8wc3w";
        const sheet = SpreadsheetApp.openById(sheetId);
        const range = sheet.getDataRange();
        result = range.getValues();
      }

      for (let i = 0; i < result.length; i++) {
        for (let k = 0; k < result[0].length; k++) {
          let tmp_val = result[i][k] + "";
          // date型を文字列型に強制的に変換させる
          if (tmp_val.indexOf("GMT+") != -1) {
            result[i][k] = result[i][k] + "";
          }
        }
      }

      for (let i = 0; i < result.length; i++) {
        for (let k = 0; k < result[0].length; k++) {
          let tmp_val = result[i][k] + "";
          // date型を文字列型に強制的に変換させる
          if (tmp_val.indexOf("GMT+") != -1) {
            result[i][k] = result[i][k] + "";
          }
        }
      }

      return result;
    } catch (e) {
      return "ERROR:" + e.message;
    }
  }

  // 年月（数値）をもとにシートを読み込み、シートのデータを返す
  function getSheetAtYearMonth(year, month) {
    try {
      if (typeof year != "number" || year < 0) {
        return new TypeError("yearの値が不正です")
      }
      if (typeof month != "number" || month <= 0 || 12 < month) {
        return new TypeError("monthの値が不正です")
      }
    } catch (e) {
      return new TypeError(`渡された引数が不正です:\n${e.message}`)
    }

    const nowDate = new Date();
    const nowYear = nowDate.getFullYear();
    const nowMonth = nowDate.getMonth() + 1;

    if (nowYear == year && nowMonth == month) {
      return getSheetValues("日報データ");
    }

    const folderId = "1Ctxw7D1jXQ5xRXKhIVnv__BIeSUJ7ba5"

    const folder = DriveApp.getFolderById(folderId);
    const nameBase = "鐵建日報データシート"
    const fileName = `${nameBase}_${year}_${month}`
    const selectedSheets = folder.getFilesByName(fileName)

    try {
      if (selectedSheets.hasNext()) {
        const selectedSheet = selectedSheets.next()
        const sheet = SpreadsheetApp.openById(selectedSheet.getId())
        const range = sheet.getDataRange()
        return range.getValues()
      } else {
        return new Error("指定した年月の日報データシートが存在しません")
      }
    } catch (e) {
      return new ReferenceError(`日報データシートの読み込み失敗:\n${e.message}`)
    }
  }


  //伝票番号に1を足すだけのコード.
  function slipNumberChenge(number) {

    try {
      const spreadsheet = SpreadsheetApp.openById('1a4xqp5yDD-zMt0jPIhCwNdyPwC7Sk5aaKjgvePC4X9U');
      const sheet = spreadsheet.getSheetByName("伝票番号"); //シート名を明確に定義しておく。getRangeが使えなくなる.

      //[書き込む行、始まる列、書き込む行数、書き込み終了列]指定して書き込み.
      sheet.getRange(sheet.getLastRow(), 1, 1, 1).setValue(++number).setNumberFormat("@");
      return true;
    } catch (e) {
      return false;
    }

  }
