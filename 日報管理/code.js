const doGet = (e) => {
    const page = (e.parameter.p || "index");
    const htmlOutput = HtmlService.createTemplateFromFile(page);
    if (e.parameter.p == "report") {
      htmlOutput.name = e.parameter.name;
      htmlOutput.id = e.parameter.id;
      htmlOutput.date = e.parameter.date;
    } else if (e.parameter.p == "calendar") {
      htmlOutput.name = e.parameter.name;
      htmlOutput.id = e.parameter.id;
    }
    //html にタイトルとファビコンを設定.
    return htmlOutput.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width,initial-scale=1')
      .setTitle("鐵建日報")
  };

  // 公開URLを取得する関数
  function getScriptUrl() {
    let url = ScriptApp.getService().getUrl();
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

      } else if (sheetName == "作業コード") {

        sheetId = "1oPnBQkWDThG4Gp94B699iNHptBX8-3IewkncPXOHDfI";
        const sheet = SpreadsheetApp.openById(sheetId);
        const range = sheet.getDataRange();
        result = range.getValues();

      } else if (sheetName == "工事No") {

        sheetId = "1gxiEYkh9w4AAQCwIjTOHVF0JJCyuh_zxYqGPJvCdE4Y";
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

      if ( sheetName == "日報データ" ) {
        return [result,sheetId];
      } else {
        return result
      }
    } catch (e) {
      return "ERROR:" + e.message;
    }
  }

  function daily_report(data, sheetId) {
    try {
      var spreadsheet = SpreadsheetApp.openById(sheetId);
      var sheet = spreadsheet.getSheetByName("日報"); //シート名を明確に定義しておく。getRangeが使えなくなる.

      //データの初期化.

      //ページを開いて、時間が経ってから書き込みだと、もしかしたらデータの欠落があるかもしれない。
      //配列の末尾にadd,report,delete等のタグをつけ、必要な箇所に必要な物だけを書き込む形にしよう。
      //addではappendrowで、reportでは、配列の順を見て、上書き、deleteでは、社員idを-にするか、全ての行を-で埋める。
      //その場合、取得するときに社員IDが-の時では、データの読み込みをしないようにする。というか、社員ID一致の時の身書いてるからできるか。
      //deleteの場合は、全ての行を-で埋める事にする。背景色も変更しましょう。

      for (let i = 0; i < data.length; i++) {

        if ( data[i].length == 12 ) {
          if ( data[i][11] == "add" ){
            data[i].pop();

            //[書き込む行、始まる列、書き込む行数、書き込み終了列]指定して書き込み.
            sheet.getRange(sheet.getLastRow()+1,1,1,data[i].length).setValues([data[i]]).setNumberFormat("@");


          } else if ( data[i][11] == "edit") {
            data[i].pop();
            //[書き込む行、始まる列、書き込む行数、書き込み終了列]指定して上書き.
            sheet.getRange(i+2,1,1,data[i].length).setValues([data[i]]).setNumberFormat("@");;

          } else if ( data[i][11] == "delete" ) {
            data[i].pop();
            //[書き込む行、始まる列、書き込む行数、書き込み終了列]指定して削除し、背景色を赤に設定.
            sheet.getRange(i+2,1,1,data[i].length).setValues([data[i]]).setNumberFormat("@");;
          }
        }
      }

      //シートを日時順に並び替え.
      sheet.getRange(2, 1, data.length, 11).sort([{column: 11, ascending: true},{column: 1, ascending: false},]);

      deleteRow(sheetId);

      return true;
    } catch (e) {
      return e.message;
    }
  }

  //一時的に-埋めされた日報のデータを削除する関数.
  function deleteRow(sheetId) {
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    var sheet = spreadsheet.getSheetByName("日報");
    const range = sheet.getDataRange();
    result = range.getValues();

    for ( let i = 1 ; i < result.length ; ++i ) {
      if ( result[i][0] == "-" ) {
        sheet.deleteRow(i+1);
      }
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
    const nowMonth = nowDate.getMonth() +1;

    if ( nowYear == year && nowMonth == month ) {
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
        return [range.getValues(), selectedSheet.getId()]
      } else {
        return new Error("指定した年月の日報データシートが存在しません")
      }
    } catch (e) {
      return new ReferenceError(`日報データシートの読み込み失敗:\n${e.message}`)
    }
  }
