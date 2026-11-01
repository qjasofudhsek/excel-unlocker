// ================= [환경 설정] =================
var BOT_TOKEN = "9194295789:AAFozPPEnaF0LeLPUClUcKCTgxhrJXIg2SI";
var TEMP_FOLDER_ID = "1J1k8_NOypmMpeze_W9tN0WebB2TQsUx_";
var RENDER_URL = "https://excel-unlocker-agq5.onrender.com";

var TELEGRAM_API = "https://api.telegram.org/bot" + BOT_TOKEN;
var TELEGRAM_FILE_API = "https://api.telegram.org/file/bot" + BOT_TOKEN;
// ==============================================

function doPost(e) {
  var chatId = null;
  try {
    if (!e || !e.postData || !e.postData.contents) return;
    var data = JSON.parse(e.postData.contents);

    if (data.callback_query) {
      chatId = data.callback_query.message.chat.id;
      handleCallback(data.callback_query);
      return;
    }

    var msg = data.message;
    if (!msg) return;
    chatId = msg.chat.id;

    if (msg.document) {
      var fileName = msg.document.file_name || "이름없음.xlsx";
      if (fileName.match(/\.(xlsx|xls)$/i)) {
        saveFileDirect(msg.document.file_id, fileName);
        var count = getTodayFileCount();
        sendMsg(chatId, "📥 [" + fileName + "] 접수 완료 (오늘 대기 파일: " + count + "개)");
      }
      return;
    }

    if (msg.text) {
      var text = msg.text.trim().split("@")[0].trim();

      if (text === "핑" || text === "테스트" || text === "test" || text === "/핑") {
        sendMsg(chatId, "🏓 퐁! 봇이 정상 작동하고 있습니다.");
        return;
      }

      if (checkPassWait(chatId)) {
        PropertiesService.getScriptProperties().setProperty("EXCEL_PASSWORD", text);
        setPassWait(chatId, false);
        sendMsg(chatId, "🔑 공통 비밀번호가 [" + text + "] 로 변경되었습니다.");
        return;
      }

      if (text === "수합" || text === "시작" || text === "메뉴" || text === "/수합" || text === "/start") {
        sendMenu(chatId);
      } else if (text === "현황" || text === "대기" || text === "/현황") {
        var c = getTodayFileCount();
        var curPass = PropertiesService.getScriptProperties().getProperty("EXCEL_PASSWORD") || "1234";
        sendMsg(chatId, "📊 [현재 현황]\n• 오늘 대기 파일: " + c + "개\n• 적용 비밀번호: " + curPass);
      } else if (text === "초기화" || text === "/초기화") {
        clearFiles();
        sendMsg(chatId, "🗑 임시 보관 중인 모든 파일이 초기화되었습니다.");
      }
    }
  } catch (err) {
    if (chatId) {
      sendMsg(chatId, "⚠️ 오류 발생: " + err.toString());
    }
  }
}

function sendMenu(chatId) {
  var curPass = PropertiesService.getScriptProperties().getProperty("EXCEL_PASSWORD") || "1234";
  var count = getTodayFileCount();

  var keyboard = {
    inline_keyboard: [
      [{ text: "📅 오늘 파일 수합 (시트 선택)", callback_data: "ASK_SHEET" }],
      [
        { text: "🔑 비밀번호 변경 (" + curPass + ")", callback_data: "SET_PW" },
        { text: "🗑 전체 초기화", callback_data: "DO_CLEAR" }
      ]
    ]
  };

  UrlFetchApp.fetch(TELEGRAM_API + "/sendMessage", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      chat_id: chatId,
      text: "📂 [부서 엑셀 수합 관리자 메뉴]\n• 오늘 대기 파일: " + count + "개\n• 적용 비밀번호: " + curPass + "\n\n원하시는 작업을 선택해주세요:",
      reply_markup: keyboard
    })
  });
}

function handleCallback(cb) {
  var chatId = cb.message.chat.id;
  var action = cb.data;

  try {
    UrlFetchApp.fetch(TELEGRAM_API + "/answerCallbackQuery", {
      method: "post",
      payload: { callback_query_id: cb.id }
    });
  } catch (e) {}

  if (action === "ASK_SHEET") {
    showSheetSelectionMenu(chatId);
  } else if (action.indexOf("SELECT_SHEET:") === 0) {
    var selectedSheetName = action.replace("SELECT_SHEET:", "");
    sendMsg(chatId, "⏳ [" + selectedSheetName + "] 시트 병합을 시작합니다. (서식 및 수식 적용 중)");
    runMergeIndividual(chatId, selectedSheetName);
  } else if (action === "SET_PW") {
    setPassWait(chatId, true);
    sendMsg(chatId, "🔑 새로운 공통 비밀번호를 채팅창에 입력해주세요.");
  } else if (action === "DO_CLEAR") {
    clearFiles();
    sendMsg(chatId, "🗑 대기 파일이 모두 삭제되었습니다.");
  }
}

function showSheetSelectionMenu(chatId) {
  var targetFiles = getTodayFiles();
  if (targetFiles.length === 0) {
    sendMsg(chatId, "⚠️ 오늘 업로드된 대기 파일이 없습니다.");
    return;
  }

  sendMsg(chatId, "🔍 템플릿 파일의 시트 목록을 조회하고 있습니다...");

  var firstFile = targetFiles[0];
  var currentPassword = PropertiesService.getScriptProperties().getProperty("EXCEL_PASSWORD") || "1234";
  var base64Data = Utilities.base64Encode(firstFile.getBlob().getBytes());

  try {
    var response = UrlFetchApp.fetch(RENDER_URL + "/get_sheets", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ file_base64: base64Data, password: currentPassword }),
      muteHttpExceptions: true
    });

    var resJson = JSON.parse(response.getContentText());
    if (resJson.status !== "success" || !resJson.sheets || resJson.sheets.length === 0) {
      sendMsg(chatId, "⚠️ 시트 목록 조회 실패: " + (resJson.message || "암호 또는 파일 확인 필요"));
      return;
    }

    var sheetButtons = [];
    for (var i = 0; i < resJson.sheets.length; i++) {
      var sName = resJson.sheets[i];
      sheetButtons.push([{ text: "📋 " + (i + 1) + ". " + sName, callback_data: "SELECT_SHEET:" + sName }]);
    }

    UrlFetchApp.fetch(TELEGRAM_API + "/sendMessage", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        chat_id: chatId,
        text: "📑 수합할 시트를 선택해주세요:\n(첫 번째 기준 파일 양식 적용)",
        reply_markup: { inline_keyboard: sheetButtons }
      })
    });

  } catch (err) {
    sendMsg(chatId, "⚠️ 시트 목록 조회 중 오류: " + err.toString());
  }
}

// 개별 파일 경량 호출 & 양식/수식 완벽 빌드 함수
function runMergeIndividual(chatId, sheetName) {
  try {
    var targetFiles = getTodayFiles();
    if (targetFiles.length === 0) {
      sendMsg(chatId, "⚠️ 수합할 파일이 없습니다.");
      return;
    }

    var currentPassword = PropertiesService.getScriptProperties().getProperty("EXCEL_PASSWORD") || "1234";
    var masterTitleRows = [];
    var masterHeaders = [];
    var allDataRows = [];
    var failedFiles = [];

    // 1. 파일별 순차 경량 통신 (502 메모리 다운 원천 방지)
    for (var i = 0; i < targetFiles.length; i++) {
      var file = targetFiles[i];
      var fName = file.getName();
      try {
        var base64Data = Utilities.base64Encode(file.getBlob().getBytes());
        var res = UrlFetchApp.fetch(RENDER_URL + "/extract_data", {
          method: "post",
          contentType: "application/json",
          payload: JSON.stringify({
            file_base64: base64Data,
            sheet_name: sheetName,
            password: currentPassword
          }),
          muteHttpExceptions: true
        });

        if (res.getResponseCode() !== 200) {
          throw new Error("HTTP " + res.getResponseCode());
        }

        var json = JSON.parse(res.getContentText());
        if (json.status !== "success") {
          throw new Error(json.message || "추출 실패");
        }

        if (masterHeaders.length === 0 && json.headers && json.headers.length > 0) {
          masterHeaders = json.headers;
          masterTitleRows = json.title_rows || [];
        }

        if (json.data_rows && json.data_rows.length > 0) {
          allDataRows = allDataRows.concat(json.data_rows);
        }

        Utilities.sleep(300);
      } catch (e) {
        failedFiles.push(fName + " (" + e.message + ")");
      }
    }

    if (masterHeaders.length === 0 || allDataRows.length === 0) {
      var errMsg = "❌ 유효한 데이터를 추출하지 못했습니다.";
      if (failedFiles.length > 0) errMsg += "\n\n[실패 목록]\n" + failedFiles.join("\n");
      sendMsg(chatId, errMsg);
      return;
    }

    // 2. 구글 스프레드시트 엔진으로 완벽한 서식/수식 엑셀 파일 생성
    var timeStamp = Utilities.formatDate(new Date(), "GMT+9", "yyyyMMdd_HHmm");
    var resultFileName = "서식통합결과_" + timeStamp;
    var newSs = SpreadsheetApp.create(resultFileName);
    var sheet = newSs.getActiveSheet();
    sheet.setName(sheetName);

    var curRow = 1;
    var colCount = masterHeaders.length;

    // 제목/설명 작성
    for (var t = 0; t < masterTitleRows.length; t++) {
      var tRow = masterTitleRows[t];
      while (tRow.length < colCount) tRow.push("");
      sheet.getRange(curRow, 1, 1, colCount).setValues([tRow.slice(0, colCount)]);
      curRow++;
    }

    // 헤더 행 작성 및 서식 적용 (회색 배경, 굵게, 중앙 정렬)
    var headerRowNum = curRow;
    sheet.getRange(headerRowNum, 1, 1, colCount)
      .setValues([masterHeaders])
      .setBackground("#F2F2F2")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
    curRow++;

    // 순수 데이터 작성 및 순번 자동 부여
    var dataStartRow = curRow;
    var normalizedData = allDataRows.map(function(row, idx) {
      var newRow = row.slice(0, colCount);
      while (newRow.length < colCount) newRow.push("");
      // 1열 순번 재정렬
      if (newRow[0] && (!isNaN(newRow[0]) || String(newRow[0]).trim() !== "")) {
        newRow[0] = idx + 1;
      }
      return newRow;
    });

    sheet.getRange(dataStartRow, 1, normalizedData.length, colCount).setValues(normalizedData);
    curRow += normalizedData.length;
    var dataEndRow = curRow - 1;

    // 하단 '계' 합계행 작성 및 SUM 수식 자동 연결
    var footerRowNum = curRow;
    sheet.getRange(footerRowNum, 1, 1, colCount)
      .setBackground("#EAEAEA")
      .setFontWeight("bold");

    sheet.getRange(footerRowNum, 1).setValue("계").setHorizontalAlignment("center");

    // 숫자 열 감지하여 SUM 수식 적용
    for (var c = 1; c <= colCount; c++) {
      var sampleVal = normalizedData[0][c - 1];
      if (sampleVal !== "" && !isNaN(sampleVal) && c > 2) {
        var colLetter = getColumnLetter(c);
        sheet.getRange(footerRowNum, c)
          .setFormula("=SUM(" + colLetter + dataStartRow + ":" + colLetter + dataEndRow + ")")
          .setHorizontalAlignment("right");
      }
    }

    // 전체 테두리 적용
    sheet.getRange(headerRowNum, 1, (footerRowNum - headerRowNum + 1), colCount)
      .setBorder(true, true, true, true, true, true);

    SpreadsheetApp.flush();

    // 3. 엑셀 파일(.xlsx)로 내보내기 및 텔레그램 전송
    var fetchUrl = "https://docs.google.com/spreadsheets/d/" + newSs.getId() + "/export?format=xlsx";
    var token = ScriptApp.getOAuthToken();
    var excelBlob = UrlFetchApp.fetch(fetchUrl, {
      headers: { Authorization: "Bearer " + token }
    }).getBlob().setName(resultFileName + ".xlsx");

    var folder = DriveApp.getFolderById(TEMP_FOLDER_ID);
    folder.createFile(excelBlob);
    DriveApp.getFileById(newSs.getId()).setTrashed(true);

    var report = "✅ [서식 및 수식 보존] 병합 완료!\n• 수합 시트: " + sheetName + "\n• 취합 파일: " + (targetFiles.length - failedFiles.length) + "개\n• 총 수합 데이터: " + allDataRows.length + "행\n• 양식, 셀 색상, 테두리 및 SUM 수식이 자동 적용되었습니다.";
    if (failedFiles.length > 0) {
      report += "\n\n⚠️ [오류 파일: " + failedFiles.length + "개]\n" + failedFiles.join("\n");
    }

    UrlFetchApp.fetch(TELEGRAM_API + "/sendDocument", {
      method: "post",
      payload: {
        chat_id: String(chatId),
        document: excelBlob,
        caption: report
      }
    });

    targetFiles.forEach(function(f) { f.setTrashed(true); });

  } catch (err) {
    sendMsg(chatId, "⚠️ 병합 처리 중 오류 발생:\n" + err.toString());
  }
}

function getColumnLetter(columnNumber) {
  var dividend = columnNumber;
  var columnName = "";
  var modulo;
  while (dividend > 0) {
    modulo = (dividend - 1) % 26;
    columnName = String.fromCharCode(65 + modulo) + columnName;
    dividend = Math.floor((dividend - modulo) / 26);
  }
  return columnName;
}

function getTodayFiles() {
  var folder = DriveApp.getFolderById(TEMP_FOLDER_ID);
  var files = folder.getFiles();
  var targetFiles = [];
  var now = new Date();

  while (files.hasNext()) {
    var f = files.next();
    var fName = f.getName();
    if (fName.indexOf("서식통합결과_") === 0 || fName.indexOf("통합결과_") === 0) continue;
    if (!fName.match(/\.(xlsx|xls)$/i)) continue;

    var d = f.getDateCreated();
    if (d.getFullYear() === now.getFullYear() && d.getMonth() === now.getMonth() && d.getDate() === now.getDate()) {
      targetFiles.push(f);
    }
  }
  return targetFiles;
}

function getTodayFileCount() {
  return getTodayFiles().length;
}

function saveFileDirect(fileId, fileName) {
  var res = UrlFetchApp.fetch(TELEGRAM_API + "/getFile?file_id=" + fileId);
  var json = JSON.parse(res.getContentText());
  var path = json.result.file_path;
  var blob = UrlFetchApp.fetch(TELEGRAM_FILE_API + "/" + path).getBlob().setName(fileName);
  
  var folder = DriveApp.getFolderById(TEMP_FOLDER_ID);
  folder.createFile(blob);
}

function clearFiles() {
  var iter = DriveApp.getFolderById(TEMP_FOLDER_ID).getFiles();
  while (iter.hasNext()) { iter.next().setTrashed(true); }
}

function setPassWait(chatId, flag) {
  PropertiesService.getScriptProperties().setProperty("WAIT_" + chatId, flag ? "1" : "0");
}

function checkPassWait(chatId) {
  return PropertiesService.getScriptProperties().getProperty("WAIT_" + chatId) === "1";
}

function sendMsg(chatId, text) {
  UrlFetchApp.fetch(TELEGRAM_API + "/sendMessage", {
    method: "post",
    payload: { chat_id: String(chatId), text: text }
  });
}
