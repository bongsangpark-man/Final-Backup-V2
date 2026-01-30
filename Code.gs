/**
 * 파일명: Code.gs
 * 기능: 장부 생성/이월 자동화, 임대 관리 및 부가세 자료 생성 메뉴 연결 (수기 관리용)
 */

// ==========================================
// [1] 메뉴 생성
// ==========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // 1. [🏢 임대현황관리]
  ui.createMenu('🏢 임대현황관리')
    .addItem('임대 관리 시스템 열기', 'showRentalSidebar') 
    .addSeparator() 
    .addItem('🔒 시트 잠금 (수정 방지)', 'lockRentalSheet') 
    .addItem('🔓 시트 잠금 해제', 'unlockRentalSheet')     
    .addToUi();

  // 2. [📂 장부만들기]
  ui.createMenu('📂 장부만들기')
    .addItem('📅 금년 장부 생성하기(1월 2일 이후 생성)', 'createNextYearSheet')
    .addSeparator()
    .addItem('⚙️ 자동화(트리거) 생성하기', 'setupTriggersForNewYear')
    .addToUi();

  // 3. [📊 부가세 신고자료]
  ui.createMenu('📊 부가세 신고자료')
    .addItem('부가세 메뉴 열기', 'showVatSidebar') 
    .addToUi();
}

// ==========================================
// [2] 금년 장부 자동 생성 & 트리거 삭제 & 바로가기
// ==========================================

function createNextYearSheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentFileName = ss.getName();
  const currentFileId = ss.getId(); 
  
  // 연도 계산
  const yearMatch = currentFileName.match(/\d{4}/);
  const currentYear = yearMatch ? parseInt(yearMatch[0]) : 2025;
  const nextYear = currentYear + 1;

  const response = ui.alert(
    `📅 ${nextYear}년 장부 생성`, 
    `현재 파일(${currentYear}년)을 마감하고\n${nextYear}년 새 장부를 생성하시겠습니까?\n\n(완료 후 현재 파일의 만기 알림 메일은 중단됩니다)`, 
    ui.ButtonSet.YES_NO
  );
  
  if (response != ui.Button.YES) return;

  try {
    // 1. 파일 복제
    const newFileName = currentFileName.replace(String(currentYear), String(nextYear)) + " (새해 장부)";
    const newFile = DriveApp.getFileById(currentFileId).makeCopy(newFileName);
    const newSS = SpreadsheetApp.openById(newFile.getId());
    const newUrl = newSS.getUrl(); 
    
    // 2. [임대 현황표] A1 셀 메모에 설정값 심기
    const targetSheet = newSS.getSheetByName('임대 현황표');
    if (targetSheet) {
      const configData = {
        prevId: currentFileId, 
        year: String(nextYear) 
      };
      targetSheet.getRange('A1').setNote(JSON.stringify(configData));
    }

    // 3. 데이터 초기화 (수기 입력 칸 비우기)
    const sheetRent = newSS.getSheetByName('임대료 납부내역');
    if (sheetRent && sheetRent.getLastRow() > 1) {
       sheetRent.getRange(2, 6, sheetRent.getLastRow()-1, sheetRent.getLastColumn()-5).clearContent().setBackground(null).clearNote();
    }
    const sheetMaint = newSS.getSheetByName('관리비 납부내역');
    if (sheetMaint && sheetMaint.getLastRow() > 1) {
       sheetMaint.getRange(2, 3, sheetMaint.getLastRow()-1, sheetMaint.getLastColumn()-2).clearContent().setBackground(null).clearNote();
    }
    const sheetExitRent = newSS.getSheetByName('임대료 납부내역(퇴실)');
    if (sheetExitRent && sheetExitRent.getLastRow() > 1) {
      sheetExitRent.deleteRows(2, sheetExitRent.getLastRow() - 1);
    }
    const sheetExitMaint = newSS.getSheetByName('관리비 납부내역(퇴실)');
    if (sheetExitMaint && sheetExitMaint.getLastRow() > 1) {
      sheetExitMaint.deleteRows(2, sheetExitMaint.getLastRow() - 1);
    }

    // 4. 구 장부의 만기 알림 트리거 삭제
    const allTriggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;
    for (let i = 0; i < allTriggers.length; i++) {
      if (allTriggers[i].getHandlerFunction() === 'sendExtensionCheckEmails') {
        ScriptApp.deleteTrigger(allTriggers[i]);
        deletedCount++;
      }
    }

    // 5. 생성 완료 팝업
    const htmlOutput = HtmlService.createHtmlOutput(
      `<div style="font-family: sans-serif; padding: 10px; text-align: center;">` +
      `  <h3 style="margin-top: 0; color: #188038;">✅ 생성 완료!</h3>` +
      `  <p>새로운 ${nextYear}년 장부 파일이 생성되었습니다.</p>` +
      `  <p>현재 파일의 자동 알림 메일은 중단되었습니다.</p>` +
      `  <div style="margin-top: 20px;">` +
      `    <a href="${newUrl}" target="_blank" style="background-color: #1a73e8; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px; font-weight: bold; display: inline-block;">🚀 새 장부로 이동하기</a>` +
      `  </div>` +
      `  <p style="margin-top: 20px; font-size: 12px; color: #666;">* 새 파일 이동 후 [📂 장부만들기] > [⚙️ 자동화 생성]을 꼭 눌러주세요!</p>` +
      `</div>`
    ).setWidth(400).setHeight(350);

    ui.showModalDialog(htmlOutput, '장부 생성 결과');

  } catch (e) {
    ui.alert('오류 발생', e.toString(), ui.ButtonSet.OK);
  }
}

function setupTriggersForNewYear() {
  const ui = SpreadsheetApp.getUi();
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  try {
    // (A) 만기 알림 메일 (매일 오전 9시)
    ScriptApp.newTrigger('sendExtensionCheckEmails')
      .timeBased()
      .atHour(9)
      .everyDays(1)
      .create();

    // (B) 현황판 자동 업데이트 (수기 수정 시)
    ScriptApp.newTrigger('autoUpdateRent')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();

    ui.alert('✅ 자동화 설정 완료', '만기 알림 메일 및 현황판 업데이트 기능이 활성화되었습니다.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('설정 실패', e.toString(), ui.ButtonSet.OK);
  }
}

// ==========================================
// [3] 부가세 전용 사이드바 호출
// ==========================================
function showVatSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('VatSidebar')
    .setTitle('📊 부가세 신고 관리')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}
