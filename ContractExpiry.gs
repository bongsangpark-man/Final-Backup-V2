/**
 * [주의] 이 스크립트는 두 가지 트리거 설정을 사용합니다.
 * 1. onEdit: 시트 수정 시 자동으로 실행 (별도 설정 불필요)
 * 2. sendExtensionCheckEmails: 평일 오전 9시~10시에 실행되도록 '트리거' 수동 추가 필요
 */

// ==========================================
// 1. 시트 데이터 갱신 (자동 실행)
// ==========================================

function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() === "임대 현황표") {
    main_UpdateContractExpiry();
  }
}

/**
 * 계약 만기 일정 메인 함수
 */
function main_UpdateContractExpiry() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const SOURCE_SHEET_NAME = "임대 현황표"; 
  const TARGET_SHEET_NAME = "계약 만기 일정";
  
  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!sourceSheet) return;

  // --- 0. 기존 입력 데이터 스마트 백업 ---
  let savedDataMap = {}; 
  let targetSheet = ss.getSheetByName(TARGET_SHEET_NAME);
  
  if (targetSheet) {
    const lastRow = targetSheet.getLastRow();
    if (lastRow > 2) { 
      const range = targetSheet.getRange(3, 1, lastRow - 2, 10);
      const existingValues = range.getValues();
      
      for (let i = 0; i < existingValues.length; i++) {
        const row = existingValues[i];
        
        const roomNo = String(row[2]).trim(); 
        const period = String(row[5]).trim(); 
        const tenant = String(row[6]).trim(); 
        const extStatus = row[8]; 
        const contStatus = row[9]; 
        
        const uniqueKey = roomNo + "_" + period + "_" + tenant;

        if (roomNo && (extStatus !== "" || contStatus !== "")) {
          savedDataMap[uniqueKey] = {
            ext: extStatus,
            cont: contStatus
          };
        }
      }
    }
  } else {
    targetSheet = ss.insertSheet(TARGET_SHEET_NAME);
  }

  // --- 1. 원본 데이터 읽기 ---
  const dataRange = sourceSheet.getDataRange();
  const values = dataRange.getValues();
  const processedData = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    const roomNo = String(row[0]).trim();       // A열
    const type = row[1];                        // B열
    const deposit = row[5];                     // F열
    let rent = row[6];                          // G열
    const periodRaw = String(row[9]).trim();    // J열
    const tenant = String(row[11]).trim();      // L열
    
    // ★ [수정] M열(12)에 주민번호가 추가되었으므로, 연락처는 N열(13)로 이동
    const contact = row[13];                    // N열
    
    if (!roomNo || !periodRaw || type === "공실") continue;

    // [수정 포인트] 날짜 파싱 함수 호출
    const dates = helper_ParseDatesUnique(periodRaw); 
    if (!dates) continue; 

    const expiryDate = dates.end;
    
    let isJeonse = false;
    const typeStr = String(type || "");
    const rentStr = String(rent || "").trim();
    if (typeStr.includes("전세") || rentStr === "" || rentStr === "-" || rent === 0) {
      isJeonse = true;
    }

    const monthsToSubtract = isJeonse ? 6 : 4;

    let checkDate = new Date(expiryDate);
    checkDate.setMonth(checkDate.getMonth() - monthsToSubtract);
    const checkDateStr = Utilities.formatDate(checkDate, Session.getScriptTimeZone(), "yy.MM");

    const currentUniqueKey = roomNo + "_" + periodRaw + "_" + tenant;
    let savedExt = "";
    let savedCont = "";
    
    if (savedDataMap[currentUniqueKey]) {
      savedExt = savedDataMap[currentUniqueKey].ext;
      savedCont = savedDataMap[currentUniqueKey].cont;
    }

    processedData.push({
      expiryDate: expiryDate,     
      checkDateStr: checkDateStr, 
      roomNo: roomNo,
      deposit: deposit,
      rent: rent,
      periodRaw: periodRaw,
      tenant: tenant,
      contact: contact,
      savedExt: savedExt,   
      savedCont: savedCont  
    });
  }

  // --- 2. 정렬 ---
  processedData.sort((a, b) => {
    if (a.checkDateStr < b.checkDateStr) return -1;
    if (a.checkDateStr > b.checkDateStr) return 1;
    return a.expiryDate - b.expiryDate;
  });

  // --- 3. 헤더 설정 ---
  targetSheet.getRange("A1:J1").merge().setValue("계약 만기 일정")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontSize(18).setFontWeight("bold").setBackground("white");
    
  const headers = ["만기 일자", "연장 여부 확인 일자", "호 수", "보증금", "월임대료", "임대기간", "계약자", "계약자연락처", "연장 유무", "계약 유무"];
  targetSheet.getRange("A2:J2").setValues([headers])
    .setHorizontalAlignment("center").setFontWeight("bold")
    .setBackground("#EFEFEF").setBorder(true, true, true, true, true, true);

  // --- 4. 데이터 출력 ---
  const lastRow = targetSheet.getLastRow();
  if (lastRow > 2) {
    targetSheet.getRange(3, 1, lastRow - 2, 10).clear({contentsOnly: true});
    targetSheet.getRange(3, 1, lastRow - 2, 10).setBorder(false, false, false, false, false, false);
  }

  if (processedData.length === 0) return;

  const outputValues = processedData.map(item => [
    item.expiryDate, item.checkDateStr, item.roomNo, item.deposit,      
    item.rent, item.periodRaw, item.tenant, item.contact,      
    item.savedExt, item.savedCont     
  ]);

  const rows = outputValues.length;
  const targetRange = targetSheet.getRange(3, 1, rows, 10);
  targetRange.setValues(outputValues);

  // --- 5. 서식 적용 ---
  targetSheet.getRange(3, 1, rows, 1).setNumberFormat("yy.MM.dd");
  targetSheet.getRange(3, 2, rows, 1).setHorizontalAlignment("center");
  targetSheet.getRange(3, 3, rows, 1).setHorizontalAlignment("center");
  targetSheet.getRange(3, 4, rows, 2).setNumberFormat("#,##0"); 
  targetSheet.getRange(3, 6, rows, 1).setHorizontalAlignment("center");
  targetSheet.getRange(3, 9, rows, 2).setHorizontalAlignment("center");
  
  targetRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  targetRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); 
  targetRange.setVerticalAlignment("middle");
}

// ==========================================
// 2. 이메일 자동 발송
// ==========================================

const MANAGER_EMAIL = "gahyeon@gahyeon.net"; 
const SENDER_NAME = "월디움상봉 계약 만기 알림";

function sendExtensionCheckEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName("계약 만기 일정");
  if (!targetSheet) return;

  const lastRow = targetSheet.getLastRow();
  if (lastRow < 3) return; 

  const dataRange = targetSheet.getRange(3, 1, lastRow - 2, 10);
  const values = dataRange.getValues();
  
  const today = new Date();
  
  // 주말(토=6, 일=0) 체크
  const dayOfWeek = today.getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    console.log("주말이라 메일을 발송하지 않습니다.");
    return;
  }

  let itemsToSend = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    
    const checkDateStr = String(row[1]); 
    const roomNo = String(row[2]);       
    const period = String(row[5]);       
    const tenant = String(row[6]);       
    const extStatus = String(row[8]);    

    if (extStatus !== "") continue;
    if (!helper_IsDatePassed(checkDateStr)) continue;

    itemsToSend.push({
      room: roomNo,
      tenant: tenant,
      period: period,
      checkDate: checkDateStr
    });
  }

  if (itemsToSend.length > 0) {
    
    let htmlBody = '<div style="font-family: Arial, sans-serif; color: #333;">';
    htmlBody += '<h2>월디움상봉 계약 만기 확인 리포트</h2>'; 
    htmlBody += `<p>현재 확인이 필요한 계약 건수는 총 <strong>${itemsToSend.length}건</strong>입니다.</p>`;
    htmlBody += '<p>아래 내역을 확인 후 <strong>\'계약 만기 일정\'</strong> 시트의 [연장 유무] 란에 입력해 주세요.</p>';
    
    htmlBody += '<table style="border-collapse: collapse; width: 100%; border: 1px solid #ddd; margin-top: 15px;">';
    htmlBody += '<tr style="background-color: #f2f2f2;">';
    htmlBody += '<th style="border: 1px solid #ddd; padding: 10px; text-align: left;">호수</th>';
    htmlBody += '<th style="border: 1px solid #ddd; padding: 10px; text-align: left;">계약자</th>';
    htmlBody += '<th style="border: 1px solid #ddd; padding: 10px; text-align: left;">임대기간</th>';
    htmlBody += '<th style="border: 1px solid #ddd; padding: 10px; text-align: left;">확인기준월</th>';
    htmlBody += '</tr>';

    itemsToSend.forEach(item => {
      htmlBody += '<tr>';
      htmlBody += `<td style="border: 1px solid #ddd; padding: 10px;"><strong>${item.room}호</strong></td>`;
      htmlBody += `<td style="border: 1px solid #ddd; padding: 10px;">${item.tenant}</td>`;
      htmlBody += `<td style="border: 1px solid #ddd; padding: 10px;">${item.period}</td>`;
      htmlBody += `<td style="border: 1px solid #ddd; padding: 10px; color: red;">${item.checkDate}</td>`;
      htmlBody += '</tr>';
    });

    htmlBody += '</table>';
    htmlBody += '<br><hr>';
    htmlBody += '<p style="font-size: 12px; color: #888;">* 본 메일은 시스템에서 자동으로 발송되었습니다.<br>';
    htmlBody += '* 시트에 조치 내용(O/X)을 입력하시면 해당 건은 내일 리포트에서 제외됩니다.</p>';
    htmlBody += '</div>';

    GmailApp.sendEmail(MANAGER_EMAIL, `[월디움상봉] 계약 만기 확인 요청 (${itemsToSend.length}건)`, "HTML을 지원하는 이메일 클라이언트를 사용해주세요.", {
      htmlBody: htmlBody,
      name: SENDER_NAME 
    });
    
    console.log(`✅ 총 ${itemsToSend.length}건 묶음 발송 완료.`);

  } else {
    console.log("📭 오늘 발송할 대상(미확인 건)이 없습니다.");
  }
}

// ==========================================
// 3. 공통 헬퍼 함수 (날짜 포맷 개선됨)
// ==========================================

function helper_ParseDatesUnique(periodStr) {
  try {
    if (!periodStr) return null;
    const str = periodStr.toString();
    const parts = str.split('~');
    if (parts.length < 2) return null;

    // [중요 수정] 하이픈(-)을 점(.)으로 먼저 치환하여 호환성 확보
    let endDateStr = parts[1].trim().replace(/-/g, ".");
    
    // 숫자와 점(.)만 남기고 나머지 제거
    endDateStr = endDateStr.replace(/[^0-9.]/g, ""); 
    
    const dateParts = endDateStr.split('.');
    
    if (dateParts.length < 3) return null;

    let year = parseInt(dateParts[0]);
    const month = parseInt(dateParts[1]) - 1;
    const day = parseInt(dateParts[2]);

    if (year < 100) year += 2000;

    return { end: new Date(year, month, day) };
  } catch (e) {
    return null;
  }
}

function helper_IsDatePassed(checkDateStr) {
  try {
    if (!checkDateStr || checkDateStr.length < 5) return false;
    
    // [중요 수정] 하이픈(-)이 들어와도 점(.)으로 치환하여 처리
    const cleanDateStr = checkDateStr.replace(/-/g, ".");
    const parts = cleanDateStr.split('.');
    
    let year = parseInt(parts[0]) + 2000;
    let month = parseInt(parts[1]) - 1; 
    
    const targetDate = new Date(year, month, 1);
    const today = new Date();
    
    return today >= targetDate;
  } catch (e) {
    return false;
  }
}