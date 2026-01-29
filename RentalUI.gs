/**
 * RentalUI.gs
 * 임대 관리 시스템의 전용 로직 파일입니다.
 * 기능: 데이터 조회/수정/변경, 타 시트 동기화, 시트 보호, **퇴실 정산 데이터 추출**
 * * [수정 사항]
 * - 임대 현황표 M열(주민번호) 추가에 따른 열 인덱스 +1 조정 (연락처 N열, 신탁 O열, 퇴실 P열, 용도 R열)
 * - 수정/변경 시 '주민등록번호' 데이터 처리 로직 추가 (인덱스 12)
 * - 임대 현황표(퇴실) N열(주민번호) 추가 대응
 */

// ==========================================
// [설정] 시트 이름 상수 정의
// ==========================================
const UI_SHEET_RENTAL = "임대 현황표";
const UI_SHEET_RENTAL_EXIT = "임대 현황표(퇴실)";
const UI_SHEET_RENT = "임대료 납부내역";
const UI_SHEET_RENT_EXIT = "임대료 납부내역(퇴실)";
const UI_SHEET_MGMT = "관리비 납부내역";
const UI_SHEET_MGMT_EXIT = "관리비 납부내역(퇴실)";

// ==========================================
// 1. 초기 데이터 조회 및 헬퍼 함수
// ==========================================

function showRentalSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('🏢 임대 관리 시스템')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getRoomListAndStatus() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UI_SHEET_RENTAL);
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return [];
  // [수정] 열 추가로 인해 읽는 범위 확장 (17 -> 18, R열까지 읽음)
  const data = sheet.getRange(2, 1, lastRow - 1, 18).getValues();
  return data.map(row => ({
    hosu: row[0],       // A열
    type: row[1],       // B열
    tenant: row[11],    // L열
    isVacancy: (row[1] === '공실' || row[11] === '공실')
  })).filter(r => r.hosu !== "");
}

function getRoomDetail(hosu) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UI_SHEET_RENTAL);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == hosu) {
      const rowData = data[i].map(cell => {
        if (Object.prototype.toString.call(cell) === '[object Date]') {
          return Utilities.formatDate(cell, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        return cell;
      });
      return rowData; 
    }
  }
  return null;
}

// ==========================================
// 2. [메뉴 A] 임대 현황 수정 로직
// ==========================================
function updateRentalInfo(formObject) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UI_SHEET_RENTAL);
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == formObject.hosu) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex === -1) throw new Error("호수를 찾을 수 없습니다.");
  
  // [수정] 열 추가로 인해 수정 범위 확장 (17 -> 18)
  const rowRange = sheet.getRange(rowIndex, 1, 1, 18);
  const currentValues = rowRange.getValues()[0];
  const newValues = [...currentValues];
  
  // 1. 사업자번호(C열, 인덱스 2)
  if('bizNum' in formObject) {
    newValues[2] = formObject.bizNum;
  }

  // 2. 계약자(L열, 인덱스 11)
  if ('tenant' in formObject) {
    newValues[11] = formObject.tenant;
  }

  // 3. 나머지 항목 매핑
  if('deposit1' in formObject && formObject.deposit1) newValues[5] = formObject.deposit1; // F열
  if('deposit2' in formObject && formObject.deposit2) newValues[5] = formObject.deposit2; // F열
  if('rent' in formObject) newValues[6] = formObject.rent;       // G열
  if('parking' in formObject) newValues[7] = formObject.parking; // H열
  if('food' in formObject) newValues[8] = formObject.food;       // I열
  if('period' in formObject) newValues[9] = formObject.period;   // J열
  if('structure' in formObject) newValues[10] = formObject.structure; // K열
  
  // ★ [신규] 주민등록번호 (M열, 인덱스 12)
  // 사이드바 수정탭에 주민번호 입력란이 없다면 기존값 유지, 있다면 업데이트
  if('resNo' in formObject) {
    newValues[12] = formObject.resNo;
  }
  
  // ★ [인덱스 수정] 기존 항목들 1칸씩 뒤로 이동
  if('phone' in formObject) newValues[13] = formObject.phone;    // N열 (12 -> 13)
  if('trust' in formObject) newValues[14] = formObject.trust;    // O열 (13 -> 14)
  if('out' in formObject) newValues[15] = formObject.out;        // P열 (14 -> 15)
  
  // 데이터 시트에 반영
  rowRange.setValues([newValues]);

  // [동기화] 계약 만기 일정 업데이트
  try {
    if (typeof main_UpdateContractExpiry === 'function') {
        main_UpdateContractExpiry();
    }
  } catch (e) {
    console.error("만기 일정 업데이트 중 오류: " + e.toString());
  }

  // [동기화] 월별 현황판 업데이트
  refreshDashboardLogic();

  return "✅ 수정이 완료되었습니다.";
}


// ==========================================
// 3. [메뉴 B] 임대 현황 변경 (퇴실 및 신규) 로직
// ==========================================
function changeRentalStatus(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRental = ss.getSheetByName(UI_SHEET_RENTAL);
  
  // 1. 호수 행 찾기
  const data = sheetRental.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == form.hosu) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex === -1) throw new Error("호수를 찾을 수 없습니다.");

  // ---------------------------------------------------
  // Step 1. 기존 데이터 이관 (Archiving) 및 정보 백업
  // ---------------------------------------------------
  // ★ [중요] 범위를 18로 늘려 R열(용도)까지 읽어야 M열(주민번호)도 포함됩니다.
  const rangeWidth = 18; // A~R열 (인덱스 0~17)
  
  const currentRowData = sheetRental.getRange(rowIndex, 1, 1, rangeWidth).getValues()[0];

  // [정산서용] 백업 변수 선언
  let oldDeposit = 0;
  let oldTenantName = "";
  let oldRent = 0;    // 월세
  let oldPayDay = ""; // 납부일
  
  if (!form.isPreviouslyVacant) { 
    const exitDate = form.exitDate;
    // A. 임대 현황표 데이터 가져오기 (이관용)
    const oldRentalType = currentRowData[1];
    
    // 기존 세입자 정보 백업
    oldDeposit = Number(currentRowData[5]); // F열: 보증금
    oldRent = Number(currentRowData[6]);    // G열: 월세
    oldTenantName = currentRowData[11];     // L열: 이름
    
    // 납부일 정보는 '임대료 납부내역' 시트 C열에 있음
    const rentSheet = ss.getSheetByName(UI_SHEET_RENT);
    const rRow = findRowIndex(rentSheet, form.hosu);
    if(rRow > 0) {
       oldPayDay = rentSheet.getRange(rRow, 3).getValue(); // C열 값
    }

    // ★ [핵심] 퇴실 시트로 데이터 이관
    // 구조: [퇴실일, 호수, 유형, ..., 이름, 주민번호, 연락처 ...]
    // currentRowData(18개)가 M열(주민번호)을 포함하고 있으므로, 
    // 맨 앞에 exitDate만 붙여주면 퇴실 시트의 N열(주민번호) 위치에 정확히 들어갑니다.
    const sheetRentalExit = ss.getSheetByName(UI_SHEET_RENTAL_EXIT);
    sheetRentalExit.appendRow([exitDate, ...currentRowData]);
    
    // B. 납부 내역 이관 (전세 아닐때만)
    if (oldRentalType !== '전세') {
       archivePaymentData(ss, UI_SHEET_RENT, UI_SHEET_RENT_EXIT, form.hosu);
    }
    
    // 관리비는 무조건 이관
    archivePaymentData(ss, UI_SHEET_MGMT, UI_SHEET_MGMT_EXIT, form.hosu);
  }

  // ---------------------------------------------------
  // Step 2. 현재 시트 초기화 & 신규 데이터 입력
  // ---------------------------------------------------
  const targetRow = sheetRental.getRange(rowIndex, 1, 1, rangeWidth);
  
  let newRow = [...currentRowData];

  // ★ [수정] 열 인덱스 밀림 및 신규 필드(주민번호) 반영
  if (form.newType === '공실') {
    newRow[1] = '공실';
    newRow[11] = '공실'; 
    newRow[2] = ''; // C열 (사업자번호)
    newRow[5] = ''; // F열 (보증금)
    newRow[6] = ''; // G열 (월세)
    newRow[7] = ''; // H열 (주차)
    newRow[8] = ''; // I열 (음식물)
    newRow[9] = ''; // J열 (기간)
    
    // L열(11)은 위에서 '공실' 처리
    
    // ★ [중요] 공실 처리 시 M열(12) 주민번호도 지워야 함!
    newRow[12] = ''; 
    
    // ★ [이동] 나머지 열 초기화
    newRow[13] = ''; // N열 (연락처 - 12->13)
    newRow[14] = ''; // O열 (신탁 - 13->14)
    newRow[15] = ''; // P열 (중도퇴실 - 14->15)
    
    // R열(17) [용도]는 건드리지 않음 (건물 속성)
  } else {
    newRow[1] = form.newType;      
    newRow[2] = form.bizNum || '';
    
    let depositVal = '';
    if(form.deposit1) depositVal = form.deposit1;
    if(form.deposit2) depositVal = form.deposit2;
    newRow[5] = depositVal; 
    
    newRow[6] = form.rent || '';
    newRow[7] = form.parking || ''; 
    newRow[8] = form.food || '';    
    newRow[9] = (form.newType === '매매') ? form.balanceDate : form.periodStart + " ~ " + form.periodEnd; 
    newRow[11] = form.tenantFinal;  
    
    // ★ [중요] 신규 계약 시 새 주민번호 입력 (없으면 빈값으로 덮어써서 기존 정보 삭제)
    newRow[12] = form.resNo || ''; 
    
    // ★ [이동] 나머지 열 입력
    newRow[13] = form.phone || '';   // N열 (연락처)
    newRow[14] = form.trust || '';   // O열 (신탁)
    newRow[15] = form.out || '';     // P열 (중도퇴실)
    
    // R열(17) [용도]는 건드리지 않음
  }
  targetRow.setValues([newRow]);

  // 납부 내역 시트 내용 지우기
  clearPaymentSheet(ss, UI_SHEET_MGMT, form.hosu, 3); 
  clearPaymentSheet(ss, UI_SHEET_RENT, form.hosu, 6);

  // ---------------------------------------------------
  // Step 3. 셀 색상 채우기 (입주 전 기간 까맣게 칠하기)
  // ---------------------------------------------------
  
  // A. 관리비
  const sheetMgmt = ss.getSheetByName(UI_SHEET_MGMT);
  const mgmtRowIdx = findRowIndex(sheetMgmt, form.hosu);
  
  if (mgmtRowIdx > 0) {
    let blackEndMonthIndex = -1;
    if (!form.isPreviouslyVacant) {
      const exitDateObj = new Date(form.exitDate);
      blackEndMonthIndex = (exitDateObj.getDate() >= 29) ? exitDateObj.getMonth() : exitDateObj.getMonth() - 1;
    } else {
      const startDateObj = new Date(form.periodStart || form.balanceDate);
      blackEndMonthIndex = startDateObj.getMonth() - 1;
    }
    if (blackEndMonthIndex >= 0) {
      const numColumnsToColor = (blackEndMonthIndex + 1) * 2;
      sheetMgmt.getRange(mgmtRowIdx, 3, 1, numColumnsToColor).setBackground("black");
    }
  }

  // B. 임대료
  if (form.newType.includes("월세")) {
    const sheetRent = ss.getSheetByName(UI_SHEET_RENT);
    const rentRowIdx = findRowIndex(sheetRent, form.hosu);
    
    if (rentRowIdx > 0) {
      const startDateObj = new Date(form.periodStart);
      const startMonth = startDateObj.getMonth();
      let blackEndMonthIndex = -1;

      if (form.newType === "월세(선불)") blackEndMonthIndex = startMonth - 1;
      else if (form.newType === "월세(후불)") blackEndMonthIndex = startMonth;

      if (blackEndMonthIndex >= 0) {
        const numColumnsToColor = (blackEndMonthIndex + 1) * 2;
        sheetRent.getRange(rentRowIdx, 6, 1, numColumnsToColor).setBackground("black");
      }
    }
  }

  // [동기화] 만기 일정 & 현황판
  try {
    if (typeof main_UpdateContractExpiry === 'function') main_UpdateContractExpiry();
  } catch (e) { console.error("만기 일정 오류: " + e.toString()); }

  refreshDashboardLogic();

  // [최종 리턴] 정산서 및 수수료 계산을 위해 신규 정보도 함께 전달
  return {
    success: true,
    message: "✅ 변경 처리(이관/초기화/색칠) 완료!",
    hosu: form.hosu,
    exitDate: form.exitDate,
    
    // 정산 대상(구 세입자)
    isPreviouslyVacant: form.isPreviouslyVacant,
    oldTenant: oldTenantName,
    oldDeposit: oldDeposit, 
    
    // 부동산 수수료 계산용 신규 정보
    newTenant: form.tenantFinal,
    newType: form.newType, 
    newDeposit: form.deposit2 ? Number(form.deposit2.replace(/,/g,'')) : 0, 
    newRent: form.rent ? Number(form.rent.replace(/,/g,'')) : 0,           
    
    // ★ [수정] 용도(R열) 인덱스 변경 16 -> 17
    fixedBizType: currentRowData[16] || "오피스텔",

    oldRent: oldRent,       
    oldPayDay: oldPayDay 
  };
}

// --- 내부 헬퍼 함수들 ---

function findRowIndex(sheet, hosu) {
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) { 
    if (data[i][0] == hosu) return i + 1;
  }
  return -1;
}

function archivePaymentData(ss, sourceName, targetName, hosu) {
  const sourceSheet = ss.getSheetByName(sourceName);
  const targetSheet = ss.getSheetByName(targetName);
  const rowIndex = findRowIndex(sourceSheet, hosu);
  
  if (rowIndex > 0) {
    const rowData = sourceSheet.getRange(rowIndex, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
    targetSheet.appendRow(rowData);
  }
}

function clearPaymentSheet(ss, sheetName, hosu, startColIndex) {
  const sheet = ss.getSheetByName(sheetName);
  const rowIndex = findRowIndex(sheet, hosu);
  if (rowIndex > 0) {
    const lastCol = sheet.getLastColumn();
    if (lastCol >= startColIndex) {
      const range = sheet.getRange(rowIndex, startColIndex, 1, lastCol - startColIndex + 1);
      range.clearContent();
      range.setBackground(null); // 색상 초기화
    }
  }
}

function refreshDashboardLogic() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName("월별 임대료 납부 현황");
  if (!dashSheet) return;
  const currentMonth = dashSheet.getRange("A1").getValue();
  try {
    if (typeof updateAndSortDashboard === 'function') updateAndSortDashboard(currentMonth);
  } catch (e) { console.error("현황판 업데이트 실패: " + e.toString());
  }
}

// ==========================================
// 4. [관리자 기능] 시트 잠금/해제 (경고창 방식)
// ==========================================

function lockRentalSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(UI_SHEET_RENTAL);

  // 1. 기존 보호 설정 초기화
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (let i = 0; i < protections.length; i++) {
    protections[i].remove();
  }
  
  PropertiesService.getScriptProperties().deleteProperty('IS_LOCKED');
  // 2. 구글 시트 자체 보호 기능 활성화 (경고창 모드)
  const protection = sheet.protect().setDescription('임대관리 시스템 보호');
  protection.setWarningOnly(true);
  SpreadsheetApp.getUi().alert('🔒 [잠금 완료]\n이제 시트를 수기로 수정하면 경고창이 뜹니다.\n(스크립트는 정상 작동합니다.)');
}

function unlockRentalSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(UI_SHEET_RENTAL);
  // 모든 보호 설정 제거
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (let i = 0; i < protections.length; i++) {
    protections[i].remove();
  }

  PropertiesService.getScriptProperties().deleteProperty('IS_LOCKED');
  SpreadsheetApp.getUi().alert('🔓 [잠금 해제]\n이제 자유롭게 수정할 수 있습니다.');
}