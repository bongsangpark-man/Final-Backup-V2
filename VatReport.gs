/**
 * 파일명: VatReport.gs
 * 기능: 부가세 신고용 '임대현황표' 양식 자동 작성
 * 수정사항: '매매' 건에 대한 잔금일 기준 필터링 및 입주일 기재 로직 추가
 * [최종 수정] 이미지 확인 결과 반영: 용도(Q열, 인덱스 16) 기준 인덱스 및 범위 조정
 */

// 메뉴 연결용 함수
function runVatRentalReport1() {
  return generateRentalStatusReport(1, 6, "1기");
}

function runVatRentalReport2() {
  return generateRentalStatusReport(7, 12, "2기");
}

/**
 * 메인 로직
 */
function generateRentalStatusReport(startMonth, endMonth, periodName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 1. 연도 파악
  const fileName = ss.getName(); 
  const yearMatch = fileName.match(/\d{4}/);
  const targetYear = yearMatch ? parseInt(yearMatch[0]) : new Date().getFullYear();
  
  // 조회 기간 설정
  const periodStart = new Date(targetYear, startMonth - 1, 1);
  const periodEnd = new Date(targetYear, endMonth, 0, 23, 59, 59);

  // 2. 시트 로드
  const sheetCurrent = ss.getSheetByName("임대 현황표");
  const sheetExit = ss.getSheetByName("임대 현황표(퇴실)"); 
  
  const targetSheetName = "부가세신고양식(임대현황표)";
  let targetSheet = ss.getSheetByName(targetSheetName);
  
  if (!targetSheet) {
    ui.alert(`⚠️ '${targetSheetName}' 시트가 없습니다.`);
    return null;
  }

  // 3. 데이터 읽기
  const lastRowCur = sheetCurrent.getLastRow();
  // ★ [수정 1] 범위 17 (A~Q열까지)로 조정 (이미지 확인 결과 Q열이 마지막)
  const curData = lastRowCur > 1 ? sheetCurrent.getRange(2, 1, lastRowCur - 1, 17).getValues() : [];

  const lastRowExit = sheetExit ? sheetExit.getLastRow() : 0;
  // ★ [수정 2] 범위 18 (A~R열까지)로 조정 (퇴실일 추가로 1칸 밀림)
  const exitData = (sheetExit && lastRowExit > 1) ? sheetExit.getRange(2, 1, lastRowExit - 1, 18).getValues() : [];

  let finalRows = [];

  // ==================================================
  // [Step 1] 현재 임대 현황
  // ==================================================
  for (let i = 0; i < curData.length; i++) {
    const row = curData[i];
    
    let item = {
      isExit: false,
      unit: String(row[0]).trim(),
      type: String(row[1]).trim(),
      bizNo: String(row[2]).trim(),
      area1: row[3],
      area2: row[4],
      deposit: row[5],
      rent: row[6],
      periodStr: String(row[9]).trim(),
      tenant: String(row[11]).trim(),
      resNo: String(row[12]).trim(), // M열(12)은 주민번호
      
      // ★ [수정 3] 용도 인덱스 16 (Q열)으로 확정
      usage: String(row[16]).trim()
    };

    if (isValidItem(item, periodStart, periodEnd, false)) {
      // 매매인 경우 잔금일을 입주일로 설정
      if (item.type.includes("매매")) {
         const balanceDate = parseDateSmart(item.periodStr);
         item.moveInDate = balanceDate ? formatDate(balanceDate) : "";
      } else {
         const dates = parsePeriodString(item.periodStr);
         item.moveInDate = dates.start ? formatDate(dates.start) : "";
      }
      item.exitDate = ""; 
      finalRows.push(item);
    }
  }

  // ==================================================
  // [Step 2] 퇴실자 현황
  // ==================================================
  for (let i = 0; i < exitData.length; i++) {
    const row = exitData[i];
    
    let exitDateRaw = row[0];
    let exitDateObj = parseDateSmart(exitDateRaw);

    // 날짜 없으면 패스
    if (!exitDateObj) continue; 

    // 1. 작년 퇴실 제외
    const fileStartObject = new Date(targetYear, 0, 1); 
    if (exitDateObj < fileStartObject) continue;

    // 2. 기수 범위 밖 제외
    if (exitDateObj < periodStart || exitDateObj > periodEnd) continue;

    let item = {
      isExit: true,
      exitDateObj: exitDateObj,
      
      unit: String(row[1]).trim(),
      type: String(row[2]).trim(),
      bizNo: String(row[3]).trim(),
      area1: row[4],
      area2: row[5],
      deposit: row[6],
      rent: row[7],
      periodStr: String(row[10]).trim(),
      tenant: String(row[12]).trim(),
      resNo: String(row[13]).trim(), // N열(13)은 주민번호 (퇴실시트라 1칸 밀림)
      
      // ★ [수정 4] 용도 인덱스 17 (R열)으로 확정 (퇴실시트라 1칸 밀림)
      usage: String(row[17]).trim()
    };

    if (isValidItem(item, periodStart, periodEnd, true)) {
      if (item.type.includes("매매")) {
         const balanceDate = parseDateSmart(item.periodStr);
         item.moveInDate = balanceDate ? formatDate(balanceDate) : "";
      } else {
         const dates = parsePeriodString(item.periodStr);
         item.moveInDate = dates.start ? formatDate(dates.start) : "";
      }
      item.exitDate = formatDate(exitDateObj);
      finalRows.push(item);
    }
  }

  // ==================================================
  // [Step 3] 정렬 및 출력
  // ==================================================
  finalRows.sort((a, b) => {
    const numA = parseInt(a.unit.replace(/[^0-9]/g, "")) || 0;
    const numB = parseInt(b.unit.replace(/[^0-9]/g, "")) || 0;
    if (numA !== numB) return numA - numB;

    if (a.isExit && !b.isExit) return -1;
    if (!a.isExit && b.isExit) return 1;
    return 0;
  });

  const lastRowTarget = targetSheet.getLastRow();
  if (lastRowTarget >= 3) {
    targetSheet.getRange(3, 1, lastRowTarget - 2, 12).clearContent();
  }

  if (finalRows.length === 0) {
    ui.alert(`조건에 맞는 데이터가 없습니다.\n(기준연도: ${targetYear})`);
    return null;
  }

  const outputValues = finalRows.map(r => {
    // 사업자 없으면 주민번호 사용
    const finalIdNum = r.bizNo ? r.bizNo : r.resNo;

    return [
      r.usage, r.unit, r.area1, r.area2, r.exitDate, r.moveInDate,
      r.tenant, finalIdNum, r.periodStr, r.type, r.deposit, r.rent
    ];
  });

  targetSheet.getRange(3, 1, outputValues.length, 12).setValues(outputValues);
  targetSheet.getRange(3, 11, outputValues.length, 2).setNumberFormat("#,##0");

  const spreadsheetId = ss.getId();
  const sheetId = targetSheet.getSheetId();
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=xlsx&gid=${sheetId}`;

  ui.alert(`✅ 생성 완료 (${targetYear}년 ${periodName})\n총 ${outputValues.length}건 작성됨`);
  
  return url;
}

// ----------------------------------------------------
// Helper Functions
// ----------------------------------------------------
function isValidItem(item, periodStart, periodEnd, isExitRow) {
  if (!item.unit) return false;
  if (item.usage.includes("근린생활")) return false;

  if (item.type.includes("매매")) {
    if (item.periodStr.toLowerCase().includes("sh")) return false;
    const balanceDate = parseDateSmart(item.periodStr);
    if (!balanceDate) return false;
    if (balanceDate < periodStart || balanceDate > periodEnd) {
      return false; 
    }
    return true;
  } else {
    if (!isExitRow) {
      const dates = parsePeriodString(item.periodStr);
      if (dates.start && dates.end) {
        if (dates.end < periodStart || dates.start > periodEnd) return false;
      }
    }
  }
  return true;
}

function parsePeriodString(str) {
  if (!str) return { start: null, end: null };
  const p = str.split("~");
  return p.length < 2 ? { start: null, end: null } : { start: parseDateSmart(p[0]), end: parseDateSmart(p[1]) };
}

function parseDateSmart(val) {
  if (!val) return null;
  if (Object.prototype.toString.call(val) === "[object Date]") {
    if (isNaN(val.getTime())) return null;
    return val;
  }
  const str = String(val).trim();
  const parts = str.replace(/[^0-9.\-]/g, "").split(/[.\-]/);
  
  if (parts.length === 3) {
    let y = parseInt(parts[0]);
    let m = parseInt(parts[1]) - 1;
    let d = parseInt(parts[2]);
    if (y < 100) y += 2000;
    const dateObj = new Date(y, m, d);
    if (!isNaN(dateObj.getTime())) return dateObj;
  }
  const simpleDate = new Date(str);
  if (!isNaN(simpleDate.getTime())) return simpleDate;
  return null;
}

function formatDate(d) {
  if(!d) return "";
  return `${d.getFullYear()}.${String(d.getMonth()+1).padStart(2,"0")}.${String(d.getDate()).padStart(2,"0")}`;
}