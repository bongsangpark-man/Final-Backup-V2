/**
 * Settlement.gs
 * 기능: '퇴실' 시트의 데이터를 조회하여 정산서 계산 및 생성
 * 수정사항: 31일(말일) 계약자 로직 유지 + 일할 계산 수식만 (월세*12/365)로 변경
 * [추가 수정] 퇴실 시트 N열(주민번호) 추가에 따른 BizType 인덱스 변경 (17 -> 18)
 */

// ==========================================
// [Part 1] 정산 데이터 자동 계산
// ==========================================
function getSettlementCalcData(hosu, exitDateStr, newInfo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const exitDate = new Date(exitDateStr);
  exitDate.setHours(0, 0, 0, 0); 
  
  // 1. 퇴실자 정보 조회
  const sheetExit = ss.getSheetByName("임대 현황표(퇴실)");
  if (!sheetExit) return { error: "'임대 현황표(퇴실)' 시트가 없습니다." };
  
  const exitData = sheetExit.getDataRange().getValues();
  let rInfo = null;
  // 가장 최근 퇴실 기록 찾기
  for (let i = exitData.length - 1; i >= 1; i--) {
    if (String(exitData[i][1]) == String(hosu)) { 
      rInfo = {
        type: String(exitData[i][2]),      
        deposit: Number(exitData[i][6]),   
        rent: Number(exitData[i][7]),      
        
        // ★ [수정] N열 추가로 인해 인덱스 17 -> 18로 변경 (R열 -> S열)
        bizType: exitData[i][17] || "도시형생활주택", 
        
        period: String(exitData[i][10])    
      };
      break; 
    }
  }
  
  if (!rInfo) return { error: "해당 호수의 퇴실 데이터를 찾을 수 없습니다." };

  // 2. 납부 내역 조회
  const rentRow = findLastRowInExitSheet(ss, "임대료 납부내역(퇴실)", hosu);
  const maintRow = findLastRowInExitSheet(ss, "관리비 납부내역(퇴실)", hosu);

  // 3. 계산 로직 수행
  const rentCalc = calculateRentLogic(rInfo, rentRow, exitDate);
  const maintCalc = calculateMaintLogic(maintRow, exitDate);

  return {
    hosu: hosu,
    date: exitDateStr,
    info: rInfo,
    rent: rentCalc,
    maint: maintCalc,
    newInfo: newInfo
  };
}

function findLastRowInExitSheet(ss, sheetName, hosu) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (String(data[i][0]) == String(hosu)) return data[i];
  }
  return [];
}

// ----------------------------------------------------
// [로직 1] 임대료 계산 (수정됨: 365일 연환산 수식 적용)
// ----------------------------------------------------
function calculateRentLogic(rInfo, rowData, exitDate) {
  const monthlyRent = rInfo.rent;
  if (!monthlyRent) return { amount: 0, period: "", prepaid: 0, prepaidRefund: 0 };

  // 1. Anchor Day (최초 계약일의 '일') 추출 - 기존 로직 유지
  let anchorDay = 1;
  if (rInfo.period) {
    try {
      const str = String(rInfo.period).split("~")[0].trim();
      if(str.indexOf('.') > -1) {
          const parts = str.split('.');
          anchorDay = Number(parts[2]);
      } else {
          const d = new Date(str);
          if(!isNaN(d.getTime())) anchorDay = d.getDate();
      }
    } catch (e) { anchorDay = 1; }
  }
  if (!anchorDay || anchorDay < 1) anchorDay = 1;

  // 2. 마지막 납부 월 확인
  let lastPaidMonth = 0;   
  let lastPaidAmount = 0;
  if (rowData && rowData.length > 5) {
    for (let c = 6; c < rowData.length; c += 2) {
      if (rowData[c] && rowData[c] !== "") {
        lastPaidMonth = (c - 6) / 2 + 1; 
        lastPaidAmount = Number(rowData[c - 1]); 
      }
    }
  }

  // 3. '납부 완료일(paidUntil)' 계산 
  let paidMonthIdx = (rInfo.type.includes("선불")) ? lastPaidMonth - 1 : lastPaidMonth - 2;
  
  let baseYear = exitDate.getFullYear();
  if (exitDate.getMonth() < 2 && paidMonthIdx > 9) {
      baseYear = exitDate.getFullYear() - 1; 
  } else if (exitDate.getMonth() > 9 && paidMonthIdx < 2) {
      baseYear = exitDate.getFullYear() + 1; 
  }

  // Anchor Day 적용하여 시작일 구하기
  let lastDayOfPaidMonth = new Date(baseYear, paidMonthIdx + 1, 0).getDate();
  let realStartDay = (anchorDay > lastDayOfPaidMonth) ? lastDayOfPaidMonth : anchorDay;
  
  let paidPeriodStart = new Date(baseYear, paidMonthIdx, realStartDay);

  // Smart End Date 로직 유지
  let paidUntil = new Date(paidPeriodStart);
  if (lastPaidMonth > 0) {
      paidUntil = getSmartEndDate(paidPeriodStart, anchorDay);
      paidUntil.setHours(0,0,0,0);
  } else {
      paidUntil = new Date(exitDate);
      paidUntil.setFullYear(paidUntil.getFullYear() - 2);
  }

  // 4. 정산 수행 (수정: 자투리 일수 계산식만 변경)
  let deductAmount = 0;
  let refundAmount = 0;
  let periodText = "";

  // [Case A] 선불 환불 (퇴실일이 납부 완료일보다 앞설 때)
  if (rInfo.type.includes("선불") && exitDate.getTime() <= paidUntil.getTime() && lastPaidMonth > 0) {
      refundAmount = lastPaidAmount > 0 ? lastPaidAmount : monthlyRent;
      
      let termStart = new Date(paidPeriodStart); 
      
      const usedDays = getDiffDays(termStart, exitDate);
      // const totalDaysInTerm = getDiffDays(termStart, paidUntil); // 기존 방식 주석 처리
      
      // ★수정 1: 환불 시 사용료 공제도 365일 기준으로 변경하여 일관성 확보
      deductAmount = Math.floor(((monthlyRent * 12) / 365) * usedDays);
      
      periodText = `${getFmt(termStart)}~${getFmt(exitDate)}`;
  } 
  // [Case B] 미납/추가 정산 (퇴실일이 납부 완료일보다 뒤일 때)
  else {
      let calcStart = new Date(paidUntil);
      calcStart.setDate(calcStart.getDate() + 1);
      
      let previousUnpaid = 0;
      if (lastPaidMonth > 0 && lastPaidAmount < monthlyRent) {
          previousUnpaid = monthlyRent - lastPaidAmount;
      }
      
      let totalRentCalc = 0;
      let finalStartStr = getFmt(calcStart);
      
      if (calcStart.getTime() > exitDate.getTime()) {
          if (previousUnpaid > 0) {
             deductAmount = previousUnpaid;
             periodText = "미납금 정산";
          } else {
             periodText = "-";
          }
      } else {
          // ★ 기존의 '한 달 단위 순회(while)' 로직 완벽 유지
          let currentCursor = new Date(calcStart);
          
          while (currentCursor.getTime() <= exitDate.getTime()) {
              let currentEnd = getSmartEndDate(currentCursor, anchorDay);
              
              if (currentEnd.getTime() <= exitDate.getTime()) {
                  // 한 달 통째로 지남 -> 월세 100% 부과 (기존 유지)
                  totalRentCalc += monthlyRent;
                  currentCursor = new Date(currentEnd);
                  currentCursor.setDate(currentCursor.getDate() + 1);
              } else {
                  // 자투리 일수 정산 -> 여기만 365일 로직 적용
                  const remainingDays = getDiffDays(currentCursor, exitDate); 
                  
                  // ★수정 2: 자투리 일수는 365일 기준으로 계산
                  // 기존: totalRentCalc += Math.floor((monthlyRent / daysInThisMonth) * remainingDays);
                  totalRentCalc += Math.floor(((monthlyRent * 12) / 365) * remainingDays);
                  break; 
              }
          }
          deductAmount = previousUnpaid + totalRentCalc;
          periodText = `${finalStartStr}~${getFmt(exitDate)}`;
      }
      refundAmount = 0;
  }

  deductAmount = Math.floor(deductAmount / 10) * 10;

  return {
    amount: deductAmount,       
    period: periodText,         
    prepaidRefund: refundAmount 
  };
}

// ----------------------------------------------------
// [로직 2] 관리비 계산 (유지)
// ----------------------------------------------------
function calculateMaintLogic(rowData, exitDate) {
  if (!rowData || rowData.length === 0) return { amount: 0, period: "" };
  
  let targetMonthIndex = 0;
  let targetAmount = 0;
  
  for (let c = 2; c < rowData.length; c += 2) {
    if (rowData[c] && rowData[c] !== "") {
      targetMonthIndex = (c - 2) / 2;
      targetAmount = Number(rowData[c]);
    }
  }
  
  const dateCellVal = rowData[targetMonthIndex * 2 + 3];
  const isPaid = (dateCellVal && dateCellVal !== "");
  
  let startDate = new Date(exitDate); 
  startDate.setHours(0,0,0,0);
  let totalDeduct = 0;

  if (isPaid) {
    startDate.setFullYear(exitDate.getFullYear());
    startDate.setMonth(targetMonthIndex + 1); 
    startDate.setDate(1);
  } else {
    startDate.setFullYear(exitDate.getFullYear());
    startDate.setMonth(targetMonthIndex); 
    startDate.setDate(1);
    totalDeduct += targetAmount;
  }
  
  const daysInExitMonth = new Date(exitDate.getFullYear(), exitDate.getMonth() + 1, 0).getDate();
  const daysUsed = exitDate.getDate();
  const proRated = Math.floor(targetAmount * (daysUsed / daysInExitMonth));
  totalDeduct += proRated;
  totalDeduct = Math.floor(totalDeduct / 10) * 10; // 10원 단위 절사

  return {
    amount: totalDeduct,
    period: `${getFmt(startDate)}~${getFmt(exitDate)}`
  };
}

// ==========================================
// [Part 2] 엑셀 생성 (유지)
// ==========================================
function createSettlementExcel(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const template = ss.getSheetByName("퇴실정산서");
  if (!template) throw new Error("'퇴실정산서' 시트가 없습니다.");
  
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
  const sheetName = `${data.hosu}호_정산서_${today}`;
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);
  sheet = template.copyTo(ss).setName(sheetName);

  // 1. 상단 기본 정보
  sheet.getRange("C2").setValue(data.hosu + "호");
  sheet.getRange("E2").setValue(data.date); 
  
  // 임차료 (4행)
  sheet.getRange("B4").setValue(data.rentType); 
  sheet.getRange("C4").setValue(data.rent.period || "-");
  sheet.getRange("D4").setValue(data.rent.amount);
  
  // 관리비 (5행)
  sheet.getRange("C5").setValue(data.maint.period || "-");
  sheet.getRange("D5").setValue(data.maint.amount);

  // 공과금 (6~8행)
  const utilPeriod = (val) => val ? `~ ${data.date}` : "";
  sheet.getRange("D6").setValue(data.utils.elec);
  if(data.utils.elec) sheet.getRange("C6").setValue(utilPeriod(true));
  
  sheet.getRange("D7").setValue(data.utils.gas);
  if(data.utils.gas) sheet.getRange("C7").setValue(utilPeriod(true));
  
  sheet.getRange("D8").setValue(data.utils.water);
  if(data.utils.water) sheet.getRange("C8").setValue(utilPeriod(true));

  // 부동산 수수료 (9행)
  if(String(data.brokerFee).trim() === "" || data.brokerFee == 0) {
      sheet.getRange("D9").setValue("");
      sheet.getRange("E9").setValue("직접 정산");
  } else {
      sheet.getRange("D9").setValue(data.brokerFee);
  }

  // 2. 하자보수 및 동적 행 추가 로직 (10행부터 시작)
  let repairItems = [];
  if (data.repair.clean) repairItems.push({ name: "청소비", amount: data.repair.clean });
  if (data.repair.filter) repairItems.push({ name: "전열교환기 필터", amount: data.repair.filter });
  if (data.repair.asItems && data.repair.asItems.length > 0) {
    data.repair.asItems.forEach(item => repairItems.push({ name: item.name, amount: item.amount }));
  }

  let currentRow = 10; 
  let addedRows = 0; 

  repairItems.forEach((item, index) => {
    if (index === 0) {
      sheet.getRange(currentRow, 2).setValue(item.name);
      sheet.getRange(currentRow, 3).setValue("");
      sheet.getRange(currentRow, 4).setValue(item.amount);
    } else {
      sheet.insertRowAfter(currentRow);
      currentRow++; 
      addedRows++;
      
      try {
        const sourceRange = sheet.getRange(currentRow - 1, 1, 1, sheet.getLastColumn());
        const targetRange = sheet.getRange(currentRow, 1);
        sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      } catch (e) {}
      
      sheet.getRange(currentRow, 2).setValue(item.name);
      sheet.getRange(currentRow, 3).setValue("");
      sheet.getRange(currentRow, 4).setValue(item.amount);
    }
  });

  if (repairItems.length === 0) {
      sheet.getRange("B10").setValue("");
      sheet.getRange("D10").setValue("");
  }

  // 3. 하단부 위치 조정
  const retainedRow = 11 + addedRows;
  sheet.getRange("D" + retainedRow).setValue(data.retained);

  const totalRow = 12 + addedRows;
  sheet.getRange("D" + totalRow).setValue(data.totalDeduct); 

  const depositRow = 13 + addedRows;
  sheet.getRange("D" + depositRow).setValue(data.deposit);

  let extraRow = 0;
  if (data.prepaidRefund && Number(String(data.prepaidRefund).replace(/,/g,'')) > 0) {
      const targetRow = depositRow + 1; 
      sheet.insertRowBefore(targetRow);
      try {
        const sourceRange = sheet.getRange(targetRow - 1, 1, 1, sheet.getLastColumn());
        const targetRange = sheet.getRange(targetRow, 1);
        sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      } catch (e) {}
      
      sheet.getRange("B" + targetRow).setValue("월세(선불)");
      sheet.getRange("C" + targetRow).setValue(""); 
      sheet.getRange("D" + targetRow).setValue(data.prepaidRefund);
      extraRow = 1; 
  }

  const bankRow = 14 + addedRows + extraRow;
  sheet.getRange("B" + bankRow).setValue(data.bank.owner); 
  sheet.getRange("C" + bankRow).setValue(`${data.bank.name} ${data.bank.account}`);
  sheet.getRange("D" + bankRow).setValue(data.finalRefund);

  const checkStartRow = 16 + addedRows + extraRow;
  const mapC = (val) => (val ? "0" : "X"); 
  const checks = data.checks;

  sheet.getRange(checkStartRow, 4).setValue(mapC(checks.key));
  sheet.getRange(checkStartRow+1, 4).setValue(mapC(checks.food));
  sheet.getRange(checkStartRow+2, 4).setValue(mapC(checks.air));
  sheet.getRange(checkStartRow+3, 4).setValue(mapC(checks.heat));
  sheet.getRange(checkStartRow+4, 4).setValue(mapC(checks.kt));

  SpreadsheetApp.flush();
  return `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=xlsx&gid=${sheet.getSheetId()}`;
}

// ----------------------------------------------------
// [Helper Functions] 날짜 포맷 및 계산
// ----------------------------------------------------
function getFmt(date) { 
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yy/MM/dd");
}

function getDiffDays(start, end) {
  const s = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  const e = new Date(end.getFullYear(), end.getMonth(), end.getDate());
  const diffTime = e - s;
  return Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
}

// ★ [신규] 말일(Anchor) 기준 정확한 종료일 계산 함수 (기존 로직 완벽 유지)
// 시작일과 Anchor(31일)를 넣으면 이번 달이 언제 끝나는지 알려줌
function getSmartEndDate(startDate, anchorDay) {
  // 1. 다음 달 1일로 이동하여 '다음 달'이 몇 월인지 파악
  let target = new Date(startDate);
  target.setMonth(target.getMonth() + 1);
  target.setDate(1);

  let year = target.getFullYear();
  let month = target.getMonth();
  
  // 2. 그 달의 마지막 날짜 확인 (28, 30, 31 중 하나)
  let lastDayOfMonth = new Date(year, month + 1, 0).getDate();
  
  // 3. AnchorDay 적용 (예: 31일이 목표인데 2월이라 28일밖에 없으면 28일로)
  let targetDay = (anchorDay > lastDayOfMonth) ? lastDayOfMonth : anchorDay;
  
  // 4. 종료일은 '다음 구간 시작일'의 하루 전날
  let nextCycleStart = new Date(year, month, targetDay);
  nextCycleStart.setDate(nextCycleStart.getDate() - 1);
  
  return nextCycleStart;
}