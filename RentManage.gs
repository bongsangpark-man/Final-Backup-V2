/**
 * 파일명: RentManage.gs
 * 기능: 수협은행 입금 내역 처리 및 연도별 자동 이월 (A1 메모 설정 방식)
 */

// ==========================================
// [1] 설정값 (임대 현황표 A1 메모에서 불러오기)
// ==========================================
function getSystemConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('임대 현황표');
  
  // A1 셀의 메모(Note)를 가져옵니다.
  const note = sheet ? sheet.getRange('A1').getNote() : "";

  // 1. 메모에 설정값(JSON)이 있는 경우 (2026년 이후 파일)
  if (note && note.startsWith('{')) {
    try {
      const config = JSON.parse(note);
      console.log("설정 로드 완료: " + config.year);
      return config;
    } catch (e) {
      console.error("설정 메모 파싱 실패: " + e.message);
    }
  }

  // 2. 메모가 없는 경우 (2025년 현재 파일) -> 기본값 사용
  // ★ 2025년은 작년 파일이 없으므로 내 ID를 그대로 사용
  return {
    prevId: ss.getId(), 
    year: '2025'
  };
}

// ==========================================
// [2] 메인 처리 함수 (Code.gs에서 호출)
// ==========================================
function processRentTransactions(transactionData) {
  
  // ★ 설정값 동적 로드
  const config = getSystemConfig();
  const PREV_YEAR_SS_ID = config.prevId;
  const CURRENT_YEAR = config.year;

  // 1. 데이터 수신 체크
  if (!transactionData || transactionData.length === 0) {
    SpreadsheetApp.getUi().alert('⚠️ 처리할 거래내역이 없습니다.');
    return;
  }

  // 2. 파일 로드 (현재 & 작년)
  const ss = SpreadsheetApp.getActiveSpreadsheet(); // 현재
  let prevSS = null;

  try {
    prevSS = SpreadsheetApp.openById(PREV_YEAR_SS_ID); // 작년
  } catch (e) {
    console.error("작년 파일 로드 실패: " + e.message);
    SpreadsheetApp.getUi().alert("⚠️ 작년 파일(" + (parseInt(CURRENT_YEAR)-1) + ")을 열 수 없습니다.\nID 권한을 확인해주세요.");
    return;
  }

  // 3. 시트 및 데이터 로드 (메모리 캐싱)
  const SHEET_RENT = '임대료 납부내역';    
  const SHEET_MAINT = '관리비 납부내역';
  
  // [현재 파일]
  const curRentSheet = ss.getSheetByName(SHEET_RENT);
  const curMaintSheet = ss.getSheetByName(SHEET_MAINT);
  const curRentData = curRentSheet.getDataRange().getValues();
  const curRentBg = curRentSheet.getDataRange().getBackgrounds();
  const curMaintData = curMaintSheet.getDataRange().getValues();
  const curMaintBg = curMaintSheet.getDataRange().getBackgrounds();
  
  // [작년 파일]
  const prevRentSheet = prevSS.getSheetByName(SHEET_RENT);
  const prevMaintSheet = prevSS.getSheetByName(SHEET_MAINT);
  const prevRentData = prevRentSheet.getDataRange().getValues();
  const prevRentBg = prevRentSheet.getDataRange().getBackgrounds();
  const prevMaintData = prevMaintSheet.getDataRange().getValues();
  const prevMaintBg = prevMaintSheet.getDataRange().getBackgrounds();

  // 설정값 상수 (코드 내부 사용)
  const THRESHOLD_MAINT = 400000;
  const IGNORE_LIMIT = 5000000;
  const IGNORE_LIST = ['이갑희', '손숙', '쏘카', '주식회사 쏘카'];

  let stats = { rent: 0, maint: 0, fail: 0, skip: 0, log: [] };

  // -----------------------------------------------------------
  // 거래내역 루프 시작
  // -----------------------------------------------------------
  transactionData.forEach((tx, index) => {
    const txDateRaw = String(tx[2]); // 예: "20260105"
    // 날짜 포맷 준비
    const dateNormal = formatDateString(txDateRaw, false); // "01/05"
    const dateWithYear = formatDateString(txDateRaw, true); // "26/01/05"

    const txType = tx[4];
    const txDesc = String(tx[5]).trim(); 
    const txAmount = Number(tx[6]);

    if (txType !== '입금') return;

    // 제외 로직
    if (IGNORE_LIST.some(name => txDesc.includes(name)) || txAmount >= IGNORE_LIMIT) {
      stats.skip++;
      return; 
    }

    // [매칭] 호수 찾기
    const matchedInfo = findBestMatchRow(curMaintData, txDesc);
    if (!matchedInfo) {
      stats.fail++;
      stats.log.push(`[실패] 매칭 불가: "${txDesc}" (${txAmount.toLocaleString()}원)`);
      return;
    }

    const unitName = curMaintData[matchedInfo.rowIndex][0]; // 호수
    
    // 인덱스 찾기
    const curMaintIdx = matchedInfo.rowIndex;
    const prevMaintIdx = findRowIndexByUnit(prevMaintData, unitName);
    const curRentIdx = findRowIndexByUnit(curRentData, unitName);
    const prevRentIdx = findRowIndexByUnit(prevRentData, unitName);

    // [분류] 임대료 vs 관리비
    let targetCategory = 'RENT';
    if (txAmount <= THRESHOLD_MAINT) {
      let isRentBackfill = false;
      if (curRentIdx !== -1 && checkIfRentGapMatches(curRentData[curRentIdx], txAmount)) isRentBackfill = true;
      if (!isRentBackfill && prevRentIdx !== -1 && checkIfRentGapMatches(prevRentData[prevRentIdx], txAmount)) isRentBackfill = true;
      if (!isRentBackfill) targetCategory = 'MAINT';
    }

    // =========================================================
    // [핵심 로직] 연도별 분기 및 이월 처리
    // =========================================================
    let moneyLeft = txAmount;

    if (targetCategory === 'MAINT') {
      // <관리비 처리>
      const hasDataInCurrentYear = checkHasData(curMaintData[curMaintIdx]);
      
      if (hasDataInCurrentYear) {
        // 올해 파일 기록
        const result = handleMaintenance(curMaintSheet, curMaintIdx, curMaintData[curMaintIdx], curMaintBg[curMaintIdx], moneyLeft, dateNormal);
        stats.maint++;
        stats.log.push(`[관리비-${CURRENT_YEAR.slice(2)}년] ${unitName}호 : ${result.msg}`);
        updateMaintMemory(curMaintData, curMaintBg, curMaintIdx, result.colIndex, moneyLeft, dateNormal, result.isPink);
      } else {
        // 작년 파일 기록
        if (prevMaintIdx !== -1) {
          const result = handleMaintenance(prevMaintSheet, prevMaintIdx, prevMaintData[prevMaintIdx], prevMaintBg[prevMaintIdx], moneyLeft, dateWithYear);
          stats.maint++;
          stats.log.push(`[관리비-${String(parseInt(CURRENT_YEAR)-1).slice(2)}년] ${unitName}호 : ${result.msg}`);
          updateMaintMemory(prevMaintData, prevMaintBg, prevMaintIdx, result.colIndex, moneyLeft, dateWithYear, result.isPink);
        } else {
           stats.fail++;
           stats.log.push(`[오류] ${unitName}호 작년 관리비 시트 행 없음`);
        }
      }

    } else {
      // <임대료 처리>
      const hasUnpaidInPrevYear = (prevRentIdx !== -1) ? checkHasUnpaidRent(prevRentData[prevRentIdx], prevRentBg[prevRentIdx]) : false;

      // 작년 미납이 있고 잔액도 있으면 -> 작년 파일 우선 변제
      if (hasUnpaidInPrevYear && moneyLeft > 0) {
        const result = handleRent(prevRentSheet, prevRentIdx, prevRentData[prevRentIdx], prevRentBg[prevRentIdx], moneyLeft, dateWithYear);
        stats.rent++;
        stats.log.push(`[임대료-${String(parseInt(CURRENT_YEAR)-1).slice(2)}년] ${unitName}호 : ${result.msg}`);
        moneyLeft = result.remaining;
      }

      // 돈이 남았으면 -> 올해 파일에 기록
      if (moneyLeft > 0) {
        if (curRentIdx !== -1) {
          const result = handleRent(curRentSheet, curRentIdx, curRentData[curRentIdx], curRentBg[curRentIdx], moneyLeft, dateNormal);
          stats.rent++;
          if(hasUnpaidInPrevYear) stats.log.push(`   └─ [이월처리-${CURRENT_YEAR.slice(2)}년] : ${result.msg}`);
          else stats.log.push(`[임대료-${CURRENT_YEAR.slice(2)}년] ${unitName}호 : ${result.msg}`);
        } else {
          stats.fail++;
          stats.log.push(`[주의] ${unitName}호 올해 임대료 행 없음 (잔액 ${moneyLeft.toLocaleString()}원)`);
        }
      }
    }
  });

  // 4. 결과 출력
  console.log("==========================================");
  console.log(`✅ 처리 완료! (임대료: ${stats.rent}건, 관리비: ${stats.maint}건, 실패: ${stats.fail}건)`);
  console.log("==========================================");
  console.log(stats.log.join('\n'));

  SpreadsheetApp.getUi().alert(
    `✅ 입금 내역 정리가 완료되었습니다.\n\n` +
    `임대료: ${stats.rent}건\n` +
    `관리비: ${stats.maint}건\n` +
    `실패/주의: ${stats.fail}건\n` +
    `(상세 내역은 실행 로그를 확인하세요)`
  );

  // 5. 현황판 강제 업데이트
  try {
     const dashSheet = ss.getSheetByName("월별 임대료 납부 현황");
     if(dashSheet) {
       const currentMonth = dashSheet.getRange("A1").getValue();
       if(typeof updateAndSortDashboard === 'function') updateAndSortDashboard(currentMonth);
     }
  } catch(e) {
    console.log("현황판 업데이트 중 경미한 오류: " + e.message);
  }
}

// ==========================================
// [3] Helper Functions (로직 처리용) - 기존과 동일
// ==========================================
// (아래 헬퍼 함수들은 기존 RentManage.gs 파일의 내용을 그대로 두시면 됩니다.)
// 만약 전체를 복사하신다면, 기존 파일의 [3] Helper Functions 부분도 모두 포함되어야 합니다.
// (편의상 생략하지 않고 핵심 헬퍼 함수들만 다시 적어드립니다. 기존 파일 내용을 그대로 쓰셔도 무방합니다.)

function getCleanDateString(cellValue) {
  if (!cellValue) return "";
  if (Object.prototype.toString.call(cellValue) === '[object Date]') {
    return Utilities.formatDate(cellValue, Session.getScriptTimeZone(), "MM/dd");
  }
  return String(cellValue);
}

function handleRent(sheet, rowIndex, rowData, rowBg, totalAmount, dateStr) {
  const standardRent = Number(rowData[4]);
  let remainingMoney = totalAmount;
  let logMessages = [];
  let loopCount = 0;
  
  let lastDateColIndex = -1;
  for (let c = 5; c < rowData.length; c += 2) {
    if (rowData[c+1] !== '' && rowData[c+1] != null) lastDateColIndex = c;
  }
  let startCol = (lastDateColIndex === -1) ? 5 : lastDateColIndex;

  while (remainingMoney > 0 && loopCount < 12) {
    loopCount++;
    let targetCol = -1;
    for (let c = startCol; c < rowData.length; c += 2) {
      const bg = rowBg[c];
      const isWritable = (bg === '#ffffff' || bg === 'white' || bg === null || bg === '#fff2cc');
      const val = Number(rowData[c]);
      const isEmpty = (rowData[c] === '' || rowData[c] == null);
      const isPartial = (val > 0 && val < standardRent); 

      if (isWritable && (isEmpty || isPartial)) { targetCol = c; break; }
    }

    if (targetCol === -1) break;

    const currentVal = Number(rowData[targetCol]) || 0;
    const needed = standardRent - currentVal;
    const payNow = Math.min(remainingMoney, needed); 

    const amountCell = sheet.getRange(rowIndex + 1, targetCol + 1);
    const dateCell = sheet.getRange(rowIndex + 1, targetCol + 2);
    const newAmount = currentVal + payNow;
    
    const oldDate = getCleanDateString(rowData[targetCol + 1]);
    let newDate = dateStr;
    if (oldDate && oldDate !== '' && !oldDate.includes(dateStr)) {
      newDate = oldDate + ", " + dateStr;
    } else if (oldDate) {
      newDate = oldDate;
    }

    amountCell.setValue(newAmount);
    dateCell.setValue(newDate);
    rowData[targetCol] = newAmount;
    rowData[targetCol + 1] = newDate;

    const unpaid = standardRent - newAmount;
    if (unpaid > 0) {
      amountCell.setBackground('#FFF2CC').setNote(`[기준] ${standardRent.toLocaleString()}\n[미납] ${unpaid.toLocaleString()}`);
      rowBg[targetCol] = '#FFF2CC';
    } else {
      amountCell.setBackground(null).clearNote();
      rowBg[targetCol] = '#ffffff'; 
      startCol = targetCol + 2;
    }
    remainingMoney -= payNow; 
    logMessages.push(`${(targetCol-3)/2}월분`);
  }
  return { msg: logMessages.join(', '), remaining: remainingMoney };
}

function handleMaintenance(sheet, rowIndex, rowData, rowBg, amount, dateStr) {
  let lastPlannedCol = -1;
  for (let c = 2; c < rowData.length; c += 2) {
    if (rowData[c] !== '' && rowData[c] != null) lastPlannedCol = c;
  }
  if (lastPlannedCol === -1) return { colIndex: -1, isPink: false, msg: "청구내역 없음", remaining: amount };

  const targetCol = lastPlannedCol;
  const dateVal = rowData[targetCol + 1];
  let mode = (dateVal !== '' && dateVal != null) ? 'UPDATE' : 'NEW';
  
  const amountCell = sheet.getRange(rowIndex + 1, targetCol + 1);
  const dateCell = sheet.getRange(rowIndex + 1, targetCol + 2);
  const cellNote = amountCell.getNote();

  let plannedAmount = 0;
  let finalAmount = 0;
  let finalDate = dateStr;

  if (mode === 'NEW') {
    plannedAmount = Number(rowData[targetCol]);
    finalAmount = amount;
  } else { 
    const currentPaid = Number(rowData[targetCol]);
    const match = cellNote.match(/\[청구\]\s*([\d,]+)/);
    plannedAmount = match ? Number(match[1].replace(/,/g, '')) : currentPaid;
    finalAmount = currentPaid + amount;
    
    const oldDate = getCleanDateString(rowData[targetCol + 1]);
    finalDate = oldDate.includes(dateStr) ? oldDate : oldDate + ", " + dateStr;
  }

  const diff = finalAmount - plannedAmount;
  if (diff === 0) {
    amountCell.setValue(finalAmount).setBackground(null).clearNote(); 
    dateCell.setValue(finalDate);
    return { colIndex: targetCol, isPink: false, msg: "완납", remaining: 0 };
  } else {
    const note = `[청구] ${plannedAmount.toLocaleString()}\n[실납] ${finalAmount.toLocaleString()}\n[차액] ${diff.toLocaleString()}`;
    amountCell.setValue(finalAmount).setBackground('#FFC7CE').setNote(note); 
    dateCell.setValue(finalDate);
    return { colIndex: targetCol, isPink: true, msg: "불일치", remaining: 0 };
  }
}

function formatDateString(rawDate, showYear) {
  var str = String(rawDate).trim();
  if (!str || str.length < 8) return rawDate;
  const mm = str.substring(4, 6);
  const dd = str.substring(6, 8);
  if (showYear) {
    const yy = str.substring(2, 4);
    return `${yy}/${mm}/${dd}`;
  } else {
    return `${mm}/${dd}`;
  }
}

function checkHasData(rowData) {
  for (let c = 2; c < rowData.length; c++) {
    if (rowData[c] !== '' && rowData[c] != null) return true;
  }
  return false;
}

function checkHasUnpaidRent(rowData, rowBg) {
  const standardRent = Number(rowData[4]);
  if (!standardRent) return false;
  for (let c = 5; c < rowData.length; c += 2) {
    const bg = rowBg[c];
    const val = Number(rowData[c]);
    const isWritable = (bg === '#ffffff' || bg === 'white' || bg === null || bg === '#fff2cc');
    if (isWritable && (rowData[c] === '' || rowData[c] == null || (val > 0 && val < standardRent))) {
      return true;
    }
  }
  return false;
}

function checkIfRentGapMatches(rowData, txAmount) {
  const standardRent = Number(rowData[4]); 
  if (!standardRent) return false;
  for (let c = 5; c < rowData.length; c += 2) {
    const val = Number(rowData[c]);
    if (val > 0 && val < standardRent) {
      if ((standardRent - val) === txAmount) return true;
    }
  }
  return false;
}

function findRowIndexByUnit(sheetData, targetUnit) {
  const cleanTarget = String(targetUnit).replace(/[^0-9a-zA-Z]/g, '').toUpperCase();
  for (let i = 1; i < sheetData.length; i++) {
    const currentUnit = String(sheetData[i][0]).replace(/[^0-9a-zA-Z]/g, '').toUpperCase();
    if (cleanTarget === currentUnit) return i;
  }
  return -1;
}

function findBestMatchRow(sheetData, txDesc) {
  let bestScore = 0;
  let bestRowIndex = -1; let bestReason = "";
  const inputRaw = txDesc.toUpperCase(); 
  const inputNumbers = inputRaw.match(/\d+/g) || [];
  const inputNameOnly = inputRaw.replace(/[0-9]/g, '').replace(/\s+/g, '');

  for (let i = 1; i < sheetData.length; i++) {
    const unitRaw = String(sheetData[i][0]).toUpperCase().trim();
    const nameStr = String(sheetData[i][1]); 
    const unitNumbers = unitRaw.match(/\d+/g) || [];
    const unitNumberStr = unitNumbers.length > 0 ? unitNumbers[0] : "";
    let score = 0; let reason = "";

    if (inputRaw === unitRaw || inputRaw === unitRaw + "호") {
      score = 100;
      reason = "호수일치";
    } else if (unitNumberStr !== "" && inputNumbers.includes(unitNumberStr)) {
      score = 90;
      reason = "호수숫자일치";
    }

    const names = nameStr.replace(/[()\/]/g, ',').split(',');
    names.forEach(n => {
      const keyName = n.trim().toUpperCase();
      const keyNameNoSpace = keyName.replace(/\s+/g, ''); 
      if (keyNameNoSpace.length < 2) return;
      if (inputNameOnly.includes(keyNameNoSpace)) {
         if (score < 95) { score = 95; reason = `이름일치(${keyName})`; }
      }
      else if (/^[A-Z]+$/.test(keyNameNoSpace) && keyNameNoSpace.startsWith(inputNameOnly) && inputNameOnly.length >= 3) {
         if (score < 88) { score = 88; reason = `이름부분일치(${keyName})`; }
      }
    });
    if (score > bestScore) { bestScore = score; bestRowIndex = i; bestReason = reason; }
  }
  return (bestScore > 0) ? { rowIndex: bestRowIndex, reason: bestReason } : null;
}

function updateMaintMemory(data, bg, r, c, amt, date, isPink) {
  if (c !== -1) {
      data[r][c] = amt;
      data[r][c+1] = date;
      if(isPink) bg[r][c] = '#FFC7CE';
      else bg[r][c] = '#ffffff';
  }
}