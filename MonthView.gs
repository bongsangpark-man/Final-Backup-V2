/**
 * [트리거 함수] 월 선택 시 OR 데이터(납부내역/원본현황) 수정 시 자동 업데이트
 * -> 트리거 설정: 이벤트 소스(스프레드시트), 이벤트 유형(수정 시)
 */
function autoUpdateRent(e) {
  var range = e.range;
  var sheet = range.getSheet();
  var sheetName = sheet.getName();
  
  // 1. '월별 임대료 납부 현황' 시트에서 '월(A1)'을 변경했을 때 실행
  if (sheetName === "월별 임대료 납부 현황" && range.getA1Notation() === "A1") {
    var selectedMonth = range.getValue();
    updateAndSortDashboard(selectedMonth);
  }

  // 2. [기능 확장] '임대료 납부내역' 또는 '임대 현황표(원본)' 수정 시 실행
  // -> 원본(임대 현황표)을 고치면 -> 납부내역의 A~E열(수식) 값이 바뀌고 -> 현황판도 즉시 업데이트됩니다.
  // ※ '임대 현황표'라는 이름은 실제 시트 탭 이름과 띄어쓰기까지 정확히 일치해야 합니다.
  if (sheetName === "임대료 납부내역" || sheetName === "임대 현황표") {
    
    var dashSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("월별 임대료 납부 현황");
    
    // 현황판 시트가 존재하는지 확인 (에러 방지)
    if (dashSheet) {
      var currentMonth = dashSheet.getRange("A1").getValue(); // 현황판에 선택된 월 확인
      updateAndSortDashboard(currentMonth); // 업데이트 실행
    }
  }
}

/**
 * [핵심 로직] 미납 현황 + 공실(배경색) + 구계약 예외(글자색) + 실시간 반영
 */
function updateAndSortDashboard(monthStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashSheet = ss.getSheetByName("월별 임대료 납부 현황");
  var dbSheet = ss.getSheetByName("임대료 납부내역");

  if (!dbSheet) return;

  // 1. 선택한 월 숫자 추출
  var currentMonth = parseInt(String(monthStr).replace(/[^0-9]/g, ''));
  if (!currentMonth || currentMonth < 1 || currentMonth > 12) return;

  // 2. DB 데이터 가져오기
  var lastRowDB = dbSheet.getLastRow();
  var maxColDB = dbSheet.getLastColumn();
  
  // 데이터가 없는 경우 방지
  if (lastRowDB < 2) return;

  var dbRange = dbSheet.getRange(2, 1, lastRowDB - 1, maxColDB);
  var dbValues = dbRange.getValues();
  var dbBackgrounds = dbRange.getBackgrounds(); // 배경색 (공실 확인용)
  var dbFontColors = dbRange.getFontColors();   // 글자색 (예외 금액 확인용)

  var processedData = [];

  // 3. 데이터 가공
  for (var i = 0; i < dbValues.length; i++) {
    var row = dbValues[i];
    var rowBackgrounds = dbBackgrounds[i];
    var rowFontColors = dbFontColors[i]; 
    
    var unit = row[0];        // A열: 호수
    var name = row[1];        // B열: 이름
    var payDay = row[2];      // C열: 납부일
    var rentTypeRaw = row[3]; // D열: 임대유형
    var stdAmount = row[4];   // E열: 기준 월세
    
    var dbRowIndex = i + 2; 

    // [필터링] 전세, 공실 이름 제외
    if (rentTypeRaw === "전세" || name === "공실") continue;

    var rentTypeClean = String(rentTypeRaw).replace(/월세|\(|\)/g, "").trim();

    // ----------------------------------------------------
    // [A] 선택 월의 납부유무 수식 (F열)
    // ----------------------------------------------------
    var targetDateColIdx = (currentMonth * 2) + 4; 
    var targetColLetter = getColumnLetter(targetDateColIdx + 1); 
    var statusFormula = "=IF('임대료 납부내역'!" + targetColLetter + dbRowIndex + "<>\"\",\"○\",\"\")";

    // ----------------------------------------------------
    // [B] 미납 현황 계산 (G열) - 과거 내역 스캔
    // ----------------------------------------------------
    var unpaidList = []; 

    for (var m = 1; m < currentMonth; m++) {
      var mAmountIdx = (m * 2) + 3; // 금액 열
      var mDateIdx = (m * 2) + 4;   // 날짜 열

      var mAmount = row[mAmountIdx];  
      var mDate = row[mDateIdx];      
      
      var mDateBgColor = rowBackgrounds[mDateIdx]; // 날짜칸 배경색 (공실체크)
      var mAmountFontColor = rowFontColors[mAmountIdx]; // 금액칸 글자색 (예외체크)
      
      // 1. 공실 체크: 날짜칸 배경이 흰색(#ffffff)이 아니면 패스
      if (mDateBgColor !== "#ffffff") continue; 

      // 2. 날짜가 없으면 -> "미납"
      if (mDate === "" || mDate == null) {
        unpaidList.push(m + "월 미납");
      } 
      else {
        // 3. 날짜가 있는데 금액을 확인해야 하는 경우
        // [예외] 금액칸의 글자 색이 검정(#000000)이 아니면 정상 납부로 간주
        if (mAmountFontColor !== "#000000") {
          continue; 
        }
        
        // 글자색이 검정이면 엄격하게 금액 비교
        if (typeof mAmount === 'number' && typeof stdAmount === 'number') {
           if (mAmount > 0 && mAmount < stdAmount) {
              unpaidList.push(m + "월 일부미납");
           }
        }
      }
    }

    var unpaidStatus = unpaidList.join(", ");

    processedData.push([unit, name, rentTypeClean, payDay, stdAmount, statusFormula, unpaidStatus]);
  }

  // 4. 정렬: 납부일 오름차순
  processedData.sort(function(a, b) {
    var dayA = parseInt(String(a[3]).replace(/[^0-9]/g, '')) || 99;
    var dayB = parseInt(String(b[3]).replace(/[^0-9]/g, '')) || 99;
    return dayA - dayB; 
  });

  // 5. 결과 입력 (4행부터, A~G열만)
  var startRow = 4; 
  var lastRowDash = dashSheet.getLastRow();
  
  // 기존 데이터 초기화 (4행부터 끝까지)
  if (lastRowDash >= startRow) {
    dashSheet.getRange(startRow, 1, lastRowDash - startRow + 1, 7).clearContent();
  }

  // 새 데이터 입력
  if (processedData.length > 0) {
    dashSheet.getRange(startRow, 1, processedData.length, 7).setValues(processedData);
  }
}

/**
 * [보조 함수] 열 번호를 알파벳으로 변환 (1 -> A, 2 -> B ...)
 */
function getColumnLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - 1) / 26 | 0;
  }
  return letter;
}