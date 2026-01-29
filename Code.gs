/**
 * 파일명: Code.gs
 * 기능: 수협은행 입금 내역 관리, 장부 생성/이월 자동화, 부가세 자료 생성 메뉴 연결
 */

// ==========================================
// [1] 사용자 설정
// ==========================================
const CLIENT_ID = '6343ea64-775d-465f-ac57-ac19e2288b79';
const CLIENT_SECRET = '3923fa02-99a7-46ed-8ac8-70c2434ab04a';

const USER_IDENTITY = '5906061'; 
const INITIAL_START_DATE = '20251215'; 

const CERT_FILE_ID = '1z3D025lX08a4BIM_myZX5rDyn53A9sc1'; 
const KEY_FILE_ID = '1NwQiDe1kbPZr3WY6yBFLywAfSK-G1a1F';
const ACCOUNTS_INFO = [
  {"bank_name": "수협은행", "code": "0007", "account": "201009440236"},
];

// ==========================================
// [2] 메뉴 생성
// ==========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // 1. [🏦 입금관리]
  ui.createMenu('🏦 입금관리')
    .addItem('입금 내역 가져오기', 'main')
    .addItem('[관리자용] 날짜 초기화', 'resetDate')
    .addToUi();

  // 2. [🏢 임대현황관리]
  ui.createMenu('🏢 임대현황관리')
    .addItem('임대 관리 시스템 열기', 'showRentalSidebar') 
    .addSeparator() 
    .addItem('🔒 시트 잠금 (수정 방지)', 'lockRentalSheet') 
    .addItem('🔓 시트 잠금 해제', 'unlockRentalSheet')     
    .addToUi();

  // 3. [📂 장부만들기]
  ui.createMenu('📂 장부만들기')
    .addItem('📅 금년 장부 생성하기(1월 2일 이후 생성)', 'createNextYearSheet')
    .addSeparator()
    .addItem('⚙️ 자동화(트리거) 생성하기', 'setupTriggersForNewYear')
    .addToUi();

  // 4. [📊 부가세 신고자료] (★수정됨: 사이드바 연결)
  ui.createMenu('📊 부가세 신고자료')
    .addItem('부가세 메뉴 열기', 'showVatSidebar') 
    .addToUi();
}

// ==========================================
// [3] 날짜 유틸리티
// ==========================================
function getFormatDate(dateObj) {
  const yyyy = dateObj.getFullYear();
  const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
  const dd = String(dateObj.getDate()).padStart(2, '0');
  return `${yyyy}${mm}${dd}`;
}

function parseDateStr(dateStr) {
  const y = parseInt(dateStr.substring(0, 4));
  const m = parseInt(dateStr.substring(4, 6)) - 1;
  const d = parseInt(dateStr.substring(6, 8));
  return new Date(y, m, d);
}

// ==========================================
// [4] Codef 클래스
// ==========================================
class Codef {
  constructor() { this.accessToken = ''; }

  requestToken(id, secret) {
    const url = "https://oauth.codef.io/oauth/token";
    const auth = Utilities.base64Encode(`${id}:${secret}`);
    try {
      const res = UrlFetchApp.fetch(url, {
        method: 'post',
        headers: {'Authorization': `Basic ${auth}`, 'Content-Type': 'application/x-www-form-urlencoded'},
        payload: {'grant_type': 'client_credentials', 'scope': 'read'},
        muteHttpExceptions: true
      });
      if (res.getResponseCode() === 200) {
        this.accessToken = JSON.parse(res.getContentText()).access_token;
        return true;
      }
    } catch (e) {}
    return false;
  }

  encryptPassword(plainText) {
    try {
      var encrypt = new JSEncrypt();
      const publicKey = `-----BEGIN PUBLIC KEY-----
      MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAjlX+sETy9SLvJdFnv4StNj5kKvrYcOIuQ2i6X+/AGJtLlfj/Tf8YeeDh9mnDaY4zf116/Up0FEqdNNpWEKdeniNVlZxLPCX97qdiFK59NJfa5pnZ+m/xixLcK8K+TxVNuEs5nkArD8RltL0XAIftbVZqYn5lwW2S+ykpwUZ7XS7u7fWMXFmo1S4AxD+YfgUWriXCrmsvKp8ZQpGUh+1MC+MHm34wjiItK5nVz3BmREpHxzeUS18V5ZgEsjRFVfYoxg/eLHLYgSuyROO4x5/yCkKH4pYG+S14N/oZt0wYyw/JcYKrUHoxZCCst6+RMp2F2CPWwg/HM3jHEqm+rGTlmQIDAQAB
      -----END PUBLIC KEY-----`;
      encrypt.setPublicKey(publicKey);
      return encrypt.encrypt(plainText);
    } catch (e) { return null; }
  }

  getFileBase64(fileId) {
    try { return Utilities.base64Encode(DriveApp.getFileById(fileId).getBlob().getBytes()); } 
    catch (e) { return null; }
  }

  createAccountCert(bankCode, account, encPw, identity, der, key) {
    const param = {
      "accountList": [{
        "countryCode": "KR", "businessType": "BK", "clientType": "P",
        "organization": bankCode, "loginType": "0", "certType": "1", 
        "derFile": der, "keyFile": key, "password": encPw, "identity": identity, "id": ""
      }]
    };
    return this.requestProduct("/v1/account/create", 1, param);
  }
  
  requestProduct(urlPath, serviceType, param) {
    if (!this.accessToken) return null;
    let domain = serviceType === 0 ? "https://api.codef.io" : "https://sandbox.codef.io"; 
    if (serviceType === 1) domain = "https://development.codef.io";
    if (!param.organization) param.organization = "";
    const options = {
      method: 'post',
      headers: {'Authorization': `Bearer ${this.accessToken}`, 'Content-Type': 'application/json'},
      payload: JSON.stringify(param),
      muteHttpExceptions: true
    };
    const res = UrlFetchApp.fetch(domain + (urlPath.startsWith('/') ? urlPath : '/' + urlPath), options);
    return decodeURIComponent(res.getContentText().replace(/\+/g, ' '));
  }
}

// ==========================================
// [5] 메인 실행 함수 (HTML 팝업)
// ==========================================
function main() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  let lastScanDateStr = props.getProperty('LAST_SCAN_DATE');
  if (!lastScanDateStr) {
    lastScanDateStr = INITIAL_START_DATE; 
    props.setProperty('LAST_SCAN_DATE', INITIAL_START_DATE);
  }

  const lastScanDateObj = parseDateStr(lastScanDateStr);
  const startDateObj = new Date(lastScanDateObj);
  startDateObj.setDate(startDateObj.getDate() + 1);

  const todayObj = new Date();
  todayObj.setDate(todayObj.getDate() - 1);
  if (startDateObj > todayObj) {
    ui.alert(`✅ 이미 최신 상태입니다.\n(마지막 조회: ${lastScanDateStr})`);
    return;
  }

  const html = HtmlService.createHtmlOutputFromFile('PasswordForm').setWidth(400).setHeight(250);
  ui.showModalDialog(html, ' ');
}

// ==========================================
// [6] 실제 조회 로직
// ==========================================
function runScraping(certPw) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  let lastScanDateStr = props.getProperty('LAST_SCAN_DATE');
  const lastScanDateObj = parseDateStr(lastScanDateStr);
  const startDateObj = new Date(lastScanDateObj);
  startDateObj.setDate(startDateObj.getDate() + 1);
  const targetStartDate = getFormatDate(startDateObj);

  const todayObj = new Date();
  todayObj.setDate(todayObj.getDate() - 1);
  const targetEndDate = getFormatDate(todayObj);

  const codef = new Codef();
  if (!codef.requestToken(CLIENT_ID, CLIENT_SECRET)) throw new Error('토큰 발급 실패');
  
  const derData = codef.getFileBase64(CERT_FILE_ID);
  const keyData = codef.getFileBase64(KEY_FILE_ID);
  const encCertPw = codef.encryptPassword(certPw);
  if(!derData || !keyData || !encCertPw) throw new Error('파일 로딩 또는 암호화 실패');

  let outputData = [];
  let log = "";
  
  for (let i = 0; i < ACCOUNTS_INFO.length; i++) {
    const info = ACCOUNTS_INFO[i];
    const createRes = codef.createAccountCert(info.code, info.account, encCertPw, USER_IDENTITY, derData, keyData);
    const createJson = JSON.parse(createRes);
    let connectedId = '';
    
    if (createJson.result.code === 'CF-00000') connectedId = createJson.data.connectedId;
    else log += `⚠️ [${info.account}] 등록 실패: ${createJson.result.message}\n`;

    if(connectedId) {
        const param = {
            "organization": info.code, "connectedId": connectedId, "account": info.account,
            "startDate": targetStartDate, "endDate": targetEndDate,
            "inquiryType": "1", "orderBy": "0"
        };
        const resText = codef.requestProduct("/v1/kr/bank/p/account/transaction-list", 1, param);
        const resJson = JSON.parse(resText);
        if (resJson.result.code === 'CF-00000') {
             let txList = resJson.data.resTrHistoryList || resJson.data || [];
             if (!Array.isArray(txList)) txList = [txList];
             txList.forEach(tx => {
                 const depositAmt = Number(tx.resAccountIn);
                 if (depositAmt > 0) {
                     outputData.push([
                         info.bank_name, info.account, tx.resAccountTrDate, tx.resAccountTrTime,
                         '입금', tx.resAccountDesc3 || tx.resUserNm, depositAmt
                     ]);
                 }
             });
        } else {
             log += `❌ [${info.account}] 조회 오류: ${resJson.result.message} (${resJson.result.code})\n`;
        }
    }
  }

  if (outputData.length > 0) {
    const sheetName = `입금_${targetStartDate}`;
    const newSheet = ss.insertSheet(sheetName);
    newSheet.appendRow(["은행", "계좌번호", "날짜", "시간", "구분", "적요", "금액"]);
    outputData.sort((a, b) => a[2].localeCompare(b[2]) || a[3].localeCompare(b[3]));
    newSheet.getRange(2, 1, outputData.length, outputData[0].length).setValues(outputData);
    newSheet.getRange(1, 1, 1, 7).setBackground("#fff2cc").setFontWeight("bold");
    newSheet.setColumnWidth(6, 150);
    ss.setActiveSheet(newSheet);

    try {
      const response = ui.alert('🏦 은행 조회 완료', `총 ${outputData.length}건을 가져왔습니다.\n[임대관리대장]에 자동 반영하시겠습니까?`, ui.ButtonSet.YES_NO);
      if (response == ui.Button.YES) processRentTransactions(outputData); 
    } catch (e) {
      console.log("RentManager 연동 중 오류 발생: " + e.message);
    }
    props.setProperty('LAST_SCAN_DATE', targetEndDate);
  } else {
    if (log) ui.alert(`⚠️ 조회 중 오류가 발생했습니다.\n\n${log}`);
    else {
      props.setProperty('LAST_SCAN_DATE', targetEndDate);
      ui.alert(`ℹ️ 알림\n\n${targetStartDate} ~ ${targetEndDate}\n기간 내에 새로운 입금 내역이 없습니다.\n(날짜는 최신으로 업데이트 되었습니다.)`);
    }
  }
}

function resetDate() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('LAST_SCAN_DATE');
  SpreadsheetApp.getUi().alert(`🔄 날짜 초기화 완료. (${INITIAL_START_DATE})부터 다시 시작합니다.`);
}


// ==========================================
// [7] 금년 장부 자동 생성 & 트리거 삭제 & 바로가기
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

    // 3. 데이터 초기화
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
    newSS.getSheets().forEach(s => {
      if (s.getName().startsWith('입금_')) newSS.deleteSheet(s);
    });

    // ★ 4. [중요] 현재 파일(구 장부)의 만기 알림 트리거 삭제
    const allTriggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;
    for (let i = 0; i < allTriggers.length; i++) {
      if (allTriggers[i].getHandlerFunction() === 'sendExtensionCheckEmails') {
        ScriptApp.deleteTrigger(allTriggers[i]); // 트리거 삭제
        deletedCount++;
      }
    }
    console.log(`기존 파일에서 알림 트리거 ${deletedCount}개 삭제됨.`);

    // 5. 생성 완료 팝업 (바로가기 버튼)
    const htmlOutput = HtmlService.createHtmlOutput(
      `<div style="font-family: sans-serif; padding: 10px; text-align: center;">` +
      `  <h3 style="margin-top: 0; color: #188038;">✅ 생성 완료!</h3>` +
      `  <p>새로운 ${nextYear}년 장부 파일이 생성되었습니다.</p>` +
      `  <p>현재 파일(${currentYear})의 자동 이메일 발송은 <strong>중단</strong>되었습니다.</p>` +
      `  <p style="background: #f1f3f4; padding: 10px; border-radius: 5px; font-size: 13px;">` +
      `    <strong>파일명:</strong> ${newFileName}` +
      `  </p>` +
      `  <div style="margin-top: 20px;">` +
      `    <a href="${newUrl}" target="_blank" style="background-color: #1a73e8; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px; font-weight: bold; display: inline-block;">` +
      `      🚀 새 장부로 이동하기` +
      `    </a>` +
      `  </div>` +
      `  <p style="margin-top: 20px; font-size: 12px; color: #666;">` +
      `    * 새 파일로 이동 후 [📂 장부만들기] > [⚙️ 자동화 생성]을 꼭 눌러주세요!` +
      `  </p>` +
      `</div>`
    ).setWidth(400).setHeight(350);

    ui.showModalDialog(htmlOutput, '장부 생성 결과');

  } catch (e) {
    ui.alert('오류 발생', e.toString(), ui.ButtonSet.OK);
  }
}

function setupTriggersForNewYear() {
  const ui = SpreadsheetApp.getUi();
  
  // 1. 기존 트리거 초기화
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  // 2. 트리거 생성
  try {
    // (A) 만기 알림 (오전 9시)
    ScriptApp.newTrigger('sendExtensionCheckEmails')
      .timeBased()
      .atHour(9)
      .everyDays(1)
      .create();

    // (B) 현황판 업데이트
    ScriptApp.newTrigger('autoUpdateRent')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();

    ui.alert(
      '✅ 자동화 설정 완료', 
      '다음 기능이 활성화되었습니다:\n\n' +
      '1. 계약 만기 알림 메일 (매일 09~10시)\n' +
      '2. 현황판 실시간 업데이트', 
      ui.ButtonSet.OK
    );

  } catch (e) {
    ui.alert('설정 실패', '권한이 부족하거나 오류가 발생했습니다.\n' + e.toString(), ui.ButtonSet.OK);
  }
}

// ==========================================
// [8] 부가세 전용 사이드바 호출 (★신규 추가)
// ==========================================
function showVatSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('VatSidebar')
    .setTitle('📊 부가세 신고 관리')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}