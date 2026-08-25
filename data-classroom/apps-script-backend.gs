// 데이터 수업 웹앱 - Google Apps Script 백엔드
// 1) 학교 Google 계정에서 새 Google Sheet를 만듭니다.
// 2) 확장 프로그램 > Apps Script에서 이 코드를 붙여넣습니다.
// 3) SHEET_ID, STUDENT_TOKEN, TEACHER_KEY를 설정합니다.
// 4) 배포 > 새 배포 > 웹 앱 > 실행 사용자: 나 / 액세스: 링크가 있는 모든 사용자(학교 정책에 맞게)로 배포합니다.

const SHEET_ID = '여기에_수집용_GOOGLE_SHEET_ID';
const SHEET_NAME = 'submissions';
const STUDENT_TOKEN = '수업용-토큰을-바꾸세요';
const TEACHER_KEY = '교사용-조회키를-길게-바꾸세요';

function sheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['timestamp','class','number','name','payload_json']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    if (String(data.token || '') !== STUDENT_TOKEN) return text_({ok:false,error:'invalid token'});
    const p = data.profile || {};
    if (!p.studentClass || !p.studentNo || !p.studentName) return text_({ok:false,error:'missing profile'});
    sheet_().appendRow([new Date(), p.studentClass, p.studentNo, p.studentName, JSON.stringify(data)]);
    return text_({ok:true});
  } catch (err) {
    return text_({ok:false,error:String(err)});
  }
}

function doGet(e) {
  try {
    const action = String(e.parameter.action || '');
    const callback = String(e.parameter.callback || '');
    if (action !== 'list') return jsonp_({ok:true,message:'data-classroom backend'}, callback);
    if (String(e.parameter.key || '') !== TEACHER_KEY) return jsonp_({ok:false,error:'invalid teacher key'}, callback);
    const sh = sheet_();
    const values = sh.getDataRange().getValues();
    const items = values.slice(1).filter(r => r[4]).map(r => ({timestamp:new Date(r[0]).toISOString(),className:r[1],number:r[2],name:r[3],payload:r[4]}));
    return jsonp_({ok:true,items:items}, callback);
  } catch (err) {
    return jsonp_({ok:false,error:String(err)}, String(e.parameter.callback || ''));
  }
}

function text_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function jsonp_(obj, callback) {
  const safe = /^[A-Za-z0-9_$.]+$/.test(callback) ? callback : '';
  if (!safe) return text_(obj);
  return ContentService.createTextOutput(safe + '(' + JSON.stringify(obj) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
}
