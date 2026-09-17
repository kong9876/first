// 수행평가 웹앱용 Google Apps Script
// 기존 데이터 수업 Apps Script 프로젝트에 아래 내용을 추가하세요.
// 기존 코드의 STUDENT_TOKEN, TEACHER_KEY를 그대로 사용합니다.
// 배포 후 기존 웹앱 URL을 그대로 assessment/index.html에서 사용합니다.

function assessmentProps_(){ return PropertiesService.getScriptProperties(); }

function assessmentStatus_(e){
  var p=e.parameter||{};
  var isTeacher=String(p.key||'')===String(TEACHER_KEY);
  var isStudent=String(p.token||'')===String(STUDENT_TOKEN);
  if(!isTeacher&&!isStudent) return {ok:false,error:'invalid access',open:false};
  return {ok:true,open:assessmentProps_().getProperty('ASSESSMENT_OPEN')==='1'};
}

function assessmentControl_(e){
  var p=e.parameter||{};
  if(String(p.key||'')!==String(TEACHER_KEY)) return {ok:false,error:'invalid teacher key',open:false};
  var open=String(p.open||'')==='1';
  assessmentProps_().setProperty('ASSESSMENT_OPEN',open?'1':'0');
  return {ok:true,open:open};
}

// 중요: 기존 doGet(e)의 try 블록에서 action을 읽은 직후 아래 두 줄을 추가하세요.
// if(action==='assessmentStatus') return jsonp_(assessmentStatus_(e),callback);
// if(action==='assessmentControl') return jsonp_(assessmentControl_(e),callback);
