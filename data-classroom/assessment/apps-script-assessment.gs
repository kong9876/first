// 수행평가 웹앱용 Google Apps Script 확장 코드
// 기존 data-classroom/apps-script-backend.gs와 같은 Apps Script 프로젝트에 추가합니다.
// 기존 상수 SHEET_ID, STUDENT_TOKEN, TEACHER_KEY 및 jsonp_()를 사용합니다.

function assessmentProps_(){ return PropertiesService.getScriptProperties(); }

function assessmentMonitorSheet_(){
  var ss=SpreadsheetApp.openById(SHEET_ID);
  var sh=ss.getSheetByName('assessment_monitor');
  if(!sh){
    sh=ss.insertSheet('assessment_monitor');
    sh.appendRow(['timestamp','class','number','name','event','leave_count','detail']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function assessmentStatus_(e){
  var p=e.parameter||{};
  var teacher=String(p.key||'')===String(TEACHER_KEY);
  var student=String(p.token||'')===String(STUDENT_TOKEN);
  if(!teacher&&!student) return {ok:false,error:'invalid access',open:false};
  return {ok:true,open:assessmentProps_().getProperty('ASSESSMENT_OPEN')==='1'};
}

function assessmentControl_(e){
  var p=e.parameter||{};
  if(String(p.key||'')!==String(TEACHER_KEY)) return {ok:false,error:'invalid teacher key',open:false};
  var open=String(p.open||'')==='1';
  assessmentProps_().setProperty('ASSESSMENT_OPEN',open?'1':'0');
  return {ok:true,open:open};
}

function assessmentEvent_(e){
  var p=e.parameter||{};
  if(String(p.token||'')!==String(STUDENT_TOKEN)) return {ok:false,error:'invalid token'};
  var c=String(p.studentClass||'').trim().slice(0,10);
  var no=String(p.studentNo||'').trim().slice(0,4);
  var name=String(p.studentName||'').trim().slice(0,20);
  if(!c||!no||!name) return {ok:false,error:'missing student'};
  var event=String(p.event||'').trim().slice(0,30);
  var count=Math.max(0,Number(p.count||0)||0);
  var detail=String(p.detail||'').trim().slice(0,120);
  assessmentMonitorSheet_().appendRow([new Date(),c,no,name,event,count,detail]);
  return {ok:true};
}

function assessmentMonitor_(e){
  var p=e.parameter||{};
  if(String(p.key||'')!==String(TEACHER_KEY)) return {ok:false,error:'invalid teacher key'};
  var sh=assessmentMonitorSheet_();
  var last=sh.getLastRow();
  if(last<2) return {ok:true,events:[]};
  var start=Math.max(2,last-199);
  var rows=sh.getRange(start,1,last-start+1,7).getValues();
  return {ok:true,events:rows.map(function(r){
    return {time:r[0] instanceof Date?r[0].toISOString():String(r[0]),studentClass:String(r[1]),studentNo:String(r[2]),studentName:String(r[3]),event:String(r[4]),count:Number(r[5])||0,detail:String(r[6]||'')};
  })};
}

// 기존 doGet(e)의 action을 읽은 직후 아래 4개 분기를 추가해야 합니다.
// if(action==='assessmentStatus') return jsonp_(assessmentStatus_(e),callback);
// if(action==='assessmentControl') return jsonp_(assessmentControl_(e),callback);
// if(action==='assessmentEvent') return jsonp_(assessmentEvent_(e),callback);
// if(action==='assessmentMonitor') return jsonp_(assessmentMonitor_(e),callback);
