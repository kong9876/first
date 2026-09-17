// 데이터 수업 웹앱 - Google Apps Script 백엔드
const SHEET_ID = '여기에_수집용_GOOGLE_SHEET_ID';
const SHEET_NAME = 'submissions';
const SHARE_SHEET_NAME = 'shares';
const LIKE_SHEET_NAME = 'likes';
const STUDENT_TOKEN = '수업용-토큰을-바꾸세요';
const TEACHER_KEY = '교사용-조회키를-길게-바꾸세요';
function ss_(){ return SpreadsheetApp.openById(SHEET_ID); }
function sheet_(){const ss=ss_();let sh=ss.getSheetByName(SHEET_NAME);if(!sh)sh=ss.insertSheet(SHEET_NAME);if(sh.getLastRow()===0){sh.appendRow(['timestamp','class','number','name','payload_json']);sh.setFrozenRows(1)}return sh}
function shareSheet_(){const ss=ss_();let sh=ss.getSheetByName(SHARE_SHEET_NAME);if(!sh)sh=ss.insertSheet(SHARE_SHEET_NAME);if(sh.getLastRow()===0){sh.appendRow(['id','timestamp','class','number','name','lesson','kind','title','text','extra_json']);sh.setFrozenRows(1)}return sh}
function likeSheet_(){const ss=ss_();let sh=ss.getSheetByName(LIKE_SHEET_NAME);if(!sh)sh=ss.insertSheet(LIKE_SHEET_NAME);if(sh.getLastRow()===0){sh.appendRow(['timestamp','post_id','class','number','name']);sh.setFrozenRows(1)}return sh}
function validProfile_(p){return p&&p.studentClass&&p.studentNo&&p.studentName}
function maskName_(name){name=String(name||'').trim();if(!name)return '';if(name.length===1)return name+'○';return name.charAt(0)+'○'.repeat(Math.min(2,name.length-1))}
function doPost(e){try{const data=JSON.parse(e.postData.contents||'{}');if(String(data.token||'')!==STUDENT_TOKEN)return text_({ok:false,error:'invalid token'});const p=data.profile||{};if(!validProfile_(p))return text_({ok:false,error:'missing profile'});const type=String(data.type||'submission');if(type==='share')return saveShare_(data,p);if(type==='like')return saveLike_(data,p);sheet_().appendRow([new Date(),p.studentClass,p.studentNo,p.studentName,JSON.stringify(data)]);return text_({ok:true})}catch(err){return text_({ok:false,error:String(err)})}}
function saveShare_(data,p){const lesson=Number(data.lesson||0),kind=String(data.kind||'idea').slice(0,30),title=String(data.title||'').trim().slice(0,80),text=String(data.text||'').trim().slice(0,500);if(!lesson||!text)return text_({ok:false,error:'missing share content'});const id=Utilities.getUuid();shareSheet_().appendRow([id,new Date(),p.studentClass,p.studentNo,p.studentName,lesson,kind,title,text,JSON.stringify(data.extra||{})]);return text_({ok:true,id:id})}
function saveLike_(data,p){const postId=String(data.postId||'').trim();if(!postId)return text_({ok:false,error:'missing post id'});const lock=LockService.getScriptLock();lock.waitLock(5000);try{const sh=likeSheet_(),values=sh.getDataRange().getValues();const exists=values.slice(1).some(r=>String(r[1])===postId&&String(r[2])===String(p.studentClass)&&String(r[3])===String(p.studentNo));if(!exists)sh.appendRow([new Date(),postId,p.studentClass,p.studentNo,p.studentName]);return text_({ok:true,added:!exists})}finally{lock.releaseLock()}}
function doGet(e){try{const action=String(e.parameter.action||''),callback=String(e.parameter.callback||'');
  if(action==='assessmentStatus')return jsonp_(assessmentStatus_(e),callback);
  if(action==='assessmentControl')return jsonp_(assessmentControl_(e),callback);
  if(action==='assessmentEvent')return jsonp_(assessmentEvent_(e),callback);
  if(action==='assessmentMonitor')return jsonp_(assessmentMonitor_(e),callback);
  if(action==='shares'){if(String(e.parameter.token||'')!==STUDENT_TOKEN)return jsonp_({ok:false,error:'invalid token'},callback);return jsonp_({ok:true,items:shareItems_(String(e.parameter.className||''),Number(e.parameter.lesson||0),false)},callback)}
  if(action==='sharesTeacher'){if(String(e.parameter.key||'')!==TEACHER_KEY)return jsonp_({ok:false,error:'invalid teacher key'},callback);return jsonp_({ok:true,items:shareItems_(String(e.parameter.className||''),Number(e.parameter.lesson||0),true)},callback)}
  if(action!=='list')return jsonp_({ok:true,message:'data-classroom backend'},callback);
  if(String(e.parameter.key||'')!==TEACHER_KEY)return jsonp_({ok:false,error:'invalid teacher key'},callback);
  const sh=sheet_(),values=sh.getDataRange().getValues(),items=values.slice(1).filter(r=>r[4]).map(r=>({timestamp:new Date(r[0]).toISOString(),className:r[1],number:r[2],name:r[3],payload:r[4]}));return jsonp_({ok:true,items},callback)
}catch(err){return jsonp_({ok:false,error:String(err)},String(e.parameter.callback||''))}}
function shareItems_(className,lesson,teacher){const sh=shareSheet_(),rows=sh.getDataRange().getValues().slice(1),likes=likeSheet_().getDataRange().getValues().slice(1),count={};likes.forEach(r=>count[String(r[1])]=(count[String(r[1])]||0)+1);return rows.filter(r=>(!className||String(r[2])===className)&&(!lesson||Number(r[5])===lesson)).slice(-200).reverse().map(r=>({id:String(r[0]),timestamp:new Date(r[1]).toISOString(),className:String(r[2]),number:String(r[3]),name:teacher?String(r[4]):maskName_(r[4]),lesson:Number(r[5]),kind:String(r[6]),title:String(r[7]),text:String(r[8]),extra:r[9]?JSON.parse(r[9]):{},likes:count[String(r[0])]||0}))}
function text_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON)}
function jsonp_(obj,callback){const safe=/^[A-Za-z0-9_$.]+$/.test(callback)?callback:'';if(!safe)return text_(obj);return ContentService.createTextOutput(safe+'('+JSON.stringify(obj)+')').setMimeType(ContentService.MimeType.JAVASCRIPT)}

// 수행평가 확장 기능
function assessmentProps_(){return PropertiesService.getScriptProperties()}
function assessmentMonitorSheet_(){const ss=ss_();let sh=ss.getSheetByName('assessment_monitor');if(!sh){sh=ss.insertSheet('assessment_monitor');sh.appendRow(['timestamp','class','number','name','event','leave_count','detail']);sh.setFrozenRows(1)}return sh}
function assessmentStatus_(e){const p=e.parameter||{},teacher=String(p.key||'')===String(TEACHER_KEY),student=String(p.token||'')===String(STUDENT_TOKEN);if(!teacher&&!student)return {ok:false,error:'invalid access',open:false};return {ok:true,open:assessmentProps_().getProperty('ASSESSMENT_OPEN')==='1'}}
function assessmentControl_(e){const p=e.parameter||{};if(String(p.key||'')!==String(TEACHER_KEY))return {ok:false,error:'invalid teacher key',open:false};const open=String(p.open||'')==='1';assessmentProps_().setProperty('ASSESSMENT_OPEN',open?'1':'0');return {ok:true,open}}
function assessmentEvent_(e){const p=e.parameter||{};if(String(p.token||'')!==String(STUDENT_TOKEN))return {ok:false,error:'invalid token'};const c=String(p.studentClass||'').trim().slice(0,10),no=String(p.studentNo||'').trim().slice(0,4),name=String(p.studentName||'').trim().slice(0,20);if(!c||!no||!name)return {ok:false,error:'missing student'};const event=String(p.event||'').slice(0,30),count=Math.max(0,Number(p.count||0)||0),detail=String(p.detail||'').slice(0,120);assessmentMonitorSheet_().appendRow([new Date(),c,no,name,event,count,detail]);return {ok:true}}
function assessmentMonitor_(e){const p=e.parameter||{};if(String(p.key||'')!==String(TEACHER_KEY))return {ok:false,error:'invalid teacher key'};const sh=assessmentMonitorSheet_(),last=sh.getLastRow();if(last<2)return {ok:true,events:[]};const start=Math.max(2,last-199),rows=sh.getRange(start,1,last-start+1,7).getValues();return {ok:true,events:rows.map(r=>({time:r[0] instanceof Date?r[0].toISOString():String(r[0]),studentClass:String(r[1]),studentNo:String(r[2]),studentName:String(r[3]),event:String(r[4]),count:Number(r[5])||0,detail:String(r[6]||'')}))}}
