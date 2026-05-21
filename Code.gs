// ═══════════════════════════════════════════════════════════════
//  후각 적응도 평가 시스템 — Google Apps Script 백엔드 v2
//  (사전 설문 포함)
// ═══════════════════════════════════════════════════════════════

const SHEET_EVAL   = '평가데이터';
const SHEET_SURVEY = '사전설문';
const CHECKPOINTS  = [1,2,3,4,5,6,7,8,9,10];

// ── 초기 시트 세팅 (최초 1회 실행) ──────────────────────────────
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ① 평가데이터 시트
  let evalSheet = ss.getSheetByName(SHEET_EVAL);
  if (!evalSheet) evalSheet = ss.insertSheet(SHEET_EVAL);
  evalSheet.clearContents();
  const evalHeaders = [
    '세션ID','이름','성별','나이','시작시간','완료여부',
    ...CHECKPOINTS.map(c=>c+'분_강도'),
    ...CHECKPOINTS.map(c=>c+'분_쾌불쾌'),
    '마지막업데이트'
  ];
  evalSheet.getRange(1,1,1,evalHeaders.length).setValues([evalHeaders])
    .setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold');
  evalSheet.setFrozenRows(1);

  // ② 사전설문 시트
  let surveySheet = ss.getSheetByName(SHEET_SURVEY);
  if (!surveySheet) surveySheet = ss.insertSheet(SHEET_SURVEY);
  surveySheet.clearContents();
  const surveyHeaders = [
    '세션ID','이름','성별','나이','전공/학년','생리주기',
    '제외_코막힘','제외_알레르기비염','제외_후각질환','제외_흡연','제외_약복용',
    '수면시간','마지막식사','카페인','컨디션','코막힘정도',
    '검사향','baseline결과','검사시각','등록시간'
  ];
  surveySheet.getRange(1,1,1,surveyHeaders.length).setValues([surveyHeaders])
    .setBackground('#1a2e1a').setFontColor('#ffffff').setFontWeight('bold');
  surveySheet.setFrozenRows(1);

  SpreadsheetApp.flush();
  return '시트 설정 완료! (평가데이터 + 사전설문)';
}

// ── CORS / 응답 헬퍼 ─────────────────────────────────────────────
function setCors(output) {
  return output.setMimeType(ContentService.MimeType.JSON)
    .setHeader('Access-Control-Allow-Origin','*')
    .setHeader('Access-Control-Allow-Methods','GET,POST')
    .setHeader('Access-Control-Allow-Headers','Content-Type');
}
function ok(data)  { return setCors(ContentService.createTextOutput(JSON.stringify({ok:true,data}))); }
function err(msg)  { return setCors(ContentService.createTextOutput(JSON.stringify({ok:false,error:msg}))); }

// ── GET ───────────────────────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action;
  if (action==='ping')        return ok('pong');
  if (action==='getSessions') return getSessions();
  return err('unknown action');
}

// ── POST ──────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;
    if (action==='register')    return registerParticipant(body);
    if (action==='submitScore') return submitScore(body);
    if (action==='complete')    return markComplete(body);
    return err('unknown action');
  } catch(ex) { return err(ex.message); }
}

// ── 참가자 등록 (설문 포함) ──────────────────────────────────────
function registerParticipant(body) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const evalSheet  = ss.getSheetByName(SHEET_EVAL);
  const surveySheet= ss.getSheetByName(SHEET_SURVEY);
  const sessionId  = body.sessionId;
  const now        = new Date().toISOString();

  // 중복 체크
  const evalData = evalSheet.getDataRange().getValues();
  for (let i=1;i<evalData.length;i++) {
    if (evalData[i][0]===sessionId) return ok({sessionId,existing:true});
  }

  // 평가데이터 행 추가
  const evalRow = [
    sessionId, body.name||'', body.gender||'', body.age||'',
    now, 'N',
    ...Array(10).fill(''), // 강도
    ...Array(10).fill(''), // 쾌불쾌
    now
  ];
  evalSheet.appendRow(evalRow);

  // 사전설문 행 추가
  if (body.survey) {
    try {
      const sv = typeof body.survey==='string' ? JSON.parse(body.survey) : body.survey;
      const excl = sv.exclItems || [];
      const surveyRow = [
        sessionId, body.name||'', sv.gender||'', sv.age||'',
        sv.major||'', sv.cycle||'',
        excl[0]||'', excl[1]||'', excl[2]||'', excl[3]||'', excl[4]||'',
        sv.sleep||'', sv.mealTime||'', sv.caffeine||'',
        sv.condition||'', sv.noseClog||'',
        sv.blSmell||'', sv.blResult||'', sv.blTime||'', now
      ];
      surveySheet.appendRow(surveyRow);
    } catch(e) {}
  }

  SpreadsheetApp.flush();
  return ok({sessionId,registered:true});
}

// ── 점수 제출 ────────────────────────────────────────────────────
function submitScore(body) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const sheet     = ss.getSheetByName(SHEET_EVAL);
  const data      = sheet.getDataRange().getValues();
  const sessionId = body.sessionId;
  const cpMin     = parseInt(body.cpMin);
  const cpIdx     = CHECKPOINTS.indexOf(cpMin);
  if (cpIdx===-1) return err('invalid cpMin');

  const intensityCol = 7  + cpIdx; // 1-based
  const comfortCol   = 17 + cpIdx;
  const lastSeenCol  = 27;

  for (let i=1;i<data.length;i++) {
    if (data[i][0]===sessionId) {
      const row = i+1;
      sheet.getRange(row,intensityCol).setValue(body.intensity);
      sheet.getRange(row,comfortCol).setValue(body.comfort);
      sheet.getRange(row,lastSeenCol).setValue(new Date().toISOString());
      SpreadsheetApp.flush();
      return ok({saved:true,cpMin,intensity:body.intensity,comfort:body.comfort});
    }
  }
  return err('session not found');
}

// ── 완료 처리 ────────────────────────────────────────────────────
function markComplete(body) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_EVAL);
  const data  = sheet.getDataRange().getValues();
  for (let i=1;i<data.length;i++) {
    if (data[i][0]===body.sessionId) {
      sheet.getRange(i+1,6).setValue('Y');
      sheet.getRange(i+1,27).setValue(new Date().toISOString());
      SpreadsheetApp.flush();
      return ok({complete:true});
    }
  }
  return err('session not found');
}

// ── 전체 세션 조회 ───────────────────────────────────────────────
function getSessions() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_EVAL);
  const data  = sheet.getDataRange().getValues();
  if (data.length<=1) return ok([]);

  const sessions = data.slice(1).map(row=>{
    const scores={}, comfort={};
    CHECKPOINTS.forEach((cp,i)=>{
      if (row[6+i]!=='')  scores[cp]  = row[6+i];
      if (row[16+i]!=='') comfort[cp] = row[16+i];
    });
    return {
      sessionId: row[0], name: row[1], gender: row[2], age: row[3],
      startTime: row[4], complete: row[5]==='Y',
      scores, comfort, lastSeen: row[26]
    };
  });
  return ok(sessions);
}
