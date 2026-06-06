/**
 * KBO 카드 컬렉션 → 구글 시트 동기화 (거래보드)
 * 사용법은 SETUP_SYNC.md 참고.
 *  - 양도가능 : 보유한 카드 (KIA 제외)  → 남에게 줄 수 있는 카드
 *  - 구함     : KIA 중 아직 없는 카드    → 받고 싶은 카드
 */
var TOKEN = 'CHANGE_ME';  // ← 앱 '동기화 설정'의 비밀키와 똑같이 바꾸세요

function doGet(e){ return json_({ok:true, msg:'KBO sync alive'}); }

function doPost(e){
  try{
    var body = JSON.parse(e.postData.contents);
    if (String(body.token) !== String(TOKEN)) return json_({ok:false, error:'bad token'});
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var head = ['연도','구단','카드종류','버전','등번호','선수명'];
    writeSheet_(ss, '양도가능(보유)', head, body.giveaway || []);
    writeSheet_(ss, '구함(KIA 미보유)', head, body.want || []);
    var info = ss.getSheetByName('요약') || ss.insertSheet('요약', 0);
    info.clear();
    info.getRange(1,1,4,2).setValues([
      ['마지막 업데이트', new Date()],
      ['양도가능(보유) 수', (body.giveaway||[]).length],
      ['구함(KIA) 수', (body.want||[]).length],
      ['앱 생성시각', body.generatedAt || '']
    ]);
    info.getRange(1,1,4,1).setFontWeight('bold');
    return json_({ok:true, giveaway:(body.giveaway||[]).length, want:(body.want||[]).length});
  }catch(err){ return json_({ok:false, error:String(err)}); }
}

function writeSheet_(ss, name, header, rows){
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear();
  var data = [header];
  for (var i=0;i<rows.length;i++){
    var r = rows[i].slice();
    while (r.length < header.length) r.push('');
    data.push(r);
  }
  sh.getRange(1,1,data.length, header.length).setValues(data);
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,header.length).setFontWeight('bold');
  sh.autoResizeColumns(1, header.length);
}

function json_(o){
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
