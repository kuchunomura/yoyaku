// ============================================================
// 予約管理アプリ（yoyaku）用 GAS Webアプリ
// 使い方: 同期用スプレッドシートを作成 → 拡張機能→Apps Script に貼り付け
//        → デプロイ→新しいデプロイ→ウェブアプリ（実行:自分／アクセス:全員）
//        → 表示された /exec URL をアプリの設定「GAS WebアプリURL」に貼る
// doPost: アプリから {type:'sync_all', reservations, stays} を受けてシートへ保存
// doGet : アプリへ {reservations, stays} を返す（端末間同期の読込用）
// シートは「人が読める列＋_json（完全復元用）」の両方を書き込む（印刷にも使える）
// ============================================================

var DAY_SHEET  = '予約_日帰り';
var STAY_SHEET = '予約_宿泊';
var DAY_COLS   = ['日付','施設','コース','時間','名前','人数','予約サイト','メモ','_id','_json'];
var STAY_COLS  = ['チェックイン','チェックアウト','カテゴリ','棟','名前','人数','予約サイト','メモ','_id','_json'];

function doPost(e){
  try{
    var data = JSON.parse(e.postData.contents);
    if(data.type === 'sync_all'){
      writeAll(data.reservations || [], data.stays || []);
      return jsonOut({status:'ok', saved:{reservations:(data.reservations||[]).length, stays:(data.stays||[]).length}});
    }
    return jsonOut({status:'error', message:'unknown type'});
  }catch(err){
    return jsonOut({status:'error', message:String(err)});
  }
}

function doGet(e){
  try{
    return jsonOut({status:'ok', reservations:readSheet(DAY_SHEET), stays:readSheet(STAY_SHEET)});
  }catch(err){
    return jsonOut({status:'error', message:String(err)});
  }
}

function jsonOut(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if(!sh) sh = ss.insertSheet(name);
  return sh;
}

function courseLabel(c){
  if(c === 0 || c === '0') return '1DAY';
  if(c) return c + '分';
  return '';
}

function writeAll(reservations, stays){
  writeRows(getSheet(DAY_SHEET), DAY_COLS, reservations.map(function(r){
    return [r.date||'', r.facility||'', courseLabel(r.course), r.startTime||'', r.name||'', r.ninzu||'', r.source||'', r.memo||'', r.id||'', JSON.stringify(r)];
  }));
  writeRows(getSheet(STAY_SHEET), STAY_COLS, stays.map(function(s){
    return [s.checkin||'', s.checkout||'', s.facGroup||'', s.facility||'', s.name||'', s.ninzu||'', s.source||'', s.memo||'', s.id||'', JSON.stringify(s)];
  }));
}

function writeRows(sh, cols, rows){
  sh.clear();
  sh.getRange(1,1,1,cols.length).setValues([cols]).setFontWeight('bold').setBackground('#f0f0f0');
  if(rows.length) sh.getRange(2,1,rows.length,cols.length).setValues(rows);
  // _json列は隠す（人が見る用ではないため）
  sh.hideColumns(cols.length);
  sh.setFrozenRows(1);
}

function readSheet(name){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if(!sh) return [];
  var last = sh.getLastRow();
  if(last < 2) return [];
  var cols = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var ji = cols.indexOf('_json');
  if(ji < 0) return [];
  var vals = sh.getRange(2, ji+1, last-1, 1).getValues();
  var out = [];
  for(var i=0;i<vals.length;i++){
    var j = vals[i][0];
    if(!j) continue;
    try{ out.push(JSON.parse(j)); }catch(e){}
  }
  return out;
}
