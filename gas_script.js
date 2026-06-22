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
var STAY_COLS  = ['チェックイン','チェックアウト','泊数','カテゴリ','棟','名前','人数','予約サイト','メモ','_id','_json'];

// 施設ID→表示名（スプレッドシートを読みやすく）
var FAC_LABELS = {
  walk:'空中ウォーク', tree:'ツリーハウス昼', bbq:'BBQスペース',
  st_th:'ツリーハウス',
  st_dm:'ドーム（ミラー）', st_dt:'ドーム（クリア）', st_dg:'ドーム（グレー）',
  st_t1:'テント①オレンジ', st_t2:'テント②ホワイト', st_t3:'テント③ベージュ',
  st_h1:'ハンモック①', st_h2:'ハンモック②'
};
var SRC_LABELS = {
  rakuten:'楽天トラベル', jalan:'じゃらん', sou:'SOUエクスペリエンス',
  other1:'外部サイト①', other2:'外部サイト②', other3:'外部サイト③',
  pass_day:'年間パス（昼）', pass_night:'年間パス（夜）'
};
function facLabel(id){ return FAC_LABELS[id] || id || ''; }
function srcLabel(k){ return k ? (SRC_LABELS[k] || k) : ''; }
// 'YYYY-MM-DD' → '6/13㈯'（曜日は丸囲み）
function fmtMD(ds){
  if(!ds) return '';
  var p = String(ds).split('-');
  if(p.length < 3) return ds;
  var dt = new Date(Number(p[0]), Number(p[1])-1, Number(p[2]));
  var w = '㈰㈪㈫㈬㈭㈮㈯'.charAt(dt.getDay());
  return Number(p[1]) + '/' + Number(p[2]) + w;
}

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
  var days = reservations.slice().sort(function(a,b){
    return String(a.date||'').localeCompare(String(b.date||'')) || String(a.startTime||'').localeCompare(String(b.startTime||''));
  });
  var st = stays.slice().sort(function(a,b){ return String(a.checkin||'').localeCompare(String(b.checkin||'')); });
  writeRows(getSheet(DAY_SHEET), DAY_COLS, days.map(function(r){
    return [fmtMD(r.date), facLabel(r.facility), courseLabel(r.course), r.startTime||'', r.name||'', r.ninzu||'', srcLabel(r.source), r.memo||'', r.id||'', JSON.stringify(r)];
  }), days.map(function(r){ return r.date||''; }));
  writeRows(getSheet(STAY_SHEET), STAY_COLS, st.map(function(s){
    return [fmtMD(s.checkin), fmtMD(s.checkout), s.nights||'', s.facGroup||'', (s.facility?facLabel(s.facility):'（棟未選択）'), s.name||'', s.ninzu||'', srcLabel(s.source), s.memo||'', s.id||'', JSON.stringify(s)];
  }), st.map(function(s){ return s.checkin||''; }));
}

function writeRows(sh, cols, rows, groupKeys){
  sh.clear();
  sh.getRange(1,1,1,cols.length).setValues([cols]).setFontWeight('bold').setBackground('#f0f0f0');
  if(rows.length){
    sh.getRange(2,1,rows.length,cols.length).setValues(rows);
    // 日付（グループキー）が変わる行の下に下線を引く
    if(groupKeys){
      for(var i=0;i<rows.length;i++){
        var isLast = (i === rows.length-1) || (groupKeys[i] !== groupKeys[i+1]);
        if(isLast){ sh.getRange(2+i,1,1,cols.length).setBorder(null,null,true,null,null,null,'#888888',SpreadsheetApp.BorderStyle.SOLID_MEDIUM); }
      }
    }
  }
  sh.hideColumns(cols.length); // _json列を隠す
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
