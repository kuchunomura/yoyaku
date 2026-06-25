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
var DAY_COLS   = ['日付','施設','コース','時間','名前','人数','予約サイト','メモ','会計','_id','_json'];
var STAY_COLS  = ['チェックイン','チェックアウト','泊数','カテゴリ','棟','ベッド','ペット','名前','人数','予約サイト','メモ','部屋準備','会計','案内','OUT','_id','_json'];

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
// ベッド構成 ww=ベッド2台 / ws=ベッド1台。表示↔内部コード変換
function bedLabelG(v){ v=String(v||'').trim(); if(v==='ws'||v==='ベッド1台') return 'ベッド1台'; if(v==='ww'||v==='ベッド2台') return 'ベッド2台'; if(/ソファ|sofa/i.test(v)) return 'ベッド1台'; if(/[2２]/.test(v)) return 'ベッド2台'; if(/[1１]/.test(v)) return 'ベッド1台'; return ''; }
function bedCodeG(v){ v=String(v||'').trim(); if(v==='ww'||v==='ws') return v; if(v==='ベッド2台') return 'ww'; if(v==='ベッド1台') return 'ws'; if(/ソファ|sofa/i.test(v)) return 'ws'; if(/[2２]/.test(v)) return 'ww'; if(/[1１]/.test(v)) return 'ws'; return ''; }
// 'YYYY-MM-DD' → '6/13㈯'（曜日は丸囲み）
function fmtMD(ds){
  if(!ds) return '';
  var p = String(ds).split('-');
  if(p.length < 3) return ds;
  var dt = new Date(Number(p[0]), Number(p[1])-1, Number(p[2]));
  var w = '㈰㈪㈫㈬㈭㈮㈯'.charAt(dt.getDay());
  return Number(p[1]) + '/' + Number(p[2]) + w;
}
// ---- 逆変換（手動編集された読みやすい列 → 内部値）。全角入力も許容 ----
function toHalf(s){
  return String(s==null?'':s)
    .replace(/[０-９]/g, function(c){ return String.fromCharCode(c.charCodeAt(0)-0xFEE0); })
    .replace(/／/g,'/').replace(/：/g,':').replace(/　/g,' ').trim();
}
function invMap(m){ var o={}; for(var k in m){ o[m[k]]=k; } return o; }
var FAC_REV = invMap(FAC_LABELS);
var SRC_REV = invMap(SRC_LABELS);
function facId(label){ label=String(label||'').trim(); if(!label || label==='（棟未選択）') return ''; return (FAC_REV[label]!==undefined)?FAC_REV[label]:label; }
function srcKey(label){ label=String(label||'').trim(); if(!label) return ''; return (SRC_REV[label]!==undefined)?SRC_REV[label]:label; }
function courseVal(label){ var s=toHalf(label); if(String(label).indexOf('1DAY')>=0||String(label).indexOf('終日')>=0||s.indexOf('1DAY')>=0) return 0; var m=s.match(/(\d+)\s*分/); return m?Number(m[1]):0; }
function numVal(v){ var s=toHalf(v).replace(/[^\d.]/g,''); var n=parseFloat(s); return isNaN(n)?0:n; }
// '6/13㈯' / 'YYYY-MM-DD' / Date / 全角 → 'YYYY-MM-DD'。年は元データから推定
function parseMD(v, yearHint){
  if(v instanceof Date){ return v.getFullYear()+'-'+('0'+(v.getMonth()+1)).slice(-2)+'-'+('0'+v.getDate()).slice(-2); }
  var s=toHalf(v); if(!s) return '';
  if(/^\d{4}-\d{1,2}-\d{1,2}/.test(s)){ var q=s.split('-'); return q[0]+'-'+('0'+q[1]).slice(-2)+'-'+('0'+q[2]).slice(-2); }
  var m=s.match(/(\d+)\/(\d+)/); if(!m) return String(v||'');
  var y=yearHint||String(new Date().getFullYear());
  return y+'-'+('0'+m[1]).slice(-2)+'-'+('0'+m[2]).slice(-2);
}
function cell2time(v){ if(v instanceof Date){ return ('0'+v.getHours()).slice(-2)+':'+('0'+v.getMinutes()).slice(-2); } return toHalf(v); }

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

// アプリと同じ並び順用：日帰り施設順／宿泊カテゴリ順
var DAY_FAC_ORDER = { walk:1, tree:0, bbq:2 };
var STAY_GRP_ORDER = { 'ツリーハウス':0, '透明ドーム':1, '空中テント':2, '空中ハンモック':3 };
function dayFacOrder(id){ return (DAY_FAC_ORDER[id]==null)?9:DAY_FAC_ORDER[id]; }
function stayGrpOrder(g){ return (STAY_GRP_ORDER[g]==null)?9:STAY_GRP_ORDER[g]; }
function writeAll(reservations, stays){
  // 日付順 → 施設順（ツリーハウス昼→空中ウォーク→BBQ）→ 予約時間順（アプリの一覧と同じ）
  var days = reservations.slice().sort(function(a,b){
    return String(a.date||'').localeCompare(String(b.date||''))
      || (dayFacOrder(a.facility) - dayFacOrder(b.facility))
      || String(a.startTime||'').localeCompare(String(b.startTime||''));
  });
  // チェックイン順 → カテゴリ順（ツリー→ドーム→テント→ハンモック）
  var st = stays.slice().sort(function(a,b){
    return String(a.checkin||'').localeCompare(String(b.checkin||''))
      || (stayGrpOrder(a.facGroup) - stayGrpOrder(b.facGroup));
  });
  writeRows(getSheet(DAY_SHEET), DAY_COLS, days.map(function(r){
    return [fmtMD(r.date), facLabel(r.facility), courseLabel(r.course), r.startTime||'', r.name||'', r.ninzu||'', srcLabel(r.source), r.memo||'', (r.done?'✅':''), r.id||'', JSON.stringify(r)];
  }), days.map(function(r){ return r.date||''; }));
  writeRows(getSheet(STAY_SHEET), STAY_COLS, st.map(function(s){
    var w=s.wf||{};
    return [fmtMD(s.checkin), fmtMD(s.checkout), s.nights||'', s.facGroup||'', (s.facility?facLabel(s.facility):'（棟未選択）'), bedLabelG(s.bed), (s.petCount||''), s.name||'', s.ninzu||'', srcLabel(s.source), s.memo||'', (w.prep?'✅':''), (w.pay?'✅':''), (w.guide?'✅':''), (w.out?'✅':''), s.id||'', JSON.stringify(s)];
  }), st.map(function(s){ return s.checkin||''; }));
}

function writeRows(sh, cols, rows, groupKeys){
  sh.clear();
  sh.getRange(1,1,1,cols.length).setValues([cols]).setFontWeight('bold').setBackground('#f0f0f0').setHorizontalAlignment('center');
  if(rows.length){
    var rng = sh.getRange(2,1,rows.length,cols.length);
    rng.setValues(rows).setHorizontalAlignment('center').setVerticalAlignment('middle');
    var memoIdx = cols.indexOf('メモ');
    if(memoIdx >= 0) sh.getRange(2, memoIdx+1, rows.length, 1).setHorizontalAlignment('left'); // メモは左寄せ
    // キャンセル行は薄グレー
    if(memoIdx >= 0){
      for(var c=0;c<rows.length;c++){
        if(String(rows[c][memoIdx]).indexOf('キャンセル') >= 0) sh.getRange(2+c,1,1,cols.length).setBackground('#f3f3f3').setFontColor('#999999');
      }
    }
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
  applyCancelVisibility(sh); // キャンセル非表示の設定を反映（同期で書き直しても維持）
}

// ===== スプレッドシートのメニュー：キャンセル予約の表示/非表示 =====
function onOpen(){
  SpreadsheetApp.getUi().createMenu('予約管理')
    .addItem('キャンセルを非表示にする','menuHideCancelled')
    .addItem('キャンセルを表示する','menuShowCancelled')
    .addToUi();
}
function menuHideCancelled(){ setHideCancelledPref(true); applyCancelVisibilityAll(); SpreadsheetApp.getActive().toast('キャンセル予約を非表示にしました'); }
function menuShowCancelled(){ setHideCancelledPref(false); applyCancelVisibilityAll(); SpreadsheetApp.getActive().toast('キャンセル予約を表示しました'); }
function setHideCancelledPref(v){ PropertiesService.getDocumentProperties().setProperty('hideCancelled', v?'1':''); }
function getHideCancelledPref(){ return PropertiesService.getDocumentProperties().getProperty('hideCancelled')==='1'; }
function applyCancelVisibilityAll(){ [DAY_SHEET,STAY_SHEET].forEach(function(n){ var sh=SpreadsheetApp.getActive().getSheetByName(n); if(sh) applyCancelVisibility(sh); }); }
function applyCancelVisibility(sh){
  var last=sh.getLastRow(); if(last<2) return;
  var lastCol=sh.getLastColumn();
  var header=sh.getRange(1,1,1,lastCol).getValues()[0];
  var memoIdx=header.indexOf('メモ'); if(memoIdx<0) return;
  var hide=getHideCancelledPref();
  var memos=sh.getRange(2,memoIdx+1,last-1,1).getValues();
  for(var i=0;i<memos.length;i++){
    var isC=String(memos[i][0]).indexOf('キャンセル')>=0;
    if(isC && hide) sh.hideRows(2+i); else sh.showRows(2+i);
  }
}

function readSheet(name){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if(!sh) return [];
  var last = sh.getLastRow();
  if(last < 2) return [];
  var lastCol = sh.getLastColumn();
  var header = sh.getRange(1,1,1,lastCol).getValues()[0];
  var idx = {}; for(var h=0;h<header.length;h++){ idx[header[h]] = h; }
  if(idx['_json'] === undefined) return [];
  var data = sh.getRange(2,1,last-1,lastCol).getValues();
  var isDay = (name === DAY_SHEET);
  var out = [];
  for(var i=0;i<data.length;i++){
    var row = data[i];
    var obj = {};
    if(row[idx['_json']]){ try{ obj = JSON.parse(row[idx['_json']]); }catch(e){} }
    // 読みやすい列の手動編集を内部値へ上書き（全角入力も許容）
    var yh = String(obj.date || obj.checkin || '').slice(0,4) || undefined;
    function col(n){ return idx[n]!==undefined ? row[idx[n]] : undefined; }
    if(isDay){
      if(idx['日付']!==undefined)     obj.date      = parseMD(col('日付'), yh);
      if(idx['施設']!==undefined)     obj.facility  = facId(col('施設'));
      if(idx['コース']!==undefined)   obj.course    = courseVal(col('コース'));
      if(idx['時間']!==undefined)     obj.startTime = cell2time(col('時間'));
      if(idx['名前']!==undefined)     obj.name      = String(col('名前')||'');
      if(idx['人数']!==undefined && String(col('人数'))!=='') obj.ninzu = String(numVal(col('人数')));
      if(idx['予約サイト']!==undefined) obj.source  = srcKey(col('予約サイト'));
      if(idx['メモ']!==undefined)     obj.memo      = String(col('メモ')||'');
      if(idx['会計']!==undefined)     obj.done      = !!String(col('会計')).trim();
    }else{
      if(idx['チェックイン']!==undefined)   obj.checkin  = parseMD(col('チェックイン'), yh);
      if(idx['チェックアウト']!==undefined) obj.checkout = parseMD(col('チェックアウト'), yh);
      if(idx['泊数']!==undefined && String(col('泊数'))!=='') obj.nights = numVal(col('泊数'))||obj.nights;
      if(idx['カテゴリ']!==undefined) obj.facGroup = String(col('カテゴリ')||'');
      if(idx['棟']!==undefined)       obj.facility = facId(col('棟'));
      if(idx['ベッド']!==undefined)   obj.bed      = bedCodeG(col('ベッド'));
      if(idx['ペット']!==undefined && String(col('ペット'))!==''){ obj.petCount = numVal(col('ペット')); obj.petFee = obj.petCount*3000*(obj.nights||1); }
      if(idx['名前']!==undefined)     obj.name     = String(col('名前')||'');
      if(idx['人数']!==undefined && String(col('人数'))!==''){ obj.ninzu = numVal(col('人数')); obj.totalPpl = obj.ninzu; }
      if(idx['予約サイト']!==undefined) obj.source = srcKey(col('予約サイト'));
      if(idx['メモ']!==undefined)     obj.memo     = String(col('メモ')||'');
      if(idx['部屋準備']!==undefined||idx['会計']!==undefined||idx['案内']!==undefined||idx['OUT']!==undefined){
        obj.wf = obj.wf || {};
        if(idx['部屋準備']!==undefined) obj.wf.prep  = !!String(col('部屋準備')).trim();
        if(idx['会計']!==undefined)     obj.wf.pay   = !!String(col('会計')).trim();
        if(idx['案内']!==undefined)     obj.wf.guide = !!String(col('案内')).trim();
        if(idx['OUT']!==undefined)      obj.wf.out   = !!String(col('OUT')).trim();
      }
    }
    if(obj.id || obj.wpId || obj.name || obj.date || obj.checkin) out.push(obj);
  }
  return out;
}
