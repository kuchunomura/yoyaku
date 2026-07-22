// ============================================================
// 予約管理アプリ（yoyaku）用 GAS Webアプリ
//
// 【独立プロジェクトとして script.google.com に置くこと】
// スプレッドシートの「拡張機能→Apps Script」からは作らない。
// 理由: バインドすると、そのスプレッドシートを消したときスクリプトごと消える。
//       2026/07/16、ドライブのストレージ削除でシートが消え、GASも道連れで消えた。
//       独立プロジェクトなら下の SS_ID を差し替えるだけで復旧できる。
//
// 使い方:
//   1. 同期用スプレッドシートを作成し、そのURLの /d/ と /edit の間のIDを下の SS_ID に貼る
//   2. script.google.com → 新しいプロジェクト → このコードを貼り付け
//   3. デプロイ → 新しいデプロイ → ウェブアプリ（実行:自分／アクセス:全員）
//   4. 表示された /exec URL を2か所に貼る:
//        yoyakuアプリ → 設定「GAS WebアプリURL」
//        POSレジ     → 設定「📅 予約連携（yoyaku）」
//   5. GASエディタで checkSpreadsheetAccess() を1回実行し、権限の承認を済ませる
//
// doPost: アプリから {type:'sync_all', reservations, stays} を受けてシートへ保存
// doGet : アプリへ {reservations, stays} を返す（端末間同期の読込用）
// シートは「人が読める列＋_json（完全復元用）」の両方を書き込む（印刷にも使える）
// ============================================================

// 同期先スプレッドシートのID（URLの /d/ と /edit の間）
var SS_ID = '1gwV7YQHA0p6pWUXjw9qAhB5Js0NK3p-QtuR063QqM54'; // yoyaku同期（2026/07/16 再作成）

function getTargetSS(){
  if(!SS_ID) throw new Error('SS_ID が未設定です。GASコード先頭の SS_ID にスプレッドシートのIDを貼ってください');
  return SpreadsheetApp.openById(SS_ID);
}

// 【GASエディタから手で実行する用】スプレッドシートを本当に開けるか確認する。
// アプリ側の疎通確認では openById の権限までは分からないので、作り直したら必ず1回実行する。
function checkSpreadsheetAccess(){
  var ss = getTargetSS();
  var names = ss.getSheets().map(function(s){ return s.getName(); });
  var msg = '✅ 開けました\n\nID: ' + SS_ID + '\n名前: ' + ss.getName() + '\nシート' + names.length + '枚: ' + names.join(', ');
  Logger.log(msg);
  return msg;
}

var DAY_SHEET  = '予約_日帰り';
var STAY_SHEET = '予約_宿泊';
var DAY_COLS   = ['日付','施設','コース','時間','名前','人数','予約サイト','メモ','会計','キャンセル','_id','_json'];
var STAY_COLS  = ['チェックイン','チェックアウト','泊数','カテゴリ','棟','ベッド','ペット','名前','人数','予約サイト','メモ','部屋準備','会計','案内','OUT','キャンセル','_id','_json'];

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
      if(data.otaack && typeof data.otaack === 'object'){
        PropertiesService.getScriptProperties().setProperty('otaack', JSON.stringify(data.otaack));
      }
      if(data.qack && typeof data.qack === 'object'){
        PropertiesService.getScriptProperties().setProperty('qack', JSON.stringify(data.qack));
      }
      // 共有メモ（全端末で同期。URL控え・引き継ぎ事項など）
      if(typeof data.sharednote === 'string'){
        PropertiesService.getScriptProperties().setProperty('sharednote', data.sharednote);
      }
      // 日付ごとの特記（イベント・取材・予定。全端末で同期）
      if(data.dayevents && typeof data.dayevents === 'object'){
        PropertiesService.getScriptProperties().setProperty('dayevents', JSON.stringify(data.dayevents));
      }
      // 最終CSV取込日時：施設ごとに新しい方を残してマージ（他デバイスの取込時刻を消さない）
      // 併せて「何月分か（csvimpmon）」も、取込時刻が新しくなった施設だけ更新して同期する
      if(data.csvimp && typeof data.csvimp === 'object'){
        var _cur={}; try{ _cur=JSON.parse(PropertiesService.getScriptProperties().getProperty('csvimp')||'{}'); }catch(_e){}
        var _curM={}; try{ _curM=JSON.parse(PropertiesService.getScriptProperties().getProperty('csvimpmon')||'{}'); }catch(_eM){}
        var _curF={}; try{ _curF=JSON.parse(PropertiesService.getScriptProperties().getProperty('csvimpfile')||'{}'); }catch(_eF){}
        var _curLM={}; try{ _curLM=JSON.parse(PropertiesService.getScriptProperties().getProperty('csvimplastmon')||'{}'); }catch(_eL){}
        var _inM=(data.csvimpmon && typeof data.csvimpmon === 'object')?data.csvimpmon:{};
        var _inF=(data.csvimpfile && typeof data.csvimpfile === 'object')?data.csvimpfile:{};
        var _inLM=(data.csvimplastmon && typeof data.csvimplastmon === 'object')?data.csvimplastmon:{};
        for(var _k in data.csvimp){ var _v=Number(data.csvimp[_k])||0; if(_v>(Number(_cur[_k])||0)){ _cur[_k]=_v; if(_inM[_k]!==undefined)_curM[_k]=_inM[_k]; if(_inF[_k]!==undefined)_curF[_k]=_inF[_k]; if(_inLM[_k]!==undefined)_curLM[_k]=_inLM[_k]; } }
        PropertiesService.getScriptProperties().setProperty('csvimp', JSON.stringify(_cur));
        PropertiesService.getScriptProperties().setProperty('csvimpmon', JSON.stringify(_curM));
        PropertiesService.getScriptProperties().setProperty('csvimpfile', JSON.stringify(_curF));
        PropertiesService.getScriptProperties().setProperty('csvimplastmon', JSON.stringify(_curLM));
      }
      // 月ごとの取込時刻（csvimpmts）：施設ごと・月ごとに新しいtsを残してマージ
      if(data.csvimpmts && typeof data.csvimpmts === 'object'){
        var _curMTS={}; try{ _curMTS=JSON.parse(PropertiesService.getScriptProperties().getProperty('csvimpmts')||'{}'); }catch(_eT){}
        for(var _fk in data.csvimpmts){ var _mm=data.csvimpmts[_fk]; if(!_mm||typeof _mm!=='object')continue; if(!_curMTS[_fk])_curMTS[_fk]={}; for(var _mk in _mm){ var _mv=Number(_mm[_mk])||0; if(_mv>(Number(_curMTS[_fk][_mk])||0))_curMTS[_fk][_mk]=_mv; } }
        PropertiesService.getScriptProperties().setProperty('csvimpmts', JSON.stringify(_curMTS));
      }
      return jsonOut({status:'ok', saved:{reservations:(data.reservations||[]).length, stays:(data.stays||[]).length}});
    }
    return jsonOut({status:'error', message:'unknown type'});
  }catch(err){
    return jsonOut({status:'error', message:String(err)});
  }
}

function doGet(e){
  try{
    var ota={}; try{ ota=JSON.parse(PropertiesService.getScriptProperties().getProperty('otaack')||'{}'); }catch(e2){}
    var qk={}; try{ qk=JSON.parse(PropertiesService.getScriptProperties().getProperty('qack')||'{}'); }catch(e3){}
    var ci={}; try{ ci=JSON.parse(PropertiesService.getScriptProperties().getProperty('csvimp')||'{}'); }catch(e4){}
    var cim={}; try{ cim=JSON.parse(PropertiesService.getScriptProperties().getProperty('csvimpmon')||'{}'); }catch(e6){}
    var cif={}; try{ cif=JSON.parse(PropertiesService.getScriptProperties().getProperty('csvimpfile')||'{}'); }catch(e8){}
    var cilm={}; try{ cilm=JSON.parse(PropertiesService.getScriptProperties().getProperty('csvimplastmon')||'{}'); }catch(e9){}
    var cimts={}; try{ cimts=JSON.parse(PropertiesService.getScriptProperties().getProperty('csvimpmts')||'{}'); }catch(e10){}
    var sn=PropertiesService.getScriptProperties().getProperty('sharednote')||'';
    var de={}; try{ de=JSON.parse(PropertiesService.getScriptProperties().getProperty('dayevents')||'{}'); }catch(e7){}
    return jsonOut({status:'ok', reservations:readSheet(DAY_SHEET), stays:readSheet(STAY_SHEET), otaack:ota, qack:qk, csvimp:ci, csvimpmon:cim, csvimpfile:cif, csvimplastmon:cilm, csvimpmts:cimts, sharednote:sn, dayevents:de});
  }catch(err){
    return jsonOut({status:'error', message:String(err)});
  }
}

function jsonOut(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name){
  var ss = getTargetSS();
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
  // キャンセルは下部へまとめず、有効予約と同じ日付順のまま（全体を時系列に）。
  // 折り畳みは applyCancelGrouping が「連続するキャンセル行の塊」ごとに行グループ化して±で畳めるようにする。
  writeRows(getSheet(DAY_SHEET), DAY_COLS, days.map(function(r){
    return [fmtMD(r.date), facLabel(r.facility), courseLabel(r.course), r.startTime||'', r.name||'', r.ninzu||'', srcLabel(r.source), r.memo||'', (r.done?'✅':''), ((r.cancelled||/キャンセル/.test(r.memo||''))?'✅':''), r.id||'', JSON.stringify(r)];
  }), days.map(function(r){ return r.date||''; }));
  writeRows(getSheet(STAY_SHEET), STAY_COLS, st.map(function(s){
    var w=s.wf||{};
    return [fmtMD(s.checkin), fmtMD(s.checkout), s.nights||'', s.facGroup||'', (s.facility?facLabel(s.facility):'（棟未選択）'), bedLabelG(s.bed), (s.petCount||''), s.name||'', s.ninzu||'', srcLabel(s.source), s.memo||'', (w.prep?'✅':''), (w.pay?'✅':''), (w.guide?'✅':''), (w.out?'✅':''), ((s.cancelled||/キャンセル/.test(s.memo||''))?'✅':''), s.id||'', JSON.stringify(s)];
  }), st.map(function(s){ return s.checkin||''; }));
}

function writeRows(sh, cols, rows, groupKeys){
  sh.clear();
  sh.getRange(1,1,1,cols.length).setValues([cols]).setFontWeight('bold').setBackground('#f0f0f0').setHorizontalAlignment('center');
  if(rows.length){
    var rng = sh.getRange(2,1,rows.length,cols.length);
    rng.setValues(rows).setHorizontalAlignment('center').setVerticalAlignment('middle');
    var memoIdx = cols.indexOf('メモ');
    var cancelIdx = cols.indexOf('キャンセル');
    if(memoIdx >= 0) sh.getRange(2, memoIdx+1, rows.length, 1).setHorizontalAlignment('left'); // メモは左寄せ
    // キャンセル行は薄グレー（メモに「キャンセル」or キャンセル列が✅）
    for(var c=0;c<rows.length;c++){
      var isC=(memoIdx>=0 && String(rows[c][memoIdx]).indexOf('キャンセル')>=0) || (cancelIdx>=0 && String(rows[c][cancelIdx]).trim()!=='');
      if(isC) sh.getRange(2+c,1,1,cols.length).setBackground('#f3f3f3').setFontColor('#999999');
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
  applyCancelGrouping(sh); // キャンセル行を行グループ化（±で折り畳み）。同期で書き直しても維持
}

// ===== キャンセル予約の折り畳み（行グループ化）=====
// キャンセル行は writeAll で各シートの下部へまとめてある。ここで1つの行グループにして、
// シート左の「−／＋」ボタンでワンタッチ折り畳み／展開できるようにする（毎回の同期でも維持）。
// 既定は「折り畳み（−で閉じた状態）」。常に開いておきたい時は下の関数をGASエディタで▶実行:
//   showCancelledRows() … 以後の同期で展開した状態にする
//   hideCancelledRows() … 以後の同期で折り畳んだ状態にする（既定）
// ※シート上の「−／＋」でその場で開閉もできる（次にアプリが同期するまで維持）。
function hideCancelledRows(){ setCancelShownPref(false); applyCancelGroupingAll(); return 'キャンセル予約を折り畳みました'; }
function showCancelledRows(){ setCancelShownPref(true);  applyCancelGroupingAll(); return 'キャンセル予約を展開しました'; }
function setCancelShownPref(v){ PropertiesService.getScriptProperties().setProperty('cancelShown', v?'1':''); }
function getCancelShownPref(){ return PropertiesService.getScriptProperties().getProperty('cancelShown')==='1'; }
function applyCancelGroupingAll(){ [DAY_SHEET,STAY_SHEET].forEach(function(n){ var sh=getTargetSS().getSheetByName(n); if(sh) applyCancelGrouping(sh); }); }
function applyCancelGrouping(sh){
  var last=sh.getLastRow();
  // 既存の行グループを一旦すべて解除（毎回作り直す）
  try{ sh.getRange(1,1,sh.getMaxRows(),1).shiftRowGroupDepth(-8); }catch(e0){}
  if(last<2) return;
  var lastCol=sh.getLastColumn();
  var header=sh.getRange(1,1,1,lastCol).getValues()[0];
  var memoIdx=header.indexOf('メモ');
  var cancelIdx=header.indexOf('キャンセル');
  if(memoIdx<0 && cancelIdx<0) return;
  var rng=sh.getRange(2,1,last-1,lastCol).getValues();
  function _isCancRow(i){ return (memoIdx>=0 && String(rng[i][memoIdx]).indexOf('キャンセル')>=0) || (cancelIdx>=0 && String(rng[i][cancelIdx]).trim()!==''); }
  // 日付順に散らばったキャンセルを「連続する塊」ごとに行グループ化（その場で±で畳める）
  var i=0;
  while(i<rng.length){
    if(_isCancRow(i)){
      var j=i; while(j<rng.length && _isCancRow(j)) j++;
      try{ sh.getRange(2+i,1,j-i,1).shiftRowGroupDepth(1); }catch(e2){}
      i=j;
    } else { i++; }
  }
  // 既定は折り畳み。showCancelledRows()を実行した時だけ展開状態にする
  try{ if(getCancelShownPref()) sh.expandAllRowGroups(); else sh.collapseAllRowGroups(); }catch(e3){}
}

function readSheet(name){
  var ss = getTargetSS();
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
      if(idx['キャンセル']!==undefined) obj.cancelled = !!String(col('キャンセル')).trim();
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
      if(idx['キャンセル']!==undefined) obj.cancelled = !!String(col('キャンセル')).trim();
    }
    if(obj.id || obj.wpId || obj.name || obj.date || obj.checkin) out.push(obj);
  }
  return out;
}
