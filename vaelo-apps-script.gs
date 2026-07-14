/**
 * VAELO — Google Sheets TWO-WAY sync, bound to YOUR real content calendar tab.
 * ------------------------------------------------------------
 * Instead of a separate tab, this reads/writes the content calendar tab you
 * already have (the one with DATE / STATUS / TOPIC / PRODUCT / FORMAT columns).
 * It finds that tab automatically by its headers.
 *
 *   • Sheet → pipeline: edit STATUS / TOPIC / PRODUCT / FORMAT in your sheet,
 *     hit ↻ Refresh in the app → the pipeline updates.
 *   • Pipeline → sheet: move a stage in the app → STATUS updates in your sheet
 *     (Topic/Product/Format too). Your other columns are never touched.
 *
 * Rows are matched by DATE. Your extra columns (caption, hashtags, visual,
 * CTA, drive link, notes) are preserved.
 *
 * STATUS ↔ stage mapping:
 *   Planned            → Calendar sent / early stages
 *   Work In Progress   → Creative in progress
 *   Ready              → Client approval
 *   Posted / Scheduled → Scheduled / Posted
 *
 * AFTER pasting/updating this file you MUST redeploy:
 *   Deploy ▸ Manage deployments ▸ (pencil ✏ on the Web app) ▸ Version: New version ▸ Deploy.
 */

var STORE_TAB = 'store';
var FIN_TAB   = 'Finance';
var CONTENT_CLIENT = 'SimpliCare';   // this spreadsheet is the SimpliCare calendar
var STAGES = ['Idea','Calendar sent','Calendar approved','Creative in progress','Internal approval','Client approval','Scheduled/Posted'];

function ss_(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function tab_(name){ var s=ss_().getSheetByName(name); if(!s) s=ss_().insertSheet(name); return s; }
function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function today_(){ return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function genId_(){ return Math.random().toString(36).slice(2,9); }
function cell_(v){ if(v instanceof Date){ return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd MMM yyyy'); } return v==null?'':String(v).trim(); }
function parseMonth_(d){ var m={jan:'01',feb:'02',mar:'03',apr:'04',may:'05',jun:'06',jul:'07',aug:'08',sep:'09',oct:'10',nov:'11',dec:'12'}; var p=String(d||'').trim().split(/\s+/); if(p.length>=3){ var mo=m[p[1].toLowerCase().slice(0,3)]; if(mo) return p[2]+'-'+mo; } return ''; }

function statusToStage_(status, fallback){
  var s=String(status||'').toLowerCase();
  if(/post|schedul|publish|live|done/.test(s)) return 6;
  if(/ready|approv/.test(s)) return 5;
  if(/progress|wip|edit|draft|film|shoot/.test(s)) return 3;
  if(/plan|idea|brief/.test(s)) return 1;
  return fallback!=null?fallback:1;
}
function stageToStatus_(stage){
  if(stage>=6) return 'Posted';
  if(stage===5) return 'Ready';
  if(stage>=3) return 'Work In Progress';
  return 'Planned';
}

/* ---------- key/value store (for the JSON blob + finance) ---------- */
function storeSheet_(){ var sh=ss_().getSheetByName(STORE_TAB); if(!sh){ sh=ss_().insertSheet(STORE_TAB); sh.appendRow(['key','value','updatedAt']); } return sh; }
function readAll_(){ var sh=storeSheet_(); var rows=sh.getDataRange().getValues(); var out={}; for(var i=1;i<rows.length;i++){ if(rows[i][0]) out[rows[i][0]]=rows[i][1]; } return out; }
function getStoreValue_(key){ var a=readAll_(); return a.hasOwnProperty(key)?a[key]:null; }
function upsertStore_(key,value){ var sh=storeSheet_(); var rows=sh.getDataRange().getValues(); for(var i=1;i<rows.length;i++){ if(rows[i][0]===key){ sh.getRange(i+1,2).setValue(value); sh.getRange(i+1,3).setValue(new Date()); return; } } sh.appendRow([key,value,new Date()]); }

/* ---------- find your content calendar tab by its headers ---------- */
function findContentSheet_(){
  var sheets=ss_().getSheets();
  for(var i=0;i<sheets.length;i++){
    var sh=sheets[i], name=sh.getName();
    if(name===STORE_TAB || name===FIN_TAB) continue;
    var rng=sh.getDataRange(); if(rng.getNumRows()<1) continue;
    var values=rng.getValues();
    for(var r=0;r<Math.min(values.length,8);r++){
      var hdr=values[r].map(function(c){ return String(c).toLowerCase(); });
      var status=hdr.findIndex(function(h){ return /status/.test(h); });
      var topic=hdr.findIndex(function(h){ return /topic/.test(h); });
      var date=hdr.findIndex(function(h){ return /date/.test(h); });
      if(status>=0 && topic>=0 && date>=0){
        return { sheet:sh, headerRow:r, values:values, colMap:{
          date:date, status:status, topic:topic,
          format: hdr.findIndex(function(h){ return /format/.test(h); }),
          product: hdr.findIndex(function(h){ return /product/.test(h); })
        }};
      }
    }
  }
  return null;
}

/* ---------- sheet → pipeline ---------- */
function contentFromSheet_(){
  var info=findContentSheet_();
  var raw=getStoreValue_('vaelo-content-state');
  var blob=raw?JSON.parse(raw):{items:[],ideas:[]};
  if(!blob.items) blob.items=[];
  if(!info) return JSON.stringify(blob);   // no content tab found → return blob unchanged

  var cm=info.colMap, values=info.values, hr=info.headerRow;
  var prevByDate={}; blob.items.forEach(function(it){ if(it.client===CONTENT_CLIENT) prevByDate[it.date]=it; });
  function v(row,i){ return i>=0?cell_(row[i]):''; }

  var sc=[];
  for(var r=hr+1;r<values.length;r++){
    var row=values[r];
    var date=v(row,cm.date); if(!date || !/\d/.test(date)) continue;  // skip non-date rows
    var prev=prevByDate[date]||{};
    var stage=statusToStage_(v(row,cm.status), prev.stage);
    var it={
      id: prev.id||genId_(), client: CONTENT_CLIENT, date: date,
      day: prev.day||'', month: prev.month||parseMonth_(date),
      format: cm.format>=0 ? v(row,cm.format) : (prev.format||''),
      topic:  v(row,cm.topic) || prev.topic || '',
      product: cm.product>=0 ? v(row,cm.product) : (prev.product||''),
      owner: prev.owner||'—', stage: stage, log: (prev.log||[]).slice()
    };
    if(prev.stage!=null && prev.stage!==stage){ it.log.push({d:today_(), t:'Stage moved to "'+STAGES[stage]+'". — by Sheet', by:'Sheet'}); }
    sc.push(it);
  }
  var others=blob.items.filter(function(it){ return it.client!==CONTENT_CLIENT; });
  blob.items=others.concat(sc);
  upsertStore_('vaelo-content-state', JSON.stringify(blob));
  return JSON.stringify(blob);
}

/* ---------- pipeline → sheet (only STATUS/Topic/Product/Format; nothing else) ---------- */
function pushContentToSheet_(state){
  var info=findContentSheet_(); if(!info) return;
  var sh=info.sheet, cm=info.colMap, values=info.values, hr=info.headerRow;
  var byDate={}; (state.items||[]).forEach(function(it){ if(it.client===CONTENT_CLIENT) byDate[it.date]=it; });
  for(var r=hr+1;r<values.length;r++){
    var date=cell_(values[r][cm.date]); if(!date || !/\d/.test(date)) continue;
    var it=byDate[date]; if(!it) continue;
    if(cm.status>=0)  sh.getRange(r+1, cm.status+1).setValue(stageToStatus_(it.stage));
    if(cm.topic>=0 && it.topic)   sh.getRange(r+1, cm.topic+1).setValue(it.topic);
    if(cm.product>=0)             sh.getRange(r+1, cm.product+1).setValue(it.product);
    if(cm.format>=0 && it.format) sh.getRange(r+1, cm.format+1).setValue(it.format);
  }
}

/* ---------- Finance (one-way readable mirror) ---------- */
function renderFinance_(data){
  var sh=tab_(FIN_TAB); sh.clear();
  sh.appendRow(['Client','Date','Item','Category','Amount','Paid','Added by']);
  Object.keys(data||{}).forEach(function(c){ (data[c]||[]).forEach(function(x){ sh.appendRow([c,x.date,x.item,x.cat,x.amt,x.paid?'Yes':'No',x.addedBy]); }); });
  sh.setFrozenRows(1);
}

function writeKey_(key,value){
  upsertStore_(key,value);
  try{
    if(key==='vaelo-content-state') pushContentToSheet_(JSON.parse(value));
    if(key==='vaelo-finance-state') renderFinance_(JSON.parse(value));
  }catch(e){}
}

/* ---------- endpoints ---------- */
function doGet(e){
  var key=e&&e.parameter&&e.parameter.key;
  if(key==='vaelo-content-state'){ return json_({key:key, value: contentFromSheet_()}); }
  if(key){ var all=readAll_(); return json_({key:key, value: all.hasOwnProperty(key)?all[key]:null}); }
  return json_(readAll_());
}
function doPost(e){
  try{
    var body=JSON.parse(e.postData.contents);
    if(body.updates){ Object.keys(body.updates).forEach(function(k){ writeKey_(k,body.updates[k]); }); }
    else if(body.key){ writeKey_(body.key, body.value); }
    return json_({ok:true});
  }catch(err){ return json_({ok:false, error:String(err)}); }
}
