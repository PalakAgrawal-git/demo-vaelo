/**
 * VAELO — Google Sheets TWO-WAY sync, MULTI-CLIENT.
 * ------------------------------------------------------------
 * Links several client content-calendar spreadsheets to the app through ONE
 * deployment. Each client's calendar tab (DATE / STATUS / TOPIC / PRODUCT /
 * FORMAT) syncs both ways with that client's pipeline in the app.
 *
 * ┌─ SET THIS UP ────────────────────────────────────────────────────────────┐
 * │ In CLIENT_SHEETS below, map each client name to their spreadsheet:        │
 * │   'THIS'  = the spreadsheet this script is attached to (its content tab). │
 * │   ''      = not linked yet — that client is skipped.                      │
 * │   a URL or ID = that client's separate spreadsheet.                       │
 * │ The client name must EXACTLY match the app's client names.               │
 * │ Every linked sheet must be owned by, or shared with, THIS Google account. │
 * └───────────────────────────────────────────────────────────────────────────┘
 *
 * After editing this file: Deploy ▸ Manage deployments ▸ ✏ ▸ New version ▸ Deploy.
 */
var CLIENT_SHEETS = {
  'SimpliCare':       'THIS',   // calendar lives in this spreadsheet
  'Zerolys':          'https://docs.google.com/spreadsheets/d/1xnC74tL-Gnj1jbb4Sz_sc6HOtYBs2hkAbg-rx_pT0Ko/edit',
  'Marigold Miraaya': '',       // paste Marigold sheet URL or ID
  'DVOC Institute':   '',       // paste URL or ID
  'Tribal Zone':      ''        // paste URL or ID
};

var STORE_TAB = 'store';
var FIN_TAB   = 'Finance';
var STAGES = ['Idea','Dhruv approval','Calendar sent','Calendar approved','Creative in progress','Internal approval','Client approval','Scheduled/Posted'];

function ss_(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function tab_(name){ var s=ss_().getSheetByName(name); if(!s) s=ss_().insertSheet(name); return s; }
function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function today_(){ return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function genId_(){ return Math.random().toString(36).slice(2,9); }
function cell_(v){ if(v instanceof Date){ return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd MMM yyyy'); } return v==null?'':String(v).trim(); }
function parseMonth_(d){ var m={jan:'01',feb:'02',mar:'03',apr:'04',may:'05',jun:'06',jul:'07',aug:'08',sep:'09',oct:'10',nov:'11',dec:'12'}; var p=String(d||'').trim().split(/\s+/); if(p.length>=3){ var mo=m[p[1].toLowerCase().slice(0,3)]; if(mo) return p[2]+'-'+mo; } return ''; }

function statusToStage_(status, fallback){
  var s=String(status||'').toLowerCase();
  if(/post|schedul|publish|live|done/.test(s)) return 7;
  if(/ready|approv/.test(s)) return 6;
  if(/progress|wip|edit|draft|film|shoot/.test(s)) return 4;
  if(/plan|idea|brief/.test(s)) return 2;
  return fallback!=null?fallback:2;
}
function stageToStatus_(stage){
  if(stage>=7) return 'Posted';
  if(stage>=6) return 'Ready';
  if(stage>=4) return 'Work In Progress';
  return 'Planned';
}

/* ---------- which spreadsheet for a client ---------- */
function sheetIdFromRef_(ref){ var m=String(ref||'').match(/\/d\/([a-zA-Z0-9-_]+)/); return m?m[1]:String(ref||'').trim(); }
function openForClient_(ref){
  if(ref==='THIS') return ss_();
  if(!ref) return null;                 // blank → not linked
  try{ return SpreadsheetApp.openById(sheetIdFromRef_(ref)); }catch(e){ return null; }
}

/* ---------- key/value store (blob + finance) ---------- */
function storeSheet_(){ var sh=ss_().getSheetByName(STORE_TAB); if(!sh){ sh=ss_().insertSheet(STORE_TAB); sh.appendRow(['key','value','updatedAt']); } return sh; }
function readAll_(){ var sh=storeSheet_(); var rows=sh.getDataRange().getValues(); var out={}; for(var i=1;i<rows.length;i++){ if(rows[i][0]) out[rows[i][0]]=rows[i][1]; } return out; }
function getStoreValue_(key){ var a=readAll_(); return a.hasOwnProperty(key)?a[key]:null; }
function upsertStore_(key,value){ var sh=storeSheet_(); var rows=sh.getDataRange().getValues(); for(var i=1;i<rows.length;i++){ if(rows[i][0]===key){ sh.getRange(i+1,2).setValue(value); sh.getRange(i+1,3).setValue(new Date()); return; } } sh.appendRow([key,value,new Date()]); }

/* ---------- find ALL content calendar tabs inside a spreadsheet ----------
   Column names are matched loosely so different clients' layouts work:
     date  ← Date
     status← Status / Stage
     topic ← Topic / Theme / Title / Hook
     format← Format / Type
     product← Product (optional) */
function findAllContentSheetsIn_(spreadsheet){
  var res=[], sheets=spreadsheet.getSheets();
  for(var i=0;i<sheets.length;i++){
    var sh=sheets[i], name=sh.getName();
    if(name===STORE_TAB || name===FIN_TAB) continue;
    var rng=sh.getDataRange(); if(rng.getNumRows()<1) continue;
    var values=rng.getValues();
    for(var r=0;r<Math.min(values.length,8);r++){
      var hdr=values[r].map(function(c){ return String(c).toLowerCase(); });
      var status=hdr.findIndex(function(h){ return /status|stage/.test(h); });
      var topic=hdr.findIndex(function(h){ return /topic|theme|title|hook/.test(h); });
      var date=hdr.findIndex(function(h){ return /date/.test(h); });
      if(status>=0 && topic>=0 && date>=0){
        res.push({ sheet:sh, headerRow:r, values:values, colMap:{
          date:date, status:status, topic:topic,
          format: hdr.findIndex(function(h){ return /format|type/.test(h); }),
          product: hdr.findIndex(function(h){ return /product/.test(h); })
        }});
        break;   // one header row per tab
      }
    }
  }
  return res;
}

/* ---------- sheet → pipeline (all linked clients) ---------- */
function itemsFromInfo_(info, client, prevByDate){
  var cm=info.colMap, values=info.values, hr=info.headerRow;
  function v(row,i){ return i>=0?cell_(row[i]):''; }
  var out=[];
  for(var r=hr+1;r<values.length;r++){
    var row=values[r];
    var date=v(row,cm.date); if(!date || !/\d/.test(date)) continue;
    var prev=prevByDate[date]||{};
    // keep the app's precise stage if the sheet STATUS still matches it
    // (so pipeline-only gates like "Dhruv approval" survive the round-trip)
    var sheetStatus=v(row,cm.status);
    var stage = (prev.stage!=null && stageToStatus_(prev.stage)===sheetStatus)
                ? prev.stage : statusToStage_(sheetStatus, prev.stage);
    var it={
      id: prev.id||genId_(), client: client, date: date,
      day: prev.day||'', month: prev.month||parseMonth_(date),
      format: cm.format>=0 ? v(row,cm.format) : (prev.format||''),
      topic:  v(row,cm.topic) || prev.topic || '',
      product: cm.product>=0 ? v(row,cm.product) : (prev.product||''),
      owner: prev.owner||'—', stage: stage, log: (prev.log||[]).slice()
    };
    if(prev.stage!=null && prev.stage!==stage){ it.log.push({d:today_(), t:'Stage moved to "'+STAGES[stage]+'". — by Sheet', by:'Sheet'}); }
    out.push(it);
  }
  return out;
}
function contentFromSheet_(){
  var raw=getStoreValue_('vaelo-content-state');
  var blob=raw?JSON.parse(raw):{items:[],ideas:[]};
  if(!blob.items) blob.items=[];
  var synced=[], syncedClients={};
  Object.keys(CLIENT_SHEETS).forEach(function(client){
    var ssx=openForClient_(CLIENT_SHEETS[client]); if(!ssx) return;
    var infos=findAllContentSheetsIn_(ssx); if(!infos.length) return;
    syncedClients[client]=true;
    var prevByDate={}; blob.items.forEach(function(it){ if(it.client===client) prevByDate[it.date]=it; });
    infos.forEach(function(info){ synced=synced.concat(itemsFromInfo_(info, client, prevByDate)); });
  });
  var others=blob.items.filter(function(it){ return !syncedClients[it.client]; });
  blob.items=others.concat(synced);
  upsertStore_('vaelo-content-state', JSON.stringify(blob));
  return JSON.stringify(blob);
}

/* ---------- pipeline → sheet (write STATUS/Topic/Product/Format back) ---------- */
function pushContentToSheet_(state){
  Object.keys(CLIENT_SHEETS).forEach(function(client){
    var ssx=openForClient_(CLIENT_SHEETS[client]); if(!ssx) return;
    var infos=findAllContentSheetsIn_(ssx); if(!infos.length) return;
    var byDate={}; (state.items||[]).forEach(function(it){ if(it.client===client) byDate[it.date]=it; });
    infos.forEach(function(info){
      var sh=info.sheet, cm=info.colMap, values=info.values, hr=info.headerRow;
      for(var r=hr+1;r<values.length;r++){
        var date=cell_(values[r][cm.date]); if(!date || !/\d/.test(date)) continue;
        var it=byDate[date]; if(!it) continue;
        if(cm.status>=0)  sh.getRange(r+1, cm.status+1).setValue(stageToStatus_(it.stage));
        if(cm.topic>=0 && it.topic)   sh.getRange(r+1, cm.topic+1).setValue(it.topic);
        if(cm.product>=0)             sh.getRange(r+1, cm.product+1).setValue(it.product);
        if(cm.format>=0 && it.format) sh.getRange(r+1, cm.format+1).setValue(it.format);
      }
    });
  });
}

/* ---------- Finance (one-way readable mirror, in the bound spreadsheet) ---------- */
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

/* ---------- diagnostics: hit ...exec?diag=1 to see per-client status ---------- */
function diag_(){
  var out={};
  Object.keys(CLIENT_SHEETS).forEach(function(client){
    var ref=CLIENT_SHEETS[client];
    var d={ linked: ref?(ref==='THIS'?'THIS':'yes'):'no' };
    if(!ref){ d.status='not linked (blank)'; out[client]=d; return; }
    var ssx; try{ ssx=openForClient_(ref); }catch(e){ d.status='OPEN ERROR — likely not shared with this account'; d.error=String(e); out[client]=d; return; }
    if(!ssx){ d.status='could NOT open (bad id, or not shared with this account)'; out[client]=d; return; }
    d.opened=ssx.getName();
    var infos=findAllContentSheetsIn_(ssx);
    if(!infos.length){ d.status='opened, but NO tab has recognizable headers'; d.tabsFound=ssx.getSheets().map(function(s){return s.getName();}); out[client]=d; return; }
    d.status='OK'; d.contentTabs=infos.map(function(i){return i.sheet.getName();}); d.dataRows=infos.reduce(function(a,i){return a+(i.values.length-i.headerRow-1);},0);
    out[client]=d;
  });
  return out;
}

/* ---------- endpoints ---------- */
function doGet(e){
  var key=e&&e.parameter&&e.parameter.key;
  if(e&&e.parameter&&e.parameter.diag){ return json_(diag_()); }
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
