"""Build the Excel-like HTML explorer.
   embed=True  -> data baked in (works as a standalone double-click file).
   embed=False -> data pulled from the local helper's /data endpoint, and the
                  in-page 'Update' button calls /update to refresh live."""
import json, math
import pandas as pd
import hap_writer as hw

def _payload(df):
    cols=list(df.columns)
    def ctype(c):
        if c in hw.DATE_COLS: return "date"
        if c in hw.NUM_COLS: return "num"
        if c in hw.TEXT_COLS: return "text"
        return "gen"
    types=[ctype(c) for c in cols]
    def cell(c,v,t):
        if v is None or (isinstance(v,float) and math.isnan(v)): return None
        if t=="date":
            try: return pd.Timestamp(v).strftime("%m/%d/%Y")
            except Exception: return str(v).strip()
        if t=="num":
            try:
                f=float(v)
                if math.isnan(f): return None
                return round(f,4) if c=="Rent-to-FMR Ratio" else round(f,2)
            except Exception: return None
        return str(v).strip()
    rows=[[cell(c,r[c],t) for c,t in zip(cols,types)] for _,r in df.iterrows()]
    return {"cols":cols,"types":types,"rows":rows}

def payload_json(df, asof):
    p=_payload(df); p["asof"]=asof
    return json.dumps(p,separators=(",",":"),default=str)

def write_html(df, path, asof, embed=True):
    if embed:
        data_js = payload_json(df, asof)
    else:
        data_js = "null"
    html=TEMPLATE.replace("/*__DATA__*/", data_js)
    with open(path,"w",encoding="utf-8") as f:
        f.write(html)
    return (len(df) if df is not None else 0)

TEMPLATE = r'''<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>HAP Database</title>
<style>
:root{--navy:#1f4e78;--navy2:#2e6da4;--line:#dfe4ec;--hi:#fff4cf;--grn:#1c7a3f;}
*{box-sizing:border-box}
html,body{margin:0;height:100%;font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#1a2230}
#top{background:var(--navy);color:#fff;padding:10px 16px;display:flex;align-items:center;gap:14px;flex-wrap:wrap}
#top h1{font-size:16px;margin:0;font-weight:650}
#top .asof{font-size:12px;opacity:.85}
#top .spacer{flex:1}
#count{font-size:13px;opacity:.9}
#top input.gsearch{padding:7px 10px;border:0;border-radius:7px;font-size:13px;min-width:200px}
#top button{border:0;border-radius:7px;padding:8px 14px;font-weight:650;font-size:13px;cursor:pointer}
#btnUpdate{background:var(--grn);color:#fff}
#btnUpdate:disabled{opacity:.6;cursor:progress}
#btnClear,#btnExport{background:#eaf0f6;color:#12324f}
#top button:hover{filter:brightness(1.07)}
#status{font-size:12.5px;background:#12324f;color:#dff0ff;padding:5px 14px;display:none}
#grid{position:absolute;top:52px;bottom:0;left:0;right:0;overflow:auto;background:#fff}
#grid.withstatus{top:79px}
table{border-collapse:separate;border-spacing:0;font-size:12px;white-space:nowrap}
thead th{position:sticky;top:0;z-index:3;background:var(--navy);color:#fff;padding:6px 9px;text-align:left;font-weight:600;cursor:pointer;border-right:1px solid #35608a}
thead th .ar{font-size:9px;opacity:.7;margin-left:3px}
thead tr.filters th{top:29px;z-index:2;background:#eef2f7;padding:3px}
thead tr.filters input{width:100%;min-width:90px;border:1px solid #c7cfdb;border-radius:5px;padding:4px 6px;font-size:12px}
th.rk,td.rk{position:sticky;left:0;z-index:1;background:#f3f6fa;color:#7a8494;text-align:right;border-right:1px solid var(--line);font-variant-numeric:tabular-nums}
thead th.rk{z-index:4;background:var(--navy);color:#fff}
tbody td{padding:5px 9px;border-bottom:1px solid #eef1f5;border-right:1px solid #f0f3f7;max-width:340px;overflow:hidden;text-overflow:ellipsis}
tbody td.num{text-align:right;font-variant-numeric:tabular-nums}
tbody tr:hover td{background:#eef4fb}
tbody tr:hover td.rk{background:#e4ecf5}
.spx td{padding:0;border:0}
</style></head><body>
<div id="top">
  <h1>HAP Database</h1><span class="asof" id="asof"></span>
  <button id="btnUpdate">⟳ Update to current month</button>
  <input class="gsearch" id="gsearch" placeholder="Search all columns…" autocomplete="off">
  <span class="spacer"></span>
  <span id="count"></span>
  <button id="btnClear">Clear filters</button>
  <button id="btnExport">Export to Excel</button>
</div>
<div id="status"></div>
<div id="grid"><table><thead id="thead"></thead><tbody id="tbody"></tbody></table></div>
<script>
let DATA = /*__DATA__*/;
let C,T,R,N,ix;
let filters=[], gq="", sortCol=0, sortDir=1, view=[];
const MONEY=new Set(["0BR Rent","1BR Rent","2BR Rent","3BR Rent","4BR Rent","5BR+ Rent","0BR UA","1BR UA","2BR UA","3BR UA","4BR UA","5BR+ UA"]);
const INT=new Set(["Total Units","Contract Units","property_total_unit_count"]);
const RATIO="Rent-to-FMR Ratio";
const ROWH=25;

function applyData(){
  if(!DATA){DATA={cols:[],types:[],rows:[],asof:"(no data yet — click Update)"};}
  C=DATA.cols;T=DATA.types;R=DATA.rows;N=C.length;ix={};C.forEach((c,i)=>ix[c]=i);
  document.getElementById('asof').textContent='Data as of '+DATA.asof;
  filters=new Array(N).fill(""); gq="";
  sortCol=(ix["Property Name"]!==undefined)?ix["Property Name"]:0; sortDir=1;
}
function disp(ci,v){
  if(v===null||v===undefined||v==="")return"";
  const c=C[ci];
  if(c===RATIO)return Number(v).toFixed(2);
  if(MONEY.has(c))return "$"+Number(v).toLocaleString();
  if(INT.has(c))return Number(v).toLocaleString();
  return v;
}
function passes(r){
  if(gq){ let hit=false; for(let i=0;i<N;i++){const v=r[i]; if(v!==null&&(''+v).toLowerCase().includes(gq)){hit=true;break;}} if(!hit)return false; }
  for(let i=0;i<N;i++){ const f=filters[i]; if(f){ const v=r[i]; if(v===null||!(''+v).toLowerCase().includes(f))return false; } }
  return true;
}
function recompute(){
  view=[]; for(let k=0;k<R.length;k++){ if(passes(R[k])) view.push(k); }
  const i=sortCol, num=(T[i]==="num");
  view.sort((a,b)=>{
    let x=R[a][i], y=R[b][i];
    const xe=(x===null||x===undefined||x===""), ye=(y===null||y===undefined||y==="");
    if(xe&&ye)return 0; if(xe)return 1; if(ye)return -1;
    if(num)return (x-y)*sortDir;
    if(T[i]==="date"){x=Date.parse(x);y=Date.parse(y); return (x-y)*sortDir;}
    return (''+x).localeCompare(''+y)*sortDir;
  });
  document.getElementById('count').textContent='Showing '+view.length.toLocaleString()+' of '+R.length.toLocaleString();
  renderHead(); renderBody(true);
}
function renderHead(){
  const th=document.getElementById('thead');
  let h1='<tr><th class="rk">#</th>';
  for(let i=0;i<N;i++){ const a=i===sortCol?(sortDir<0?'▼':'▲'):''; h1+='<th data-i="'+i+'">'+C[i]+' <span class="ar">'+a+'</span></th>'; }
  h1+='</tr>';
  let h2='<tr class="filters"><th class="rk"></th>';
  for(let i=0;i<N;i++){ h2+='<th><input data-f="'+i+'" value="'+(filters[i].replace(/"/g,'&quot;'))+'" placeholder="filter"></th>'; }
  h2+='</tr>';
  th.innerHTML=h1+h2;
  th.querySelectorAll('th[data-i]').forEach(el=>el.onclick=()=>{const i=+el.dataset.i; if(i===sortCol)sortDir*=-1; else{sortCol=i;sortDir=(T[i]==="num"||T[i]==="date")?-1:1;} recompute();});
  th.querySelectorAll('input[data-f]').forEach(inp=>{
    inp.oninput=()=>{filters[+inp.dataset.f]=inp.value.trim().toLowerCase(); clearTimeout(inp._t); inp._t=setTimeout(recompute,180);};
    inp.onclick=e=>e.stopPropagation();
  });
}
const grid=document.getElementById('grid'), tbody=document.getElementById('tbody');
function renderBody(reset){
  const total=view.length;
  const vh=grid.clientHeight, scrollTop=grid.scrollTop;
  const start=Math.max(0,Math.floor(scrollTop/ROWH)-6);
  const vis=Math.ceil(vh/ROWH)+14;
  const end=Math.min(total,start+vis);
  let html='<tr class="spx"><td colspan="'+(N+1)+'" style="height:'+(start*ROWH)+'px"></td></tr>';
  for(let p=start;p<end;p++){
    const r=R[view[p]];
    html+='<tr><td class="rk">'+(p+1)+'</td>';
    for(let i=0;i<N;i++){ const num=(T[i]==="num"); html+='<td class="'+(num?'num':'')+'">'+ (r[i]===null?'':esc(disp(i,r[i]))) +'</td>'; }
    html+='</tr>';
  }
  html+='<tr class="spx"><td colspan="'+(N+1)+'" style="height:'+((total-end)*ROWH)+'px"></td></tr>';
  tbody.innerHTML=html;
  if(reset)grid.scrollTop=scrollTop;
}
function esc(s){return (''+s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
let raf=null;
grid.addEventListener('scroll',()=>{ if(raf)return; raf=requestAnimationFrame(()=>{raf=null;renderBody(false);}); });
window.addEventListener('resize',()=>renderBody(false));
document.getElementById('gsearch').addEventListener('input',e=>{gq=e.target.value.trim().toLowerCase(); clearTimeout(window._g); window._g=setTimeout(recompute,200);});
document.getElementById('btnClear').onclick=()=>{filters=new Array(N).fill("");gq="";document.getElementById('gsearch').value="";renderHead();recompute();};

/* ---- status helper + Update button ---- */
function setStatus(msg,show=true){const s=document.getElementById('status');s.textContent=msg;s.style.display=show?'block':'none';grid.classList.toggle('withstatus',show);renderBody(false);}
const served = location.protocol.startsWith('http');
document.getElementById('btnUpdate').onclick=doUpdate;
async function doUpdate(){
  if(!served){
    setStatus("This is the offline file. To pull the current month, open the HAP app (it starts the helper), then click Update there.");
    return;
  }
  const btn=document.getElementById('btnUpdate'); btn.disabled=true;
  setStatus("Downloading this month's HUD files and rebuilding… this takes about 5 minutes. You can leave this tab open.");
  try{
    const r=await fetch('/update',{method:'POST'});
    const j=await r.json();
    if(!j.ok) throw new Error(j.error||'update failed');
    const d=await fetch('/data',{cache:'no-store'}); DATA=await d.json();
    applyData(); recompute();
    setStatus("✓ Updated. Data is now current as of "+DATA.asof+"  ("+Number(j.rows).toLocaleString()+" properties).");
  }catch(e){ setStatus("Update failed: "+e.message+"  — check the helper window / internet and try again."); }
  btn.disabled=false;
}

/* ---------- Export to .xlsx (pure JS, offline, formatted like the master) ---------- */
document.getElementById('btnExport').onclick=exportXlsx;
function xmlesc(s){return (''+s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function dserial(mdy){ const p=(''+mdy).split('/'); if(p.length!==3)return null; const d=new Date(Date.UTC(+p[2],+p[0]-1,+p[1])); const epoch=Date.UTC(1899,11,30); return Math.round((d-epoch)/86400000); }
function styleFor(i){ const c=C[i], t=T[i]; if(t==="date")return 2; if(t==="num")return (c===RATIO?4:3); if(t==="text")return 1; return 0; }
function colLetter(n){let s="";n++;while(n>0){let m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=(n-n%26)/26;}return s;}
function buildSheetXml(){
  let rows='';
  let hc=''; for(let i=0;i<N;i++){ hc+='<c r="'+colLetter(i)+'1" t="inlineStr"><is><t xml:space="preserve">'+xmlesc(C[i])+'</t></is></c>'; }
  rows+='<row r="1">'+hc+'</row>';
  for(let p=0;p<view.length;p++){
    const r=R[view[p]], rn=p+2; let cc='';
    for(let i=0;i<N;i++){
      const v=r[i]; if(v===null||v===undefined||v===""){continue;}
      const ref=colLetter(i)+rn, st=styleFor(i), t=T[i];
      if(t==="num"){ cc+='<c r="'+ref+'" s="'+st+'"><v>'+v+'</v></c>'; }
      else if(t==="date"){ const s=dserial(v); if(s!==null)cc+='<c r="'+ref+'" s="'+st+'"><v>'+s+'</v></c>'; else cc+='<c r="'+ref+'" t="inlineStr"><is><t>'+xmlesc(v)+'</t></is></c>'; }
      else { cc+='<c r="'+ref+'" s="'+st+'" t="inlineStr"><is><t xml:space="preserve">'+xmlesc(v)+'</t></is></c>'; }
    }
    rows+='<row r="'+rn+'">'+cc+'</row>';
  }
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetData>'+rows+'</sheetData><autoFilter ref="A1:'+colLetter(N-1)+(view.length+1)+'"/></worksheet>';
}
const STYLES='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
 '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'+
 '<numFmts count="1"><numFmt numFmtId="164" formatCode="m/d/yyyy"/></numFmts>'+
 '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'+
 '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'+
 '<borders count="1"><border/></borders>'+
 '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'+
 '<cellXfs count="5">'+
 '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'+
 '<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'+
 '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'+
 '<xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'+
 '<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'+
 '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
function xlsxFiles(){
  return {
   "[Content_Types].xml":'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
   "_rels/.rels":'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
   "xl/workbook.xml":'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="MasterTable" sheetId="1" r:id="rId1"/></sheets></workbook>',
   "xl/_rels/workbook.xml.rels":'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
   "xl/styles.xml":STYLES,
   "xl/worksheets/sheet1.xml":buildSheetXml()
  };
}
const crcTable=(()=>{const t=[];for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=c&1?0xEDB88320^(c>>>1):c>>>1;t[n]=c>>>0;}return t;})();
function crc32(bytes){let c=0xFFFFFFFF;for(let i=0;i<bytes.length;i++)c=crcTable[(c^bytes[i])&0xFF]^(c>>>8);return (c^0xFFFFFFFF)>>>0;}
function strBytes(s){return new TextEncoder().encode(s);}
function exportXlsx(){
  const files=xlsxFiles();
  let offset=0; const central=[]; const chunks=[];
  const names=Object.keys(files);
  function u16(a){return [a&255,(a>>8)&255];}
  function u32(a){return [a&255,(a>>8)&255,(a>>16)&255,(a>>24)&255];}
  for(const name of names){
    const nameB=strBytes(name), data=strBytes(files[name]), crc=crc32(data);
    const local=[].concat(u32(0x04034b50),u16(20),u16(0),u16(0),u16(0),u16(0),u32(crc),u32(data.length),u32(data.length),u16(nameB.length),u16(0));
    chunks.push(new Uint8Array(local)); chunks.push(nameB); chunks.push(data);
    const cen=[].concat(u32(0x02014b50),u16(20),u16(20),u16(0),u16(0),u16(0),u16(0),u32(crc),u32(data.length),u32(data.length),u16(nameB.length),u16(0),u16(0),u16(0),u16(0),u32(0),u32(offset));
    central.push({head:new Uint8Array(cen),name:nameB});
    offset+=local.length+nameB.length+data.length;
  }
  const cstart=offset; let csize=0; const cchunks=[];
  for(const c of central){cchunks.push(c.head);cchunks.push(c.name);csize+=c.head.length+c.name.length;}
  const eocd=[].concat(u32(0x06054b50),u16(0),u16(0),u16(names.length),u16(names.length),u32(csize),u32(cstart),u16(0));
  const all=chunks.concat(cchunks,[new Uint8Array(eocd)]);
  let len=0; all.forEach(a=>len+=a.length); const out=new Uint8Array(len); let o=0; all.forEach(a=>{out.set(a,o);o+=a.length;});
  const blob=new Blob([out],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob);
  const d=new Date(); const stamp=d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
  a.download='HAP_Export_'+stamp+'.xlsx'; a.click(); setTimeout(()=>URL.revokeObjectURL(a.href),4000);
}
/* ---- boot: use served data if available, else embedded ---- */
async function boot(){
  if(served && DATA===null){
    try{ const r=await fetch('/data',{cache:'no-store'}); if(r.ok){ DATA=await r.json(); } }catch(e){}
  }
  applyData(); recompute();
  if(served && (!R || R.length===0)){ setStatus("No data yet. Click “⟳ Update to current month” to pull the latest HUD numbers."); }
}
boot();
</script></body></html>'''
