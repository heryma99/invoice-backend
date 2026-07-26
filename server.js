/**
 * 发票系统后端代理服务（自适应：云端 HTTP 直连 / 本机 lark-cli）
 *
 * 云端（Railway 等）：设置环境变量 APP_ID + APP_SECRET + SPREADSHEET_TOKEN + SHEET_ID
 *   -> 用 tenant_access_token 直接调飞书 API，不依赖本机 lark-cli / Windows 钥匙串
 * 本机：不设 APP_ID，则走原 lark-cli 封装（保留你的 user 身份，零改动）
 */
const express = require('express');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const app = express();
const PORT = process.env.PORT || (process.env.RAILWAY ? 3000 : 3460);
const CLOUD = !!(process.env.APP_ID && process.env.APP_SECRET); // 云端走 HTTP 直连
const SPREADSHEET_TOKEN = process.env.SPREADSHEET_TOKEN || 'RWUDsL3PGh0J07tnAPocCm8xnrb';
const SHEET_ID = process.env.SHEET_ID || '08e89c';
const DOMAIN = 'https://open.feishu.cn';

app.use(express.json({ limit: '50mb' }));
app.use((req,res,next)=>{
  res.header('Access-Control-Allow-Origin','*');
  res.header('Access-Control-Allow-Methods','GET,POST,OPTIONS');
  res.header('Access-Control-Allow-Headers','Content-Type');
  if(req.method==='OPTIONS') return res.sendStatus(200);
  next();
});

/* ===================== 云端：纯 HTTP 直连飞书 ===================== */
let _token=null, _exp=0;
async function cloudToken(){
  if(_token && Date.now()<_exp) return _token;
  const r = await fetch(DOMAIN+'/open-apis/auth/v3/tenant_access_token/internal', {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({app_id:process.env.APP_ID, app_secret:process.env.APP_SECRET})
  });
  const j = await r.json();
  if(!j.tenant_access_token) throw new Error('token fail: '+JSON.stringify(j));
  _token=j.tenant_access_token; _exp=Date.now()+(j.expire-120)*1000;
  return _token;
}
async function feishu(method, p, body){
  const t = await cloudToken();
  const opt = {method, headers:{Authorization:`Bearer ${t}`,'Content-Type':'application/json'}};
  if(body) opt.body=JSON.stringify(body);
  const r = await fetch(DOMAIN+p, opt);
  const text = await r.text();
  let j; try{ j=JSON.parse(text); }catch(e){ throw new Error(`${p} HTTP ${r.status} 非JSON响应`); }
  if(j.code && j.code!==0){ const e=new Error(`${p} -> ${j.code} ${j.msg}`); e.code=j.code; throw e; }
  return j;
}
async function cloudFindRow(fid){
  // 用 values 接口读 B 列后在本地搜索（避免 find 接口路径差异）
  const vals = await cloudCsvColumn('B', 1200);
  const key = String(fid).toUpperCase();
  for(let i=0;i<vals.length;i++){ if(String(vals[i]).toUpperCase()===key) return i+1; }
  return null;
}
async function cloudGetAttachment(row, col){
  // 镜像表附件列存的是富文本字符串 "[{\"fileToken\":\"...\"}]"，从中解析 token
  const range = `${SHEET_ID}!${col}${row}:${col}${parseInt(row)+1}`;
  const j = await feishu('GET', `/open-apis/sheets/v2/spreadsheets/${SPREADSHEET_TOKEN}/values/${range}`);
  const vals = j.data && j.data.valueRange && j.data.valueRange.values;
  const raw = vals && vals[0] && vals[0][0];
  if(!raw) return null;
  const str = typeof raw==='string'?raw:JSON.stringify(raw);
  const m = str.match(/fileToken["']?\s*[:"]\s*([A-Za-z0-9]+)/i);
  return m?{token:m[1], name:''}:null;
}
async function cloudDownload(token, outPath){
  const t = await cloudToken();
  const r = await fetch(DOMAIN+`/open-apis/drive/v1/medias/${token}/download`, {headers:{Authorization:`Bearer ${t}`}});
  if(!r.ok) throw new Error('download HTTP '+r.status);
  const buf = Buffer.from(await r.arrayBuffer());
  fs.writeFileSync(outPath, buf);
  return buf.length;
}
async function cloudCsvColumn(col, maxRow){
  // 批量读 B 列用于搜索缓存
  const range = `${col}1:${col}${maxRow}`;
  const j = await feishu('GET', `/open-apis/sheets/v2/spreadsheets/${SPREADSHEET_TOKEN}/values/${range}`);
  const rows = (j.data && j.data.valueRange && j.data.valueRange.values) || [];
  return rows.map(r=>r[0]?String(r[0]):'');
}

/* ===================== 本机：lark-cli ===================== */
const BASH = 'C:/Program Files/Git/bin/bash.exe';
const SHELL = { shell: BASH, encoding: 'utf-8' };
const IDENT = process.env.LARK_AS || 'bot';
function larkRaw(cmd){
  return JSON.parse(execSync(`${cmd} --as ${IDENT} --json`, {...SHELL, maxBuffer:50*1024*1024, timeout:30000}));
}
function larkFindRow(fid){
  const r=larkRaw(`lark-cli sheets +cells-search --url "https://plbrands.feishu.cn/sheets/${SPREADSHEET_TOKEN}" --sheet-id ${SHEET_ID} --find "${fid}" --range B1:B1200`);
  if(r.ok && r.data && r.data.matches && r.data.matches.length){
    return parseInt(r.data.matches[0].address.replace(/[A-Z]/g,''));
  }
  return null;
}
function larkGetAttachment(row, col){
  const r=larkRaw(`lark-cli sheets +cells-get --url "https://plbrands.feishu.cn/sheets/${SPREADSHEET_TOKEN}" --sheet-id ${SHEET_ID} --range ${col}${row} --include value,formula,comment`);
  const c=r.data && r.data.ranges && r.data.ranges[0] && r.data.ranges[0].cells && r.data.ranges[0].cells[0] && r.data.ranges[0].cells[0][0];
  if(!c) return null;
  if(c.attachmentToken) return {token:c.attachmentToken, name:c.value||''};
  if(c.rich_text) for(const rt of c.rich_text){
    if(rt.attachment_token) return {token:rt.attachment_token, name:rt.text||c.value||''};
    if(rt.attachment && rt.attachment.file_token) return {token:rt.attachment.file_token, name:rt.attachment.file_name||rt.text||c.value||''};
  }
  return null;
}
function larkDownload(token, outRel){
  const full=path.join(TMP_DIR, outRel);
  try{ larkRaw(`lark-cli drive +download --file-token ${token} --output "tmp/${outRel}" --overwrite`); if(fs.existsSync(full)) return fs.statSync(full).size; }catch(e){ console.log('[dl] drive fail', e.message.substring(0,60)); }
  try{ execSync(`lark-cli api GET "/open-apis/drive/v1/medias/${token}/download" --as ${IDENT} --output "tmp/${outRel}"`, {...SHELL, maxBuffer:100*1024*1024, timeout:30000}); if(fs.existsSync(full)&&fs.statSync(full).size>100) return fs.statSync(full).size; }catch(e){ console.log('[dl] media fail', e.message.substring(0,60)); }
  return 0;
}

/* ===================== 搜索缓存 ===================== */
let rowCache=null;
async function buildCache(){
  console.log('[cache] 预加载货件号索引...');
  try{
    const vals = CLOUD ? await cloudCsvColumn('B',960) : (()=>{ const r=larkRaw(`lark-cli sheets +csv-get --url "https://plbrands.feishu.cn/sheets/${SPREADSHEET_TOKEN}" --sheet-id ${SHEET_ID} --range B1:B960`); return (r.data&&r.data.annotated_csv||'').split('\n').map(l=>{const m=l.match(/\[row=(\d+)\]\s*(.+)/); return m?m[2].replace(/^"|"$/g,'').trim():'';}); })();
    rowCache={};
    vals.forEach((v,i)=>{ const key=String(v).toUpperCase(); if(key) rowCache[key]=i+1; });
    console.log(`[cache] 索引就绪: ${Object.keys(rowCache).length} 个货件号`);
  }catch(e){ console.log('[cache] 预加载失败:', e.message.substring(0,100)); rowCache={}; }
}
async function findRow(fid){
  if(!rowCache) await buildCache();
  const key=String(fid).toUpperCase();
  if(rowCache[key]){ console.log(`[search] 缓存命中 row=${rowCache[key]}`); return rowCache[key]; }
  try{
    const n = CLOUD ? await cloudFindRow(fid) : larkFindRow(fid);
    if(n>0){ rowCache[key]=n; return n; }
  }catch(e){ console.log('[search] fallback fail:', e.message.substring(0,80)); }
  return null;
}
async function getAttachment(row,col){ return CLOUD ? await cloudGetAttachment(row,col) : larkGetAttachment(row,col); }
async function download(token,outRel){ return CLOUD ? await cloudDownload(token,path.join(TMP_DIR,outRel)) : larkDownload(token,outRel); }

/* ===================== 解析 xlsx ===================== */
const TMP_DIR = path.join(__dirname,'tmp');
if(!fs.existsSync(TMP_DIR)) fs.mkdirSync(TMP_DIR,{recursive:true});

async function parseXlsx(xlsxPath, fnskuPath){
  const wb=new ExcelJS.Workbook(); await wb.xlsx.readFile(xlsxPath);
  const ws=wb.worksheets[0]; if(!ws) return {items:[], error:'空表'};
  let fnskuMap={};
  if(fnskuPath && fs.existsSync(fnskuPath)){
    try{ const wb2=new ExcelJS.Workbook(); await wb2.xlsx.readFile(fnskuPath); const ws2=wb2.worksheets[0];
      if(ws2) for(let r=2;r<=ws2.rowCount;r++){ const rr=ws2.getRow(r); const ms=String(rr.getCell(1).value||'').trim(), fn=String(rr.getCell(2).value||'').trim(), nm=String(rr.getCell(3).value||'').trim(); if(ms&&nm){fnskuMap[ms.toUpperCase()]=nm;fnskuMap[fn.toUpperCase()]=nm;} }
    }catch(e){ console.log('[fnsku] fail', e.message.substring(0,60)); }
  }
  let headerRow=1, isAmazon=false;
  for(let r=1;r<=Math.min(ws.rowCount,5);r++){
    const row=ws.getRow(r); let txt='';
    for(let c=1;c<=row.cellCount;c++) txt+=String(row.getCell(c).value||'');
    const low=txt.toLowerCase();
    if(low.includes('序号')&&low.includes('msku')&&low.includes('单箱数量')){headerRow=r; isAmazon=true; break;}
    if(low.includes('箱号')&&low.includes('sku')){headerRow=r; break;}
    if(low.includes('序号')&&low.includes('sku')){headerRow=r; break;}
  }
  const hdr=ws.getRow(headerRow); const headers=[];
  for(let c=1;c<=Math.max(hdr.cellCount,30);c++) headers.push(String(hdr.getCell(c).value||'').trim());
  const low=headers.map(h=>h.toLowerCase());
  const find=cands=>{const i=low.findIndex(h=>cands.some(k=>h===k||h.includes(k))); return i>=0?i+1:0;};
  const col={seq:find(['序号','no','number','num','row']), sku:find(['msku','sku','型号','产品型号']), fnsku:find(['fnsku']), nameCn:find(['中文品名','品名','名称','中文名称']), nameEn:find(['英文品名','英文名称','product name','nameen']), qty:find(['单箱数量','数量','qty','quantity']), ctns:find(['箱数','ctns','cartons']), boxName:find(['箱子名称','box label','箱号名称','label']), boxNo:find(['箱号','box no','carton','boxno','ctn']), weight:find(['箱子毛重','单箱毛重','箱重','重量','weight']), declare:find(['申报价','申报价值','declare','申报单价','unit price']), cost:find(['成本','采购价','cost']), material:find(['材质','material']), hs:find(['hs','海关编码','hscode']), brand:find(['品牌','brand']), len:find(['箱子长度','长','length']), wid:find(['箱子宽度','宽','width']), hgt:find(['箱子高度','高','height']), elec:find(['带电','elec']), magnet:find(['带磁','magnet']), saleUrl:find(['销售链接','sale url','saleurl'])};
  console.log(`[parse] headerRow=${headerRow} isAmazon=${isAmazon} sku=${col.sku} fnsku=${col.fnsku} qty=${col.qty} boxName=${col.boxName}`);
  const items=[];
  for(let r=headerRow+1;r<=ws.rowCount;r++){
    const row=ws.getRow(r); const vals=[];
    for(let c=1;c<=Math.max(row.cellCount,30);c++) vals.push(row.getCell(c).value);
    if(vals.every(v=>v===null||v===undefined||v==='')) continue;
    const get=idx=>idx>0&&vals[idx-1]!==null&&vals[idx-1]!==undefined?String(vals[idx-1]).trim():'';
    if(isAmazon){
      const sku=get(col.sku), fnsku=get(col.fnsku), qtyPB=parseInt(get(col.qty))||0, ctns=parseInt(get(col.ctns))||1, bn=get(col.boxName), en=get(col.nameEn);
      let labels=[];
      if(bn){ const last=bn.includes(' - ')?bn.split(' - ').pop():bn; const m=last.match(/^([A-Za-z]?)(\d+)\s*[～~\-]\s*([A-Za-z]?\d+)$/);
        if(m){ const prefix=m[1]||'B'; const start=parseInt(m[2]), end=parseInt(m[3].replace(/^[A-Za-z]/,'')); for(let b=start;b<=end;b++) labels.push(prefix+b); }
        else if(last.match(/^[A-Za-z]?\d+$/)) labels.push(last);
        else labels=bn.split(/[；;，,、\s]+/).filter(Boolean);
      }
      if(labels.length===0) for(let b=1;b<=ctns;b++) labels.push('B'+b);
      for(const lb of labels){ const it={boxNo:lb,sku,fnsku,nameCn:get(col.nameCn),nameEn:en,qty:qtyPB,declare:get(col.declare),cost:get(col.cost),material:get(col.material),hs:get(col.hs),brand:get(col.brand),weight:get(col.weight),len:get(col.len),wid:get(col.wid),hgt:get(col.hgt),elec:(get(col.elec)||'').toUpperCase().startsWith('Y')?'Y':'N',magnet:(get(col.magnet)||'').toUpperCase().startsWith('Y')?'Y':'N',saleUrl:get(col.saleUrl)};
        if(!it.nameEn){const k=(it.fnsku||it.sku||'').toUpperCase(); if(fnskuMap[k]) it.nameEn=fnskuMap[k];} items.push(it); }
    } else {
      const sku=get(col.sku)||get(col.model);
      const it={boxNo:get(col.boxNo),sku,fnsku:get(col.fnsku),nameCn:get(col.nameCn),nameEn:get(col.nameEn),qty:parseInt(get(col.qty))||1,declare:get(col.declare),cost:get(col.cost),material:get(col.material),hs:get(col.hs),brand:get(col.brand),weight:get(col.weight),len:get(col.len),wid:get(col.wid),hgt:get(col.hgt),elec:(get(col.elec)||'').toUpperCase().startsWith('Y')?'Y':'N',magnet:(get(col.magnet)||'').toUpperCase().startsWith('Y')?'Y':'N',saleUrl:get(col.saleUrl)};
      if(!it.nameEn){const k=(it.fnsku||it.sku||'').toUpperCase(); if(fnskuMap[k]) it.nameEn=fnskuMap[k];} items.push(it);
    }
  }
  return {items, error:null};
}

/* ===================== 路由 ===================== */
app.get('/api/health', (req,res)=>res.json({ok:true, mode:CLOUD?'cloud':'local', time:new Date().toISOString(), pid:process.pid}));
app.get('/api/debug-cache', async (req,res)=>{
  const fid=req.query.fid;
  if(!rowCache) await buildCache();
  if(!rowCache) return res.json({ok:false,msg:'cache not built',fid});
  const key=String(fid).toUpperCase();
  res.json({ok:true, mode:CLOUD?'cloud':'local', cacheSize:Object.keys(rowCache).length, hasKey:!!rowCache[key], row:rowCache[key], key, fid, sampleKeys:Object.keys(rowCache).slice(0,5)});
});
app.post('/api/fetch-packing-list', async (req,res)=>{
  const {fid}=req.body||{}; if(!fid) return res.status(400).json({ok:false, error:'缺 fid'});
  console.log(`\n========== ${fid} ==========`);
  try{
    const row=await findRow(fid);
    if(!row) return res.json({ok:false, error:`${fid} 不在FBA表`, code:'NOT_FOUND'});
    const pack=await getAttachment(row,'P');
    if(!pack) return res.json({ok:false, error:`行${row} P列无附件`, code:'NO_ATTACHMENT'});
    console.log(`[info] token=${pack.token} name=${pack.name}`);
    const fnsku=await getAttachment(row,'Q');
    const pRel=`${fid}_p.xlsx`;
    const n=await download(pack.token, pRel);
    if(!n) return res.json({ok:false, error:'下载失败', code:'DL_FAIL'});
    let fRel=null;
    if(fnsku){ fRel=`${fid}_f.xlsx`; await download(fnsku.token, fRel); }
    const xp=path.join(TMP_DIR,pRel); const fp=fRel?path.join(TMP_DIR,fRel):null;
    const {items,error:pe}=await parseXlsx(xp,fp);
    try{fs.unlinkSync(xp)}catch(_){} if(fp) try{fs.unlinkSync(fp)}catch(_){}
    res.json({ok:true, fid, row, mode:CLOUD?'cloud':'local', packingName:pack.name, items, itemCount:items.length, parseError:pe});
  }catch(e){ console.error('[err]',e.message); res.status(500).json({ok:false, error:e.message, code:e.code}); }
});

app.listen(PORT, ()=>{
  console.log(`\n  🚀 发票后端启动 mode=${CLOUD?'cloud(HTTP直连)':'local(lark-cli)'} port=${PORT}`);
  setTimeout(()=>{ try{ buildCache(); }catch(_){} }, 200);
});
