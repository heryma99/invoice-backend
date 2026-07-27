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
const SHEET_ID = (process.env.SHEET_ID && process.env.SHEET_ID!=='SHEET_ID') ? process.env.SHEET_ID : 'a11447';
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
  // 镜像表附件列存的是 JSON 字符串 "[{\"fileToken\":\"...\"}]"，从中解析 token
  const range = `${SHEET_ID}!${col}${row}:${col}${row}`;
  const j = await feishu('GET', `/open-apis/sheets/v2/spreadsheets/${SPREADSHEET_TOKEN}/values/${range}`);
  const vals = j.data && j.data.valueRange && j.data.valueRange.values;
  const raw = vals && vals[0] && vals[0][0];
  if(!raw) return null;
  let arr = raw;
  if(typeof raw === 'string'){ try{ arr = JSON.parse(raw); }catch(_){ arr = null; } }
  if(Array.isArray(arr)){ for(const x of arr){ if(x && x.fileToken) return {token:x.fileToken, name:''}; } }
  if(arr && typeof arr==='object' && arr.fileToken) return {token:arr.fileToken, name:''};
  const m = String(raw).match(/fileToken["']?\s*[:"]\s*([A-Za-z0-9]+)/i);
  return m?{token:m[1], name:''}:null;
}
async function cloudDownload(token, outPath){
  // 注意：用 drive /files/ 端点（bot 身份对 /medias/ 端点会 403）
  const t = await cloudToken();
  const url = DOMAIN+`/open-apis/drive/v1/files/${token}/download`;
  console.log('[dl] GET', url, 'tok=', t.slice(0,8), 'proxy=', process.env.HTTPS_PROXY||process.env.https_proxy||'none');
  const r = await fetch(url, {headers:{Authorization:`Bearer ${t}`}});
  console.log('[dl] status=', r.status, 'redirected=', r.redirected, 'type=', r.headers.get('content-type'));
  if(!r.ok) throw new Error(r.status===403 ? '附件尚未同步到镜像库(旧token), 请稍后重试或等待同步完成' : 'download HTTP '+r.status);
  const buf = Buffer.from(await r.arrayBuffer());
  fs.writeFileSync(outPath, buf);
  return buf.length;
}
async function cloudCsvColumn(col, maxRow){
  // 批量读某列用于搜索缓存（range 必须带 SHEET_ID 前缀）
  const range = `${SHEET_ID}!${col}1:${col}${maxRow}`;
  const j = await feishu('GET', `/open-apis/sheets/v2/spreadsheets/${SPREADSHEET_TOKEN}/values/${range}`);
  const rows = (j.data && j.data.valueRange && j.data.valueRange.values) || [];
  return rows.map(r=>r[0]?String(r[0]):'');
}

/* ===================== 本机：lark-cli ===================== */
const BASH = 'C:/Program Files/Git/bin/bash.exe';
const SHELL = { shell: BASH, encoding: 'utf-8' };
const IDENT = process.env.LARK_AS || 'user'; // 本机模式跨租户读源表必须user身份(bot无sheets scope且跨租户)
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
let rowCache=null, lastCacheErr=null;
async function buildCache(){
  console.log('[cache] 预加载货件号索引...');
  try{
    const vals = CLOUD ? await cloudCsvColumn('B',960) : (()=>{ const r=larkRaw(`lark-cli sheets +csv-get --url "https://plbrands.feishu.cn/sheets/${SPREADSHEET_TOKEN}" --sheet-id ${SHEET_ID} --range B1:B960`); return (r.data&&r.data.annotated_csv||'').split('\n').map(l=>{const m=l.match(/\[row=(\d+)\]\s*(.+)/); return m?m[2].replace(/^"|"$/g,'').trim():'';}); })();
    rowCache={};
    vals.forEach((v,i)=>{ const key=String(v).toUpperCase(); if(key) rowCache[key]=i+1; });
    console.log(`[cache] 索引就绪: ${Object.keys(rowCache).length} 个货件号`);
    lastCacheErr=null;
  }catch(e){ console.log('[cache] 预加载失败:', e.message.substring(0,100)); rowCache={}; lastCacheErr=e.message; }
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

/* --- 格式自适应读取: PK头=xlsx, 否则CSV(UTF-8/GBK自动) --- */
function decodeBuf(buf){
  let txt=buf.toString('utf8');
  if(txt.includes('\uFFFD')){ try{ txt=new TextDecoder('gbk').decode(buf); console.log('[read] GBK编码'); }catch(_){} }
  return txt.replace(/^\uFEFF/,'');
}
function csvParse(txt){
  const rows=[]; let row=[], cur='', inQ=false;
  for(let i=0;i<txt.length;i++){ const ch=txt[i];
    if(inQ){ if(ch==='"'){ if(txt[i+1]==='"'){cur+='"';i++;} else inQ=false; } else cur+=ch; }
    else{ if(ch==='"') inQ=true;
      else if(ch===',') { row.push(cur); cur=''; }
      else if(ch==='\n'){ row.push(cur); rows.push(row); row=[]; cur=''; }
      else if(ch!=='\r') cur+=ch; } }
  if(cur!==''||row.length){ row.push(cur); rows.push(row); }
  return rows;
}
async function readRows(fp){
  const buf=fs.readFileSync(fp);
  if(buf.length>1&&buf[0]===0x50&&buf[1]===0x4B){ // xlsx
    const wb=new ExcelJS.Workbook(); await wb.xlsx.load(buf);
    const ws=wb.worksheets[0]; if(!ws) return [];
    const rows=[];
    for(let r=1;r<=ws.rowCount;r++){ const row=ws.getRow(r); const vals=[];
      for(let c=1;c<=Math.max(row.cellCount,30);c++){ let v=row.getCell(c).value;
        if(v&&typeof v==='object') v = v.richText? v.richText.map(t=>t.text).join('') : (v.text!==undefined? v.text : (v.result!==undefined? v.result : ''));
        vals.push(v===null||v===undefined?'':v); }
      rows.push(vals); }
    return rows;
  }
  console.log('[read] 非zip, 按CSV解析');
  return csvParse(decodeBuf(buf));
}
/* --- 亚马逊官方"箱内容清单"CSV (工作流导出, 分区+箱号列表) --- */
function isAmzWorkflow(rows){
  return rows.some(r=>{const a=String((r&&r[0])||''); return a.includes('货件编号')||a.includes('工作流程名称');})
      && rows.some(r=>{const j=(r||[]).join(','); return String((r&&r[0])||'').trim()==='SKU'&&j.includes('FNSKU');});
}
function parseAmzWorkflow(rows){
  const items=[]; let mode=null, ci=null, boxCols=null;
  const boxDims={}; // boxLabel -> {weight,len,wid,hgt} (矩阵分区底部箱规)
  const dimKey=s=> s.includes('重量')?'weight': s.includes('长度')?'len': s.includes('宽度')?'wid': s.includes('高度')?'hgt': null;
  const mkItem=o=>({boxNo:'',boxLabel:'',sku:'',fnsku:'',nameCn:'',nameEn:'',qty:0,declare:'',cost:'',material:'',hs:'',brand:'',weight:'',len:'',wid:'',hgt:'',elec:'N',magnet:'N',saleUrl:'',asin:'',...o});
  for(const raw of rows){
    const vals=(raw||[]).map(v=>String(v==null?'':v).trim());
    const joined=vals.join(',');
    /* 矩阵分区底部箱规行: ",,,,,,,,包装箱重量（千克）：,13,8.75,..." */
    const dimCell=vals.findIndex(v=>v.startsWith('包装箱')&&dimKey(v));
    if(dimCell>=0 && boxCols){
      const k=dimKey(vals[dimCell]);
      for(const bc of boxCols){ const v=vals[bc.idx]; if(v) (boxDims[bc.boxLabel]=boxDims[bc.boxLabel]||{})[k]=v; }
      continue;
    }
    /* 表头行 */
    if(vals[0]==='SKU' && joined.includes('FNSKU')){
      const f=names=>vals.findIndex(h=>names.some(n=>h===n||h.includes(n)));
      if(joined.includes('每箱件数')&&joined.includes('箱号')){ /* 列表分区(原厂包装) */
        mode='list';
        ci={sku:0, name:f(['商品名称']), asin:f(['ASIN']), fnsku:f(['FNSKU']),
            weight:f(['包装箱重量','箱重']), len:f(['箱子长度']), wid:f(['箱子宽度']), hgt:f(['箱子高度']),
            qty:f(['每箱件数']), ctns:f(['箱子总数']), box:f(['箱号'])};
      } else { /* 矩阵分区(单件/混装): 箱号FBA...U000035做列头 */
        boxCols=[];
        vals.forEach((h,i)=>{
          const m=String(h).match(/U0*(\d+)$/);
          if(m){ const n=parseInt(m[1]); boxCols.push({idx:i, boxNo:h, boxLabel:'B'+n}); }
        });
        if(boxCols.length){ mode='matrix'; ci={sku:0, name:f(['商品名称']), asin:f(['ASIN']), fnsku:f(['FNSKU'])}; }
      }
      continue;
    }
    if(!mode) continue;
    const sku=vals[0]; if(!sku||sku==='SKU') continue;
    const nameEn=(vals[ci.name]||'').replace(/''/g,"'");
    if(mode==='list'){
      const qty=parseInt(vals[ci.qty])||0; if(!qty) continue;
      let labels=String(vals[ci.box]||'').split(/[,，\s]+/).filter(Boolean);
      if(labels.length===0){ const n=parseInt(vals[ci.ctns])||1; for(let b=1;b<=n;b++) labels.push('B'+b); }
      for(const lb of labels){
        const m=lb.match(/U0*(\d+)$/);
        const boxLabel = m ? ('B'+parseInt(m[1])) : lb;
        const boxNo = m ? lb : '';
        items.push(mkItem({boxNo, boxLabel, sku, fnsku:vals[ci.fnsku]||'', nameEn, qty,
          weight:vals[ci.weight]||'', len:vals[ci.len]||'', wid:vals[ci.wid]||'', hgt:vals[ci.hgt]||'', asin:vals[ci.asin]||''}));
      }
    } else { /* matrix: 每个箱列一个数量 */
      for(const bc of boxCols){
        const q=parseInt(vals[bc.idx])||0; if(!q) continue;
        items.push(mkItem({boxNo:bc.boxNo, boxLabel:bc.boxLabel, sku, fnsku:vals[ci.fnsku]||'', nameEn, qty:q, asin:vals[ci.asin]||''}));
      }
    }
  }
  for(const it of items){ const d=boxDims[it.boxLabel]; if(d){ it.weight=it.weight||d.weight||''; it.len=it.len||d.len||''; it.wid=it.wid||d.wid||''; it.hgt=it.hgt||d.hgt||''; } }
  return items;
}

/* --- 亚马逊/物流商「每箱多款SKU」矩阵 (列头: 第1箱/第2箱..., 底部行: 箱号/箱子名称/尺寸) --- */
function isMskuPerBoxMatrix(rows){
  return rows.some(r=>(r||[]).some(v=>/每箱多款/.test(String(v==null?'':v))))
      || rows.some(r=>{
          const joined=(r||[]).map(v=>String(v==null?'':v)).join(',');
          return /第\s*\d+\s*箱/.test(joined) && (joined.includes('MSKU')||joined.includes('FNSKU'));
        });
}
function parseMskuPerBoxMatrix(rows){
  const items=[];
  const mkItem=o=>({boxNo:'',boxLabel:'',sku:'',fnsku:'',nameCn:'',nameEn:'',qty:0,declare:'',cost:'',material:'',hs:'',brand:'JW PEI',weight:'',len:'',wid:'',hgt:'',elec:'N',magnet:'N',saleUrl:'',asin:'',...o});
  // 1) find header row & box columns
  let hi=-1, boxCols=[]; // {idx, boxIndex}
  for(let i=0;i<rows.length;i++){
    const found=[];
    const r=rows[i]||[];
    for(let j=0;j<r.length;j++){
      const m=String(r[j]==null?'':r[j]).trim().match(/^第\s*(\d+)\s*箱$/);
      if(m) found.push({idx:j, boxIndex:parseInt(m[1])});
    }
    if(found.length){ hi=i; boxCols=found; break; }
  }
  if(hi<0 || boxCols.length===0) return [];
  // 2) scan bottom rows for box info
  const boxInfo={}; // boxIndex -> {boxNo, boxLabel, weight, len, wid, hgt}
  for(let i=hi+1;i<rows.length;i++){
    const r=rows[i]||[];
    if(!r.some(v=>v!==null&&v!==undefined&&String(v).trim()!=='')) continue;
    const left=r.slice(0,5).map(v=>String(v==null?'':v).trim().toLowerCase()).join(' ');
    const dimKey = left.includes('weight of box') ? 'weight'
                 : left.includes('box length') ? 'len'
                 : left.includes('box width') ? 'wid'
                 : left.includes('box height') ? 'hgt' : null;
    if(dimKey){
      for(const bc of boxCols){ const v=r[bc.idx]; if(v!==null&&v!==undefined&&String(v).trim()!=='') (boxInfo[bc.boxIndex]=boxInfo[bc.boxIndex]||{})[dimKey]=String(v).trim(); }
      continue;
    }
    if(left.includes('箱号') && !left.includes('箱子名称')){
      for(const bc of boxCols){ const v=r[bc.idx]; if(v!==null&&v!==undefined) (boxInfo[bc.boxIndex]=boxInfo[bc.boxIndex]||{}).boxNo=String(v).trim(); }
      continue;
    }
    if(left.includes('箱子名称')){
      for(const bc of boxCols){ const v=r[bc.idx]; if(v!==null&&v!==undefined) (boxInfo[bc.boxIndex]=boxInfo[bc.boxIndex]||{}).boxLabel=String(v).trim(); }
    }
  }
  // 3) find column indexes
  const H=(rows[hi]||[]).map(v=>String(v==null?'':v).trim().toLowerCase());
  const find=cands=>{for(let i=0;i<H.length;i++){for(const k of cands){if(H[i]===k||H[i].includes(k)) return i;}} return -1;};
  const ci={sku:find(['msku','sku']), fnsku:find(['fnsku']), name:find(['品名','名称','商品名称'])};
  // 4) parse SKU rows
  for(let i=hi+1;i<rows.length;i++){
    const r=rows[i]||[];
    if(!r.some(v=>v!==null&&v!==undefined&&String(v).trim()!=='')) continue;
    const seq=String(r[0]==null?'':r[0]).trim();
    if(!/^\d+$/.test(seq)) continue;
    const sku=ci.sku>=0?String(r[ci.sku]==null?'':r[ci.sku]).trim():'';
    const fnsku=ci.fnsku>=0?String(r[ci.fnsku]==null?'':r[ci.fnsku]).trim():'';
    if(!sku && !fnsku) continue;
    const name=ci.name>=0?String(r[ci.name]==null?'':r[ci.name]).trim():'';
    for(const bc of boxCols){
      const v=r[bc.idx];
      const q=(v===null||v===undefined||String(v).trim()==='')?0:parseInt(String(v).trim());
      if(!q || isNaN(q)) continue;
      const info=boxInfo[bc.boxIndex]||{};
      const boxNo=info.boxNo || info.boxLabel || ('B'+bc.boxIndex);
      const boxLabel=info.boxLabel || ('B'+bc.boxIndex);
      items.push(mkItem({boxNo, boxLabel, sku, fnsku, nameEn:name, qty:q,
        weight:info.weight||'', len:info.len||'', wid:info.wid||'', hgt:info.hgt||''}));
    }
  }
  return items;
}

async function parseXlsx(xlsxPath, fnskuPath){
  const rows=await readRows(xlsxPath);
  if(!rows.length) return {items:[], error:'空表'};
  let fnskuMap={};
  if(fnskuPath && fs.existsSync(fnskuPath)){
    try{ const rows2=await readRows(fnskuPath);
      for(let r=1;r<rows2.length;r++){ const rr=rows2[r]||[]; const ms=String(rr[0]||'').trim(), fn=String(rr[1]||'').trim(), nm=String(rr[2]||'').trim(); if(ms&&nm){fnskuMap[ms.toUpperCase()]=nm;fnskuMap[fn.toUpperCase()]=nm;} }
    }catch(e){ console.log('[fnsku] fail', e.message.substring(0,60)); }
  }
  /* 亚马逊工作流CSV: 专用解析 */
  if(isAmzWorkflow(rows)){
    const items=parseAmzWorkflow(rows);
    for(const it of items){ if(!it.nameEn){const k=(it.fnsku||it.sku||'').toUpperCase(); if(fnskuMap[k]) it.nameEn=fnskuMap[k];} }
    console.log(`[parse] 亚马逊箱内容清单CSV items=${items.length}`);
    const meta=parseHeaderMeta(rows);
    return {items, meta, error: items.length?null:'亚马逊CSV解析0行'};
  }
  /* 「每箱多款SKU」矩阵 (列头: 第1箱/第2箱..., 底部: 箱号/箱子名称/尺寸) */
  if(isMskuPerBoxMatrix(rows)){
    const items=parseMskuPerBoxMatrix(rows);
    for(const it of items){ if(!it.nameEn){const k=(it.fnsku||it.sku||'').toUpperCase(); if(fnskuMap[k]) it.nameEn=fnskuMap[k];} }
    console.log(`[parse] 每箱多款SKU矩阵 items=${items.length}`);
    const meta=parseHeaderMeta(rows);
    return {items, meta, error: items.length?null:'每箱多款SKU矩阵解析0行'};
  }
  /* 通用表格 (xlsx或普通CSV) */
  const getCell=(r,c)=>{ const v=(rows[r-1]||[])[c-1]; return v===null||v===undefined?'':v; };
  const rowCount=rows.length;
  let headerRow=1, isAmazon=false;
  for(let r=1;r<=Math.min(rowCount,5);r++){
    let txt=(rows[r-1]||[]).map(v=>String(v==null?'':v)).join('');
    const low=txt.toLowerCase();
    if(low.includes('序号')&&low.includes('msku')&&low.includes('单箱数量')){headerRow=r; isAmazon=true; break;}
    if(low.includes('箱号')&&low.includes('sku')){headerRow=r; break;}
    if(low.includes('序号')&&low.includes('sku')){headerRow=r; break;}
  }
  const headers=[];
  for(let c=1;c<=Math.max((rows[headerRow-1]||[]).length,30);c++) headers.push(String(getCell(headerRow,c)).trim());
  const low=headers.map(h=>h.toLowerCase());
  const find=cands=>{const i=low.findIndex(h=>cands.some(k=>h===k||h.includes(k))); return i>=0?i+1:0;};
  const col={seq:find(['序号','no','number','num','row']), sku:find(['msku','sku','型号','产品型号']), fnsku:find(['fnsku']), nameCn:find(['中文品名','品名','名称','中文名称']), nameEn:find(['英文品名','英文名称','product name','nameen']), qty:find(['单箱数量','数量','qty','quantity']), ctns:find(['箱数','ctns','cartons']), boxName:find(['箱子名称','box label','箱号名称','label']), boxNo:find(['箱号','box no','carton','boxno','ctn']), weight:find(['箱子毛重','单箱毛重','箱重','重量','weight']), declare:find(['申报价','申报价值','declare','申报单价','unit price']), cost:find(['成本','采购价','cost']), material:find(['材质','material']), hs:find(['hs','海关编码','hscode']), brand:find(['品牌','brand']), len:find(['箱子长度','长','length']), wid:find(['箱子宽度','宽','width']), hgt:find(['箱子高度','高','height']), elec:find(['带电','elec']), magnet:find(['带磁','magnet']), saleUrl:find(['销售链接','sale url','saleurl'])};
  console.log(`[parse] headerRow=${headerRow} isAmazon=${isAmazon} sku=${col.sku} fnsku=${col.fnsku} qty=${col.qty} boxName=${col.boxName}`);
  const items=[];
  for(let r=headerRow+1;r<=rowCount;r++){
    const vals=[];
    for(let c=1;c<=Math.max((rows[r-1]||[]).length,30);c++){ const v=getCell(r,c); vals.push(v===''?null:v); }
    if(vals.every(v=>v===null||v===undefined||v==='')) continue;
    const get=idx=>idx>0&&vals[idx-1]!==null&&vals[idx-1]!==undefined?String(vals[idx-1]).trim():'';
    if(isAmazon){
      const sku=get(col.sku), fnsku=get(col.fnsku), qtyPB=parseInt(get(col.qty))||0, ctns=parseInt(get(col.ctns))||1, bn=get(col.boxName), fbaBoxStr=get(col.boxNo), en=get(col.nameEn);
      // 箱子名称 -> 箱标 (B1/B2...)
      let labels=[];
      if(bn){ const last=bn.includes(' - ')?bn.split(' - ').pop():bn; const m=last.match(/^([A-Za-z]?)(\d+)\s*[～~\-]\s*([A-Za-z]?\d+)$/);
        if(m){ const prefix=m[1]||'B'; const start=parseInt(m[2]), end=parseInt(m[3].replace(/^[A-Za-z]/,'')); for(let b=start;b<=end;b++) labels.push(prefix+b); }
        else if(last.match(/^[A-Za-z]?\d+$/)) labels.push(last);
        else labels=bn.split(/[；;，,、\s]+/).filter(Boolean);
      }
      if(labels.length===0) for(let b=1;b<=ctns;b++) labels.push('B'+b);
      // 箱号 -> FBA 箱 ID (FBA19...U000001)
      const parseFbaRange=str=>{
        const s=String(str||'').replace(/[；;]$/,'').trim();
        const m=s.match(/^(.+?)(\d+)～(\d+)$/);
        if(m){ const prefix=m[1], start=parseInt(m[2]), end=parseInt(m[3]), pad=m[2].length; const arr=[]; for(let i=start;i<=end;i++) arr.push(prefix+String(i).padStart(pad,'0')); return arr; }
        return s?[s]:[];
      };
      let fbaIds=parseFbaRange(fbaBoxStr);
      if(fbaIds.length===0) for(let b=1;b<=ctns;b++) fbaIds.push('');
      const count=Math.max(labels.length, fbaIds.length, ctns);
      for(let k=0;k<count;k++){
        const boxLabel=labels[k]||('B'+(k+1));
        const boxNo=fbaIds[k]||'';
        const it={boxNo, boxLabel, sku, fnsku, nameCn:get(col.nameCn), nameEn:en, qty:qtyPB, declare:get(col.declare), cost:get(col.cost), material:get(col.material), hs:get(col.hs), brand:get(col.brand), weight:get(col.weight), len:get(col.len), wid:get(col.wid), hgt:get(col.hgt), elec:(get(col.elec)||'').toUpperCase().startsWith('Y')?'Y':'N', magnet:(get(col.magnet)||'').toUpperCase().startsWith('Y')?'Y':'N', saleUrl:get(col.saleUrl)};
        if(!it.nameEn){const k2=(it.fnsku||it.sku||'').toUpperCase(); if(fnskuMap[k2]) it.nameEn=fnskuMap[k2];} items.push(it);
      }
    } else {
      const sku=get(col.sku)||get(col.model);
      const it={boxNo:get(col.boxNo),boxLabel:get(col.boxName),sku,fnsku:get(col.fnsku),nameCn:get(col.nameCn),nameEn:get(col.nameEn),qty:parseInt(get(col.qty))||1,declare:get(col.declare),cost:get(col.cost),material:get(col.material),hs:get(col.hs),brand:get(col.brand),weight:get(col.weight),len:get(col.len),wid:get(col.wid),hgt:get(col.hgt),elec:(get(col.elec)||'').toUpperCase().startsWith('Y')?'Y':'N',magnet:(get(col.magnet)||'').toUpperCase().startsWith('Y')?'Y':'N',saleUrl:get(col.saleUrl)};
      if(!it.nameEn){const k=(it.fnsku||it.sku||'').toUpperCase(); if(fnskuMap[k]) it.nameEn=fnskuMap[k];} items.push(it);
    }
  }
  const meta=parseHeaderMeta(rows);
  return {items, meta, error:null};
}

/* --- 解析装箱清单表头元数据: 货件名称(含FC代码)/配送地址/箱数等 --- */
function parseHeaderMeta(rows){
  const meta={};
  const KEYMAP={'工作流程名称':'workflowName','货件编号':'shipmentNo','货件名称':'shipmentName','配送地址':'destAddress','箱子数量':'boxes','SKU 数量':'skuCount','商品数量':'qty'};
  for(let i=0;i<Math.min(rows.length,14);i++){
    const r=(rows[i]||[]).map(v=>String(v==null?'':v).trim());
    if(r.length>=2 && KEYMAP[r[0]]) meta[KEYMAP[r[0]]]=r[1];
  }
  // FC 代码: 货件名称形如 "FBA STA (...) - RUH8" 或 "FBA15...-RUH8"
  if(meta.shipmentName){
    const m=meta.shipmentName.match(/([A-Z]{3,4}\d{1,3})\s*$/);
    if(m) meta.fcCode=m[1];
  }
  return meta;
}
/* ===================== 路由 ===================== */
app.get('/api/health', (req,res)=>res.json({ok:true, mode:CLOUD?'cloud':'local', time:new Date().toISOString(), pid:process.pid}));
app.get('/api/debug-cache', async (req,res)=>{
  const fid=req.query.fid;
  if(!rowCache) await buildCache();
  let probe={};
  try{
    const t=await cloudToken(); probe.tokenOk=true; probe.tokenHead=t.slice(0,10);
    const j=await feishu('GET',`/open-apis/sheets/v2/spreadsheets/${SPREADSHEET_TOKEN}/values/${SHEET_ID}!B1:B3`);
    probe.sample=(j.data&&j.data.valueRange&&j.data.valueRange.values)||[];
  }catch(e){ probe.err=e.message; }
  if(!rowCache) return res.json({ok:false,msg:'cache not built',fid, table:SPREADSHEET_TOKEN, sheet:SHEET_ID, lastCacheErr, probe});
  const key=String(fid).toUpperCase();
  res.json({ok:true, mode:CLOUD?'cloud':'local', cacheSize:Object.keys(rowCache).length, hasKey:!!rowCache[key], row:rowCache[key], key, fid, sampleKeys:Object.keys(rowCache).slice(0,5), table:SPREADSHEET_TOKEN, sheet:SHEET_ID, probe});
});
/* ===================== 实时搜索交接清单（倒着填用） ===================== */
// 本机：读某列（annotated_csv 解析）
function localColumn(col, maxRow){
  const r = larkRaw(`lark-cli sheets +csv-get --url "https://plbrands.feishu.cn/sheets/${SPREADSHEET_TOKEN}" --sheet-id ${SHEET_ID} --range ${col}1:${col}${maxRow}`);
  return (r.data&&r.data.annotated_csv||'').split('\n').map(l=>{const m=l.match(/\[row=(\d+)\]\s*(.+)/); return m?m[2].replace(/^"|"$/g,'').trim():'';});
}
// 读一整行 A-X，返回值的数组（云端/本机自适应）
async function readRowValues(row){
  if(CLOUD){
    const j = await feishu('GET', `/open-apis/sheets/v2/spreadsheets/${SPREADSHEET_TOKEN}/values/${SHEET_ID}!A${row}:X${row}`);
    const vals = j.data && j.data.valueRange && j.data.valueRange.values && j.data.valueRange.values[0];
    return vals || [];
  }
  const r = larkRaw(`lark-cli sheets +cells-get --url "https://plbrands.feishu.cn/sheets/${SPREADSHEET_TOKEN}" --sheet-id ${SHEET_ID} --range A${row}:X${row}`);
  const cells = r.data && r.data.ranges && r.data.ranges[0] && r.data.ranges[0].cells;
  if(!cells || !cells[0]) return [];
  return cells[0].map(c=>{
    if(c==null) return '';
    if(c.value!==undefined && c.value!==null) return c.value;
    if(c.attachmentToken) return JSON.stringify([{fileToken:c.attachmentToken}]);
    if(c.rich_text) for(const rt of c.rich_text){ if(rt.attachment_token) return JSON.stringify([{fileToken:rt.attachment_token}]); }
    return '';
  });
}
// 把一整行映射成与前端 HANDOVER_INDEX 一致的字段结构
async function readRowFull(row){
  const vals = await readRowValues(row);
  if(!vals.length) return null;
  const g = i => (vals[i]!==undefined && vals[i]!==null) ? String(vals[i]) : '';
  const pl = g(15), fk = g(16);
  return {
    row,
    internal_no: g(0), fba_shipment: g(1), created: g(2), country: g(3), air_sea: g(4),
    qty: g(5), boxes: g(6), carrier: g(10), pickup_addr: g(11), fba_label: g(14),
    packing_list: pl, fnsku_file: fk, invoice_drop: g(23),
    hasPacking: !!pl.trim(), hasFnsku: !!fk.trim()
  };
}
app.get('/api/search-handover', async (req,res)=>{
  const q=(req.query.q||'').trim().toUpperCase();
  if(q.length<2) return res.json({ok:true, results:[], mode:CLOUD?'cloud':'local'});
  try{
    const bCol = CLOUD ? await cloudCsvColumn('B',1200) : localColumn('B',1200);
    const rows=[];
    bCol.forEach((v,i)=>{ if(String(v).toUpperCase().includes(q)) rows.push(i+1); });
    const aCol = CLOUD ? await cloudCsvColumn('A',1200) : localColumn('A',1200);
    aCol.forEach((v,i)=>{ if(String(v).toUpperCase().includes(q) && !rows.includes(i+1)) rows.push(i+1); });
    rows.sort((a,b)=>a-b);
    const results=[];
    for(const row of rows.slice(0,25)){
      const r = await readRowFull(row);
      if(r) results.push(r);
    }
    res.json({ok:true, results, mode:CLOUD?'cloud':'local', table:SPREADSHEET_TOKEN, sheet:SHEET_ID, source: CLOUD?'镜像表(实时)':'源表(实时)'});
  }catch(e){ console.error('[search-handover]',e.message); res.status(500).json({ok:false, error:e.message, code:e.code}); }
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
    if(fnsku){ try{ fRel=`${fid}_f.xlsx`; const fn=await download(fnsku.token, fRel); if(!fn){ fRel=null; console.log('[fnsku] 下载失败, 跳过(仅影响品名映射)'); } }catch(e){ fRel=null; console.log('[fnsku] 下载异常, 跳过:', e.message.slice(0,60)); } }
    const xp=path.join(TMP_DIR,pRel); const fp=fRel?path.join(TMP_DIR,fRel):null;
    const {items,meta,error:pe}=await parseXlsx(xp,fp);
    try{fs.unlinkSync(xp)}catch(_){} if(fp) try{fs.unlinkSync(fp)}catch(_){}
    res.json({ok:true, fid, row, mode:CLOUD?'cloud':'local', packingName:pack.name, items, itemCount:items.length, meta:meta||{}, parseError:pe});
  }catch(e){ console.error('[err]',e.message); res.status(500).json({ok:false, error:e.message, code:e.code}); }
});

/* ===================== 全量货件列表（生成预装脚本用） ===================== */
app.get('/api/list-handovers', async (req,res)=>{
  try{
    let rows;
    if(CLOUD){
      const j = await feishu('GET', `/open-apis/sheets/v2/spreadsheets/${SPREADSHEET_TOKEN}/values/${SHEET_ID}!A1:X960`);
      rows = (j.data && j.data.valueRange && j.data.valueRange.values) || [];
    } else {
      const r = larkRaw(`lark-cli sheets +csv-get --url "https://plbrands.feishu.cn/sheets/${SPREADSHEET_TOKEN}" --sheet-id ${SHEET_ID} --range A1:X960`);
      rows = (r.data&&r.data.annotated_csv||'').split('\n').map(l=>{const m=l.match(/\[row=(\d+)\]\s*(.+)/); return m?m[2].replace(/^"|"$/g,'').split(','):[];});
    }
    const results=[];
    for(let i=1;i<rows.length;i++){
      const vals=rows[i]||[];
      if(!vals.length || !String(vals[1]||'').trim()) continue; // 跳过 B 列(货件号)空行
      const g=idx=>(vals[idx]!==undefined&&vals[idx]!==null)?String(vals[idx]).trim():'';
      const pl=g(15), fk=g(16);
      results.push({row:i+1, internal_no:g(0), fba_shipment:g(1), created:g(2), country:g(3), air_sea:g(4), qty:g(5), boxes:g(6), carrier:g(10), pickup_addr:g(11), fba_label:g(14), packing_list:pl, fnsku_file:fk, invoice_drop:g(23), hasPacking:!!pl.trim(), hasFnsku:!!fk.trim()});
    }
    res.json({ok:true, count:results.length, results, mode:CLOUD?'cloud':'local', table:SPREADSHEET_TOKEN, sheet:SHEET_ID});
  }catch(e){ console.error('[list-handovers]',e.message); res.status(500).json({ok:false, error:e.message}); }
});

if(require.main===module){
  app.listen(PORT, ()=>{
    console.log(`\n  🚀 发票后端启动 mode=${CLOUD?'cloud(HTTP直连)':'local(lark-cli)'} port=${PORT}`);
    setTimeout(()=>{ try{ buildCache(); }catch(_){} }, 200);
  });
}
module.exports={parseXlsx, readRows};
