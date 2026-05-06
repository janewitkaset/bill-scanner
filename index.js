const express = require("express");
const axios = require("axios");
const crypto = require("crypto");
const { execFileSync } = require("child_process");
const fs = require("fs");
const path = require("path");
require("dotenv").config();

const app = express();
app.use(express.json({ verify: (req, _res, buf) => { req.rawBody = buf; } }));
app.use(express.urlencoded({ extended: true }));

const DATA_DIR    = path.join(__dirname, "data");
const BILLS_FILE  = path.join(DATA_DIR, "bills.xlsx");
const DB_FILE     = path.join(DATA_DIR, "bills.json");
const SAVE_SCRIPT = path.join(__dirname, "save_bill.py");
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

const BUYERS = {
  florish: { label:"บริษัทฟลอริช", sheet:"florish", keywords:["ฟลอริช","florish","florrish"], taxIds:["0105564006991"] },
  janewit: { label:"นายเจนวิทย์",  sheet:"janewit", keywords:["เจนวิทย์","janewit","เจนวิท"],  taxIds:["1100800169546"] },
};

// ─── JSON DB ─────────────────────────────────
function loadDB() {
  if (!fs.existsSync(DB_FILE)) return { bills: [] };
  try { return JSON.parse(fs.readFileSync(DB_FILE,"utf8")); } catch { return { bills: [] }; }
}
function saveDB(db) { fs.writeFileSync(DB_FILE, JSON.stringify(db,null,2),"utf8"); }
function addBillToDB(bill) {
  const db = loadDB();
  bill.id = Date.now().toString();
  bill.created_at = new Date().toISOString();
  db.bills.unshift(bill);
  saveDB(db);
  return bill.id;
}

// ─── MATCH BUYER ─────────────────────────────
function matchBuyer(customerName, taxId) {
  const text = (customerName||"").toLowerCase();
  const tid  = (taxId||"").replace(/\D/g,"");
  for (const [key,cfg] of Object.entries(BUYERS)) {
    if (tid && cfg.taxIds.some(t => tid.includes(t)||t.includes(tid))) return key;
    if (cfg.keywords.some(kw => text.includes(kw.toLowerCase()))) return key;
  }
  return null;
}

// ─── FIX BUDDHIST YEAR ───────────────────────
function fixYear(dateStr) {
  if (!dateStr) return dateStr;
  const m = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    let y = parseInt(m[1]);
    if (y > 2400) y -= 543;
    return `${y}-${m[2]}-${m[3]}`;
  }
  return dateStr;
}

// ─── EXCEL ───────────────────────────────────
function saveBillToExcel(data, sheetName) {
  const payload = JSON.stringify({ sheet: sheetName, data });
  try {
    const r = execFileSync("python3",[SAVE_SCRIPT,BILLS_FILE],{ input:payload, encoding:"utf8" }).trim();
    return r === "ok";
  } catch(e) { console.error("Excel:",e.message); return false; }
}

// ─── LINE ────────────────────────────────────
function verifySignature(req) {
  const sig = req.headers["x-line-signature"];
  const hash = crypto.createHmac("sha256",process.env.LINE_CHANNEL_SECRET).update(req.rawBody).digest("base64");
  return hash === sig;
}
const lh = () => ({ Authorization:`Bearer ${process.env.LINE_CHANNEL_ACCESS_TOKEN}`, "Content-Type":"application/json" });
async function replyLine(tok,text) { await axios.post("https://api.line.me/v2/bot/message/reply",{replyToken:tok,messages:[{type:"text",text}]},{headers:lh()}); }
async function pushLine(uid,text)  { await axios.post("https://api.line.me/v2/bot/message/push", {to:uid,messages:[{type:"text",text}]},{headers:lh()}); }

// ─── CLAUDE ──────────────────────────────────
async function downloadImage(id) {
  const r = await axios.get(`https://api-data.line.me/v2/bot/message/${id}/content`,{headers:{Authorization:`Bearer ${process.env.LINE_CHANNEL_ACCESS_TOKEN}`},responseType:"arraybuffer"});
  return Buffer.from(r.data).toString("base64");
}
async function extractBillData(b64) {
  const r = await axios.post("https://api.anthropic.com/v1/messages",{
    model:"claude-sonnet-4-5", max_tokens:1000,
    messages:[{role:"user",content:[
      {type:"image",source:{type:"base64",media_type:"image/jpeg",data:b64}},
      {type:"text",text:`อ่านใบกำกับภาษี/ใบเสร็จนี้ ตอบเป็น JSON เท่านั้น ไม่มีข้อความอื่น:
{
  "vendor_name": "ชื่อผู้ขาย",
  "customer_name": "ชื่อลูกค้า/ผู้ซื้อ",
  "customer_tax_id": "เลขผู้เสียภาษีลูกค้า 13 หลัก ตัวเลขเท่านั้น",
  "bill_number": "เลขที่บิล",
  "amount_before_vat": "ยอดก่อน VAT ตัวเลขเท่านั้น",
  "vat_amount": "ยอด VAT ตัวเลขเท่านั้น",
  "total_amount": "ยอดรวมสุทธิ ตัวเลขเท่านั้น",
  "bill_date": "วันที่ YYYY-MM-DD (พ.ศ.ให้แปลงเป็น ค.ศ. เช่น 2569=2026)",
  "confidence": "high/medium/low"
}
ถ้าไม่พบข้อมูลใด ให้ใส่ null`}
    ]}]
  },{headers:{"x-api-key":process.env.ANTHROPIC_API_KEY,"anthropic-version":"2023-06-01","content-type":"application/json"}});
  const text = r.data.content[0].text.trim().replace(/```json|```/g,"").trim();
  return JSON.parse(text);
}

function processBill(data) {
  data.bill_date = fixYear(data.bill_date);
  const buyerKey   = matchBuyer(data.customer_name, data.customer_tax_id);
  const cfg        = buyerKey ? BUYERS[buyerKey] : null;
  const sheetName  = cfg ? cfg.sheet : "unmatched";
  const buyerLabel = cfg ? cfg.label : "ไม่ระบุ";
  const record = {
    vendor_name:data.vendor_name||"", customer_name:data.customer_name||"",
    customer_tax_id:data.customer_tax_id||"", bill_number:data.bill_number||"",
    bill_date:data.bill_date||"", amount_before_vat:data.amount_before_vat||"",
    vat_amount:data.vat_amount||"", total_amount:data.total_amount||"",
    confidence:data.confidence||"", buyer:buyerKey||"unmatched", buyer_label:buyerLabel,
    timestamp:new Date().toLocaleString("th-TH",{timeZone:"Asia/Bangkok"}),
  };
  const id = addBillToDB(record);
  saveBillToExcel({...record,id}, sheetName);
  return {...record, id, buyerKey, sheetName, saved:true};
}

function formatReply(result) {
  const fmt = v => v ? parseFloat(v).toLocaleString("th-TH",{minimumFractionDigits:2}) : "-";
  const icon = result.buyerKey==="florish"?"🏢":result.buyerKey==="janewit"?"👤":"📂";
  return `🧾 อ่านบิลสำเร็จ
${icon} บันทึกแล้ว: ${result.buyer_label}${!result.buyerKey?"\n⚠️ ไม่พบชื่อผู้ซื้อ — บันทึกใน unmatched":""}

🏷️  Vendor:     ${result.vendor_name||"-"}
👤  ลูกค้า:     ${result.customer_name||"-"}
🔢  เลขที่บิล:  ${result.bill_number||"-"}
📅  วันที่:     ${result.bill_date||"-"}

💰  ก่อน VAT:   ${fmt(result.amount_before_vat)} ฿
💸  VAT:        ${fmt(result.vat_amount)} ฿
💵  รวมสุทธิ:   ${fmt(result.total_amount)} ฿

📊  ความแม่นยำ: ${result.confidence==="high"?"สูง ✅":result.confidence==="medium"?"ปานกลาง ⚠️":"ต่ำ ❌"}`;
}

// ─── WEBHOOK ─────────────────────────────────
app.post("/webhook", async (req,res) => {
  if (!verifySignature(req)) return res.status(401).send("Unauthorized");
  res.sendStatus(200);
  for (const event of req.body.events||[]) {
    if (event.type!=="message") continue;
    const {replyToken} = event;
    const userId = event.source.userId;
    if (event.message.type==="image") {
      try {
        await replyLine(replyToken,"⏳ กำลังอ่านบิล รอสักครู่...");
        const b64    = await downloadImage(event.message.id);
        const data   = await extractBillData(b64);
        const result = processBill(data);
        await pushLine(userId, formatReply(result));
      } catch(e) {
        console.error(e.message);
        await pushLine(userId,"❌ เกิดข้อผิดพลาด: "+e.message);
      }
    } else if (event.message.type==="text") {
      const txt = (event.message.text||"").trim().toLowerCase();
      if (txt==="help"||txt==="ช่วยเหลือ") {
        await replyLine(replyToken,"📖 วิธีใช้งาน\n\n📸 ส่งรูปบิล → บันทึกอัตโนมัติ\n\n🗂️ แยก Sheet:\n🏢 ฟลอริช (TAX: 0105564006991)\n👤 เจนวิทย์ (TAX: 1100800169546)\n📂 ไม่ระบุ\n\n🌐 จัดการข้อมูล:\nhttps://bill-scanner-production.up.railway.app/admin");
      } else {
        await replyLine(replyToken,"📸 กรุณาส่งรูปถ่ายบิล\nพิมพ์ 'help' เพื่อดูวิธีใช้");
      }
    }
  }
});

// ─── REST API ─────────────────────────────────
app.get("/api/bills", (req,res) => {
  const db = loadDB();
  const buyer = req.query.buyer||"all";
  const bills = buyer==="all" ? db.bills : db.bills.filter(b=>b.buyer===buyer);
  res.json({total:bills.length,bills});
});
app.put("/api/bills/:id", (req,res) => {
  const db = loadDB();
  const idx = db.bills.findIndex(b=>b.id===req.params.id);
  if (idx===-1) return res.status(404).json({error:"ไม่พบบิล"});
  db.bills[idx] = {...db.bills[idx],...req.body,id:req.params.id};
  saveDB(db); res.json({ok:true,bill:db.bills[idx]});
});
app.delete("/api/bills/:id", (req,res) => {
  const db = loadDB();
  const before = db.bills.length;
  db.bills = db.bills.filter(b=>b.id!==req.params.id);
  if (db.bills.length===before) return res.status(404).json({error:"ไม่พบบิล"});
  saveDB(db); res.json({ok:true});
});
app.post("/api/bills", (req,res) => {
  const data = req.body;
  data.bill_date = fixYear(data.bill_date);
  if (!data.buyer) {
    const k = matchBuyer(data.customer_name,data.customer_tax_id);
    data.buyer = k||"unmatched";
    data.buyer_label = k ? BUYERS[k].label : "ไม่ระบุ";
  } else {
    data.buyer_label = BUYERS[data.buyer]?.label||"ไม่ระบุ";
  }
  data.timestamp = new Date().toLocaleString("th-TH",{timeZone:"Asia/Bangkok"});
  const id = addBillToDB(data);
  saveBillToExcel({...data,id}, data.buyer);
  res.json({ok:true,id});
});

// ─── ADMIN UI ─────────────────────────────────
app.get("/admin", (_req,res) => { res.send(`<!DOCTYPE html>
<html lang="th"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Bill Scanner Admin</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:sans-serif;background:#0f172a;color:#e2e8f0;min-height:100vh}
.hdr{background:#1e293b;padding:14px 20px;border-bottom:1px solid #334155;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px}
.logo{font-size:18px;font-weight:700;color:#38bdf8}
.tabs{display:flex;background:#1e293b;border-bottom:1px solid #334155;overflow-x:auto}
.tab{padding:11px 18px;cursor:pointer;font-size:14px;color:#94a3b8;border:none;background:none;border-bottom:3px solid transparent;white-space:nowrap}
.tab.active{color:#38bdf8;border-bottom-color:#38bdf8}
.wrap{max-width:1300px;margin:0 auto;padding:16px}
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin-bottom:16px}
.stat{background:#1e293b;border:1px solid #334155;border-radius:10px;padding:14px}
.sl{font-size:11px;color:#64748b;text-transform:uppercase;margin-bottom:5px}
.sv{font-size:22px;font-weight:700;font-family:monospace}
.blue{color:#38bdf8}.green{color:#34d399}.purple{color:#a78bfa}
.bar{display:flex;gap:10px;margin-bottom:14px;flex-wrap:wrap}
.srch{flex:1;min-width:180px;padding:8px 12px;background:#1e293b;border:1px solid #334155;border-radius:7px;color:#e2e8f0;font-size:14px;outline:none}
.srch:focus{border-color:#38bdf8}
.btn{padding:8px 14px;border-radius:7px;font-size:13px;cursor:pointer;border:none;font-weight:500}
.bp{background:#38bdf8;color:#0f172a}.bp:hover{background:#7dd3fc}
.bd{background:#ef4444;color:#fff}.bd:hover{background:#dc2626}
.bg{background:transparent;color:#94a3b8;border:1px solid #334155}.bg:hover{border-color:#38bdf8;color:#38bdf8}
.bs{padding:4px 9px;font-size:12px}
table{width:100%;border-collapse:collapse;background:#1e293b;border-radius:10px;overflow:hidden}
th{padding:10px 12px;font-size:11px;color:#64748b;text-transform:uppercase;background:#0f172a;text-align:left;white-space:nowrap}
td{padding:11px 12px;font-size:13px;border-bottom:1px solid #0f172a}
tr:last-child td{border-bottom:none}
tr:hover td{background:#1a2744}
.bdg{display:inline-flex;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:600}
.b-florish{background:rgba(52,211,153,.15);color:#34d399}
.b-janewit{background:rgba(167,139,250,.15);color:#a78bfa}
.b-unmatched{background:rgba(251,146,60,.15);color:#fb923c}
.b-high{background:rgba(52,211,153,.15);color:#34d399}
.b-medium{background:rgba(251,191,36,.15);color:#fbbf24}
.b-low{background:rgba(248,113,113,.15);color:#f87171}
.mono{font-family:monospace;font-size:12px}
.amt{font-family:monospace;font-weight:600;color:#38bdf8}
.mb{display:none;position:fixed;inset:0;background:rgba(0,0,0,.75);z-index:100;align-items:center;justify-content:center;padding:16px}
.mb.show{display:flex}
.md{background:#1e293b;border:1px solid #334155;border-radius:12px;padding:22px;width:100%;max-width:500px;max-height:90vh;overflow-y:auto}
.md h3{font-size:15px;font-weight:700;margin-bottom:14px}
.fr{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:10px}
.fg{margin-bottom:10px}
.fg label{font-size:12px;color:#94a3b8;display:block;margin-bottom:4px}
.fg input,.fg select{width:100%;padding:8px 11px;background:#0f172a;border:1px solid #334155;border-radius:6px;color:#e2e8f0;font-size:13px;outline:none}
.fg input:focus,.fg select:focus{border-color:#38bdf8}
.ma{display:flex;gap:8px;justify-content:flex-end;margin-top:14px}
.empty{padding:40px;text-align:center;color:#64748b}
</style></head><body>
<div class="hdr">
  <div class="logo">🧾 Bill Scanner Admin</div>
  <div style="display:flex;gap:8px">
    <button class="btn bg" onclick="location='/download'">⬇ Excel</button>
    <button class="btn bp" onclick="openAdd()">+ เพิ่มบิล</button>
  </div>
</div>
<div class="tabs">
  <button class="tab active" onclick="setTab('all',this)">📋 ทั้งหมด <span id="c-all"></span></button>
  <button class="tab" onclick="setTab('florish',this)">🏢 ฟลอริช <span id="c-florish"></span></button>
  <button class="tab" onclick="setTab('janewit',this)">👤 เจนวิทย์ <span id="c-janewit"></span></button>
  <button class="tab" onclick="setTab('unmatched',this)">📂 ไม่ระบุ <span id="c-unmatched"></span></button>
</div>
<div class="wrap">
  <div class="stats">
    <div class="stat"><div class="sl">บิล</div><div class="sv blue" id="s-n">0</div></div>
    <div class="stat"><div class="sl">ก่อน VAT</div><div class="sv" id="s-b">฿0</div></div>
    <div class="stat"><div class="sl">VAT</div><div class="sv purple" id="s-v">฿0</div></div>
    <div class="stat"><div class="sl">รวมสุทธิ</div><div class="sv green" id="s-t">฿0</div></div>
  </div>
  <div class="bar">
    <input class="srch" id="q" placeholder="🔍 ค้นหา..." oninput="render()">
    <button class="btn bg" onclick="load()">🔄</button>
  </div>
  <div style="overflow-x:auto">
    <table><thead><tr>
      <th>วันที่</th><th>Vendor</th><th>ลูกค้า</th><th>เลขบิล</th>
      <th>ก่อน VAT</th><th>VAT</th><th>รวม</th><th>ผู้ซื้อ</th><th>แม่นยำ</th><th>จัดการ</th>
    </tr></thead><tbody id="tb"></tbody></table>
    <div class="empty" id="emp" style="display:none">ไม่มีข้อมูล</div>
  </div>
</div>
<div class="mb" id="modal">
  <div class="md">
    <h3 id="mt">เพิ่มบิล</h3>
    <input type="hidden" id="eid">
    <div class="fr">
      <div class="fg"><label>Vendor</label><input id="fv" placeholder="บริษัท..."></div>
      <div class="fg"><label>ลูกค้า</label><input id="fc" placeholder="บริษัท/นาย..."></div>
    </div>
    <div class="fr">
      <div class="fg"><label>เลขผู้เสียภาษี</label><input id="ft" placeholder="0105564006991"></div>
      <div class="fg"><label>เลขบิล</label><input id="fb" placeholder="INV-001"></div>
    </div>
    <div class="fr">
      <div class="fg"><label>วันที่ (YYYY-MM-DD)</label><input id="fd" placeholder="2026-05-06"></div>
      <div class="fg"><label>ผู้ซื้อ</label>
        <select id="fbu">
          <option value="">อัตโนมัติ</option>
          <option value="florish">บริษัทฟลอริช</option>
          <option value="janewit">นายเจนวิทย์</option>
          <option value="unmatched">ไม่ระบุ</option>
        </select>
      </div>
    </div>
    <div class="fr">
      <div class="fg"><label>ก่อน VAT</label><input id="fa" type="number" placeholder="0.00"></div>
      <div class="fg"><label>VAT</label><input id="fva" type="number" placeholder="0.00"></div>
    </div>
    <div class="fg"><label>รวมสุทธิ</label><input id="fto" type="number" placeholder="0.00"></div>
    <div class="ma">
      <button class="btn bg" onclick="closeM()">ยกเลิก</button>
      <button class="btn bp" onclick="save()">บันทึก</button>
    </div>
  </div>
</div>
<script>
let all=[],tab='all';
const fmt=v=>v?parseFloat(v).toLocaleString('th-TH',{minimumFractionDigits:2}):'-';
const fmtB=v=>v?'฿'+parseFloat(v).toLocaleString('th-TH',{maximumFractionDigits:0}):'฿0';

async function load(){
  const r=await fetch('/api/bills?buyer=all');
  const j=await r.json();
  all=j.bills||[];
  document.getElementById('c-all').textContent=all.length;
  ['florish','janewit','unmatched'].forEach(k=>{
    document.getElementById('c-'+k).textContent=all.filter(b=>b.buyer===k).length;
  });
  render();
}

function setTab(t,el){
  tab=t;
  document.querySelectorAll('.tab').forEach(x=>x.classList.remove('active'));
  el.classList.add('active');
  render();
}

function getFilt(){
  const q=document.getElementById('q').value.toLowerCase();
  return all
    .filter(b=>tab==='all'||b.buyer===tab)
    .filter(b=>!q||[b.vendor_name,b.customer_name,b.bill_number].some(v=>(v||'').toLowerCase().includes(q)));
}

function render(){
  const data=getFilt();
  const tb=document.getElementById('tb');
  const emp=document.getElementById('emp');
  const sn=data.reduce((s,b)=>s+parseFloat(b.amount_before_vat||0),0);
  const sv=data.reduce((s,b)=>s+parseFloat(b.vat_amount||0),0);
  const st=data.reduce((s,b)=>s+parseFloat(b.total_amount||0),0);
  document.getElementById('s-n').textContent=data.length;
  document.getElementById('s-b').textContent=fmtB(sn);
  document.getElementById('s-v').textContent=fmtB(sv);
  document.getElementById('s-t').textContent=fmtB(st);
  if(!data.length){tb.innerHTML='';emp.style.display='block';return;}
  emp.style.display='none';
  tb.innerHTML=data.map(b=>\`<tr>
    <td class="mono">\${b.bill_date||'-'}</td>
    <td><strong>\${b.vendor_name||'-'}</strong></td>
    <td style="color:#94a3b8;font-size:12px">\${b.customer_name||'-'}</td>
    <td class="mono">\${b.bill_number||'-'}</td>
    <td class="amt">฿\${fmt(b.amount_before_vat)}</td>
    <td style="color:#fbbf24;font-family:monospace">฿\${fmt(b.vat_amount)}</td>
    <td class="amt">฿\${fmt(b.total_amount)}</td>
    <td><span class="bdg b-\${b.buyer||'unmatched'}">\${b.buyer_label||'ไม่ระบุ'}</span></td>
    <td><span class="bdg b-\${b.confidence||'low'}">\${b.confidence==='high'?'✅':b.confidence==='medium'?'⚠️':'❌'}</span></td>
    <td style="display:flex;gap:5px">
      <button class="btn bg bs" onclick='openEdit(\${JSON.stringify(b)})'>✏️</button>
      <button class="btn bd bs" onclick="del('\${b.id}')">🗑</button>
    </td>
  </tr>\`).join('');
}

function openAdd(){
  document.getElementById('mt').textContent='เพิ่มบิล';
  document.getElementById('eid').value='';
  ['fv','fc','ft','fb','fd','fa','fva','fto'].forEach(id=>document.getElementById(id).value='');
  document.getElementById('fbu').value='';
  document.getElementById('modal').classList.add('show');
}
function openEdit(b){
  document.getElementById('mt').textContent='แก้ไขบิล';
  document.getElementById('eid').value=b.id;
  document.getElementById('fv').value=b.vendor_name||'';
  document.getElementById('fc').value=b.customer_name||'';
  document.getElementById('ft').value=b.customer_tax_id||'';
  document.getElementById('fb').value=b.bill_number||'';
  document.getElementById('fd').value=b.bill_date||'';
  document.getElementById('fbu').value=b.buyer||'';
  document.getElementById('fa').value=b.amount_before_vat||'';
  document.getElementById('fva').value=b.vat_amount||'';
  document.getElementById('fto').value=b.total_amount||'';
  document.getElementById('modal').classList.add('show');
}
function closeM(){ document.getElementById('modal').classList.remove('show'); }

async function save(){
  const id=document.getElementById('eid').value;
  const body={
    vendor_name:document.getElementById('fv').value,
    customer_name:document.getElementById('fc').value,
    customer_tax_id:document.getElementById('ft').value,
    bill_number:document.getElementById('fb').value,
    bill_date:document.getElementById('fd').value,
    buyer:document.getElementById('fbu').value||undefined,
    amount_before_vat:document.getElementById('fa').value,
    vat_amount:document.getElementById('fva').value,
    total_amount:document.getElementById('fto').value,
  };
  const url=id?'/api/bills/'+id:'/api/bills';
  const method=id?'PUT':'POST';
  await fetch(url,{method,headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
  closeM(); await load();
}

async function del(id){
  if(!confirm('ลบบิลนี้?')) return;
  await fetch('/api/bills/'+id,{method:'DELETE'});
  await load();
}

load();
setInterval(load,30000);
</script></body></html>`); });

// ─── DOWNLOAD / TEST / STATUS ─────────────────
app.get("/download", (_req,res) => {
  if (!fs.existsSync(BILLS_FILE)) return res.status(404).send("ยังไม่มีข้อมูล");
  res.download(BILLS_FILE,"bills.xlsx");
});
app.get("/test-api", async (_req,res) => {
  try {
    const r = await axios.post("https://api.anthropic.com/v1/messages",
      {model:"claude-sonnet-4-5",max_tokens:10,messages:[{role:"user",content:"say hi"}]},
      {headers:{"x-api-key":process.env.ANTHROPIC_API_KEY,"anthropic-version":"2023-06-01","content-type":"application/json"}});
    res.send("✅ API ใช้ได้: "+JSON.stringify(r.data.content));
  } catch(e) { res.send("❌ Error: "+e.message+" | "+JSON.stringify(e.response?.data)); }
});
app.get("/", (_req,res) => {
  const db = loadDB();
  res.send(`Bill Scanner Bot 🤖<br>บิล: ${db.bills.length} รายการ<br><a href="/admin">🗄️ Admin</a> | <a href="/download">⬇️ Excel</a>`);
});

const PORT = process.env.PORT||3000;
app.listen(PORT, () => {
  console.log(`✅ Server on port ${PORT}`);
  console.log(`🗄️  Admin: /admin | 💾 DB: ${DB_FILE}`);
});
