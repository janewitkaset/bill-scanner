const express = require("express");
const axios = require("axios");
const crypto = require("crypto");
const { execFileSync } = require("child_process");
const fs = require("fs");
const path = require("path");
require("dotenv").config();

const app = express();

// ───────────────────────────────────────────────
// 📁 ไฟล์ Excel เดียว — หลาย Sheet ในไฟล์เดียวกัน
// ───────────────────────────────────────────────
const DATA_DIR   = path.join(__dirname, "data");
const BILLS_FILE = path.join(DATA_DIR, "bills.xlsx");   // ← ไฟล์เดียว
const SAVE_SCRIPT = path.join(__dirname, "save_bill.py");

const BUYERS = {
  florish: {
    label: "บริษัทฟลอริช",
    sheet: "florish",
    keywords: ["ฟลอริช", "florish", "florrish"],
  },
  janewit: {
    label: "นายเจนวิทย์",
    sheet: "janewit",
    keywords: ["เจนวิทย์", "janewit", "เจนวิท"],
  },
};

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

// ─── LINE signature ───────────────────────────
app.use(express.json({
  verify: (req, _res, buf) => { req.rawBody = buf; },
}));

function verifySignature(req) {
  const sig  = req.headers["x-line-signature"];
  const hash = crypto
    .createHmac("sha256", process.env.LINE_CHANNEL_SECRET)
    .update(req.rawBody)
    .digest("base64");
  return hash === sig;
}

// ─── Match ชื่อผู้ซื้อ ────────────────────────
function matchBuyer(customerName) {
  if (!customerName) return null;
  const lower = customerName.toLowerCase();
  for (const [key, cfg] of Object.entries(BUYERS)) {
    if (cfg.keywords.some((kw) => lower.includes(kw.toLowerCase()))) return key;
  }
  return null;
}

// ─── บันทึกลง Excel (เรียก Python) ───────────
function saveBill(data) {
  const buyerKey = matchBuyer(data.customer_name);
  const cfg       = buyerKey ? BUYERS[buyerKey] : null;
  const sheetName = cfg ? cfg.sheet : "unmatched";
  const buyerLabel = cfg ? cfg.label : "ไม่ระบุ";

  const payload = JSON.stringify({
    sheet: sheetName,
    data: {
      timestamp:        new Date().toLocaleString("th-TH", { timeZone: "Asia/Bangkok" }),
      vendor_name:      data.vendor_name      || "",
      customer_name:    data.customer_name    || "",
      bill_number:      data.bill_number      || "",
      bill_date:        data.bill_date        || "",
      amount_before_vat: data.amount_before_vat || "",
      vat_amount:       data.vat_amount       || "",
      total_amount:     data.total_amount     || "",
      confidence:       data.confidence       || "",
    },
  });

  try {
    const result = execFileSync("python3", [SAVE_SCRIPT, BILLS_FILE], {
      input: payload,
      encoding: "utf8",
    }).trim();
    return { saved: result === "ok", buyerKey, buyerLabel, sheetName };
  } catch (err) {
    console.error("Excel error:", err.message);
    return { saved: false, buyerKey, buyerLabel, sheetName };
  }
}

// ─── ดาวน์โหลดรูปจาก LINE ────────────────────
async function downloadImage(messageId) {
  const res = await axios.get(
    `https://api-data.line.me/v2/bot/message/${messageId}/content`,
    {
      headers: { Authorization: `Bearer ${process.env.LINE_CHANNEL_ACCESS_TOKEN}` },
      responseType: "arraybuffer",
    }
  );
  return Buffer.from(res.data).toString("base64");
}

// ─── Claude อ่านบิล ───────────────────────────
async function extractBillData(base64Image) {
  const res = await axios.post(
    "https://api.anthropic.com/v1/messages",
    {
      model: "claude-3-5-sonnet-20241022",
      max_tokens: 1000,
      messages: [{
        role: "user",
        content: [
          {
            type: "image",
            source: { type: "base64", media_type: "image/jpeg", data: base64Image },
          },
          {
            type: "text",
            text: `อ่านใบกำกับภาษี/ใบเสร็จนี้ ตอบเป็น JSON เท่านั้น ไม่มีข้อความอื่น:
{
  "vendor_name": "ชื่อผู้ขาย/บริษัทที่ออกบิล",
  "customer_name": "ชื่อลูกค้า/ผู้ซื้อ (ดึงให้ครบ)",
  "bill_number": "เลขที่บิล/ใบกำกับ",
  "amount_before_vat": "ยอดก่อน VAT (ตัวเลขเท่านั้น)",
  "vat_amount": "ยอด VAT (ตัวเลขเท่านั้น)",
  "total_amount": "ยอดรวมสุทธิ (ตัวเลขเท่านั้น)",
  "bill_date": "วันที่ในบิล (YYYY-MM-DD)",
  "confidence": "high/medium/low"
}
ถ้าไม่พบข้อมูลใด ให้ใส่ null`,
          },
        ],
      }],
    },
    {
      headers: {
        "x-api-key": process.env.ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
      },
    }
  );
  const text = res.data.content[0].text.trim().replace(/```json|```/g, "").trim();
  return JSON.parse(text);
}

// ─── สร้างข้อความตอบกลับ ──────────────────────
function formatReply(data, { saved, buyerKey, buyerLabel, sheetName }) {
  const fmt = (v) =>
    v ? parseFloat(v).toLocaleString("th-TH", { minimumFractionDigits: 2 }) : "-";

  const icon = buyerKey === "florish" ? "🏢" : buyerKey === "janewit" ? "👤" : "📂";
  const savedLine = saved
    ? `${icon} บันทึกแล้ว: ${buyerLabel} (Sheet: ${sheetName})\n📊 ไฟล์: data/bills.xlsx`
    : `⚠️ บันทึกไม่สำเร็จ`;

  return `🧾 อ่านบิลสำเร็จ
${savedLine}${!buyerKey ? "\n⚠️ ไม่พบชื่อผู้ซื้อ — บันทึกใน Sheet: unmatched" : ""}

🏷️  Vendor:     ${data.vendor_name    || "-"}
👤  ลูกค้า:     ${data.customer_name  || "-"}
🔢  เลขที่บิล:  ${data.bill_number    || "-"}
📅  วันที่:     ${data.bill_date      || "-"}

💰  ก่อน VAT:   ${fmt(data.amount_before_vat)} ฿
💸  VAT:        ${fmt(data.vat_amount)} ฿
💵  รวมสุทธิ:   ${fmt(data.total_amount)} ฿

📊  ความแม่นยำ: ${
    data.confidence === "high"   ? "สูง ✅" :
    data.confidence === "medium" ? "ปานกลาง ⚠️" : "ต่ำ ❌"
  }`;
}

// ─── LINE helpers ─────────────────────────────
const lineHeaders = () => ({
  Authorization: `Bearer ${process.env.LINE_CHANNEL_ACCESS_TOKEN}`,
  "Content-Type": "application/json",
});

async function replyLine(replyToken, text) {
  await axios.post(
    "https://api.line.me/v2/bot/message/reply",
    { replyToken, messages: [{ type: "text", text }] },
    { headers: lineHeaders() }
  );
}

async function pushLine(userId, text) {
  await axios.post(
    "https://api.line.me/v2/bot/message/push",
    { to: userId, messages: [{ type: "text", text }] },
    { headers: lineHeaders() }
  );
}

// ─── Webhook ──────────────────────────────────
app.post("/webhook", async (req, res) => {
  if (!verifySignature(req)) return res.status(401).send("Unauthorized");
  res.sendStatus(200);

  for (const event of req.body.events || []) {
    if (event.type !== "message") continue;
    const { replyToken } = event;
    const userId = event.source.userId;

    if (event.message.type === "image") {
      try {
        await replyLine(replyToken, "⏳ กำลังอ่านบิล รอสักครู่...");
        const base64     = await downloadImage(event.message.id);
        const data       = await extractBillData(base64);
        const saveResult = saveBill(data);
        await pushLine(userId, formatReply(data, saveResult));
      } catch (err) {
        console.error(err.message);
        await pushLine(userId, "❌ เกิดข้อผิดพลาด กรุณาส่งรูปบิลที่ชัดเจนอีกครั้ง");
      }

    } else if (event.message.type === "text") {
      const txt = (event.message.text || "").trim().toLowerCase();
      if (txt === "help" || txt === "ช่วยเหลือ") {
        await replyLine(replyToken,
          `📖 วิธีใช้งาน Bill Scanner Bot\n\n` +
          `📸 ส่งรูปบิล → บันทึกลง Excel อัตโนมัติ\n\n` +
          `📊 ไฟล์: data/bills.xlsx\n` +
          `🏢 Sheet "florish"   → บริษัทฟลอริช\n` +
          `👤 Sheet "janewit"   → นายเจนวิทย์\n` +
          `📂 Sheet "unmatched" → ระบุไม่ได้\n\n` +
          `⬇️ ดาวน์โหลด:\nhttps://your-app.railway.app/download`
        );
      } else {
        await replyLine(replyToken,
          "📸 กรุณาส่งรูปถ่ายบิล/ใบกำกับภาษี\nพิมพ์ 'help' เพื่อดูวิธีใช้งาน"
        );
      }
    }
  }
});

// ─── Download bills.xlsx ──────────────────────
app.get("/download", (_req, res) => {
  if (!fs.existsSync(BILLS_FILE)) return res.status(404).send("ยังไม่มีข้อมูล");
  res.download(BILLS_FILE, "bills.xlsx");
});

// ─── Status ───────────────────────────────────
app.get("/", (_req, res) => {
  const exists = fs.existsSync(BILLS_FILE);
  res.send(
    `Bill Scanner Bot 🤖<br>` +
    `ไฟล์: data/bills.xlsx — ${exists ? "✅ มีอยู่" : "⏳ ยังไม่มีข้อมูล"}<br>` +
    `<a href="/download">⬇️ ดาวน์โหลด bills.xlsx</a>`
  );
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`✅ Server on port ${PORT}`);
  console.log(`📊 Excel: ${BILLS_FILE}`);
  console.log(`   Sheets: florish | janewit | unmatched`);
});

// ─── Test Anthropic API ───────────────────────
app.get("/test-api", async (_req, res) => {
  try {
    const response = await axios.post(
      "https://api.anthropic.com/v1/messages",
      {
        model: "claude-3-5-sonnet-20241022",
        max_tokens: 10,
        messages: [{ role: "user", content: "say hi" }],
      },
      {
        headers: {
          "x-api-key": process.env.ANTHROPIC_API_KEY,
          "anthropic-version": "2023-06-01",
          "content-type": "application/json",
        },
      }
    );
    res.send("✅ API ใช้ได้: " + JSON.stringify(response.data.content));
  } catch (err) {
    res.send("❌ Error: " + err.message + " | Status: " + (err.response?.status) + " | " + JSON.stringify(err.response?.data));
  }
});
