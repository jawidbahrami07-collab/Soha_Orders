// /api/order.js
import ExcelJS from "exceljs";
import FormData from "form-data";

let orderCounter = 0; // شمارنده موقتی افزایشی (در هر اجرای سرور از صفر)

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ ok: false, message: "POST only" });
  }

  const BOT_TOKEN = process.env.BOT_TOKEN;
  const CHAT_ID = process.env.CHAT_ID;

  try {
    const data = req.body;
    const orderCode = `ORD-${String(orderCounter).padStart(6, "0")}`;
    orderCounter++;

    const now = new Date().toLocaleString("fa-IR");

    // 📋 ساخت پیام متنی برای تلگرام
    const lines = [];
    lines.push("📦 سفارش جدید دریافت شد");
    lines.push(`👤 نام و نام خانوادگی: ${data.name || "-"}`);
    lines.push(`📱 شماره تماس: ${data.phone || "-"}`);
    lines.push(`🏠 آدرس: ${data.address || "-"}`);
    if (data.postal_code) lines.push(`📮 کد پستی: ${data.postal_code}`);

    lines.push("\n🛍️ محصولات سفارش داده‌شده:");
    const productList = [];

    function addProd(qty, unit, title) {
      if (qty && Number(qty) > 0) {
        productList.push(`• ${title} — ${qty} ${unit}`);
      }
    }

    addProd(data.saha500_qty, data.saha500_unit, "سها ۵۰۰ گرمی سبز");
    addProd(data.saha250_qty, data.saha250_unit, "سها ۲۵۰ گرمی ساشه");
    addProd(data.box1kg_qty, data.box1kg_unit, "باکس پوچ یک کیلویی");
    addProd(data.goldPack_qty, data.goldPack_unit, "پاکت طلایی پنجره‌دار");
    addProd(data.plainPack_qty, data.plainPack_unit, "پاکت ساده یک کیلویی");

    if (productList.length === 0) productList.push("هیچ محصولی انتخاب نشده ❗");
    lines.push(productList.join("\n"));

    if (data.note) lines.push(`\n📝 توضیحات: ${data.note}`);
    lines.push(`🕒 زمان ثبت سفارش: ${now}`);
    lines.push(`🔢 کد سفارش: ${orderCode}`);

    const text = lines.join("\n");

    // ✉️ ارسال پیام متنی سفارش به گروه تلگرام
    await fetch(`https://api.telegram.org/bot${BOT_TOKEN}/sendMessage`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ chat_id: CHAT_ID, text }),
    });

    // 📊 ساخت فایل Excel برای همین سفارش
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("سفارش");

    sheet.columns = [
      { header: "فیلد", key: "field", width: 25 },
      { header: "مقدار", key: "value", width: 50 },
    ];

    const addRow = (key, value) => sheet.addRow({ field: key, value: value || "-" });

    addRow("کد سفارش", orderCode);
    addRow("نام و نام خانوادگی", data.name);
    addRow("شماره تماس", data.phone);
    addRow("آدرس", data.address);
    addRow("کد پستی", data.postal_code);
    addRow("زمان ثبت", now);
    addRow("توضیحات", data.note || "-");

    sheet.addRow({ field: "", value: "" });
    sheet.addRow({ field: "محصولات سفارش داده‌شده", value: "" });
    for (const p of productList) {
      sheet.addRow({ field: "", value: p });
    }

    // 🧾 تبدیل فایل اکسل به بایت و ارسال به تلگرام
    const buffer = await workbook.xlsx.writeBuffer();
    const formData = new FormData();
    formData.append("chat_id", CHAT_ID);
    formData.append("caption", `📄 فایل اکسل سفارش ${orderCode}`);
    formData.append("document", new Blob([buffer]), `${orderCode}.xlsx`);

    await fetch(`https://api.telegram.org/bot${BOT_TOKEN}/sendDocument`, {
      method: "POST",
      body: formData,
    });

    // ✅ پاسخ نهایی به مرورگر (فقط تأیید ساده)
    res.status(200).json({ ok: true });
  } catch (err) {
    console.error("Order error:", err);
    res.status(500).json({ ok: false, message: err.message });
  }
}
