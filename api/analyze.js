// api/analyze.js — Vercel Serverless Function (Node runtime).
// Trung gian giữ ANTHROPIC_API_KEY ở phía server (KHÔNG lộ ra trình duyệt).
// Trình duyệt gọi /api/analyze với đúng body Anthropic Messages API cần
// (model, max_tokens, messages) — hàm này chỉ gắn thêm x-api-key + anthropic-version
// rồi chuyển tiếp NGUYÊN VĂN sang api.anthropic.com, trả nguyên văn kết quả về.
// Đặt ANTHROPIC_API_KEY trong Vercel Project Settings → Environment Variables.

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: { message: "Method not allowed" } });
    return;
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    res.status(500).json({ error: { message: "Server thiếu biến môi trường ANTHROPIC_API_KEY." } });
    return;
  }

  try {
    const upstream = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify(req.body),
    });
    const data = await upstream.json();
    res.status(upstream.status).json(data);
  } catch (err) {
    res.status(502).json({ error: { message: "Lỗi khi gọi Anthropic API: " + String(err) } });
  }
}
