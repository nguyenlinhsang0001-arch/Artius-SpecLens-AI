// api/detect.js — Tầng ĐỊNH VỊ (grounding / Pass-1) cho bản web trên Vercel.
//
// Vai trò: nhận 1 ảnh phối cảnh -> gọi Gemini khoanh vùng vật thể rời
//          -> trả box đã chuẩn hoá về (x1,y1,x2,y2) trong [0..1]
//             (ĐÚNG convention parseBoxes của bạn, nên client không phải sửa format).
//
// Đặt file này TẠI: api/detect.js  (cùng chỗ với api/analyze.js) trong repo web.
// Vercel sẽ tự expose thành endpoint POST /api/detect.
//
// Yêu cầu:
//   - package.json:  thêm "@google/genai" vào dependencies  (SDK cần Node >= 22)
//   - Vercel env var: GEMINI_API_KEY  (KHÔNG commit key vào repo)
//   - Nếu api/analyze.js của bạn viết CommonJS (module.exports/require) thì đổi file này
//     cho khớp, hoặc đổi tên thành api/detect.mjs để chạy ESM độc lập.
//
// LƯU Ý model string: chuỗi model Gemini đổi khá nhanh. Mặc định dưới đây theo docs
//   hiện hành; hãy xác nhận lại tên model flash-tier đang phát hành trước khi chạy thật.

import { GoogleGenAI, Type } from "@google/genai";

const MODEL = process.env.GEMINI_DETECT_MODEL || "gemini-3.6-flash"; // TODO: xác nhận chuỗi model hiện hành

// Prompt: chỉ khoanh VẬT THỂ RỜI (nội thất, đèn, thiết bị vệ sinh, cửa, decor).
// KHÔNG khoanh bề mặt khuếch tán (sơn/đá mảng/gạch/giấy dán) — phần đó để Claude đọc theo vùng.
// PASS-1 = ĐỊNH VỊ TỪNG CÁ THỂ. Không để Gemini tự gộp đồ giống nhau nữa:
// việc "vật nào là cùng một sản phẩm" do tầng C (montage clustering ở client) quyết định,
// dựa trên so sánh thị giác cạnh nhau — chính xác hơn Gemini nhiều. Ở đây chỉ cần: THẤY ĐỦ + KHOANH ĐÚNG.
const DEFAULT_PROMPT =
  "Detect the 2D bounding box of EACH individual visible movable/fixed furniture item, " +
  "lighting fixture, sanitary fixture, door/window, and decor object in this interior render. " +
  "Be thorough with LIGHTING — include every distinct fixture: pendant/hanging lamps, ceiling spots/downlights, " +
  "wall sconces, wall-mounted reading lamps, bedside/table lamps, floor lamps, and picture lights. " +
  "Return ONE box PER physical instance. Do NOT merge repeated identical items into a single box: " +
  "if there are 6 identical chairs, return 6 separate boxes, one tightly on each chair. " +
  "Do NOT box diffuse surface finishes (paint, stone slab areas, tiles, wallpaper) NOR flat floor rugs/area rugs/carpets (those are handled as surfaces elsewhere). " +
  "Do NOT box any of these (ignore them completely): people/persons/staff, computer monitors or screens, desktop computers, laptops, keyboards, and books. " +
  "Never return masks. Limit to 40 objects. " +
  "Return a JSON array; each item has box_2d as [y_min, x_min, y_max, x_max] normalized 0-1000, " +
  "a short English label, and count = 1 for a single instance (only > 1 if this ONE box genuinely covers several identical items you could not separate).";

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }
  try {
    // Body kỳ vọng: { image: { media_type, data(base64) }, target_classes?: string[], prompt?: string }
    const { image, target_classes, prompt } = req.body || {};
    if (!image || !image.data) {
      res.status(400).json({ error: "Missing image (base64)" });
      return;
    }
    if (!process.env.GEMINI_API_KEY) {
      res.status(500).json({ error: "Missing GEMINI_API_KEY on server" });
      return;
    }

    const promptText =
      prompt ||
      (Array.isArray(target_classes) && target_classes.length
        ? `Detect the 2D bounding boxes of these object types if present: ${target_classes.join(", ")}. ` +
          "Return ONE box PER physical instance — do NOT merge identical repeats into one box. " +
          "Never return masks. Limit to 40 objects. Each item: box_2d [y_min,x_min,y_max,x_max] normalized 0-1000, a short label, and count = 1 for a single instance (only > 1 if one box unavoidably covers several identical items)."
        : DEFAULT_PROMPT);

    const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

    const response = await ai.models.generateContent({
      model: MODEL,
      contents: [
        {
          role: "user",
          parts: [
            { inlineData: { mimeType: image.media_type || "image/jpeg", data: image.data } },
            { text: promptText },
          ],
        },
      ],
      config: {
        temperature: 0.2,
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.ARRAY,
          items: {
            type: Type.OBJECT,
            properties: {
              box_2d: { type: Type.ARRAY, items: { type: Type.INTEGER } }, // [y_min,x_min,y_max,x_max] 0..1000
              label: { type: Type.STRING },
              count: { type: Type.INTEGER }, // số bản y hệt trong toàn ảnh (>=1)
            },
            required: ["box_2d", "label"],
          },
        },
      },
    });

    let raw = [];
    try {
      raw = JSON.parse(response.text || "[]");
    } catch {
      raw = [];
    }

    const clamp01 = (v) => (v < 0 ? 0 : v > 1 ? 1 : v);
    const regions = [];
    for (const b of raw) {
      const bb = b && b.box_2d;
      if (!Array.isArray(bb) || bb.length !== 4) continue;
      let [ymin, xmin, ymax, xmax] = bb;
      // 0..1000 -> 0..1, map sang (x1,y1,x2,y2)
      let x1 = xmin / 1000, y1 = ymin / 1000, x2 = xmax / 1000, y2 = ymax / 1000;
      if (x2 < x1) [x1, x2] = [x2, x1]; // Gemini đôi khi đảo min/max
      if (y2 < y1) [y1, y2] = [y2, y1];
      x1 = clamp01(x1); y1 = clamp01(y1); x2 = clamp01(x2); y2 = clamp01(y2);
      if (x2 - x1 < 0.003 || y2 - y1 < 0.003) continue; // bỏ box suy biến
      const count = Number.isInteger(b.count) && b.count > 0 ? b.count : 1;
      regions.push({ label: (b.label || "").trim(), count, x1, y1, x2, y2 });
    }

    res.status(200).json({ regions: nms(regions, 0.6), model: MODEL });
  } catch (err) {
    res.status(500).json({ error: (err && err.message) || "detect failed" });
  }
}

// --- NMS: khử box chồng nhau (hạng mục B4) ---
function iou(a, b) {
  const ix1 = Math.max(a.x1, b.x1), iy1 = Math.max(a.y1, b.y1);
  const ix2 = Math.min(a.x2, b.x2), iy2 = Math.min(a.y2, b.y2);
  const iw = Math.max(0, ix2 - ix1), ih = Math.max(0, iy2 - iy1);
  const inter = iw * ih;
  if (inter <= 0) return 0;
  const areaA = (a.x2 - a.x1) * (a.y2 - a.y1);
  const areaB = (b.x2 - b.x1) * (b.y2 - b.y1);
  return inter / (areaA + areaB - inter);
}

function nms(regions, thresh = 0.6) {
  const kept = [];
  // ưu tiên giữ box diện tích lớn (thường là instance rõ nhất)
  const sorted = [...regions].sort(
    (p, q) => (q.x2 - q.x1) * (q.y2 - q.y1) - (p.x2 - p.x1) * (p.y2 - p.y1)
  );
  for (const r of sorted) {
    if (kept.every((k) => iou(r, k) < thresh)) kept.push(r);
  }
  return kept;
}
