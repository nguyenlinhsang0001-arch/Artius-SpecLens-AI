/*
  Inventory Extractor — ARTIUS reskin (v2).
  FUNCTIONALITY is identical to uploads/inventory-extractor.jsx (v4):
  same parsing, material-code table, marker editing, thumbnails,
  Claude image analysis, Excel/TSV/annotated-image/JSON-bundle export.
  Visual layer: ARTIUS navy / steel-blue design, ARTIUS logo mark, and the
  material table is now grouped by "Nhóm" with a color-coded left rail and
  inline thumbnails instead of a flat spreadsheet-style table.
  Place "artius-logo-white.png" alongside this file.
*/
import React, { useState, useRef, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import logoUrl from "./artius-logo-white.png";

/* ---------- inline lucide-style icons (replace lucide-react) ---------- */
const S = { width: 15, height: 15, viewBox: "0 0 24 24", fill: "none", stroke: "currentColor", strokeWidth: 2, strokeLinecap: "round", strokeLinejoin: "round" };
const svg = (props, ...kids) => React.createElement("svg", { ...S, ...(props || {}) }, ...kids);
const P = (d) => React.createElement("path", { d });
const C = (cx, cy, r) => React.createElement("circle", { cx, cy, r });
const R = (x, y, w, h, rx) => React.createElement("rect", { x, y, width: w, height: h, rx });
const Upload = (p) => svg(p, P("M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"), P("M17 8l-5-5-5 5"), P("M12 3v12"));
const Download = (p) => svg(p, P("M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"), P("M7 10l5 5 5-5"), P("M12 15V3"));
const Copy = (p) => svg(p, R(9, 9, 13, 13, 2), P("M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"));
const Plus = (p) => svg(p, P("M12 5v14"), P("M5 12h14"));
const Trash2 = (p) => svg(p, P("M3 6h18"), P("M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"), P("M10 11v6"), P("M14 11v6"), P("M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"));
const AlertTriangle = (p) => svg(p, P("M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"), P("M12 9v4"), P("M12 17h.01"));
const MapPin = (p) => svg(p, P("M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"), C(12, 10, 3));
const ImageDown = (p) => svg(p, R(3, 3, 18, 18, 2), C(8.5, 8.5, 1.5), P("M21 15l-5-5L5 21"), P("M12 22v-4"));
const Hash = (p) => svg(p, P("M4 9h16"), P("M4 15h16"), P("M10 3 8 21"), P("M16 3l-2 18"));
const Package = (p) => svg(p, P("M16.5 9.4 7.5 4.21"), P("M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"), P("M3.27 6.96 12 12.01l8.73-5.05"), P("M12 22.08V12"));
const Loader2 = (p) => svg(p, P("M21 12a9 9 0 1 1-6.219-8.56"));
const ChevronLeft = (p) => svg(p, P("M15 18l-6-6 6-6"));
const PanelLeft = (p) => svg(p, R(3, 3, 18, 18, 2), P("M9 3v18"));
const Check = (p) => svg(p, P("M20 6L9 17l-5-5"));
const Undo2 = (p) => svg(p, P("M9 14 4 9l5-5"), P("M4 9h11a4 4 0 0 1 0 8h-1"));
const Search = (p) => svg(p, C(11, 11, 8), P("M21 21l-4.3-4.3"));
const XIcon = (p) => svg(p, P("M18 6 6 18"), P("M6 6l12 12"));
const Combine = (p) => svg(p, R(3, 3, 8, 8, 2), P("M9 14a2 2 0 0 0 2 2h1"), P("M14 9h6a1 1 0 0 1 1 1v10a1 1 0 0 1-1 1H10a1 1 0 0 1-1-1v-1"));
const Scissors = (p) => svg(p, C(6, 6, 3), C(6, 18, 3), P("M20 4 8.12 15.88"), P("M14.47 14.48 20 20"), P("M8.12 8.12 12 12"));
const ZoomIn = (p) => svg(p, C(11, 11, 8), P("M21 21l-4.3-4.3"), P("M11 8v6"), P("M8 11h6"));
const ChevronDown = (p) => svg(p, P("M6 9l6 6 6-6"));
const ChevronUp = (p) => svg(p, P("M18 15l-6-6-6 6"));

/* ================= UNCHANGED LOGIC (verbatim) ================= */
const NHOM_OPTS = ["Nội thất", "Đèn", "Vật liệu bề mặt", "Cửa & Vách kính", "Hardware", "Trang trí"];
const TINCAY_OPTS = ["Cao", "Trung bình", "Thấp"];

const CODE_NAME = {
  AW: "TRANH & TRANG TRÍ", AC: "NHỰA TẤM (MICA)", BL: "MÀN SÁO", BR: "TẤM XUYÊN SÁNG",
  CA: "THẢM CUỘN", CP: "TẤM COMPACT", CT: "RÈM", DC: "DECAL / STICKER", FA: "VẢI",
  D: "CỬA ĐI", W: "CỬA SỔ & KHUNG KÍNH",
  F: "NỘI THẤT RỜI", J: "NỘI THẤT CỐ ĐỊNH",
  GL: "KÍNH", GW: "MẢNG XANH & CÂY CẢNH", LT: "ĐÈN", LM: "LAMINATE",
  LE: "DA THUỘC", MD: "PHÀO", MF: "VÁN MFC", MI: "GƯƠNG", MT: "KIM LOẠI", PL: "CHẤT DẺO",
  PU: "GỖ SƠN PU", PT: "SƠN NƯỚC", PE: "SƠN HIỆU ỨNG", RB: "CAO SU", RUG: "THẢM RỜI",
  RT: "TẤM ĐAN MÂY", SA: "THIẾT BỊ VỆ SINH", SF: "HOÀN THIỆN ĐẶC BIỆT", SI: "SIMILI / GIẢ DA",
  ST: "ĐÁ", SS: "INOX", TL: "GẠCH", VL: "VINYL", VN: "VÁN LẠNG", WD: "GỖ ĐẶC",
  WF: "SÀN GỖ", WP: "GIẤY DÁN TƯỜNG",
};
const VALID_PREFIX = new Set(Object.keys(CODE_NAME));

// Tên sheet khi xuất Excel. F (nội thất rời) + J (nội thất cố định) gộp chung 1 sheet "NỘI THẤT".
// Các prefix khác giữ nguyên tên nhóm vật liệu như cũ.
function sheetCategory(pfx) {
  const p = String(pfx || "").toUpperCase();
  if (p === "F" || p === "J") return "NỘI THẤT";
  return VALID_PREFIX.has(p) ? CODE_NAME[p] : "KHÁC (chưa có mã)";
}

const CODE_LIST_TEXT =
  "AW=Tranh/trang trí, AC=Nhựa tấm mica, BL=Màn sáo, BR=Tấm xuyên sáng, CA=Thảm cuộn, CP=Tấm compact, " +
  "CT=Rèm, D=Cửa đi, DC=Decal, F=Nội thất rời (sofa/ghế/bàn di dời được), FA=Vải, GL=Kính, GW=Mảng xanh & cây cảnh (gộp cả chậu), " +
  "J=Nội thất cố định (tủ liền tường/quầy), LT=Đèn, LM=Laminate, LE=Da thuộc, " +
  "MD=Phào, MF=Ván MFC, MI=Gương, MT=Kim loại, PL=Chất dẻo, PU=Gỗ sơn PU, PT=Sơn, PE=Sơn hiệu ứng, " +
  "RB=Cao su, RUG=Thảm rời, RT=Tấm đan mây, SA=Thiết bị vệ sinh, SF=Hoàn thiện đặc biệt, SI=Simili/giả da, " +
  "ST=Đá, SS=Inox (thép không gỉ), TL=Gạch, VL=Vinyl, VN=Ván lạng, W=Cửa sổ/khung kính kiến trúc, WD=Gỗ đặc, WF=Ván sàn, WP=Giấy dán tường";

const PROMPT = `Bạn là chuyên gia bóc tách nội thất & vật liệu từ ảnh phối cảnh (render) nội thất, phục vụ báo giá. Hãy bóc CHI TIẾT.

Nhiệm vụ (LÀM CẢ HAI):
(1) LIỆT KÊ MỌI món NỘI THẤT nhìn thấy — mỗi món (hoặc nhóm giống hệt nhau) là 1 dòng, gán mã F (rời) hoặc J (cố định).
(2) Phát hiện từng LỚP VẬT LIỆU / CHI TIẾT hoàn thiện nhìn thấy (sơn, đá, gỗ, nẹp, đèn, kính…), định vị bằng bounding box và gán MÃ VẬT LIỆU.
Cả hai đều BẮT BUỘC — không được chỉ trả về vật liệu bề mặt mà bỏ đồ nội thất.

THỨ TỰ TRẢ VỀ (QUAN TRỌNG): xuất TẤT CẢ dòng NỘI THẤT (F/J) TRƯỚC TIÊN, ngay sau header; rồi mới đến đèn (LT), vật liệu bề mặt và các mã khác. Mục đích: nếu kết quả dài bị cắt bớt thì vẫn không mất đồ nội thất.

CHỈ trả về các dòng phân tách bằng "|". KHÔNG lời dẫn, markdown, code block.
Header (dòng đầu tiên):
ma|nhom|mon|vat_lieu_finish|so_luong|vi_tri|do_tin_cay|ghi_chu|boxes

Cột:
- ma: tiền tố mã vật liệu, CHỌN đúng 1 tiền tố từ BẢNG MÃ bên dưới theo vật liệu/loại của dòng. Chỉ ghi tiền tố (vd PT, ST, LT, D, W, F, J), KHÔNG kèm số.
  + Nội thất RỜI, di dời được (sofa, armchair, ghế đôn, ghế, bàn rời, đôn, kệ/tủ rời…) → F.
  + Nội thất CỐ ĐỊNH, gắn liền công trình (tủ liền tường, tủ bếp, quầy/quầy bar, vách tủ, kệ âm tường…) → J.
  + Chỉ để TRỐNG khi thật sự không xác định được loại.
- nhom: 1 trong [Nội thất, Đèn, Vật liệu bề mặt, Cửa & Vách kính, Hardware, Trang trí].
- mon: tên ngắn gọn (tiếng Việt), KHÔNG kèm số lượng trong tên (đừng viết "2 ghế", "bộ 4 đèn", "3 chậu"…) — số lượng ghi ở cột so_luong. Cá thể giống hệt nhau gộp 1 dòng, tên nhất quán.
- vat_lieu_finish, vi_tri: ngắn gọn.
- so_luong: SỐ NGUYÊN ≥ 1 — ĐẾM số vật thể CÙNG LOẠI (giống nhau) nhìn thấy trong ảnh cho dòng này (vd 4 ghế ăn giống nhau → 4). Nếu chỉ có 1 thì ghi 1.
- do_tin_cay: Cao / Trung bình / Thấp.
- ghi_chu: cảnh báo ngắn hoặc để trống.
- boxes: CHỈ ĐÚNG 1 box "x1,y1,x2,y2" (0..1; (x1,y1) trên-trái, (x2,y2) dưới-phải, ôm SÁT vật). QUAN TRỌNG: dù so_luong > 1, mỗi dòng CHỈ đặt 1 box duy nhất — đặt lên đúng cá thể NHÌN RÕ NHẤT (ít bị che, đủ sáng, ở tiền cảnh) trong tất cả cá thể cùng loại. KHÔNG đặt nhiều box cho các cá thể lặp lại.

BẢNG MÃ VẬT LIỆU (tiền tố = loại):
${CODE_LIST_TEXT}

ĐỊNH VỊ CHÍNH XÁC:
- Box nằm ĐÚNG trên vật thật; tâm box rơi trúng vật.
- Đèn (tường/trần/hắt): box lên CHÍNH thiết bị đèn, KHÔNG lên quầng sáng. Đèn tường đúng độ cao thật.

BÓC CHI TIẾT VẬT LIỆU (không mô tả chung chung):
- Tách một bề mặt/món thành TỪNG lớp vật liệu & chi tiết RIÊNG, mỗi lớp có mã riêng.
  Ví dụ: tường sơn đen có nẹp đồng => (1) PT "Sơn đen" + (2) MT "Nẹp đồng".
  Tủ gỗ tay nắm inox mặt đá => WD/PU thân gỗ + SS tay nắm + ST mặt đá.
- Nẹp, phào, chỉ, ron, tay nắm, bản lề tách thành dòng riêng.
- Cửa đi → mã D; cửa sổ / khung kính kiến trúc → mã W. Phụ kiện GẮN LIỀN cửa/cửa sổ (tay nắm, bản lề, khóa) gộp vào D/W. Phụ kiện kim loại rời khác → MT (kim loại) hoặc SS (inox) tùy chất liệu.

ĐỒ NỘI THẤT (BẮT BUỘC — đừng bỏ sót):
- Mỗi món nội thất (sofa, armchair, ghế, đôn, bàn, kệ, giường, đầu giường, tủ, quầy, vách tủ, kệ tivi…) PHẢI có 1 dòng riêng, coi CẢ MÓN là một dòng — KHÔNG được chỉ tách thành vật liệu rồi bỏ qua món.
- F = nội thất RỜI, di dời được. J = nội thất CỐ ĐỊNH, gắn liền công trình.
- Được phép THÊM dòng vật liệu bề mặt của món (vd vải bọc FA, mặt đá ST, thân gỗ WD) nếu nhìn rõ, nhưng dòng F/J của món vẫn là bắt buộc và KHÔNG bị thay thế.

CÂY XANH:
- Chậu cây CÓ cây cảnh → GỘP cả cụm (cây + chậu) thành 1 dòng, mã GW. KHÔNG tách "chậu" và "cây" thành 2 mã.
- Chỉ mảng cây/tường cây không chậu cũng dùng GW.

QUY TẮC ĐẾM:
- Chỉ tạo box cho cá thể NHÌN THẤY RÕ; mỗi cá thể 1 box. KHÔNG phỏng đoán vật bị khuất (ghi ở ghi_chu, không tăng số).
- Vật liệu bề mặt / chi tiết dạng đường: 1 box đại diện (không dùng để đếm).
- Không bịa. Không dùng "|" hay ";" trong chữ.

Ví dụ dòng:
PT|Vật liệu bề mặt|Sơn đen mờ|Tường|Vách tivi|Cao||0.05,0.10,0.40,0.85
MT|Vật liệu bề mặt|Nẹp đồng|Nẹp dọc trang trí|Vách tivi|Trung bình|Theo m dài|0.44,0.20,0.46,0.70
LT|Đèn|Đèn tường|Gắn tường, ánh sáng ấm|Tường trái|Cao||0.18,0.42,0.21,0.52; 0.61,0.42,0.64,0.52
F|Nội thất|Sofa băng|Bọc vải, khung gỗ|Phòng khách|Cao||0.30,0.55,0.72,0.86
J|Nội thất|Tủ tivi liền tường|Gỗ phủ laminate|Vách tivi|Cao||0.05,0.62,0.45,0.86`;

// Prompt cho TẦNG 2 (đọc từng crop). Dùng lại toàn bộ quy tắc/taxonomy của PROMPT,
// nhưng ép: ảnh là 1 crop chỉ chứa 1 vật -> trả ĐÚNG 1 dòng, cột boxes để TRỐNG
// (toạ độ đã có từ tầng detect Gemini).
const CROP_PROMPT = PROMPT +
  "\n\n[CHẾ ĐỘ CROP] Ảnh dưới đây là 1 ẢNH CẮT (crop) chỉ chứa MỘT vật thể/chi tiết chính. " +
  "Chỉ trả về ĐÚNG 1 DÒNG cho vật thể chính đó, theo đúng thứ tự cột đã nêu. " +
  "Cột boxes để TRỐNG (hệ thống đã có toạ độ). so_luong=1 (trừ khi trong chính crop này thấy rõ nhiều bản y hệt). " +
  "KHÔNG in header, KHÔNG markdown, KHÔNG giải thích.\n" +
  "PHÂN NHÓM (cột nhom) CHO ĐÚNG — chỉ dùng đúng một trong các giá trị: " +
  "\"Nội thất\", \"Đèn\", \"Vật liệu bề mặt\", \"Cửa & Vách kính\", \"Hardware\", \"Trang trí\".\n" +
  "- Phụ kiện TRANG TRÍ / bày biện: gối tựa & gối trang trí, khay, bình/lọ/chậu, giỏ/rổ, tranh, đồng hồ để bàn, sách bày, nến, tượng, cây & chậu cây, vật trang trí nhỏ -> nhom = \"Trang trí\" (TUYỆT ĐỐI KHÔNG xếp các thứ này vào \"Nội thất\").\n" +
  "- Đồ NỘI THẤT chức năng: giường, ghế, sofa, tủ, kệ, bàn, táp đầu giường, đôn -> \"Nội thất\".\n" +
  "- Đèn / thiết bị chiếu sáng -> \"Đèn\".\n" +
  "Nếu có dòng 'Gợi ý loại vật' bên dưới, hãy dùng nó để xác định đúng vật và đúng nhóm.";

// Prompt cho PASS BỀ MẶT — ĐỘC LẬP (KHÔNG kế thừa PROMPT), chỉ lấy vật liệu bề mặt/hoàn thiện.
// Đồ rời (nội thất/đèn/thiết bị/gương/thảm) đã do tầng detect Gemini + đọc crop lo -> ở đây CẤM liệt kê
// để tránh TRÙNG LẶP và để model tập trung đúng phần bề mặt mà detector hay bỏ sót.
const SURFACE_PROMPT =
`Bạn là chuyên gia bóc tách VẬT LIỆU BỀ MẶT / HOÀN THIỆN từ ảnh phối cảnh (render) nội thất, phục vụ báo giá.
CHỈ liệt kê các LỚP HOÀN THIỆN CỐ ĐỊNH gắn với sàn – tường – trần – cửa:
- Sàn: gỗ (WF), gạch (TL), đá (ST)
- Tường: ốp gỗ/lam (WD), veneer (VN), laminate (LM), MFC (MF), giấy dán tường (WP), sơn (PT), sơn hiệu ứng (PE), đá ốp (ST), kính/gương ốp mảng (GL)
- Trần: sơn/tấm trần, mảng đèn hắt (đặc điểm trần)
- Rèm/màn: rèm vải (CT), màn sáo (BL)
- Nẹp/phào (MD); vách kính/lan can kính (GL)

TUYỆT ĐỐI KHÔNG liệt kê đồ NỘI THẤT RỜI (giường, ghế, sofa, tủ, bàn, ottoman, kệ...), ĐÈN trang trí, THIẾT BỊ, GƯƠNG soi rời, THẢM rời, cây/chậu, vật trang trí. Những thứ đó đã xử lý ở lượt khác — nếu liệt kê sẽ gây TRÙNG LẶP. Nếu ảnh không có mảng bề mặt nào rõ thì trả về rỗng.

CHỈ trả về các dòng phân tách bằng "|" — KHÔNG header, KHÔNG markdown, KHÔNG giải thích. Mỗi dòng gồm 9 cột theo đúng thứ tự:
ma|nhom|mon|vat_lieu_finish|so_luong|vi_tri|do_tin_cay|ghi_chu|boxes
- ma: tiền tố mã theo loại bề mặt (WF, TL, ST, WD, VN, LM, MF, WP, PT, PE, GL, SF, CT, BL, MD).
- nhom: LUÔN ghi đúng "Vật liệu bề mặt".
- mon: tên mảng bề mặt ngắn gọn (vd "Sàn gỗ", "Ốp tường lam gỗ", "Rèm cửa", "Trần thạch cao", "Sơn tường").
- vat_lieu_finish: mô tả vật liệu/màu/vân ngắn gọn.
- so_luong: 1.
- vi_tri: khu vực (vd "Sàn phòng ngủ", "Tường đầu giường", "Cửa sổ").
- do_tin_cay: Cao / Trung bình / Thấp.
- ghi_chu: ngắn gọn hoặc để trống.
- boxes: 1 box "x1,y1,x2,y2" (0..1; (x1,y1) trên-trái, (x2,y2) dưới-phải) của 1 vùng ĐẠI DIỆN cho mảng bề mặt đó.`;

let _uid = 0;
const nextId = () => (_uid += 1);
const clamp01 = (v) => Math.max(0, Math.min(1, v));

// ---- A4: chuẩn hoá key gộp XUYÊN ẢNH (bảo thủ) ----
// Bỏ dấu tiếng Việt + hạ chữ, tách token rồi SORT -> "bọc vải" == "vải bọc", "Sofa Vải" == "vai sofa".
// CHỈ dùng để so khớp GỘP; chữ hiển thị trong bảng vẫn giữ nguyên như AI trả về.
function stripVN(s) {
  return String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d").replace(/Đ/g, "D").toLowerCase();
}
function normTokens(s) {
  return stripVN(s).replace(/[^a-z0-9\s]/g, " ").split(/\s+/).filter(Boolean).sort().join(" ");
}
// Key gộp: prefix + token(món) + token(vật liệu). Hai dòng cùng bộ token -> cùng 1 mã.
function mergeKey(pfx, mon, vat_lieu) {
  return String(pfx || "").toUpperCase() + "||" + normTokens(mon) + "||" + normTokens(vat_lieu);
}
// Tách chuỗi vị trí thành các phần rời (loại trùng) khi gộp dòng.
const splitLocs = (s) => String(s || "").split(/[,;]/).map((x) => x.trim()).filter(Boolean);

// --- Khử trùng ĐỒ RỜI trong CÙNG 1 ẢNH (khi Gemini khoanh nhiều box cho cùng 1 vật) ---
function _boxInter(a, b) {
  const ix1 = Math.max(a.x1, b.x1), iy1 = Math.max(a.y1, b.y1);
  const ix2 = Math.min(a.x2, b.x2), iy2 = Math.min(a.y2, b.y2);
  return Math.max(0, ix2 - ix1) * Math.max(0, iy2 - iy1);
}
function _boxArea(a) { return Math.max(0, a.x2 - a.x1) * Math.max(0, a.y2 - a.y1); }
function boxesOverlap(a, b) {
  if (!a || !b) return false;
  const inter = _boxInter(a, b); if (inter <= 0) return false;
  const areaA = _boxArea(a), areaB = _boxArea(b);
  const iou = inter / (areaA + areaB - inter);
  const contain = inter / Math.max(1e-9, Math.min(areaA, areaB)); // phần nhỏ nằm trong phần lớn bao nhiêu
  return iou > 0.35 || contain > 0.6;
}
// Cùng prefix + cùng tên món + box CHỒNG nhau -> gộp 1 dòng (giữ box lớn hơn, SL lấy MAX, gộp vị trí).
// Đồ trùng tên nhưng box TÁCH RỜI (2 vật thật) -> giữ riêng.
function dedupeObjects(items) {
  const out = [];
  for (const it of items) {
    const box = it.instances && it.instances[0];
    const key = String(it.prefix || "").toUpperCase() + "||" + normTokens(it.mon);
    const hit = out.find((d) => d._k === key && boxesOverlap(d.instances && d.instances[0], box));
    if (hit) {
      const hb = hit.instances[0];
      if (box && _boxArea(box) > _boxArea(hb)) hit.instances[0] = box;
      const s1 = parseInt(hit.soLuong, 10), s2 = parseInt(it.soLuong, 10);
      hit.soLuong = Math.max(Number.isFinite(s1) ? s1 : 1, Number.isFinite(s2) ? s2 : 1);
      const locs = new Set(splitLocs(hit.vi_tri)); splitLocs(it.vi_tri).forEach((s) => locs.add(s));
      hit.vi_tri = Array.from(locs).join(", ");
    } else {
      out.push({ ...it, _k: key });
    }
  }
  return out.map((it) => { const { _k, ...rest } = it; return rest; });
}

// Gom tập id ẢNH NGUỒN của 1 item: từ srcImgs (đã gộp trước đó), srcImg (lúc phân tích), và imgId của các box.
// Nhờ cái này, kể cả dòng CHƯA có box vẫn biết được nó bóc ra từ ảnh nào để gắn lại đúng chỗ.
function srcIdsOf(it) {
  const s = new Set();
  (it.srcImgs || []).forEach((x) => s.add(x));
  if (it.srcImg != null) s.add(it.srcImg);
  (it.instances || []).forEach((b) => { if (b.imgId != null) s.add(b.imgId); });
  return s;
}

function parseBoxes(s) {
  if (!s) return [];
  const out = [];
  for (const chunk of s.split(";")) {
    const n = chunk.trim().split(",").map((x) => parseFloat(x));
    if (n.length >= 4 && n.slice(0, 4).every((v) => !isNaN(v))) {
      let [x1, y1, x2, y2] = n.map(clamp01);
      if (x2 < x1) [x1, x2] = [x2, x1];
      if (y2 < y1) [y1, y2] = [y2, y1];
      out.push({ x1, y1, x2, y2 });
    }
  }
  return out;
}

function parseItems(text) {
  const clean = String(text || "").replace(/```/g, "").trim();
  const lines = clean.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  const items = [];
  for (const line of lines) {
    if (!line.includes("|")) continue;
    const parts = line.split("|");
    const f0 = (parts[0] || "").trim().toLowerCase();
    const f1 = (parts[1] || "").trim().toLowerCase();
    if (f0 === "ma" || f1 === "nhom" || f1 === "nhóm") continue;
    if (parts.length < 3) continue;
    const prefix = (parts[0] || "").trim();
    const nhom = (parts[1] || "").trim();
    const mon = (parts[2] || "").trim();
    const vat_lieu = (parts[3] || "").trim();
    const soLuongRaw = (parts[4] || "").trim();
    const vi_tri = (parts[5] || "").trim();
    const do_tin_cay = (parts[6] || "").trim() || "Trung bình";
    const ghi_chu = (parts[7] || "").trim();
    const boxStr = parts.slice(8).join("|").trim();
    if (!mon && !nhom) continue;
    const boxes = parseBoxes(boxStr).slice(0, 1); // CHỈ giữ 1 box = đối tượng nhìn rõ nhất (1 ký hiệu / dòng)
    const slNum = parseInt(soLuongRaw, 10);
    const soLuong = Number.isFinite(slNum) && slNum > 0 ? slNum : (boxes.length || 1);
    items.push({ id: nextId(), prefix, ma: "", nhom, mon, vat_lieu, soLuong, vi_tri, do_tin_cay, ghi_chu, instances: boxes, thumb: null });
  }
  return items;
}

function codeItems(items, getPrefix) {
  const counters = {}, seen = {};
  return items.map((it) => {
    const pfx = String(getPrefix(it) || "").toUpperCase().replace(/[^A-Z]/g, "");
    if (!VALID_PREFIX.has(pfx)) return { ...it, ma: "" };
    // Gộp mã khi TRÙNG prefix + tên món + vật liệu/finish (đã chuẩn hoá bỏ dấu & sort token — xem mergeKey).
    // Nhờ vậy cách viết lệch nhau giữa các ảnh vẫn nhận CÙNG 1 mã; còn khác finish thật thì vẫn tách mã.
    const key = mergeKey(pfx, it.mon, it.vat_lieu);
    if (seen[key]) return { ...it, ma: seen[key] };
    counters[pfx] = (counters[pfx] || 0) + 1;
    const code = pfx + "-" + String(counters[pfx]).padStart(2, "0");
    seen[key] = code;
    return { ...it, ma: code };
  });
}

// Gộp các dòng TRÙNG (cùng prefix + tên món + vật liệu/màu) thành MỘT dòng duy nhất.
// - instances (các bounding box) được nối lại → 1 mã nhưng vẫn giữ đủ vị trí trên ảnh.
// - vi_tri / ghi_chu gộp lại (loại trùng); do_tin_cay lấy mức cao nhất.
// - Dòng chưa có mã hợp lệ thì KHÔNG gộp (giữ riêng), vì không đủ căn cứ để coi là trùng.
function mergeRows(items, getPrefix) {
  const map = new Map(), out = [], srcMap = new Map();
  const rank = { "Cao": 3, "Trung bình": 2, "Thấp": 1 };
  for (const it of items) {
    const pfx = String(getPrefix(it) || "").toUpperCase().replace(/[^A-Z]/g, "");
    if (!pfx || !VALID_PREFIX.has(pfx)) { out.push(it); continue; }
    const key = mergeKey(pfx, it.mon, it.vat_lieu);
    if (!map.has(key)) {
      const base = { ...it, instances: [...(it.instances || [])] };
      map.set(key, base); out.push(base); srcMap.set(base, srcIdsOf(it));
    } else {
      const base = map.get(key);
      // Giữ 1 ký hiệu (đối tượng rõ nhất): nếu base chưa có box mà dòng mới có thì lấy 1 box của dòng mới.
      if ((!base.instances || base.instances.length === 0) && it.instances && it.instances.length) base.instances = [it.instances[0]];
      // so_luong khi gộp trùng: lấy MAX (tránh cộng dồn trùng khi cùng 1 cụm vật xuất hiện ở nhiều ảnh).
      const slBase = parseInt(base.soLuong, 10), slIt = parseInt(it.soLuong, 10);
      base.soLuong = Math.max(Number.isFinite(slBase) ? slBase : 0, Number.isFinite(slIt) ? slIt : 0) || base.soLuong || 1;
      const locs = new Set(splitLocs(base.vi_tri)); splitLocs(it.vi_tri).forEach((s) => locs.add(s));
      base.vi_tri = Array.from(locs).join(", ");
      const note = String(it.ghi_chu || "").trim();
      if (note && String(base.ghi_chu || "").indexOf(note) < 0) base.ghi_chu = [base.ghi_chu, note].filter(Boolean).join("; ");
      if ((rank[it.do_tin_cay] || 0) > (rank[base.do_tin_cay] || 0)) base.do_tin_cay = it.do_tin_cay;
      const s = srcMap.get(base); srcIdsOf(it).forEach((x) => s.add(x));
    }
  }
  // Ghi lại nguồn ảnh lên từng dòng (quan trọng cho dòng chưa gắn box: vẫn biết bóc từ ảnh nào).
  for (const base of out) { if (srcMap.has(base)) base.srcImgs = Array.from(srcMap.get(base)); }
  return out;
}

// Gộp trùng TRONG CÙNG 1 ẢNH: khác mergeRows ở chỗ CỘNG DỒN so_luong (không lấy MAX).
// Lý do: nếu tầng C lỡ cho ra 2 dòng cùng mô tả trong 1 ảnh (vd 2 cụm ghế bị mô tả y hệt),
// đó là 2 cụm vật THẬT nên số lượng phải cộng lại. Còn xuyên-ảnh vẫn để mergeRows lấy MAX
// (tránh cộng dồn khi cùng cụm vật xuất hiện lại ở ảnh khác). Chạy TRƯỚC mergeRows, trên
// đúng các item vừa bóc từ 1 ảnh (nên toàn bộ đều same-image).
function mergeSameImage(items) {
  const map = new Map(), out = [];
  const rank = { "Cao": 3, "Trung bình": 2, "Thấp": 1 };
  for (const it of items) {
    const pfx = String(it.prefix || "").toUpperCase().replace(/[^A-Z]/g, "");
    if (!pfx || !VALID_PREFIX.has(pfx)) { out.push(it); continue; }
    const key = mergeKey(pfx, it.mon, it.vat_lieu);
    if (!map.has(key)) {
      const base = { ...it, instances: [...(it.instances || [])], memberBoxes: [...(it.memberBoxes || [])] };
      map.set(key, base); out.push(base);
    } else {
      const base = map.get(key);
      const s1 = parseInt(base.soLuong, 10), s2 = parseInt(it.soLuong, 10);
      const n1 = Number.isFinite(s1) ? s1 : 1, n2 = Number.isFinite(s2) ? s2 : 1;
      // Chỉ CỘNG DỒN cho cụm đồ rời đã qua clustering (có memberBoxes) — số lượng là ĐẾM được.
      // Vật liệu bề mặt (không memberBoxes) tính theo m², không đếm -> giữ MAX như cũ.
      const countable = (base.memberBoxes && base.memberBoxes.length) || (it.memberBoxes && it.memberBoxes.length);
      base.soLuong = countable ? (n1 + n2) : Math.max(n1, n2);
      if ((!base.instances || !base.instances.length) && it.instances && it.instances.length) base.instances = [it.instances[0]];
      if (it.memberBoxes && it.memberBoxes.length) base.memberBoxes = [...(base.memberBoxes || []), ...it.memberBoxes];
      const locs = new Set(splitLocs(base.vi_tri)); splitLocs(it.vi_tri).forEach((s) => locs.add(s)); base.vi_tri = Array.from(locs).join(", ");
      const note = String(it.ghi_chu || "").trim();
      if (note && String(base.ghi_chu || "").indexOf(note) < 0) base.ghi_chu = [base.ghi_chu, note].filter(Boolean).join("; ");
      if ((rank[it.do_tin_cay] || 0) > (rank[base.do_tin_cay] || 0)) base.do_tin_cay = it.do_tin_cay;
    }
  }
  return out;
}

// Sắp xếp bảng theo MÃ vật liệu: A→Z, phần số theo thứ tự tự nhiên (D-02 trước D-10).
// Tên sheet (nhóm) của một dòng — dùng cho tab sheet trong bảng, khớp với lúc export Excel.
const rowSheet = (r) => sheetCategory((r.ma || "").split("-")[0] || r.prefix || "");

// numeric:true để "PT-2" xếp trước "PT-10"; dòng chưa có mã (ma rỗng) luôn đẩy xuống cuối.
function sortRows(rs) {
  return [...rs].sort((a, b) => {
    const A = (a.ma || "").trim(), B = (b.ma || "").trim();
    if (!A && !B) return 0;
    if (!A) return 1;
    if (!B) return -1;
    return A.localeCompare(B, undefined, { numeric: true, sensitivity: "base" });
  });
}

function makeThumb(imgEl, box, px = 72) {
  try {
    const nw = imgEl.naturalWidth, nh = imgEl.naturalHeight;
    if (!nw || !nh) return null;
    let sx = box.x1 * nw, sy = box.y1 * nh;
    let sw = Math.max(6, (box.x2 - box.x1) * nw), sh = Math.max(6, (box.y2 - box.y1) * nh);
    const padX = sw * 0.15, padY = sh * 0.15;
    sx = Math.max(0, sx - padX); sy = Math.max(0, sy - padY);
    sw = Math.min(nw - sx, sw + 2 * padX); sh = Math.min(nh - sy, sh + 2 * padY);
    const c = document.createElement("canvas"); c.width = px; c.height = px;
    const ctx = c.getContext("2d");
    ctx.fillStyle = "#0d1119"; ctx.fillRect(0, 0, px, px);
    const scale = Math.max(px / sw, px / sh);
    const dw = sw * scale, dh = sh * scale;
    ctx.drawImage(imgEl, sx, sy, sw, sh, (px - dw) / 2, (px - dh) / 2, dw, dh);
    return c.toDataURL("image/jpeg", 0.82);
  } catch (e) { return null; }
}

function makeExportCrop(imgEl, box, maxPx = 440) {
  try {
    const nw = imgEl.naturalWidth, nh = imgEl.naturalHeight;
    if (!nw || !nh) return null;
    let sx = box.x1 * nw, sy = box.y1 * nh, sw = (box.x2 - box.x1) * nw, sh = (box.y2 - box.y1) * nh;
    const padX = sw * 0.12, padY = sh * 0.12;
    sx = Math.max(0, sx - padX); sy = Math.max(0, sy - padY);
    sw = Math.min(nw - sx, sw + 2 * padX); sh = Math.min(nh - sy, sh + 2 * padY);
    const scale = Math.min(1, maxPx / Math.max(sw, sh));
    const dw = Math.max(1, Math.round(sw * scale)), dh = Math.max(1, Math.round(sh * scale));
    const c = document.createElement("canvas"); c.width = dw; c.height = dh;
    c.getContext("2d").drawImage(imgEl, sx, sy, sw, sh, 0, 0, dw, dh);
    return { data: c.toDataURL("image/jpeg", 0.85), w: dw, h: dh };
  } catch (e) { return null; }
}

// Bỏ tiền tố "data:image/...;base64," -> chỉ còn base64 (để gửi cho API).
function b64of(dataUrl) { const i = String(dataUrl || "").indexOf(","); return i >= 0 ? dataUrl.slice(i + 1) : dataUrl; }

/* ===================== TẦNG C: gom nhóm bằng thị giác (montage clustering) =====================
   Bài toán: Gemini (B) khoanh TỪNG cá thể -> nhiều ghế giống nhau ra nhiều box. Nếu đọc mỗi crop
   ĐỘC LẬP thì Claude mô tả lệch nhau ("Ghế armchair" vs "Armchair (ghế bành đơn)") -> không gộp được.
   Cách xử lý: ghép mọi crop thành 1 ẢNH LƯỚI đánh số rồi hỏi Claude "ô nào là CÙNG 1 sản phẩm?".
   So sánh cạnh nhau nên phân biệt được "cùng mẫu khác ánh sáng" vs "hai mẫu khác nhau" — điều mà
   string-matching và cả Gemini đều làm dở. Sau đó mỗi nhóm CHỈ đọc 1 cá thể đại diện (full-res). */

// Ghép các vùng (regions) thành 1 canvas lưới, mỗi ô là crop của 1 vật + số thứ tự (1..n) góc trên-trái.
function buildMontage(imgEl, regions, cell = 240, pad = 8, maxEdge = 2200) {
  try {
    const n = regions.length;
    if (!n || !imgEl || !imgEl.naturalWidth) return null;
    const cols = Math.ceil(Math.sqrt(n));
    const rows = Math.ceil(n / cols);
    let cW = cell, cH = cell;
    let W = cols * cW + (cols + 1) * pad, H = rows * cH + (rows + 1) * pad;
    const sc = Math.min(1, maxEdge / Math.max(W, H)); // co lại nếu lưới quá lớn
    if (sc < 1) { cW = Math.max(120, Math.floor(cW * sc)); cH = cW; W = cols * cW + (cols + 1) * pad; H = rows * cH + (rows + 1) * pad; }
    const c = document.createElement("canvas"); c.width = W; c.height = H;
    const ctx = c.getContext("2d");
    ctx.fillStyle = "#0d1119"; ctx.fillRect(0, 0, W, H);
    const nw = imgEl.naturalWidth, nh = imgEl.naturalHeight;
    for (let i = 0; i < n; i++) {
      const r = regions[i];
      const col = i % cols, row = Math.floor(i / cols);
      const x = pad + col * (cW + pad), y = pad + row * (cH + pad);
      let sx = r.x1 * nw, sy = r.y1 * nh, sw = Math.max(4, (r.x2 - r.x1) * nw), sh = Math.max(4, (r.y2 - r.y1) * nh);
      const padX = sw * 0.08, padY = sh * 0.08;
      sx = Math.max(0, sx - padX); sy = Math.max(0, sy - padY);
      sw = Math.min(nw - sx, sw + 2 * padX); sh = Math.min(nh - sy, sh + 2 * padY);
      const scale = Math.min(cW / sw, cH / sh), dw = sw * scale, dh = sh * scale;
      ctx.fillStyle = "#0d1119"; ctx.fillRect(x, y, cW, cH);
      ctx.drawImage(imgEl, sx, sy, sw, sh, x + (cW - dw) / 2, y + (cH - dh) / 2, dw, dh);
      const label = String(i + 1);
      ctx.font = "bold 18px sans-serif";
      const tw = ctx.measureText(label).width;
      ctx.fillStyle = "rgba(0,0,0,0.78)"; ctx.fillRect(x + 3, y + 3, tw + 12, 24);
      ctx.fillStyle = "#fff"; ctx.textBaseline = "top"; ctx.fillText(label, x + 9, y + 6);
      ctx.strokeStyle = "rgba(255,255,255,0.16)"; ctx.lineWidth = 1; ctx.strokeRect(x + 0.5, y + 0.5, cW - 1, cH - 1);
    }
    return c.toDataURL("image/jpeg", 0.85);
  } catch (e) { return null; }
}

// Parse JSON nhóm từ text model trả về. An toàn: số nào thiếu -> TÁCH thành nhóm riêng (thà tách còn hơn gộp nhầm).
function parseGroups(text, n) {
  let s = String(text || "").replace(/```json/gi, "").replace(/```/g, "").trim();
  const i = s.indexOf("{"), j = s.lastIndexOf("}");
  if (i >= 0 && j > i) s = s.slice(i, j + 1);
  let obj; try { obj = JSON.parse(s); } catch { return null; }
  const groups = obj && Array.isArray(obj.groups) ? obj.groups : null;
  if (!groups) return null;
  const out = [], seen = new Set();
  for (const g of groups) {
    const mem = (g && Array.isArray(g.members) ? g.members : [])
      .map((x) => parseInt(x, 10)).filter((x) => x >= 1 && x <= n && !seen.has(x));
    mem.forEach((x) => seen.add(x));
    if (mem.length) out.push({ members: mem, loai: String((g && g.loai) || "").trim(), nhom: String((g && g.nhom) || "").trim() });
  }
  for (let k = 1; k <= n; k++) if (!seen.has(k)) { out.push({ members: [k], loai: "", nhom: "" }); seen.add(k); }
  return out.length ? out : null;
}

// Gọi Claude 1 lần với ảnh montage -> danh sách nhóm [{members:[1based..], loai, nhom}].
async function clusterRegions(montageB64, n) {
  const prompt =
`Ảnh kèm theo là 1 lưới (contact sheet) gồm ${n} ô, mỗi ô đánh SỐ ở góc trên-trái (1..${n}). Mỗi ô là ảnh cắt của MỘT vật thể trong CÙNG một ảnh nội thất.
Nhiệm vụ: gom những ô là CÙNG MỘT sản phẩm (cùng kiểu dáng/thiết kế/chất liệu; bỏ qua khác biệt nhỏ do góc nhìn, khoảng cách hay ánh sáng) vào chung 1 nhóm. Hai vật KHÁC kiểu dáng phải ở 2 nhóm khác nhau. Mỗi ô thuộc ĐÚNG 1 nhóm; mọi số 1..${n} xuất hiện đúng một lần.
CHỈ trả về JSON (không markdown, không giải thích), dạng:
{"groups":[{"members":[1,3,5],"loai":"ghế armchair","nhom":"Nội thất"},{"members":[2],"loai":"bàn trà tròn","nhom":"Nội thất"}]}
"nhom" chọn 1 trong ["Nội thất","Đèn","Vật liệu bề mặt","Cửa & Vách kính","Hardware","Trang trí"]. "loai" là tên loại ngắn gọn tiếng Việt.`;
  try {
    const res = await fetch("/api/analyze", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: "claude-sonnet-4-6", max_tokens: 1200,
        messages: [{ role: "user", content: [
          { type: "image", source: { type: "base64", media_type: "image/jpeg", data: montageB64 } },
          { type: "text", text: prompt },
        ] }],
      }),
    });
    if (!res.ok) return null;
    const data = await res.json();
    if (data && data.type === "error") return null;
    const textOut = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n");
    return parseGroups(textOut, n);
  } catch (e) { return null; }
}

// map có giới hạn số lời gọi song song (tránh dội rate-limit khi đọc nhiều crop cùng lúc).
async function mapLimit(arr, limit, fn) {
  const out = new Array(arr.length); let i = 0;
  async function worker() { while (i < arr.length) { const idx = i++; out[idx] = await fn(arr[idx], idx); } }
  await Promise.all(Array.from({ length: Math.min(Math.max(1, limit), arr.length || 1) }, worker));
  return out;
}

function separate(pts, minDist, w, h, iters) {
  const p = pts.map((o) => ({ x: o.x, y: o.y }));
  const n = p.length, N = iters || 70;
  for (let it = 0; it < N; it++) {
    let moved = false;
    for (let i = 0; i < n; i++) {
      for (let j = i + 1; j < n; j++) {
        let dx = p[j].x - p[i].x, dy = p[j].y - p[i].y;
        let d = Math.sqrt(dx * dx + dy * dy);
        if (d < minDist) {
          if (d < 0.001) { const a = i * 2.399963; dx = Math.cos(a); dy = Math.sin(a); d = 1; }
          const push = (minDist - d) / 2, ux = dx / d, uy = dy / d;
          p[i].x -= ux * push; p[i].y -= uy * push; p[j].x += ux * push; p[j].y += uy * push; moved = true;
        }
      }
    }
    for (let k = 0; k < n; k++) { p[k].x = Math.max(minDist / 2, Math.min(w - minDist / 2, p[k].x)); p[k].y = Math.max(minDist / 2, Math.min(h - minDist / 2, p[k].y)); }
    if (!moved) break;
  }
  return p;
}

/* ================= ARTUS navy theme (visual layer only) ================= */
const css = `
@import url('https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@400;500;600;700&family=Newsreader:wght@400;500&display=swap');
html, body, #root { margin:0; padding:0; height:100%; background:#070a11; }
.ax { --bg:#070a11; --side1:#0d1220; --side2:#090c15; --panel:#101725; --panel2:#111827;
  --card1:#121a2a; --card2:#0d131f; --input:#0d1119; --line:rgba(255,255,255,0.07); --line2:rgba(255,255,255,0.12);
  --ac:#7ba3cf; --ac2:#9dc0e6; --ac3:#bcd6f0; --acink:#0c1524; --acsoft:rgba(123,163,207,0.14);
  --tx:#eef2f8; --tx2:#e6eaf2; --tx3:#aab4c4; --mut:#8a93a3; --mut2:#6f7889; --faint:#5f6878;
  --amber:#e0a44a; --amber2:#e0b57a; --amber-soft:rgba(224,164,74,0.1); --green:#7fd8ab;
  --sans:'Be Vietnam Pro',system-ui,sans-serif; --serif:'Newsreader',serif;
  display:flex; min-height:100vh; background:radial-gradient(1200px 800px at 78% -10%, rgba(123,163,207,0.08), transparent 60%), var(--bg);
  color:var(--tx2); font-family:var(--sans); }
.ax *, .ax *::before, .ax *::after { box-sizing:border-box; }

/* sidebar */
.ax-side { width:300px; flex:0 0 300px; background:linear-gradient(180deg,var(--side1),var(--side2)); border-right:1px solid rgba(255,255,255,0.06);
  padding:24px 20px 32px; position:sticky; top:0; height:100vh; overflow-y:auto; transition:flex-basis .2s ease, width .2s ease, padding .2s ease; }
.ax-side.collapsed { width:0; flex-basis:0; padding-left:0; padding-right:0; overflow:hidden; border-right:none; }
.ax-side.collapsed > * { display:none; }
.ax-collapse { position:absolute; top:16px; right:14px; width:28px; height:28px; display:inline-flex; align-items:center; justify-content:center;
  background:#131a2a; border:1px solid rgba(255,255,255,0.1); border-radius:8px; color:var(--tx3); cursor:pointer; transition:border-color .15s,color .15s; z-index:2; }
.ax-collapse:hover { border-color:rgba(123,163,207,0.5); color:var(--ac2); }
.ax-toprow { display:flex; align-items:center; gap:12px; }
.ax-expand { padding:8px 10px; }
.ax-brand { text-align:center; padding:6px 0 20px; border-bottom:1px solid rgba(255,255,255,0.06); margin-bottom:22px; }
.ax-word { font-family:var(--serif); font-size:26px; letter-spacing:.42em; font-weight:400; color:var(--tx); padding-left:.42em; }
.ax-tag { font-size:8.5px; letter-spacing:.34em; color:#667186; margin-top:3px; }
.ax-pill { display:inline-block; margin-top:12px; font-size:9px; letter-spacing:.28em; color:var(--ac); border:1px solid rgba(123,163,207,0.32); border-radius:999px; padding:4px 12px; }
.ax-sec { display:flex; align-items:center; gap:8px; font-size:11px; font-weight:700; letter-spacing:.12em; text-transform:uppercase; color:var(--ac2); margin-bottom:12px; }
.ax-sec svg { color:var(--ac); }
.ax-field { margin-bottom:11px; }
.ax-flabel { font-size:9.5px; letter-spacing:.1em; text-transform:uppercase; color:var(--faint); margin-bottom:4px; display:block; }
.ax-input { width:100%; background:var(--input); border:1px solid rgba(255,255,255,0.09); border-radius:9px; padding:9px 11px; color:var(--tx2); font-family:var(--sans); font-size:12.5px; outline:none; transition:border-color .15s; }
.ax-input:focus { border-color:rgba(123,163,207,0.55); }
.ax-input::placeholder { color:var(--faint); }
.ax-date { font-size:11px; color:var(--mut2); margin-top:2px; }
.ax-stats { display:grid; grid-template-columns:1fr 1fr 1fr; gap:8px; }
.ax-stat { background:var(--panel2); border:1px solid rgba(255,255,255,0.06); border-radius:10px; padding:11px 6px; text-align:center; }
.ax-stat b { font-size:23px; font-weight:700; display:block; line-height:1; }
.ax-stat span { font-size:8.5px; letter-spacing:.04em; color:var(--faint); margin-top:5px; display:block; }

/* main */
.ax-main { flex:1; min-width:0; display:flex; flex-direction:column; }
.ax-top { display:flex; align-items:center; justify-content:space-between; padding:16px 32px; border-bottom:1px solid rgba(255,255,255,0.06);
  position:sticky; top:0; background:rgba(7,10,17,0.82); backdrop-filter:blur(12px); z-index:20; }
.ax-h1 { margin:0; font-size:24px; font-weight:700; color:var(--tx); letter-spacing:-.01em; }
.ax-sub { margin:3px 0 0; font-size:10.5px; letter-spacing:.16em; text-transform:uppercase; color:var(--mut2); }
.ax-scroll { flex:1; overflow-y:auto; padding:24px 32px 60px; }
.stack > * + * { margin-top:22px; }

/* section blocks */
.block-head { display:flex; align-items:center; justify-content:space-between; margin-bottom:12px; gap:12px; }
.section-label { font-size:11px; font-weight:700; letter-spacing:.14em; text-transform:uppercase; color:var(--ac2); }
.count { font-size:11.5px; color:var(--mut2); }
.panel { border:1px solid var(--line); border-radius:16px; background:linear-gradient(180deg,var(--card1),var(--card2)); padding:18px; }

/* controls */
.ctl-row { display:flex; align-items:center; gap:10px; flex-wrap:wrap; margin-bottom:14px; }
.ctl-row .spacer { flex:1; }
.file-meta { font-size:11.5px; color:var(--mut); word-break:break-all; }
.btn { font-family:var(--sans); font-size:12.5px; font-weight:600; border-radius:9px; padding:9px 15px; border:1px solid transparent; cursor:pointer; display:inline-flex; align-items:center; gap:7px; line-height:1; transition:background .15s,border-color .15s,color .15s,filter .15s; }
.btn:focus-visible { outline:2px solid var(--ac); outline-offset:2px; }
.btn-primary { background:linear-gradient(135deg,var(--ac2),var(--ac)); color:var(--acink); box-shadow:0 5px 16px rgba(123,163,207,0.28); }
.btn-primary:hover { filter:brightness(1.06); } .btn-primary:disabled { background:#26364d; color:#6a7688; box-shadow:none; cursor:not-allowed; }
.btn-ghost { background:#131a2a; color:var(--tx2); border-color:rgba(255,255,255,0.1); }
.btn-ghost:hover { border-color:rgba(123,163,207,0.5); } .btn-ghost:disabled { color:#525b6b; cursor:not-allowed; border-color:rgba(255,255,255,0.06); }
.btn-ghost.on { background:var(--amber-soft); border-color:rgba(224,164,74,0.4); color:var(--amber2); }

/* dropdown "Tải Ảnh" */
.dl-menu-wrap { position:relative; display:inline-flex; }
.dl-menu { position:absolute; top:calc(100% + 6px); left:0; z-index:30; min-width:210px; padding:5px;
  background:var(--panel2); border:1px solid var(--line2); border-radius:11px; box-shadow:0 10px 28px rgba(0,0,0,0.5); }
.dl-menu.dl-menu-up { top:auto; bottom:calc(100% + 6px); box-shadow:0 -10px 28px rgba(0,0,0,0.5); }
.dl-row { margin-top:14px; margin-bottom:0; }
.dl-menu-item { width:100%; display:flex; align-items:center; gap:9px; text-align:left; font-family:var(--sans); font-size:12.5px; font-weight:600;
  color:var(--tx2); background:transparent; border:0; border-radius:8px; padding:9px 11px; cursor:pointer; line-height:1.2; transition:background .12s,color .12s; }
.dl-menu-item:hover { background:rgba(123,163,207,0.14); color:#fff; }
.dl-menu-item:disabled { color:#525b6b; cursor:not-allowed; }
.dl-menu-item:disabled:hover { background:transparent; color:#525b6b; }

/* image */
.imgwrap { position:relative; border:1px solid var(--line2); border-radius:14px; overflow:hidden; background:#0c1119; user-select:none; touch-action:none; }
.imgwrap.edit { cursor:crosshair; }
.imgwrap img.base { display:block; width:100%; height:auto; pointer-events:none; }
.img-placeholder { display:flex; flex-direction:column; align-items:center; justify-content:center; gap:10px; min-height:260px; color:var(--mut2);
  font-size:12.5px; border:1.5px dashed rgba(123,163,207,0.32); border-radius:14px; background:var(--panel); cursor:pointer; transition:border-color .15s, background .15s; }
.img-placeholder:hover { border-color:rgba(123,163,207,0.6); background:#0f1626; }
.img-placeholder.dragover { border-color:var(--ac2); background:rgba(123,163,207,0.12); }
.img-placeholder .ph-title { color:var(--tx3); font-size:13px; }
.img-placeholder .ph-sub { color:var(--mut2); font-size:11px; letter-spacing:.02em; }
.img-placeholder .ph-btn { margin-top:2px; }
.imgwrap.dragover { outline:2px dashed var(--ac2); outline-offset:-6px; }

/* viền vùng crop khi chọn ký hiệu — thân trong suốt sự kiện, chỉ handle nhận kéo */
.cropbox { position:absolute; border:1px dashed rgba(157,192,230,0.45); background:rgba(123,163,207,0.05); border-radius:3px;
  pointer-events:none; z-index:3; }
.cropbox.unlocked { border:1.5px dashed var(--ac3,#9dc0e6); background:rgba(123,163,207,0.10); box-shadow:0 0 0 1px rgba(157,192,230,0.25) inset;
  pointer-events:auto; cursor:grab; touch-action:none; }
.cropbox.unlocked:active { cursor:grabbing; }
/* nút Xong (✓) & Xóa (✗) đặt NGOÀI vùng crop — góc trên-phải */
.crop-tools { position:absolute; right:0; top:0; transform:translate(0,-118%); display:flex; gap:6px; pointer-events:auto; z-index:8; }
.crop-tick, .crop-del { width:20px; height:20px; border-radius:50%; border:none; display:flex; align-items:center; justify-content:center;
  cursor:pointer; padding:0; box-shadow:0 2px 6px rgba(0,0,0,.5); touch-action:none; }
.crop-tick { background:var(--green); color:#0c1a12; }
.crop-del { background:#e08a8a; color:#2a0c0c; }
.crop-tick:hover, .crop-del:hover { filter:brightness(1.08); }
.crop-h { position:absolute; width:6px; height:6px; background:rgba(157,192,230,0.85); border:1px solid rgba(12,21,36,0.85); border-radius:2px;
  pointer-events:auto; touch-action:none; z-index:4; box-shadow:0 1px 2px rgba(0,0,0,.4); }
.crop-h::before { content:""; position:absolute; inset:-8px; } /* mở rộng vùng bắt kéo mà không phình phần nhìn thấy */
.crop-h:hover { background:var(--ac3); }
.marker-done { position:absolute; left:100%; top:100%; transform:translate(-35%,-35%); width:18px; height:18px; border-radius:50%;
  background:var(--green); color:#0c1a12; border:none; display:flex; align-items:center; justify-content:center;
  pointer-events:auto; cursor:pointer; z-index:6; box-shadow:0 2px 6px rgba(0,0,0,.5); padding:0; }
.marker-done:hover { filter:brightness(1.08); }
.h-nw { left:0; top:0; transform:translate(-50%,-50%); cursor:nwse-resize; }
.h-ne { right:0; top:0; transform:translate(50%,-50%); cursor:nesw-resize; }
.h-se { right:0; bottom:0; transform:translate(50%,50%); cursor:nwse-resize; }
.h-sw { left:0; bottom:0; transform:translate(-50%,50%); cursor:nesw-resize; }
.h-n { left:50%; top:0; transform:translate(-50%,-50%); cursor:ns-resize; }
.h-s { left:50%; bottom:0; transform:translate(-50%,50%); cursor:ns-resize; }
.h-w { left:0; top:50%; transform:translate(-50%,-50%); cursor:ew-resize; }
.h-e { right:0; top:50%; transform:translate(50%,-50%); cursor:ew-resize; }
.img-placeholder .ico { width:52px; height:52px; border-radius:14px; background:rgba(123,163,207,0.12); display:flex; align-items:center; justify-content:center; color:var(--ac2); }
.marker { position:absolute; transform:translate(-50%,-50%); min-width:24px; height:24px; padding:0 6px; border-radius:12px; background:rgba(255,255,255,0.10); color:#fff;
  font-size:11px; font-weight:700; display:flex; align-items:center; justify-content:center; border:1.5px solid rgba(255,255,255,0.35); backdrop-filter:blur(5px); -webkit-backdrop-filter:blur(5px); box-shadow:0 2px 8px rgba(0,0,0,.5); cursor:pointer; line-height:1; touch-action:none; opacity:.75; }
.marker.dim { opacity:.42; } .marker.active { background:rgba(255,255,255,0.12); color:#fff; border:1.5px solid rgba(127,216,171,0.5); transform:translate(-50%,-50%) scale(1.15); z-index:6; opacity:1; box-shadow:0 2px 8px rgba(0,0,0,.5); }
.marker.active::before { content:""; position:absolute; inset:-4px; border-radius:999px; pointer-events:none; padding:2.5px;
  background:conic-gradient(from 0deg, rgba(127,216,171,0) 0deg, rgba(127,216,171,0.12) 140deg, #7fd8ab 275deg, #d8ffe9 330deg, #7fd8ab 360deg);
  -webkit-mask:linear-gradient(#000 0 0) content-box, linear-gradient(#000 0 0); -webkit-mask-composite:xor;
  mask:linear-gradient(#000 0 0) content-box, linear-gradient(#000 0 0); mask-composite:exclude;
  animation:mk-ring-spin 1.15s linear infinite; }
@keyframes mk-ring-spin { to { transform:rotate(360deg); } }
@media (prefers-reduced-motion: reduce) { .marker.active::before { animation:none; background:#7fd8ab; } }
.marker.cropping { cursor:grab; background:rgba(255,255,255,0.14); color:#fff; border-color:rgba(255,255,255,0.6); opacity:.72; transform:translate(-50%,-50%);
  box-shadow:0 0 0 2px rgba(157,192,230,0.55), 0 2px 8px rgba(0,0,0,.45); z-index:7; }
.marker.cropping:active { cursor:grabbing; }
.marker.cropping.hl { opacity:1; }
.marker-del { position:absolute; left:100%; top:0; transform:translate(-35%,-60%); width:18px; height:18px; border-radius:50%;
  background:#e08a8a; color:#2a0c0c; border:none; display:flex; align-items:center; justify-content:center;
  pointer-events:auto; cursor:pointer; z-index:6; box-shadow:0 2px 6px rgba(0,0,0,.5); padding:0; }
.marker-del:hover { filter:brightness(1.08); }
.imgwrap.edit .marker { cursor:grab; }
.hint { font-size:11.5px; color:var(--mut); margin-top:11px; line-height:1.5; } .hint.edit-on { color:var(--amber2); }

/* schedule table */
.sheet-tabs { display:flex; flex-wrap:wrap; gap:6px; margin:14px 0 12px; padding:0; }
.sheet-tab { display:inline-flex; align-items:center; gap:6px; padding:6px 12px; font-size:12px; font-weight:600; color:var(--tx3);
  background:#0f1626; border:1px solid rgba(255,255,255,0.08); border-radius:9px; cursor:pointer; transition:background .15s,color .15s,border-color .15s; }
.sheet-tab:hover { color:var(--tx); border-color:rgba(255,255,255,0.18); }
.sheet-tab.on { background:rgba(123,163,207,0.14); color:var(--ac3); border-color:rgba(123,163,207,0.5); }
.sheet-tab .tab-n { font-size:10.5px; font-weight:700; color:var(--mut2); background:rgba(255,255,255,0.06); border-radius:20px; padding:1px 7px; }
.sheet-tab.on .tab-n { color:var(--ac2); }
.img-filter { display:flex; flex-wrap:wrap; align-items:center; gap:6px; margin:0 0 12px; }
.imgf-label { font-size:11px; font-weight:600; color:var(--mut2); margin-right:2px; }
.imgf-chip { display:inline-flex; align-items:center; gap:6px; padding:5px 11px; font-size:11.5px; font-weight:600; color:var(--tx3);
  background:#0f1626; border:1px solid rgba(255,255,255,0.08); border-radius:20px; cursor:pointer; transition:background .15s,color .15s,border-color .15s; }
.imgf-chip:hover { color:var(--tx); border-color:rgba(255,255,255,0.18); }
.imgf-chip.on { background:rgba(123,163,207,0.16); color:var(--ac3); border-color:rgba(123,163,207,0.55); }
.imgf-chip .tab-n { font-size:10px; font-weight:700; color:var(--mut2); background:rgba(255,255,255,0.06); border-radius:20px; padding:1px 6px; }
.imgf-chip.on .tab-n { color:var(--ac2); }
.sched-wrap { border:1px solid var(--line); border-radius:14px; background:var(--panel2); overflow:hidden; }
.sched-scroll { overflow-x:auto; }
table.sched { border-collapse:collapse; width:100%; min-width:1000px; font-family:var(--sans); font-size:12.5px; }
.sched thead th { text-align:left; font-weight:700; font-size:9.5px; letter-spacing:.08em; text-transform:uppercase; color:var(--mut2); background:#0e1526;
  padding:11px 11px; border-bottom:1px solid var(--line2); border-right:1px solid var(--line); white-space:nowrap; }
.sched thead th:last-child { border-right:none; }
.sched tbody td { padding:0; border-bottom:1px solid var(--line); border-right:1px solid var(--line); vertical-align:middle; }
.sched tbody td:last-child { border-right:none; }
.sched tbody tr:last-child td { border-bottom:none; }
.sched tbody tr { transition:background .12s; }
.sched tbody tr.row-low td:first-child { box-shadow:inset 3px 0 0 var(--amber); }
.sched tbody tr.row-low td { background:rgba(224,164,74,0.06); }
.sched tbody tr.active-row td { background:rgba(123,163,207,0.12); }
.cell-input, .cell-select { width:100%; border:none; background:transparent; font-family:var(--sans); font-size:12.5px; color:var(--tx2); padding:9px 10px; }
.cell-input::placeholder { color:var(--faint); }
.cell-input:focus, .cell-select:focus { outline:none; background:var(--input); box-shadow:inset 0 0 0 2px var(--ac); border-radius:6px; }
.cell-select { appearance:none; -webkit-appearance:none; cursor:pointer; }
.cell-select option { background:#101725; color:var(--tx2); }
.cell-code { font-weight:700; color:var(--ac3); letter-spacing:.03em; }
.col-stt { width:44px; } .stt-cell { text-align:center; color:var(--mut2); font-size:12.5px; padding:9px 4px; }
.col-thumb { width:58px; } .thumb-cell { padding:5px; text-align:center; }
.thumb-cell img { width:44px; height:44px; object-fit:cover; border-radius:7px; border:1px solid var(--line2); display:inline-block; vertical-align:middle; }
.thumb-ph { width:44px; height:44px; border-radius:7px; border:1px dashed var(--line2); background:var(--input); display:inline-block; }
.col-code { width:82px; }
.qty-cell { text-align:center; font-weight:700; color:var(--tx); padding:9px 6px; }
.col-act { width:44px; } td.act { text-align:center; }
.icon-danger { background:transparent; border:none; color:var(--mut2); cursor:pointer; padding:7px; border-radius:7px; display:inline-flex; }
.icon-danger:hover { color:#e08a8a; background:rgba(214,79,79,0.12); }
.grp-row .icon-danger { flex:0 0 30px; width:30px; justify-content:center; align-items:center; padding:7px 0; }
.grp-row .axchk { margin:0; flex:0 0 auto; }
.toolbar { display:flex; align-items:center; gap:10px; flex-wrap:wrap; margin-top:14px; } .toolbar .spacer { flex:1; }

.note { display:flex; gap:10px; align-items:flex-start; background:var(--amber-soft); border:1px solid rgba(224,164,74,0.28); color:var(--amber2); border-radius:14px; padding:14px 16px; font-size:12px; line-height:1.55; }
.note svg { flex:0 0 auto; margin-top:1px; }
.note b { color:#f0cf9a; }
.status { font-size:12px; color:var(--ac2); margin-top:10px; display:flex; align-items:center; gap:7px; }
.error { margin:0; padding:12px 15px; background:rgba(214,79,79,0.1); border:1px solid rgba(214,79,79,0.35); color:#e79b9b; border-radius:12px; font-size:12.5px; }
.empty { padding:44px 20px; text-align:center; } .empty .eyebrow { font-size:11px; letter-spacing:.12em; text-transform:uppercase; color:var(--mut2); } .empty .msg { color:var(--mut2); font-size:12.5px; margin-top:8px; line-height:1.55; }
.spin { animation:axspin 1s linear infinite; } @keyframes axspin { to { transform:rotate(360deg); } }
@media (prefers-reduced-motion: reduce) { .spin { animation:none; } }
`;

/* ---------- CSS bổ sung: dải nhiều ảnh (dùng lại token ARTUS) ---------- */
const cssExtra = `
.imgcount { font-size:11.5px; color:var(--mut2); }
.imgstrip { display:flex; gap:10px; flex-wrap:wrap; margin-bottom:14px; }
.imgtile { position:relative; width:94px; height:70px; border-radius:10px; overflow:hidden; border:1px solid var(--line2);
  background:#0c1119; cursor:pointer; flex:0 0 auto; transition:border-color .15s, box-shadow .15s; }
.imgtile img { width:100%; height:100%; object-fit:cover; display:block; }
.imgtile:hover { border-color:rgba(123,163,207,0.55); }
.imgtile.on { border-color:var(--ac2); box-shadow:0 0 0 2px rgba(123,163,207,0.45); }
.imgtile .idx { position:absolute; left:5px; top:5px; min-width:18px; height:18px; padding:0 5px; border-radius:9px;
  background:rgba(12,21,36,0.85); color:var(--ac3); font-size:10.5px; font-weight:700; display:flex; align-items:center; justify-content:center; }
.imgtile .st { position:absolute; right:5px; top:5px; width:10px; height:10px; border-radius:50%; border:1.5px solid #0c1524; }
.imgtile .st.idle { background:var(--mut2); }
.imgtile .st.analyzing { background:var(--amber); animation:axspin 1s linear infinite; }
.imgtile .st.done { background:var(--green); }
.imgtile .st.error { background:#d64f4f; }
.imgtile .cnt { position:absolute; left:5px; bottom:5px; padding:1px 6px; border-radius:8px; background:rgba(12,21,36,0.85);
  color:var(--tx3); font-size:9.5px; font-weight:600; }
.imgtile .rm { position:absolute; right:4px; bottom:4px; width:20px; height:20px; border-radius:6px; border:none;
  background:rgba(12,21,36,0.82); color:#e08a8a; display:flex; align-items:center; justify-content:center; cursor:pointer; opacity:0; transition:opacity .15s; }
.imgtile:hover .rm { opacity:1; }
.imgtile .rm:hover { background:rgba(214,79,79,0.28); }
.imgadd { width:94px; height:70px; border-radius:10px; border:1.5px dashed rgba(123,163,207,0.4); background:var(--panel);
  color:var(--ac2); display:flex; flex-direction:column; align-items:center; justify-content:center; gap:3px; cursor:pointer; flex:0 0 auto;
  font-size:10.5px; font-weight:600; transition:border-color .15s, background .15s; }
.imgadd:hover { border-color:var(--ac2); background:#0f1626; }

/* C1 — hover đồng bộ hàng <-> ký hiệu (đặt SAU .marker.active để thắng khi trùng) */
.marker.hl { opacity:1; transform:translate(-50%,-50%) scale(1.3); z-index:60; box-shadow:0 0 0 3px rgba(157,192,230,0.65), 0 3px 12px rgba(0,0,0,0.55); }
.sched tbody tr.hl-row td { background:rgba(157,192,230,0.16) !important; }

/* C2 — thanh lọc + tìm */
.filterbar { display:flex; align-items:center; gap:10px; flex-wrap:wrap; margin:2px 0 10px; }
.searchbox { position:relative; display:inline-flex; align-items:center; flex:0 1 300px; }
.searchbox svg { position:absolute; left:10px; color:var(--mut2); pointer-events:none; }
.searchbox input { width:100%; background:var(--input); border:1px solid rgba(255,255,255,0.1); border-radius:9px; padding:8px 30px 8px 32px; color:var(--tx2); font-family:var(--sans); font-size:12.5px; outline:none; transition:border-color .15s; }
.searchbox input:focus { border-color:rgba(123,163,207,0.5); }
.searchbox .clr { position:absolute; right:6px; width:22px; height:22px; border:none; background:transparent; color:var(--mut2); cursor:pointer; display:flex; align-items:center; justify-content:center; border-radius:6px; }
.searchbox .clr:hover { color:var(--tx2); background:rgba(255,255,255,0.06); }
.chip-toggle { display:inline-flex; align-items:center; gap:7px; padding:7px 12px; font-size:12px; font-weight:600; color:var(--tx3); background:#131a2a; border:1px solid rgba(255,255,255,0.1); border-radius:9px; cursor:pointer; transition:all .15s; }
.chip-toggle:hover { color:var(--tx); border-color:rgba(255,255,255,0.18); }
.chip-toggle.on { background:var(--amber-soft); border-color:rgba(224,164,74,0.45); color:var(--amber2); }
.chip-toggle .dotc { width:8px; height:8px; border-radius:50%; background:var(--amber); }
.filter-note { font-size:11px; color:var(--mut2); margin-left:auto; }

/* B1 — cột chọn + thanh hành động chọn */
.col-sel { width:34px; } .sel-cell { text-align:center; padding:0 4px; }
.axchk { width:15px; height:15px; accent-color:var(--ac); cursor:pointer; vertical-align:middle; }
.selbar { display:flex; align-items:center; gap:10px; flex-wrap:wrap; padding:9px 12px; margin:2px 0 10px; border:1px solid rgba(123,163,207,0.3);
  background:var(--acsoft); border-radius:11px; }
.selbar .seln { font-size:12.5px; font-weight:700; color:var(--ac3); }
.selbar .spacer { flex:1; }

/* C3 — lightbox soi crop */
.lb-backdrop { position:fixed; inset:0; z-index:9000; background:rgba(5,8,14,0.82); backdrop-filter:blur(3px); display:flex; align-items:center; justify-content:center; padding:24px; }
.lb-card { background:linear-gradient(180deg,var(--card1),var(--card2)); border:1px solid var(--line2); border-radius:16px; max-width:min(920px,94vw); max-height:92vh; overflow:auto; box-shadow:0 20px 60px rgba(0,0,0,0.6); }
.lb-head { display:flex; align-items:center; gap:10px; padding:14px 16px; border-bottom:1px solid var(--line); position:sticky; top:0; background:var(--card1); }
.lb-code { font-family:var(--serif); font-size:15px; letter-spacing:.05em; color:var(--ac3); }
.lb-title { font-size:13px; color:var(--tx2); }
.lb-close { margin-left:auto; width:30px; height:30px; border:none; border-radius:8px; background:#131a2a; color:var(--tx3); cursor:pointer; display:flex; align-items:center; justify-content:center; }
.lb-close:hover { color:var(--tx); }
.lb-body { padding:16px; display:flex; flex-direction:column; gap:12px; align-items:center; }
.lb-body img { max-width:100%; border-radius:10px; border:1px solid var(--line2); background:#0c1119; }
.lb-meta { font-size:11.5px; color:var(--mut2); text-align:center; line-height:1.6; }

/* A+B — dòng chưa gắn ký hiệu (SL=0) */
.pin-btn { display:inline-flex; flex-direction:column; align-items:center; justify-content:center; gap:3px; width:72px; height:72px; flex:0 0 auto; padding:6px 6px; font-size:9px; font-weight:700; color:var(--ac2);
  background:transparent; border:1.5px dashed rgba(123,163,207,0.5); border-radius:8px; cursor:pointer; white-space:nowrap; line-height:1.15; text-align:center; }
.pin-btn:hover { background:var(--acsoft); border-color:var(--ac2); color:var(--ac3); }
.qty-cell.qty-zero { color:var(--amber2); font-weight:700; }
.sched tbody tr.row-unpinned td.stt-cell { box-shadow:inset 3px 0 0 var(--ac); }

/* ===== Bố cục CHIA ĐÔI + banner Thông tin dự án trên cùng ===== */
.ax2 { flex-direction:column; height:100vh; min-height:100vh; overflow:hidden; }

.topbar { flex:0 0 auto; background:linear-gradient(180deg,var(--side1),var(--side2)); border-bottom:1px solid var(--line); }
.topbar-main { display:flex; align-items:center; gap:18px; padding:11px 22px; flex-wrap:wrap; }
.tb-brand { display:flex; align-items:center; gap:14px; }
.tb-word { font-family:var(--serif); font-size:21px; letter-spacing:.4em; color:var(--tx); padding-left:.4em; line-height:1; }
.tb-pill { font-size:8.5px; letter-spacing:.24em; color:var(--ac); border:1px solid rgba(123,163,207,0.32); border-radius:999px; padding:4px 11px; white-space:nowrap; }
.tb-stats { display:flex; align-items:stretch; gap:8px; flex-wrap:wrap; }
.tb-stat { background:var(--panel2); border:1px solid rgba(255,255,255,0.06); border-radius:10px; padding:6px 12px; text-align:center; min-width:56px; }
.tb-stat b { font-size:17px; font-weight:700; display:block; line-height:1.1; color:var(--tx2); }
.tb-stat span { font-size:8px; letter-spacing:.04em; color:var(--faint); display:block; margin-top:2px; white-space:nowrap; }
.tb-stat.warn b { color:var(--amber2); }
.tb-stat.info b { color:var(--ac2); }
.tb-toggle { margin-left:auto; }
.tb-fields { display:flex; gap:12px; flex-wrap:wrap; margin-top:8px; padding:16px 22px 16px; border-top:1px solid rgba(255,255,255,0.05); }
.tb-field { flex:1 1 170px; min-width:150px; }
.tb-field label { font-size:9px; letter-spacing:.1em; text-transform:uppercase; color:var(--faint); margin-bottom:4px; display:block; }
.tb-field input { width:100%; background:var(--input); border:1px solid rgba(255,255,255,0.09); border-radius:9px; padding:8px 11px; color:var(--tx2); font-family:var(--sans); font-size:12.5px; outline:none; transition:border-color .15s; }
.tb-field input:focus { border-color:rgba(123,163,207,0.55); }
.tb-field input::placeholder { color:var(--faint); }

.ax-split { flex:1 1 auto; min-height:0; display:flex; align-items:stretch; }
.pane { display:flex; flex-direction:column; min-width:0; min-height:0; }
.pane-image { flex:0 0 var(--splitpct, 45%); order:1; min-width:0; }
.split-divider { order:2; flex:0 0 7px; align-self:stretch; cursor:col-resize; position:relative; background:transparent; }
.split-divider::before { content:""; position:absolute; left:2px; top:0; bottom:0; width:2px; background:var(--line); border-radius:2px; transition:background .15s, box-shadow .15s; }
.split-divider::after { content:""; position:absolute; left:50%; top:50%; transform:translate(-50%,-50%); width:4px; height:34px; border-radius:3px; background:rgba(123,163,207,0.28); }
.split-divider:hover::before, .split-divider.dragging::before { background:var(--ac); box-shadow:0 0 0 1px rgba(123,163,207,0.35); }
.pane-table { flex:1 1 0; order:3; }
.pane-body { flex:1 1 auto; min-height:0; overflow-y:auto; padding:18px 22px 44px; }

@media (max-width: 1080px) {
  .ax2 { height:auto; overflow:visible; }
  .ax-split { flex-direction:column; }
  .pane-table, .pane-image { flex:1 1 auto; order:0; border-right:none; }
  .pane-table { border-bottom:1px solid var(--line); }
  .split-divider { display:none; }
  .pane-body { overflow:visible; }
}

/* Scrollbar đồng bộ theme (WebKit + Firefox) cho bảng Inventory & các vùng cuộn */
.pane-body, .sched-scroll, .lb-card { scrollbar-width:thin; scrollbar-color:rgba(123,163,207,0.45) var(--input); }
.pane-body::-webkit-scrollbar, .sched-scroll::-webkit-scrollbar, .lb-card::-webkit-scrollbar { width:11px; height:11px; }
.pane-body::-webkit-scrollbar-track, .sched-scroll::-webkit-scrollbar-track, .lb-card::-webkit-scrollbar-track { background:var(--input); border-radius:8px; }
.pane-body::-webkit-scrollbar-thumb, .sched-scroll::-webkit-scrollbar-thumb, .lb-card::-webkit-scrollbar-thumb {
  background-color:rgba(123,163,207,0.4); border-radius:8px; border:2px solid var(--input); }
.pane-body::-webkit-scrollbar-thumb:hover, .sched-scroll::-webkit-scrollbar-thumb:hover, .lb-card::-webkit-scrollbar-thumb:hover { background-color:var(--ac); }
.pane-body::-webkit-scrollbar-thumb:active, .sched-scroll::-webkit-scrollbar-thumb:active, .lb-card::-webkit-scrollbar-thumb:active { background-color:var(--ac2); }
.pane-body::-webkit-scrollbar-corner, .sched-scroll::-webkit-scrollbar-corner, .lb-card::-webkit-scrollbar-corner { background:var(--input); }

/* Logo ARTIUS trong banner */
.tb-logo { height:44px; width:auto; display:block; }

/* ===== Bảng vật liệu GỘP THEO NHÓM (thay cho <table class="sched">) ===== */
.grp-wrap { border:1px solid var(--line); border-radius:14px; background:var(--panel2); overflow:hidden; }
.grp-head { padding:7px 12px; background:#0e1526; font-size:9px; font-weight:700; letter-spacing:.08em; text-transform:uppercase; border-bottom:1px solid var(--line2); display:flex; align-items:center; gap:7px; position:sticky; top:0; z-index:2; }
.grp-dot { width:8px; height:8px; border-radius:2px; display:inline-block; flex:0 0 auto; }
.grp-row { display:flex; align-items:center; gap:9px; padding:7px 12px; border-bottom:1px solid var(--line); border-left:3px solid transparent; cursor:pointer; transition:background .12s; }
.grp-row.active-row { background:rgba(123,163,207,0.12); }
.grp-row.hl-row { background:rgba(157,192,230,0.16) !important; }
.grp-row.row-low { background:rgba(224,164,74,0.06); }
.grp-thumb { width:72px; height:72px; object-fit:cover; border-radius:8px; border:1px solid var(--line2); cursor:zoom-in; flex:0 0 auto; }
.grp-thumb-ph { width:72px; height:72px; border-radius:8px; border:1px dashed var(--line2); background:var(--input); display:inline-block; flex:0 0 auto; }
.grp-input { background:transparent; border:none; font-family:var(--sans); color:var(--tx2); outline:none; }
.grp-input::placeholder { color:var(--faint); }
.grp-stt { width:26px; flex:0 0 auto; text-align:center; font-size:11px; font-weight:700; color:var(--mut2); }
.grp-cap-stt { width:26px; flex:0 0 auto; text-align:center; }
.grp-code { width:60px; flex:0 0 auto; font-weight:700; color:var(--ac3); font-size:12px; padding:5px 0; }
.grp-main { flex:1; min-width:120px; display:flex; flex-direction:column; }
.grp-mon { font-size:12px; padding:2px 0; }
.grp-vl { font-size:10px; color:var(--mut2); padding:1px 0; }
.grp-vitri { width:110px; flex:0 0 auto; font-size:11px; color:var(--tx3); padding:5px 4px; }
.grp-sl { width:34px; flex:0 0 auto; text-align:center; font-size:11px; font-weight:700; color:var(--tx2); padding:4px 0; -moz-appearance:textfield; }
.grp-sl::-webkit-outer-spin-button, .grp-sl::-webkit-inner-spin-button { -webkit-appearance:none; margin:0; }
.grp-sl.qty-zero { color:var(--amber2); }
.grp-select { width:104px; flex:0 0 auto; background:transparent; border:none; font-family:var(--sans); color:var(--tx3); font-size:11px; padding:5px 0; cursor:pointer; }
.grp-select option { background:#101725; color:var(--tx2); }
.grp-note { width:110px; flex:0 0 auto; font-size:10.5px; color:var(--faint); padding:5px 4px; }
.grp-caption { display:flex; align-items:center; gap:9px; padding:6px 12px; background:#0e1526; border-bottom:1px solid var(--line2); font-size:9px; font-weight:700; letter-spacing:.06em; text-transform:uppercase; color:var(--mut2); position:sticky; top:0; z-index:3; }
.grp-cap-sel { width:15px; flex:0 0 auto; }
.grp-cap-thumb { width:72px; flex:0 0 auto; }
.grp-cap-code { width:60px; flex:0 0 auto; }
.grp-cap-main { flex:1; min-width:120px; }
.grp-cap-vitri { width:110px; flex:0 0 auto; }
.grp-cap-sl { width:34px; flex:0 0 auto; text-align:center; }
.grp-cap-select { width:104px; flex:0 0 auto; }
.grp-cap-act { width:30px; flex:0 0 auto; }
`;

// Lấy tiền tố nhóm mã của 1 dòng: ưu tiên phần chữ trong "ma" (vd "F-03" -> "F"),
// nếu chưa có mã thì dùng "prefix" do AI trả về. Dùng chung cho merge + đánh mã.
const gpOf = (r) => ((r.ma && String(r.ma).split("-")[0]) || r.prefix || "");

// Số lượng để hiển thị / xuất file: ưu tiên so_luong do AI đếm (đã cho phép sửa tay);
// nếu để trống thì lấy theo số ký hiệu (box) đã gắn.
const qtyOf = (r) => {
  if (r && r.soLuong != null && String(r.soLuong).trim() !== "") { const n = parseInt(r.soLuong, 10); return Number.isFinite(n) ? n : (r.instances ? r.instances.length : 0); }
  return r && r.instances ? r.instances.length : 0;
};

// Gỡ mọi bounding box thuộc ảnh imgId ra khỏi các dòng.
// Dùng khi phân tích LẠI 1 ảnh (tránh nhân đôi box) hoặc khi gỡ hẳn 1 ảnh.
// - Dòng chỉ tồn tại nhờ ảnh này (giờ rỗng) => loại bỏ.
// - Dòng nhập tay chưa có box, hoặc còn box từ ảnh khác => giữ lại.
function stripImageInstances(rows, imgId) {
  const out = [];
  for (const r of rows) {
    const kept = r.instances.filter((b) => b.imgId !== imgId);
    const hadFromImg = kept.length !== r.instances.length;
    if (kept.length === 0 && hadFromImg) continue;
    out.push({ ...r, instances: kept });
  }
  return out;
}

// Màu dải trái theo Nhóm — dùng cho bảng gộp nhóm (grp-head / grp-row).
const GROUP_COLOR = { "Nội thất": "#7ba3cf", "Đèn": "#e0a44a", "Vật liệu bề mặt": "#aab4c4", "Cửa & Vách kính": "#bcd6f0", "Hardware": "#6f7889", "Trang trí": "#7fd8ab" };

// Gom danh sách dòng (đã lọc) thành các nhóm theo NHOM_OPTS (giữ thứ tự cố định),
// nhóm lạ (không nằm trong NHOM_OPTS) được xếp cuối. Chỉ trả về nhóm có dòng.
function groupRowsByNhom(rows) {
  const extra = Array.from(new Set(rows.map((r) => r.nhom).filter((n) => n && !NHOM_OPTS.includes(n))));
  const order = [...NHOM_OPTS, ...extra];
  return order.map((k) => ({ key: k, rows: rows.filter((r) => r.nhom === k) })).filter((g) => g.rows.length > 0);
}

// Đánh SỐ ký hiệu theo ĐÚNG thứ tự hiển thị trong bảng: gom theo NHÓM (hạng mục, thứ tự NHOM_OPTS),
// trong mỗi nhóm giữ nguyên thứ tự dòng (đã sort theo mã tăng dần). Nhờ vậy số ký hiệu trên ảnh &
// STT trong bảng luôn TĂNG DẦN theo loại hạng mục, không lộn xộn. Trả về Map: rowId -> số thứ tự (1..N).
function buildDisplayNo(rows) {
  const m = new Map(); let n = 0;
  groupRowsByNhom(rows).forEach((g) => g.rows.forEach((r) => { n += 1; m.set(r.id, n); }));
  return m;
}

function InventoryExtractor() {
  const [rows, setRows] = useState([]);
  // Danh sách nhiều ảnh. Mỗi ảnh: { id, preview(dataURL), imgData(base64), mediaType, fileName, status }
  // status: "idle" | "analyzing" | "done" | "error"
  const [images, setImages] = useState([]);
  const [activeImgId, setActiveImgId] = useState(null);
  const [readyTick, setReadyTick] = useState(0); // bump khi 1 ảnh offscreen load xong -> tính lại thumbnail còn thiếu

  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [status, setStatus] = useState(null);
  const [activeId, setActiveId] = useState(null);
  const [markerEdit, setMarkerEdit] = useState(false);
  const [dlMenuOpen, setDlMenuOpen] = useState(false);   // dropdown nút "Tải Ảnh" (1 ảnh / tất cả)
  const [cropRowId, setCropRowId] = useState(null);      // dòng đang MỞ KHOÁ kéo vùng crop qua double-click ký hiệu
  const [dispSize, setDispSize] = useState({ w: 0, h: 0 });
  const [sideCollapsed, setSideCollapsed] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  const [activeSheet, setActiveSheet] = useState("__all__");

  const [selected, setSelected] = useState(() => new Set()); // B1: id các dòng đang tick
  const [hoverId, setHoverId] = useState(null);              // C1: dòng đang rê chuột
  const [search, setSearch] = useState("");                  // C2: tìm nhanh
  const [onlyLow, setOnlyLow] = useState(false);             // C2: chỉ hiện tin cậy Thấp
  const [onlyUnpinned, setOnlyUnpinned] = useState(false);   // A: chỉ hiện dòng chưa gắn ký hiệu (SL=0)
  const [imgFilter, setImgFilter] = useState(null);          // lọc bảng: null = tất cả ảnh; hoặc imgId của 1 ảnh phối cảnh
  const [undoStack, setUndoStack] = useState([]);            // B5: hoàn tác thao tác xoá/gộp/tách
  const [lightbox, setLightbox] = useState(null);            // C3: { code, title, src, meta }
  const [infoOpen, setInfoOpen] = useState(true);            // banner Thông tin dự án: hiện/ẩn
  const [splitPct, setSplitPct] = useState(45);               // % bề rộng cột Ảnh (trái); kéo divider để đổi
  const [splitDragging, setSplitDragging] = useState(false);

  const [projectName, setProjectName] = useState("");
  const [client, setClient] = useState("");
  const [location, setLocation] = useState("");
  const [author, setAuthor] = useState("");
  const [dateStr, setDateStr] = useState(new Date().toLocaleDateString("vi-VN"));

  const fileRef = useRef(null);
  const imgRef = useRef(null);   // <img> đang HIỂN THỊ (ảnh active) — chỉ dùng để lấy kích thước hiển thị
  const wrapRef = useRef(null);
  const imgElMap = useRef(new Map()); // id -> HTMLImageElement offscreen (đã decode) cho MỌI ảnh; dùng cắt thumbnail/crop
  const splitRef = useRef(null);      // container chia đôi, để tính tỉ lệ khi kéo divider
  const dlMenuRef = useRef(null);     // wrapper dropdown "Tải Ảnh" — để đóng khi click ra ngoài

  const activeImage = images.find((im) => im.id === activeImgId) || null;

  const getEl = (id) => imgElMap.current.get(id) || null;
  const elReady = (id) => { const e = getEl(id); return !!(e && e.naturalWidth); };
  const imgIndex = (id) => images.findIndex((x) => x.id === id);

  // Chuẩn bị ảnh để GỬI API: thu nhỏ cạnh dài về <=maxEdge và nén JPEG.
  // Lý do: ảnh render 2K-4K gửi thẳng base64 rất nặng -> API từ chối (HTTP 400/413) hoặc request fail.
  // Thu nhỏ về ~1568px là đủ để AI đọc vật liệu/nội thất mà payload giảm mạnh, đồng thời chuẩn hoá webp/png -> jpeg.
  function apiImageFor(imgId, maxEdge = 1568, quality = 0.85) {
    const im = images.find((x) => x.id === imgId);
    const el = getEl(imgId);
    if (el && el.naturalWidth) {
      try {
        const nw = el.naturalWidth, nh = el.naturalHeight;
        const scale = Math.min(1, maxEdge / Math.max(nw, nh));
        const dw = Math.max(1, Math.round(nw * scale)), dh = Math.max(1, Math.round(nh * scale));
        const c = document.createElement("canvas"); c.width = dw; c.height = dh;
        c.getContext("2d").drawImage(el, 0, 0, dw, dh);
        const url = c.toDataURL("image/jpeg", quality);
        return { media_type: "image/jpeg", data: url.substring(url.indexOf(",") + 1) };
      } catch (e) { /* nếu canvas lỗi thì rơi xuống dùng ảnh gốc */ }
    }
    return { media_type: (im && im.mediaType) || "image/jpeg", data: im ? im.imgData : "" };
  }

  // Tên file xuất: ưu tiên tên dự án, không thì tên ảnh active. (Trước đây hàm này thiếu -> lỗi khi export.)
  const safeName = () => {
    const base = String(projectName || (activeImage && activeImage.fileName) || "inventory");
    return base.replace(/\.[a-z0-9]+$/i, "").replace(/[\\/:*?"<>|]+/g, " ").trim().slice(0, 60) || "inventory";
  };

  // Kích thước vùng ảnh hiển thị (đổi khi chuyển ảnh vì tỉ lệ khác nhau)
  useEffect(() => {
    const el = wrapRef.current;
    if (!el) return;
    const update = () => setDispSize({ w: el.clientWidth, h: el.clientHeight });
    update();
    let ro;
    if (typeof ResizeObserver !== "undefined") { ro = new ResizeObserver(update); ro.observe(el); }
    window.addEventListener("resize", update);
    return () => { if (ro) ro.disconnect(); window.removeEventListener("resize", update); };
  }, [activeImgId]);

  // Đóng dropdown "Tải Ảnh" khi click ra ngoài hoặc nhấn Esc
  useEffect(() => {
    if (!dlMenuOpen) return;
    const onDown = (e) => { if (dlMenuRef.current && !dlMenuRef.current.contains(e.target)) setDlMenuOpen(false); };
    const onKey = (e) => { if (e.key === "Escape") setDlMenuOpen(false); };
    window.addEventListener("pointerdown", onDown);
    window.addEventListener("keydown", onKey);
    return () => { window.removeEventListener("pointerdown", onDown); window.removeEventListener("keydown", onKey); };
  }, [dlMenuOpen]);

  // Tính thumbnail cho các dòng còn thiếu, mỗi khi có ảnh offscreen vừa load xong
  useEffect(() => {
    setRows((rs) => {
      let changed = false;
      const next = rs.map((r) => {
        const b0 = r.instances[0];
        if (!r.thumb && b0 && elReady(b0.imgId)) {
          const t = makeThumb(getEl(b0.imgId), b0);
          if (t) { changed = true; return { ...r, thumb: t }; }
        }
        return r;
      });
      return changed ? next : rs;
    });
  }, [readyTick]);

  // Dán ảnh từ clipboard (screenshot hoặc file đã copy) — hỗ trợ dán nhiều ảnh cùng lúc
  useEffect(() => {
    function onPaste(e) {
      const dt = e.clipboardData; if (!dt) return;
      const list = [];
      if (dt.items) for (const it of Array.from(dt.items)) { if (it.type && it.type.indexOf("image") === 0) { const f = it.getAsFile(); if (f) list.push(f); } }
      if (!list.length && dt.files) for (const f of Array.from(dt.files)) { if (/^image\//.test(f.type)) list.push(f); }
      if (list.length) { addFiles(list); e.preventDefault(); }
    }
    window.addEventListener("paste", onPaste);
    return () => window.removeEventListener("paste", onPaste);
  }, []);

  // Số thứ tự ký hiệu theo thứ tự hiển thị (gom nhóm) — dùng cho marker trên ảnh & STT trong bảng.
  const displayNo = useMemo(() => buildDisplayNo(rows), [rows]);

  // Ký hiệu (marker) hiển thị trên ảnh — CHỈ những box thuộc ảnh đang xem
  const markerLayout = useMemo(() => {
    const flat = [];
    rows.forEach((r, idx) => r.instances.forEach((b, j) => {
      if (b.imgId !== activeImgId) return;
      flat.push({ rowId: r.id, rowIdx: idx, instIdx: j, cx: (b.x1 + b.x2) / 2, cy: (b.y1 + b.y2) / 2 });
    }));
    const W = dispSize.w || 1000, H = dispSize.h || 660;
    const pts = flat.map((f) => ({ x: f.cx * W, y: f.cy * H }));
    const sep = separate(pts, 26, W, H);
    return flat.map((f, i) => ({ ...f, leftPct: (sep[i].x / W) * 100, topPct: (sep[i].y / H) * 100 }));
  }, [rows, dispSize, activeImgId]);

  // Đọc 1 File -> tạo metadata ảnh + phần tử Image offscreen (để cắt thumbnail/crop mà không cần gắn vào DOM)
  function readImageFile(f) {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = () => {
        const url = String(reader.result);
        const semi = url.indexOf(";"), comma = url.indexOf(",");
        const meta = {
          id: nextId(),
          preview: url,
          imgData: url.substring(comma + 1),
          mediaType: url.substring(5, semi) || "image/jpeg",
          fileName: f.name || ("anh-dan-" + Date.now() + ".png"),
          status: "idle",
        };
        const el = new Image();
        el.onload = () => setReadyTick((t) => t + 1);
        el.src = url;
        resolve({ meta, el });
      };
      reader.onerror = () => resolve(null);
      reader.readAsDataURL(f);
    });
  }

  // Thêm 1..n ảnh vào danh sách (dùng chung: chọn file, kéo-thả, dán clipboard)
  async function addFiles(fileList) {
    const files = Array.from(fileList || []).filter((f) => /^image\//.test(f.type || ""));
    if (!files.length) { setError("Không có ảnh hợp lệ. Hãy chọn file ảnh (jpg, png, webp…)."); return; }
    setError(null);
    const loaded = (await Promise.all(files.map(readImageFile))).filter(Boolean);
    if (!loaded.length) return;
    loaded.forEach((l) => imgElMap.current.set(l.meta.id, l.el));
    setImages((prev) => [...prev, ...loaded.map((l) => l.meta)]);
    setActiveImgId((prev) => (prev == null ? loaded[0].meta.id : prev));
  }

  function onPickFiles(e) { addFiles(e.target.files); if (e.target) e.target.value = ""; }
  function openPicker() { if (fileRef.current) fileRef.current.click(); }

  function onDragOver(e) { e.preventDefault(); e.stopPropagation(); if (!dragOver) setDragOver(true); }
  function onDragLeave(e) { e.preventDefault(); e.stopPropagation(); setDragOver(false); }
  function onDrop(e) {
    e.preventDefault(); e.stopPropagation(); setDragOver(false);
    const fs = e.dataTransfer && e.dataTransfer.files;
    if (fs && fs.length) addFiles(fs);
  }

  function selectImage(id) { setActiveImgId(id); setMarkerEdit(false); setCropRowId(null); }

  // Chọn 1 dòng trong bảng. Nếu có >1 ảnh và ảnh đang xem KHÔNG chứa ký hiệu của dòng này,
  // tự chuyển sang ảnh phối cảnh đầu tiên có ký hiệu của dòng đó.
  function selectRow(r) {
    setActiveId(r.id);
    if (images.length > 1 && r.instances && r.instances.length) {
      const imgIds = r.instances.map((b) => b.imgId);
      if (imgIds.indexOf(activeImgId) < 0) {
        const target = imgIds.find((id) => images.some((im) => im.id === id));
        if (target != null && target !== activeImgId) { setActiveImgId(target); setMarkerEdit(false); setCropRowId(null); }
      }
    }
  }

  // Gỡ 1 ảnh: xóa khỏi danh sách, gỡ box của nó khỏi bảng rồi gộp + đánh mã lại.
  // Không xoá element khỏi imgElMap (giữ để Hoàn tác nếu cần), chỉ ẩn khỏi UI.
  function removeImage(id) {
    pushUndo("gỡ ảnh");
    setImages((prev) => {
      const next = prev.filter((im) => im.id !== id);
      setActiveImgId((cur) => (cur === id ? (next[0] ? next[0].id : null) : cur));
      return next;
    });
    setRows((rs) => sortRows(codeItems(mergeRows(stripImageInstances(rs, id), gpOf), gpOf)));
    setImgFilter((f) => (f === id ? null : f));
    setSelected(new Set());
    setStatus("Đã gỡ 1 ảnh và cập nhật lại bảng (gộp trùng, đánh mã lại).");
  }

  // Gọi API phân tích 1 ảnh -> danh sách item (chưa gắn imgId). Báo lỗi CHI TIẾT để dễ chẩn đoán.
  // ĐÃ ĐỔI: gọi qua /api/analyze (hàm backend trên Vercel giữ ANTHROPIC_API_KEY phía server)
  // thay vì gọi thẳng api.anthropic.com từ trình duyệt — cách cũ CHỈ chạy được bên trong
  // claude.ai (nền tảng tự chèn quyền truy cập); ra ngoài (deploy Vercel) sẽ bị CORS chặn
  // và không có API key hợp lệ để gửi kèm. Body gửi lên GIỮ NGUYÊN như cũ, /api/analyze
  // chỉ chuyển tiếp nguyên văn sang Anthropic rồi trả nguyên văn kết quả về — không đổi gì
  // ở phần xử lý response bên dưới.
  // FALLBACK: Claude đọc CẢ ẢNH trong 1 lượt (luồng cũ) — dùng khi /api/detect chưa sẵn sàng.
  async function callAnalyzeSingle(imgId) {
    const pic = apiImageFor(imgId);
    if (!pic.data) throw new Error("Ảnh chưa nạp xong, thử lại sau 1-2 giây.");
    let res;
    try {
      res = await fetch("/api/analyze", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-6", max_tokens: 1000,
          messages: [{ role: "user", content: [
            { type: "image", source: { type: "base64", media_type: pic.media_type, data: pic.data } },
            { type: "text", text: PROMPT },
          ] }],
        }),
      });
    } catch (netErr) {
      throw new Error("Lỗi mạng khi gọi API (fetch thất bại).");
    }
    if (!res.ok) {
      let detail = "";
      try { const j = await res.json(); detail = (j && j.error && j.error.message) || ""; }
      catch { try { detail = (await res.text()).slice(0, 200); } catch (e) { /* bỏ qua */ } }
      throw new Error("HTTP " + res.status + (detail ? " · " + detail : ""));
    }
    const data = await res.json();
    if (data && data.type === "error") throw new Error(data.error ? data.error.message : "API trả về lỗi.");
    const textOut = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n");
    return parseItems(textOut);
  }

  // TẦNG 1 (grounding): gọi /api/detect (Gemini) -> danh sách vùng {label,count,x1,y1,x2,y2} trong [0..1].
  async function detectRegionsApi(pic) {
    const res = await fetch("/api/detect", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ image: { media_type: pic.media_type, data: pic.data } }),
    });
    if (!res.ok) throw new Error("detect HTTP " + res.status);
    const j = await res.json();
    return Array.isArray(j.regions) ? j.regions : [];
  }

  // TẦNG 2 (đọc): gửi 1 CROP cho Claude -> 1 item (dùng lại parseItems, lấy dòng đầu).
  // hintLabel = nhãn tiếng Anh do Gemini gán (vase/tray/basket/pillow...) -> giúp phân nhóm đúng.
  async function readCrop(cropB64, hintLabel) {
    const promptText = CROP_PROMPT + (hintLabel ? ("\n\nGợi ý loại vật (từ hệ phát hiện, tiếng Anh): " + hintLabel) : "");
    const res = await fetch("/api/analyze", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: "claude-sonnet-4-6", max_tokens: 400,
        messages: [{ role: "user", content: [
          { type: "image", source: { type: "base64", media_type: "image/jpeg", data: cropB64 } },
          { type: "text", text: promptText },
        ] }],
      }),
    });
    if (!res.ok) return null;
    const data = await res.json();
    if (data && data.type === "error") return null;
    const textOut = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n");
    const items = parseItems(textOut);
    return items[0] || null;
  }

  // PASS BỀ MẶT: đọc TOÀN ẢNH, chỉ lấy vật liệu bề mặt/hoàn thiện (sàn/tường/trần/rèm/sơn/đá ốp...).
  async function readSurfaces(pic) {
    const res = await fetch("/api/analyze", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: "claude-sonnet-4-6", max_tokens: 1000,
        messages: [{ role: "user", content: [
          { type: "image", source: { type: "base64", media_type: pic.media_type, data: pic.data } },
          { type: "text", text: SURFACE_PROMPT },
        ] }],
      }),
    });
    if (!res.ok) return [];
    const data = await res.json();
    if (data && data.type === "error") return [];
    const textOut = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n");
    return parseItems(textOut);
  }

  // Two-stage: detect (Gemini) -> crop hi-res -> read (Claude) từng vùng, ĐỒNG THỜI đọc pass bề mặt.
  // Nếu /api/detect lỗi hoặc không có vùng nào -> FALLBACK về callAnalyzeSingle (Claude đọc cả ảnh).
  async function callAnalyze(imgId) {
    const pic = apiImageFor(imgId);
    if (!pic.data) throw new Error("Ảnh chưa nạp xong, thử lại sau 1-2 giây.");
    const el = getEl(imgId);
    try {
      if (el && el.naturalWidth) {
        const regions = await detectRegionsApi(pic);   // B: Gemini khoanh TỪNG cá thể (không tự gộp)
        if (regions.length) {
          // TẦNG C: gom nhóm bằng thị giác. Ghép mọi crop -> montage -> hỏi Claude "ô nào cùng 1 sản phẩm".
          // Với >=2 vùng mới cần gom; 1 vùng thì khỏi. Cluster lỗi -> fallback mỗi vùng 1 nhóm (như cũ).
          let groups = null;
          if (regions.length >= 2) {
            const montage = buildMontage(el, regions, 240);
            if (montage) groups = await clusterRegions(b64of(montage), regions.length);
          }
          if (!groups) groups = regions.map((_, i) => ({ members: [i + 1], loai: "", nhom: "" }));

          // Chạy SONG SONG: (a) mỗi NHÓM đọc 1 cá thể ĐẠI DIỆN (full-res); (b) pass bề mặt toàn ảnh.
          const [read, surfaceItems] = await Promise.all([
            mapLimit(groups, 4, async (g) => {
              const idxs = (g.members || []).map((m) => m - 1).filter((i) => i >= 0 && i < regions.length);
              if (!idxs.length) return null;
              // đại diện = cá thể có box LỚN NHẤT (thường rõ nhất, ít bị che)
              let repI = idxs[0], repA = _boxArea(regions[repI]);
              for (const i of idxs) { const a = _boxArea(regions[i]); if (a > repA) { repA = a; repI = i; } }
              const rg = regions[repI];
              const crop = makeExportCrop(el, rg, 1000);
              if (!crop || !crop.data) return null;
              const one = await readCrop(b64of(crop.data), g.loai || rg.label);   // hint = loại do C gom được
              if (!one) return null;
              one.instances = [{ x1: rg.x1, y1: rg.y1, x2: rg.x2, y2: rg.y2 }];    // 1 ký hiệu = cá thể đại diện
              // so_luong = tổng count các thành viên trong nhóm (mỗi cá thể count=1 -> = số thành viên).
              // Bền với cả trường hợp B lỡ gộp (1 box count=k): vẫn cộng đúng.
              one.soLuong = idxs.reduce((s, i) => s + (Number.isInteger(regions[i].count) && regions[i].count > 0 ? regions[i].count : 1), 0);
              if (g.nhom && !one.nhom) one.nhom = g.nhom;
              // lưu box mọi thành viên (không hiển thị marker) — để dành cho việc tách nhóm / xuất sau này
              one.memberBoxes = idxs.map((i) => ({ x1: regions[i].x1, y1: regions[i].y1, x2: regions[i].x2, y2: regions[i].y2 }));
              return one;
            }),
            readSurfaces(pic).catch(() => []),
          ]);
          const objectItems = dedupeObjects(read.filter(Boolean));   // an toàn: gộp nốt nếu 2 đại diện lỡ chồng box
          const combined = [...objectItems, ...surfaceItems];   // đồ rời (đã gom nhóm) + bề mặt (Claude toàn ảnh)
          if (combined.length) return combined;
        }
      }
    } catch (e) {
      // detect/crop lỗi -> rơi xuống fallback bên dưới (không chặn quy trình cũ)
    }
    return callAnalyzeSingle(imgId);
  }

  // Phân tích 1 ảnh rồi GỘP kết quả VÀO bảng hiện có (không thay thế).
  // Đây là mấu chốt để món/vật liệu giống nhau ở NHIỀU ảnh gộp về 1 mã duy nhất:
  //   [dòng cũ (đã bỏ box của ảnh này)] + [item mới] -> mergeRows theo (prefix|món|vật liệu) -> codeItems.
  async function analyzeOne(imgId) {
    const im = images.find((x) => x.id === imgId);
    if (!im) return 0;
    setImages((prev) => prev.map((x) => (x.id === imgId ? { ...x, status: "analyzing", err: "" } : x)));
    try {
      let items = await callAnalyze(imgId);
      items = mergeSameImage(items);   // C: gộp trùng TRONG cùng ảnh, CỘNG DỒN số lượng (trước khi gộp xuyên ảnh lấy MAX)
      items = items.map((it) => ({ ...it, srcImg: imgId, instances: (it.instances || []).map((b) => ({ ...b, imgId })) }));
      setRows((prev) => {
        const stripped = stripImageInstances(prev, imgId);        // idempotent nếu phân tích lại cùng ảnh
        let combined = mergeRows([...stripped, ...items], gpOf);   // GỘP TRÙNG XUYÊN ẢNH
        combined = codeItems(combined, gpOf);                      // đánh mã tuần tự toàn cục
        combined = combined.map((r) => {
          if (r.thumb) return r;
          const b0 = r.instances[0];
          return { ...r, thumb: b0 && elReady(b0.imgId) ? makeThumb(getEl(b0.imgId), b0) : null };
        });
        return sortRows(combined);
      });
      setImages((prev) => prev.map((x) => (x.id === imgId ? { ...x, status: "done" } : x)));
      return items.length;
    } catch (err) {
      const msg = (err && err.message) || "Lỗi không xác định";
      setImages((prev) => prev.map((x) => (x.id === imgId ? { ...x, status: "error", err: msg } : x)));
      throw err;
    }
  }

  // Phân tích riêng ảnh đang xem
  async function analyzeActive() {
    if (!activeImage) { setError("Chưa có ảnh. Hãy tải ảnh phối cảnh lên trước."); return; }
    setLoading(true); setError(null); setStatus(null); setSelected(new Set());
    try {
      const n = await analyzeOne(activeImgId);
      if (!n) setError("Không đọc được món nào từ ảnh này. Thử ảnh rõ hơn, hoặc crop sát khu vực cần bóc rồi phân tích lại.");
      else setStatus("Đã phân tích ảnh #" + (imgIndex(activeImgId) + 1) + " · gộp trùng xuyên ảnh & gán mã lại toàn bộ.");
    } catch (e) { setError("Không phân tích được ảnh #" + (imgIndex(activeImgId) + 1) + ": " + ((e && e.message) || "lỗi không xác định")); }
    finally { setLoading(false); }
  }

  // Phân tích LẦN LƯỢT tất cả ảnh (tuần tự để việc gộp cộng dồn nhất quán)
  async function analyzeAll() {
    if (!images.length) { setError("Chưa có ảnh nào để phân tích."); return; }
    setLoading(true); setError(null); setStatus(null); setSelected(new Set());
    let ok = 0, fail = 0; const errs = [];
    const list = images.slice();
    for (let i = 0; i < list.length; i++) {
      const im = list[i];
      setActiveImgId(im.id);
      try { await analyzeOne(im.id); ok++; }
      catch (e) { fail++; errs.push("Ảnh #" + (i + 1) + ": " + ((e && e.message) || "lỗi")); }
    }
    setLoading(false);
    if (!ok) setError("Không phân tích được ảnh nào — " + (errs[0] || "kiểm tra kết nối rồi thử lại.") + (errs.length > 1 ? " (và " + (errs.length - 1) + " ảnh khác)" : ""));
    else {
      setStatus("Đã phân tích " + ok + "/" + list.length + " ảnh" + (fail ? " · " + fail + " ảnh lỗi" : "") + " · gộp trùng xuyên ảnh & gán mã.");
      if (fail) setError("Một số ảnh lỗi — " + errs.join(" | "));
    }
  }

  // ---- B5: Hoàn tác (undo) cho các thao tác phá huỷ (xoá/gộp/tách/gỡ ảnh) ----
  function pushUndo(label) {
    setUndoStack((s) => [...s.slice(-9), { rows, images, activeImgId, label }]); // giữ tối đa 10 mức
  }
  function undo() {
    setUndoStack((s) => {
      if (!s.length) return s;
      const last = s[s.length - 1];
      setRows(last.rows); setImages(last.images); setActiveImgId(last.activeImgId);
      setSelected(new Set());
      setStatus("Đã hoàn tác: " + last.label);
      return s.slice(0, -1);
    });
  }

  function updateRow(id, field, value) { setRows((rs) => rs.map((r) => (r.id === id ? { ...r, [field]: value } : r))); }
  function deleteRow(id) { pushUndo("xoá dòng"); setRows((rs) => rs.filter((r) => r.id !== id)); if (activeId === id) setActiveId(null); setSelected((s) => { const n = new Set(s); n.delete(id); return n; }); }
  function addRow() { const r = { id: nextId(), prefix: "", ma: "", nhom: "Nội thất", mon: "", vat_lieu: "", soLuong: 1, vi_tri: "", do_tin_cay: "Trung bình", ghi_chu: "", instances: [], thumb: null }; setRows((rs) => sortRows([...rs, r])); setActiveId(r.id); }
  function recode() {
    setRows((rs) => sortRows(codeItems(mergeRows(rs, gpOf), gpOf)));
    setActiveId(null); setSelected(new Set());
    setStatus("Đã gộp dòng trùng, đánh mã lại và sắp xếp theo bảng vật liệu.");
  }
  function toggleEdit() { const next = !markerEdit; setMarkerEdit(next); if (next && !activeId && rows.length) setActiveId(rows[0].id); }

  // Kéo đường chia giữa 2 cột để đổi tỉ lệ (giới hạn 28%–72%).
  function startSplitDrag(e) {
    e.preventDefault();
    const cont = splitRef.current; if (!cont) return;
    const rect = cont.getBoundingClientRect();
    setSplitDragging(true);
    const onMove = (ev) => { const pct = ((ev.clientX - rect.left) / rect.width) * 100; setSplitPct(Math.max(28, Math.min(72, pct))); };
    const onUp = () => { setSplitDragging(false); window.removeEventListener("pointermove", onMove); window.removeEventListener("pointerup", onUp); };
    window.addEventListener("pointermove", onMove); window.addEventListener("pointerup", onUp);
  }

  // Số thứ tự ảnh nguồn của 1 dòng (vd ["#2"] hoặc ["#1","#2"]) — dùng để hiện & để gắn đúng ảnh.
  function srcNums(r) {
    const ids = (r.instances.length ? r.instances.map((b) => b.imgId) : (r.srcImgs || []));
    return Array.from(new Set(ids)).map((id) => imgIndex(id)).filter((i) => i >= 0).sort((a, b) => a - b).map((i) => "#" + (i + 1));
  }

  // B: bắt đầu gắn ký hiệu cho 1 dòng chưa có box — TỰ NHẢY sang ảnh nguồn rồi bật marker-edit.
  function pinRow(r) {
    const ids = (r.instances.length ? r.instances.map((b) => b.imgId) : (r.srcImgs || []));
    const existing = Array.from(new Set(ids)).filter((id) => images.some((im) => im.id === id));
    const targetImg = existing.length
      ? (existing.indexOf(activeImgId) >= 0 ? activeImgId : existing[0])
      : (activeImgId != null ? activeImgId : (images[0] ? images[0].id : null));
    if (targetImg == null) { setStatus("Chưa có ảnh để gắn ký hiệu. Hãy tải ảnh phối cảnh trước."); return; }
    setActiveImgId(targetImg);
    setActiveId(r.id);
    setMarkerEdit(true);
    const n = imgIndex(targetImg) + 1;
    const note = existing.length ? "" : " (không rõ ảnh nguồn — đang mở ảnh đang xem)";
    setStatus("Đang gắn “" + (r.mon || r.ma || "dòng này") + "” trên ảnh #" + n + note + ": bấm lên ảnh tại đúng vị trí." + (r.vi_tri ? " Gợi ý vị trí AI ghi: " + r.vi_tri + "." : ""));
  }

  // ---- B1: chọn dòng + gộp/tách thủ công ----
  function toggleSelect(id) { setSelected((s) => { const n = new Set(s); n.has(id) ? n.delete(id) : n.add(id); return n; }); }
  function clearSelect() { setSelected(new Set()); }

  // Gộp các dòng đã tick thành 1 (dồn box, gộp vị trí/ghi chú, lấy tin cậy cao nhất), rồi đánh mã lại.
  // Dùng khi AI tách nhầm 1 món thành nhiều dòng do chữ lệch nhau mà chuẩn hoá tự động chưa bắt được.
  function mergeSelected() {
    const chosen = rows.filter((r) => selected.has(r.id));
    if (chosen.length < 2) return;
    pushUndo("gộp dòng");
    const rank = { "Cao": 3, "Trung bình": 2, "Thấp": 1 };
    const base = { ...chosen[0], instances: [...chosen[0].instances] };
    const sset = srcIdsOf(chosen[0]);
    const locs = new Set(splitLocs(base.vi_tri));
    const notes = base.ghi_chu ? [String(base.ghi_chu).trim()] : [];
    let maxSL = parseInt(base.soLuong, 10); if (!Number.isFinite(maxSL)) maxSL = base.instances.length || 0;
    for (let i = 1; i < chosen.length; i++) {
      const c = chosen[i];
      // Giữ 1 ký hiệu (đối tượng rõ nhất): nếu base chưa có box mà dòng gộp có thì lấy 1 box.
      if ((!base.instances || base.instances.length === 0) && c.instances && c.instances.length) base.instances = [c.instances[0]];
      const slc = parseInt(c.soLuong, 10); if (Number.isFinite(slc)) maxSL = Math.max(maxSL, slc);
      srcIdsOf(c).forEach((x) => sset.add(x));
      splitLocs(c.vi_tri).forEach((s) => locs.add(s));
      const nt = String(c.ghi_chu || "").trim();
      if (nt && notes.indexOf(nt) < 0) notes.push(nt);
      if ((rank[c.do_tin_cay] || 0) > (rank[base.do_tin_cay] || 0)) base.do_tin_cay = c.do_tin_cay;
    }
    base.soLuong = maxSL || 1;
    base.srcImgs = Array.from(sset);
    base.vi_tri = Array.from(locs).join(", ");
    base.ghi_chu = notes.filter(Boolean).join("; ");
    const dropIds = new Set(chosen.slice(1).map((r) => r.id));
    setRows((rs) => sortRows(codeItems(rs.filter((r) => !dropIds.has(r.id)).map((r) => (r.id === base.id ? base : r)), gpOf)));
    setSelected(new Set()); setActiveId(base.id);
    setStatus("Đã gộp " + chosen.length + " dòng thành 1 (SL=" + qtyOf(base) + ") và đánh mã lại.");
  }

  // Tách mỗi dòng đã tick (có >=2 ký hiệu) thành từng dòng SL=1, giữ nguyên chữ.
  // Không tự đánh mã lại (tránh nhiều dòng cùng mã) — sửa tên khác nhau rồi bấm "Gộp trùng & đánh mã lại".
  function splitSelected() {
    const targets = rows.filter((r) => selected.has(r.id) && r.instances.length > 1);
    if (!targets.length) { setStatus("Chọn dòng có từ 2 ký hiệu trở lên để tách."); return; }
    pushUndo("tách dòng");
    setRows((rs) => {
      const out = [];
      for (const r of rs) {
        if (selected.has(r.id) && r.instances.length > 1) {
          r.instances.forEach((b) => out.push({ ...r, id: nextId(), soLuong: 1, instances: [b], thumb: elReady(b.imgId) ? makeThumb(getEl(b.imgId), b) : null }));
        } else out.push(r);
      }
      return sortRows(out);
    });
    setSelected(new Set());
    setStatus("Đã tách thành các dòng SL=1. Đổi tên cho khác nhau rồi bấm “Gộp trùng & đánh mã lại”.");
  }

  // Thêm ký hiệu (box) cho dòng đang chọn — box gắn với ẢNH ĐANG XEM (activeImgId)
  function onImageClick(e) {
    if (!markerEdit) {
      // Ngoài chế độ Thêm ký hiệu: bấm vùng trống -> bỏ tô màu (bỏ chọn dòng) & khoá lại vùng crop.
      // Marker có stopPropagation nên handler này chỉ chạy khi bấm đúng vùng trống / nền ảnh.
      if (activeId != null) setActiveId(null);
      if (cropRowId != null) setCropRowId(null);
      return;
    }
    if (activeId == null) { setStatus("Hãy chọn một dòng trong bảng trước khi thêm ký hiệu."); return; }
    if (activeImgId == null) return;
    const rect = e.currentTarget.getBoundingClientRect();
    const nx = clamp01((e.clientX - rect.left) / rect.width), ny = clamp01((e.clientY - rect.top) / rect.height);
    const box = { imgId: activeImgId, x1: clamp01(nx - 0.03), y1: clamp01(ny - 0.03), x2: clamp01(nx + 0.03), y2: clamp01(ny + 0.03) };
    setRows((rs) => rs.map((r) => {
      if (r.id !== activeId) return r;
      const thumb = r.thumb || (elReady(activeImgId) ? makeThumb(getEl(activeImgId), box) : null);
      return { ...r, instances: [...r.instances, box], thumb };
    }));
  }

  function startMarker(e, rowId, instIdx) {
    e.stopPropagation(); setActiveId(rowId);
    const inCrop = cropRowId === rowId;          // đang mở khoá kéo vùng crop cho đúng dòng này (double-click)
    const canDrag = markerEdit || inCrop;         // chế độ Thêm ký hiệu HOẶC mở khoá crop -> cho phép kéo move
    if (!canDrag) return;
    const wrap = wrapRef.current; if (!wrap) return;
    const rect = wrap.getBoundingClientRect();
    const row = rows.find((r) => r.id === rowId); const b = row && row.instances[instIdx]; if (!b) return;
    const startCx = (b.x1 + b.x2) / 2, startCy = (b.y1 + b.y2) / 2, w = b.x2 - b.x1, h = b.y2 - b.y1;
    const startPX = e.clientX, startPY = e.clientY; let moved = false;
    const onMove = (ev) => {
      if (Math.abs(ev.clientX - startPX) > 4 || Math.abs(ev.clientY - startPY) > 4) moved = true;
      if (!moved) return;
      const cx = clamp01(startCx + (ev.clientX - startPX) / rect.width), cy = clamp01(startCy + (ev.clientY - startPY) / rect.height);
      const nb = { ...b, x1: clamp01(cx - w / 2), y1: clamp01(cy - h / 2), x2: clamp01(cx + w / 2), y2: clamp01(cy + h / 2) };
      setRows((rs) => rs.map((r) => (r.id === rowId ? { ...r, instances: r.instances.map((it, i) => (i === instIdx ? nb : it)) } : r)));
    };
    const onUp = () => {
      window.removeEventListener("pointermove", onMove); window.removeEventListener("pointerup", onUp);
      // Chỉ CHẾ ĐỘ Thêm ký hiệu mới xóa khi click-không-kéo. Chế độ mở khoá crop: click không xóa (dùng nút "x").
      if (!moved) { if (markerEdit && !inCrop) setRows((rs) => rs.map((r) => (r.id === rowId ? { ...r, instances: r.instances.filter((_, i) => i !== instIdx) } : r))); }
      else if (instIdx === 0 && elReady(b.imgId)) { setRows((rs) => rs.map((r) => { if (r.id !== rowId) return r; const first = r.instances[0]; return { ...r, thumb: first ? makeThumb(getEl(first.imgId), first) : r.thumb }; })); }
    };
    window.addEventListener("pointermove", onMove); window.addEventListener("pointerup", onUp);
  }

  // Kéo giãn/thu vùng crop (bounding box) theo hướng dir. Giữ nguyên imgId của box.
  function startResize(e, rowId, instIdx, dir) {
    e.stopPropagation(); e.preventDefault(); setActiveId(rowId);
    const wrap = wrapRef.current; if (!wrap) return;
    const rect = wrap.getBoundingClientRect();
    const row = rows.find((r) => r.id === rowId); const b = row && row.instances[instIdx]; if (!b) return;
    const base = { x1: b.x1, y1: b.y1, x2: b.x2, y2: b.y2 };
    const MIN = 0.01;
    const onMove = (ev) => {
      const nx = clamp01((ev.clientX - rect.left) / rect.width);
      const ny = clamp01((ev.clientY - rect.top) / rect.height);
      let { x1, y1, x2, y2 } = base;
      if (dir.indexOf("w") >= 0) x1 = nx;
      if (dir.indexOf("e") >= 0) x2 = nx;
      if (dir.indexOf("n") >= 0) y1 = ny;
      if (dir.indexOf("s") >= 0) y2 = ny;
      if (x2 - x1 < MIN) { if (dir.indexOf("w") >= 0) x1 = x2 - MIN; else x2 = x1 + MIN; }
      if (y2 - y1 < MIN) { if (dir.indexOf("n") >= 0) y1 = y2 - MIN; else y2 = y1 + MIN; }
      const nb = { ...b, x1: clamp01(x1), y1: clamp01(y1), x2: clamp01(x2), y2: clamp01(y2) };
      setRows((rs) => rs.map((r) => (r.id === rowId ? { ...r, instances: r.instances.map((it, i) => (i === instIdx ? nb : it)) } : r)));
    };
    const onUp = () => {
      window.removeEventListener("pointermove", onMove); window.removeEventListener("pointerup", onUp);
      if (instIdx === 0 && elReady(b.imgId)) { setRows((rs) => rs.map((r) => { if (r.id !== rowId) return r; const first = r.instances[0]; return { ...r, thumb: first ? makeThumb(getEl(first.imgId), first) : r.thumb }; })); }
    };
    window.addEventListener("pointermove", onMove); window.addEventListener("pointerup", onUp);
  }

  // Xóa 1 ký hiệu (box) cụ thể của dòng — dùng nút "x" khi đang chỉnh vùng crop.
  function deleteInstance(rowId, instIdx) {
    setRows((rs) => rs.map((r) => {
      if (r.id !== rowId) return r;
      const insts = r.instances.filter((_, i) => i !== instIdx);
      const first = insts[0];
      const thumb = first ? (elReady(first.imgId) ? makeThumb(getEl(first.imgId), first) : r.thumb) : null;
      return { ...r, instances: insts, thumb };
    }));
  }

  function buildCategorySheet(cat, list) {
    const aoa = [], merges = [];
    aoa.push(["CHỈ DẪN KỸ THUẬT VẬT LIỆU / MATERIAL SPECIFICATION"]);
    aoa.push(["Dự án / Project:", projectName]);
    aoa.push(["Chủ đầu tư / Client:", client]);
    aoa.push(["Địa điểm / Location:", location]);
    aoa.push(["Ngày / Date:", dateStr, "", "Người bóc / By:", author]);
    aoa.push([]);
    const headR = aoa.length;
    aoa.push(["STT / NO.", "KÝ HIỆU / SYMBOL", "THÔNG TIN / INFORMATION", "", "HÌNH ẢNH MẪU 3D / 3D", "HÌNH ẢNH DUYỆT / APPROVED", "LINK"]);
    const catR = aoa.length;
    aoa.push([cat]);
    merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: 6 } });
    merges.push({ s: { r: headR, c: 2 }, e: { r: headR, c: 3 } });
    merges.push({ s: { r: catR, c: 0 }, e: { r: catR, c: 6 } });
    let stt = 0;
    list.forEach((r) => {
      stt++; const s = aoa.length;
      const desc = [r.mon, r.vat_lieu].filter(Boolean).join(" — ");
      aoa.push([stt, r.ma || "", "Mô tả / Description:", desc, "", "", ""]);
      aoa.push(["", "", "Vị trí sử dụng / Area:", r.vi_tri || "", "", "", ""]);
      aoa.push(["", "", "Số lượng (SL) / Qty:", String(qtyOf(r)), "", "", ""]);
      aoa.push(["", "", "Kích thước / Dimension:", "", "", "", ""]);
      aoa.push(["", "", "Nhãn hiệu / Brand:", "", "", "", ""]);
      aoa.push(["", "", "Ghi chú / Note:", r.ghi_chu || "", "", "", ""]);
      [0, 1, 4, 5, 6].forEach((c) => merges.push({ s: { r: s, c }, e: { r: s + 5, c } }));
    });
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!merges"] = merges;
    ws["!cols"] = [{ wch: 7 }, { wch: 12 }, { wch: 26 }, { wch: 42 }, { wch: 20 }, { wch: 20 }, { wch: 16 }];
    return ws;
  }

  function exportExcel() {
    try {
      if (!rows.length) return;
      const wb = XLSX.utils.book_new();
      const used = new Set();
      const addSheet = (name, ws) => {
        let n = String(name || "SHEET").replace(/[\\/?*\[\]:]/g, " ").trim().slice(0, 28) || "SHEET";
        let base = n, k = 2;
        while (used.has(n.toLowerCase())) { n = (base.slice(0, 24) + " " + k).trim(); k++; }
        used.add(n.toLowerCase()); XLSX.utils.book_append_sheet(wb, ws, n);
      };
      const H = ["STT", "Mã", "Nhóm", "Món", "Vật liệu / Finish", "Vị trí", "SL", "Độ tin cậy", "Ghi chú"];
      const sum = [H, ...rows.map((r, i) => [i + 1, r.ma, r.nhom, r.mon, r.vat_lieu, r.vi_tri, qtyOf(r), r.do_tin_cay, r.ghi_chu])];
      const wsSum = XLSX.utils.aoa_to_sheet(sum);
      wsSum["!cols"] = [{ wch: 6 }, { wch: 10 }, { wch: 15 }, { wch: 24 }, { wch: 26 }, { wch: 16 }, { wch: 6 }, { wch: 12 }, { wch: 26 }];
      addSheet("TỔNG HỢP", wsSum);
      const groups = {}, order = [];
      rows.forEach((r) => {
        const pfx = (r.ma.split("-")[0] || "").toUpperCase();
        const cat = sheetCategory(pfx);
        if (!groups[cat]) { groups[cat] = []; order.push(cat); }
        groups[cat].push(r);
      });
      order.forEach((cat) => addSheet(cat, buildCategorySheet(cat, groups[cat])));
      XLSX.writeFile(wb, safeName() + ".xlsx");
      setStatus("Đã xuất " + safeName() + ".xlsx — sheet TỔNG HỢP + mỗi nhóm vật liệu 1 sheet theo template spec.");
    } catch (e) {
      setStatus("Trình duyệt chặn tải file. Dùng “Sao chép bảng” rồi dán vào Excel.");
    }
  }

  async function copyTSV() {
    const H = ["STT", "Mã", "Nhóm", "Món", "Vật liệu / Finish", "Vị trí", "SL", "Độ tin cậy", "Ghi chú"];
    const tsv = [H.join("\t"), ...rows.map((r, i) => [i + 1, r.ma, r.nhom, r.mon, r.vat_lieu, r.vi_tri, qtyOf(r), r.do_tin_cay, r.ghi_chu].join("\t"))].join("\n");
    try { await navigator.clipboard.writeText(tsv); setStatus("Đã sao chép bảng — dán (Ctrl/Cmd + V) vào Excel."); }
    catch {
      const ta = document.createElement("textarea"); ta.value = tsv; document.body.appendChild(ta); ta.select();
      try { document.execCommand("copy"); setStatus("Đã sao chép bảng."); } catch { setStatus("Không sao chép được tự động — hãy chọn bảng và copy thủ công."); }
      document.body.removeChild(ta);
    }
  }

  // Vẽ ảnh có đánh số ký hiệu (STT theo số dòng toàn cục) -> dataURL, hoặc null nếu ảnh chưa sẵn sàng / không có ký hiệu.
  function annotatedDataUrl(imgId) {
    const img = getEl(imgId);
    if (!img || !img.naturalWidth) return null;
    const nw = img.naturalWidth, nh = img.naturalHeight;
    const c = document.createElement("canvas"); c.width = nw; c.height = nh;
    const ctx = c.getContext("2d"); ctx.drawImage(img, 0, 0);
    const r = Math.max(13, Math.round(Math.min(nw, nh) * 0.02));
    const dispNo = buildDisplayNo(rows);
    const flat = [];
    rows.forEach((row) => row.instances.forEach((b) => { if (b.imgId !== imgId) return; flat.push({ no: dispNo.get(row.id) || 0, x: ((b.x1 + b.x2) / 2) * nw, y: ((b.y1 + b.y2) / 2) * nh }); }));
    if (!flat.length) return null;
    const sep = separate(flat.map((f) => ({ x: f.x, y: f.y })), r * 2.3, nw, nh);
    ctx.font = "bold " + Math.round(r * 1.15) + "px sans-serif"; ctx.textAlign = "center"; ctx.textBaseline = "middle";
    flat.forEach((f, i) => {
      const cx = sep[i].x, cy = sep[i].y;
      ctx.beginPath(); ctx.arc(cx, cy, r, 0, Math.PI * 2); ctx.fillStyle = "rgba(123,163,207,0.95)"; ctx.fill();
      ctx.lineWidth = Math.max(1, r * 0.14); ctx.strokeStyle = "#0c1524"; ctx.stroke();
      ctx.fillStyle = "#0c1524"; ctx.fillText(String(f.no), cx, cy);
    });
    return c.toDataURL("image/jpeg", 0.9);
  }

  // Tải ảnh đánh số cho ẢNH ĐANG XEM (1 file)
  function downloadAnnotated() {
    if (!activeImgId) return;
    if (!elReady(activeImgId)) { setStatus("Ảnh chưa sẵn sàng."); return; }
    const url = annotatedDataUrl(activeImgId);
    if (!url) { setStatus("Ảnh này chưa có ký hiệu nào để đánh số."); return; }
    const idx = imgIndex(activeImgId) + 1;
    try { const a = document.createElement("a"); a.href = url; a.download = safeName() + "-anh" + idx + "-danh-so.jpg"; a.click(); setStatus("Đã tải ảnh #" + idx + " (đã đánh số)."); }
    catch (e) { setStatus("Không tải được ảnh đánh số trong môi trường này."); }
  }

  // B3: tải TẤT CẢ ảnh đã đánh số (mỗi ảnh 1 file, tải liên tiếp)
  async function downloadAllAnnotated() {
    if (!images.length) return;
    let n = 0;
    for (let i = 0; i < images.length; i++) {
      const url = annotatedDataUrl(images[i].id);
      if (!url) continue;
      try { const a = document.createElement("a"); a.href = url; a.download = safeName() + "-anh" + (i + 1) + "-danh-so.jpg"; a.click(); n++; }
      catch (e) { /* bỏ qua ảnh lỗi */ }
      await new Promise((res) => setTimeout(res, 350)); // giãn nhịp để trình duyệt không chặn tải hàng loạt
    }
    setStatus(n ? ("Đã tải " + n + " ảnh đã đánh số. Nếu trình duyệt hỏi, hãy cho phép tải nhiều file.") : "Không có ảnh nào có ký hiệu để đánh số.");
  }

  // C3: crop lớn (tối đa ~1000px) để soi trong lightbox
  function bigCrop(box) {
    if (!box || !elReady(box.imgId)) return null;
    return makeExportCrop(getEl(box.imgId), box, 1000);
  }
  function openLightbox(r) {
    const b0 = r.instances[0];
    const crop = b0 ? bigCrop(b0) : null;   // bigCrop trả về { data, w, h } — phải lấy .data cho <img src>
    const src = crop && crop.data ? crop.data : null;
    if (!src) { setStatus("Chưa có crop cho dòng này (thiếu ký hiệu hoặc ảnh chưa sẵn sàng)."); return; }
    const imgN = b0 ? (imgIndex(b0.imgId) + 1) : "?";
    setLightbox({ code: r.ma || "—", title: [r.mon, r.vat_lieu].filter(Boolean).join(" — ") || "(chưa đặt tên)", src, meta: "SL " + qtyOf(r) + " · crop từ ảnh #" + imgN + (r.vi_tri ? " · " + r.vi_tri : "") });
  }

  function exportBundle() {
    if (!rows.length) return;
    const items = rows.map((r, i) => {
      const b0 = r.instances[0];
      const img = b0 && elReady(b0.imgId) ? makeExportCrop(getEl(b0.imgId), b0) : null;
      return { stt: i + 1, ma: r.ma, nhom: r.nhom, mon: r.mon, vat_lieu: r.vat_lieu, vi_tri: r.vi_tri, sl: qtyOf(r), do_tin_cay: r.do_tin_cay, ghi_chu: r.ghi_chu, image: img };
    });
    const bundle = { meta: { project: projectName, client, location, author, date: dateStr, images: images.length, generatedAt: new Date().toISOString() }, items };
    try {
      const blob = new Blob([JSON.stringify(bundle)], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a"); a.href = url; a.download = safeName() + "-bundle.json"; a.click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      const withImg = items.filter((x) => x.image).length;
      setStatus("Đã xuất gói .json (" + items.length + " dòng, " + withImg + " ảnh crop) — gửi file này vào chat để tạo .xlsx có ảnh nhúng.");
    } catch (e) { setStatus("Không tạo được gói .json trong môi trường này."); }
  }

  const hasRows = rows.length > 0;
  const hasImages = images.length > 0;
  const activeIdx = rows.findIndex((r) => r.id === activeId);
  const totalSyms = rows.reduce((s, r) => s + r.instances.length, 0);
  const lowN = rows.filter((r) => r.do_tin_cay === "Thấp").length;
  const analyzedN = images.filter((im) => im.status === "done").length;

  // Tab chia theo 6 NHÓM CHỨC NĂNG ở cột Nhóm (NHOM_OPTS), thứ tự cố định, chỉ hiện nhóm có dòng.
  const ALL_SHEET = "__all__";
  const sheetTabs = NHOM_OPTS
    .map((g) => ({ key: g, count: rows.filter((r) => r.nhom === g).length }))
    .filter((t) => t.count > 0);
  const curSheet = sheetTabs.some((t) => t.key === activeSheet) ? activeSheet : ALL_SHEET;

  // C2: hàm khớp tìm kiếm (bỏ dấu) trên món/mã/vật liệu/vị trí
  const q = stripVN(search).trim();
  const matchSearch = (r) => !q || stripVN([r.ma, r.mon, r.vat_lieu, r.vi_tri, r.nhom].join(" ")).indexOf(q) >= 0;
  const passFilter = (r) => (curSheet === ALL_SHEET || r.nhom === curSheet) && (!onlyLow || r.do_tin_cay === "Thấp") && (!onlyUnpinned || r.instances.length === 0) && (imgFilter == null || r.instances.some((b) => b.imgId === imgFilter)) && matchSearch(r);
  const visibleCount = rows.filter(passFilter).length;
  const unpinnedN = rows.filter((r) => r.instances.length === 0).length; // A: số dòng chưa gắn ký hiệu
  const selCount = rows.filter((r) => selected.has(r.id)).length;
  const canSplit = rows.some((r) => selected.has(r.id) && r.instances.length > 1);

  // B5 + C3: phím tắt Ctrl/Cmd+Z hoàn tác; Esc đóng lightbox
  useEffect(() => {
    function onKey(e) {
      if (e.key === "Escape" && lightbox) { setLightbox(null); return; }
      const z = (e.key === "z" || e.key === "Z");
      if (z && (e.ctrlKey || e.metaKey) && !e.shiftKey) {
        const t = e.target, tag = t && t.tagName;
        if (tag === "INPUT" || tag === "SELECT" || tag === "TEXTAREA") return; // để undo gõ chữ trong ô hoạt động bình thường
        if (undoStack.length) { e.preventDefault(); undo(); }
      }
    }
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [undoStack, lightbox]);

  return (
    <div className="ax ax2">
      <style dangerouslySetInnerHTML={{ __html: css + cssExtra }} />

      {/* ===== BANNER THÔNG TIN DỰ ÁN (đầu trang · ẩn/hiện) ===== */}
      <header className="topbar">
        <div className="topbar-main">
          <div className="tb-brand">
            <img src={logoUrl} alt="ARTIUS" className="tb-logo" />
            <span className="tb-pill">SPEC MATERIAL AGENT</span>
          </div>
          <div className="tb-stats">
            <div className="tb-stat"><b style={{ color: hasRows ? "var(--ac2)" : undefined }}>{rows.length}</b><span>MÓN</span></div>
            <div className="tb-stat"><b>{totalSyms}</b><span>KÝ HIỆU</span></div>
            {lowN > 0 && <div className="tb-stat warn"><b>{lowN}</b><span>TIN CẬY THẤP</span></div>}
            {unpinnedN > 0 && <div className="tb-stat info"><b>{unpinnedN}</b><span>CHƯA GẮN KH</span></div>}
            <div className="tb-stat"><b style={{ color: hasImages ? "var(--ac2)" : undefined }}>{analyzedN}/{images.length}</b><span>ẢNH ĐÃ BÓC</span></div>
          </div>
          <button className="btn btn-ghost tb-toggle" onClick={() => setInfoOpen((v) => !v)} title={infoOpen ? "Ẩn thông tin dự án" : "Hiện thông tin dự án"}>
            {infoOpen ? <ChevronUp size={15} /> : <ChevronDown size={15} />} Thông tin dự án
          </button>
        </div>
        {infoOpen && (
          <div className="tb-fields">
            <div className="tb-field"><label>Dự án</label><input placeholder="Tên dự án…" value={projectName} onChange={(e) => setProjectName(e.target.value)} /></div>
            <div className="tb-field"><label>Chủ đầu tư</label><input placeholder="CĐT…" value={client} onChange={(e) => setClient(e.target.value)} /></div>
            <div className="tb-field"><label>Địa điểm</label><input placeholder="Địa điểm…" value={location} onChange={(e) => setLocation(e.target.value)} /></div>
            <div className="tb-field"><label>Người bóc</label><input placeholder="Tên bạn…" value={author} onChange={(e) => setAuthor(e.target.value)} /></div>
            <div className="tb-field"><label>Ngày lập</label><input placeholder="dd/mm/yyyy…" value={dateStr} onChange={(e) => setDateStr(e.target.value)} /></div>
          </div>
        )}
      </header>

      {/* ===== CHIA ĐÔI MÀN HÌNH: trái = Bảng bóc · phải = Ảnh phối cảnh ===== */}
      <div className="ax-split" ref={splitRef} style={{ "--splitpct": splitPct + "%" }}>
        <section className="pane pane-image">
          <div className="pane-body">

          {/* 01 · images */}
          <div>
            <div className="block-head">
              <span className="section-label">02 · Ảnh phối cảnh (nhiều ảnh)</span>
              {hasImages && <span className="imgcount">{images.length} ảnh · đã phân tích {analyzedN}</span>}
            </div>
            <div className="panel">
              <div className="ctl-row">
                <input ref={fileRef} type="file" accept="image/*" multiple onChange={onPickFiles} style={{ display: "none" }} />
                <div className="spacer" />
                {hasRows && (<button className={"btn btn-ghost" + (markerEdit ? " on" : "")} onClick={toggleEdit}><MapPin size={15} /> {markerEdit ? "Đang chỉnh ký hiệu" : "Thêm ký hiệu"}</button>)}
                <button className="btn btn-ghost" onClick={analyzeActive} disabled={loading || !activeImage}><MapPin size={15} /> Phân tích ảnh này</button>
                <button className="btn btn-primary" onClick={analyzeAll} disabled={loading || !hasImages}>{loading && <Loader2 size={15} className="spin" />}{loading ? "Đang phân tích…" : "Phân tích tất cả (" + images.length + ")"}</button>
              </div>

              {hasImages && (
                <div className="imgstrip">
                  {images.map((im, i) => {
                    const cnt = rows.filter((r) => r.instances.some((b) => b.imgId === im.id)).length;
                    return (
                      <div key={im.id} className={"imgtile" + (im.id === activeImgId ? " on" : "")} onClick={() => selectImage(im.id)} title={im.status === "error" && im.err ? ("LỖI — " + im.err) : im.fileName}>
                        <img src={im.preview} alt={im.fileName} />
                        <span className="idx">{i + 1}</span>
                        <span className={"st " + im.status} title={im.status === "error" && im.err ? im.err : im.status} />
                        {cnt > 0 && <span className="cnt">{cnt} mã</span>}
                        <button className="rm" title="Gỡ ảnh này" aria-label="Gỡ ảnh" onClick={(e) => { e.stopPropagation(); removeImage(im.id); }}><Trash2 width={12} height={12} /></button>
                      </div>
                    );
                  })}
                  <button className="imgadd" onClick={openPicker}><Plus size={16} /> Thêm ảnh</button>
                </div>
              )}

              {activeImage ? (
                <div ref={wrapRef} className={"imgwrap" + (markerEdit ? " edit" : "") + (dragOver ? " dragover" : "")} onClick={onImageClick}
                  onDragOver={onDragOver} onDragEnter={onDragOver} onDragLeave={onDragLeave} onDrop={onDrop}>
                  <img ref={imgRef} className="base" src={activeImage.preview} alt="Ảnh phối cảnh" onLoad={() => { const el = wrapRef.current; if (el) setDispSize({ w: el.clientWidth, h: el.clientHeight }); }} />
                  {markerLayout.map((m) => {
                    // Mở khoá kéo vùng crop (double-click): ẩn TẤT CẢ ký hiệu (kể cả ký hiệu bên trong vùng crop).
                    if (!markerEdit && cropRowId != null) return null;
                    // Chế độ Thêm ký hiệu: chỉ hiện ký hiệu của dòng đang chọn.
                    if (markerEdit && activeId != null && m.rowId !== activeId) return null;
                    let stateCls = "";
                    if (activeId != null) stateCls = (m.rowId === activeId ? " active" : " dim");
                    const cls = "marker" + stateCls + (m.rowId === hoverId ? " hl" : "");
                    const showDone = markerEdit && m.rowId === activeId;
                    const no = displayNo.get(m.rowId) || (m.rowIdx + 1);
                    return (<div key={m.rowId + "-" + m.instIdx} className={cls} style={{ left: m.leftPct + "%", top: m.topPct + "%" }}
                      title={no + ". " + ((rows[m.rowIdx] && rows[m.rowIdx].mon) || "—")}
                      onPointerEnter={() => setHoverId(m.rowId)} onPointerLeave={() => setHoverId((h) => (h === m.rowId ? null : h))}
                      onPointerDown={(e) => startMarker(e, m.rowId, m.instIdx)} onClick={(e) => e.stopPropagation()}
                      onDoubleClick={(e) => { e.stopPropagation(); setActiveId(m.rowId); if (!markerEdit) setCropRowId((prev) => (prev === m.rowId ? null : m.rowId)); }}>
                      {no}
                      {showDone && (
                        <button className="marker-del" title="Xóa bớt ký hiệu này" aria-label="Xóa ký hiệu này"
                          onPointerDown={(e) => e.stopPropagation()}
                          onClick={(e) => { e.stopPropagation(); deleteInstance(m.rowId, m.instIdx); }}><XIcon width={11} height={11} /></button>
                      )}
                      {showDone && (
                        <button className="marker-done" title="Xong — thoát chỉnh crop" aria-label="Xong"
                          onPointerDown={(e) => e.stopPropagation()}
                          onClick={(e) => { e.stopPropagation(); setMarkerEdit(false); }}><Check width={12} height={12} /></button>
                      )}
                    </div>);
                  })}
                  {(() => {
                    const cropId = markerEdit ? activeId : cropRowId;   // markerEdit dùng dòng đang chọn; ngoài ra dùng dòng vừa double-click
                    if (cropId == null) return null;
                    const r = rows.find((x) => x.id === cropId);
                    if (!r) return null;
                    const unlocked = !markerEdit; // mở khoá qua double-click (không ở chế độ Thêm ký hiệu)
                    return r.instances.map((b, i) => ({ b, i })).filter(({ b }) => b.imgId === activeImgId).map(({ b, i }) => (
                      <div key={"crop-" + r.id + "-" + i} className={"cropbox" + (unlocked ? " unlocked" : "")}
                        style={{ left: (b.x1 * 100) + "%", top: (b.y1 * 100) + "%", width: ((b.x2 - b.x1) * 100) + "%", height: ((b.y2 - b.y1) * 100) + "%" }}
                        onPointerDown={unlocked ? (e) => startMarker(e, r.id, i) : undefined}
                        onClick={unlocked ? (e) => e.stopPropagation() : undefined}>
                        {["nw", "n", "ne", "e", "se", "s", "sw", "w"].map((d) => (
                          <span key={d} className={"crop-h h-" + d}
                            onPointerDown={(e) => startResize(e, r.id, i, d)} onClick={(e) => e.stopPropagation()} />
                        ))}
                        {unlocked && (
                          <div className="crop-tools">
                            <button className="crop-del" title="Xóa bớt ký hiệu này" aria-label="Xóa ký hiệu này"
                              onPointerDown={(e) => e.stopPropagation()}
                              onClick={(e) => { e.stopPropagation(); const remaining = (r.instances ? r.instances.length : 0) - 1; deleteInstance(r.id, i); if (remaining <= 0) setCropRowId(null); }}><XIcon width={12} height={12} /></button>
                            <button className="crop-tick" title="Xong — khoá lại vùng crop" aria-label="Xong"
                              onPointerDown={(e) => e.stopPropagation()}
                              onClick={(e) => { e.stopPropagation(); setCropRowId(null); }}><Check width={13} height={13} /></button>
                          </div>
                        )}
                      </div>
                    ));
                  })()}
                </div>
              ) : (
                <div className={"img-placeholder" + (dragOver ? " dragover" : "")}
                  onClick={openPicker}
                  onDragOver={onDragOver} onDragEnter={onDragOver} onDragLeave={onDragLeave} onDrop={onDrop}>
                  <div className="ph-title">Chọn hoặc kéo-thả NHIỀU ảnh phối cảnh để bắt đầu</div>
                  <div className="ph-sub">Chọn nhiều file cùng lúc · kéo-thả · hoặc dán ảnh (Ctrl/Cmd + V)</div>
                </div>
              )}

              {activeImage && markerEdit && (
                <div className="hint edit-on">
                  {activeIdx >= 0
                    ? "Đang chọn dòng #" + (activeIdx + 1) + " (" + (rows[activeIdx].mon || "—") + ") — bấm vùng trống để thêm ký hiệu · KÉO ký hiệu để di chuyển · bấm (không kéo) để xóa. Ký hiệu thêm sẽ gắn vào ẢNH ĐANG XEM (#" + (imgIndex(activeImgId) + 1) + ")."
                    : "Chọn một dòng trong bảng, rồi bấm lên ảnh để đặt ký hiệu."}
                </div>
              )}

              <div className="ctl-row dl-row">
                <div className="dl-menu-wrap" ref={dlMenuRef}>
                  <button className={"btn btn-ghost" + (dlMenuOpen ? " on" : "")} onClick={() => setDlMenuOpen((v) => !v)} disabled={!hasRows || !hasImages} aria-haspopup="menu" aria-expanded={dlMenuOpen}>
                    <ImageDown size={15} /> Tải Ảnh <ChevronUp size={13} />
                  </button>
                  {dlMenuOpen && (
                    <div className="dl-menu dl-menu-up" role="menu">
                      <button className="dl-menu-item" role="menuitem" disabled={!hasRows || !activeImage}
                        onClick={() => { setDlMenuOpen(false); downloadAnnotated(); }}>
                        <ImageDown size={14} /> <span>Tải 1 ảnh (ảnh đang xem)</span>
                      </button>
                      <button className="dl-menu-item" role="menuitem" disabled={!hasRows || !hasImages}
                        onClick={() => { setDlMenuOpen(false); downloadAllAnnotated(); }}>
                        <ImageDown size={14} /> <span>Tải tất cả ảnh</span>
                      </button>
                    </div>
                  )}
                </div>
              </div>
            </div>
          </div>

          {error && <p className="error">{error}</p>}
          </div>
        </section>

        <div className={"split-divider" + (splitDragging ? " dragging" : "")} onPointerDown={startSplitDrag} role="separator" aria-label="Kéo để chỉnh tỉ lệ hai cột" title="Kéo để chỉnh tỉ lệ hai cột" />

        <section className="pane pane-table">
          <div className="pane-body">

          {/* 02 · table */}
          <div>
            <div className="block-head">
              <span className="section-label">01 · Bảng Inventory & Mã vật liệu</span>
              {hasRows && <span className="count">{rows.length} món · {totalSyms} ký hiệu</span>}
            </div>

            {hasRows && (
              <div className="sheet-tabs" role="tablist">
                <button className={"sheet-tab" + (curSheet === ALL_SHEET ? " on" : "")} onClick={() => setActiveSheet(ALL_SHEET)}>
                  Tất cả <span className="tab-n">{rows.length}</span>
                </button>
                {sheetTabs.map((t) => (
                  <button key={t.key} className={"sheet-tab" + (curSheet === t.key ? " on" : "")} onClick={() => setActiveSheet(t.key)} title={t.key}>
                    {t.key} <span className="tab-n">{t.count}</span>
                  </button>
                ))}
              </div>
            )}

            {hasRows && images.length > 1 && (
              <div className="img-filter" role="tablist" aria-label="Lọc theo ảnh phối cảnh">
                <span className="imgf-label">Lọc theo ảnh:</span>
                <button className={"imgf-chip" + (imgFilter == null ? " on" : "")} onClick={() => setImgFilter(null)}>Tất cả ảnh</button>
                {images.map((im, i) => {
                  const cnt = rows.filter((r) => r.instances.some((b) => b.imgId === im.id)).length;
                  return (
                    <button key={im.id} className={"imgf-chip" + (imgFilter === im.id ? " on" : "")}
                      onClick={() => setImgFilter((v) => (v === im.id ? null : im.id))}
                      title={"Chỉ hiện ký hiệu thuộc ảnh phối cảnh #" + (i + 1)}>
                      Ảnh {i + 1}<span className="tab-n">{cnt}</span>
                    </button>
                  );
                })}
              </div>
            )}

            {hasRows && (
              <div className="filterbar">
                <div className="searchbox">
                  <Search width={15} height={15} />
                  <input placeholder="Tìm theo mã, món, vật liệu, vị trí…" value={search} onChange={(e) => setSearch(e.target.value)} />
                  {search && <button className="clr" title="Xoá tìm" aria-label="Xoá tìm" onClick={() => setSearch("")}><XIcon width={14} height={14} /></button>}
                </div>
                <button className={"chip-toggle" + (onlyLow ? " on" : "")} onClick={() => setOnlyLow((v) => !v)} title="Chỉ hiện dòng độ tin cậy Thấp">
                  <span className="dotc" /> Chỉ tin cậy Thấp{lowN ? " (" + lowN + ")" : ""}
                </button>
                <button className={"chip-toggle" + (onlyUnpinned ? " on" : "")} onClick={() => setOnlyUnpinned((v) => !v)} title="Chỉ hiện dòng chưa gắn ký hiệu (SL=0, không có hình/chấm)">
                  <span className="dotc" style={{ background: "var(--ac)" }} /> Chưa gắn KH{unpinnedN ? " (" + unpinnedN + ")" : ""}
                </button>
                {(q || onlyLow || onlyUnpinned || imgFilter != null || curSheet !== ALL_SHEET) && <span className="filter-note">Hiện {visibleCount}/{rows.length} dòng</span>}
              </div>
            )}

            {selCount > 0 && (
              <div className="selbar">
                <span className="seln">Đã chọn {selCount} dòng</span>
                <button className="btn btn-ghost" onClick={mergeSelected} disabled={selCount < 2}><Combine size={15} /> Gộp dòng đã chọn</button>
                <button className="btn btn-ghost" onClick={splitSelected} disabled={!canSplit}><Scissors size={15} /> Tách dòng</button>
                <div className="spacer" />
                <button className="btn btn-ghost" onClick={clearSelect}><XIcon size={15} /> Bỏ chọn</button>
              </div>
            )}

            <div className="grp-wrap">
              {hasRows && (
                <div className="grp-caption">
                  <span className="grp-cap-sel" />
                  <span className="grp-cap-stt">#</span>
                  <span className="grp-cap-thumb">Ảnh</span>
                  <span className="grp-cap-code">Mã</span>
                  <span className="grp-cap-main">Món / Vật liệu</span>
                  <span className="grp-cap-sl">SL</span>
                  <span className="grp-cap-select">Tin cậy</span>
                  <span className="grp-cap-act" />
                </div>
              )}
              {hasRows ? (
                groupRowsByNhom(rows.filter(passFilter)).map((g) => (
                  <div key={g.key}>
                    <div className="grp-head" style={{ color: GROUP_COLOR[g.key] || "var(--tx3)" }}>
                      <span className="grp-dot" style={{ background: GROUP_COLOR[g.key] || "var(--tx3)" }} />
                      {g.key} · {g.rows.length}
                    </div>
                    {g.rows.map((r) => (
                      <div key={r.id}
                        className={"grp-row" + (r.do_tin_cay === "Thấp" ? " row-low" : "") + (r.id === activeId ? " active-row" : "") + (r.id === hoverId ? " hl-row" : "")}
                        style={{ borderLeftColor: GROUP_COLOR[g.key] || "var(--tx3)" }}
                        onClick={() => selectRow(r)}
                        onMouseEnter={() => setHoverId(r.id)} onMouseLeave={() => setHoverId((h) => (h === r.id ? null : h))}>
                        <input type="checkbox" className="axchk" aria-label="Chọn dòng" checked={selected.has(r.id)} onClick={(e) => e.stopPropagation()} onChange={() => toggleSelect(r.id)} />
                        <span className="grp-stt">{displayNo.get(r.id) || (rows.indexOf(r) + 1)}</span>
                        {r.thumb
                          ? <img className="grp-thumb" src={r.thumb} alt={r.mon} title="Bấm để phóng to soi crop" onClick={(e) => { e.stopPropagation(); openLightbox(r); }} />
                          : (r.instances.length === 0
                            ? <button className="pin-btn" title={srcNums(r).length ? ("Bóc từ ảnh " + srcNums(r).join(", ") + " — bấm để mở đúng ảnh đó và gắn ký hiệu") : "Chưa rõ ảnh nguồn — bấm để gắn lên ảnh đang xem"} onClick={(e) => { e.stopPropagation(); pinRow(r); }}><Plus width={12} height={12} /> gắn{srcNums(r).length ? " · " + srcNums(r).join(",") : ""}</button>
                            : <span className="grp-thumb-ph" />)}
                        <input className="grp-input grp-code" aria-label="Mã" placeholder="—" value={r.ma} onClick={(e) => e.stopPropagation()} onChange={(e) => updateRow(r.id, "ma", e.target.value)} onBlur={() => setRows((rs) => sortRows(rs))} />
                        <div className="grp-main">
                          <input className="grp-input grp-mon" aria-label="Món" placeholder="Tên món…" value={r.mon} onClick={(e) => e.stopPropagation()} onChange={(e) => updateRow(r.id, "mon", e.target.value)} />
                          <input className="grp-input grp-vl" aria-label="Vật liệu" placeholder="Vật liệu / finish…" value={r.vat_lieu} onClick={(e) => e.stopPropagation()} onChange={(e) => updateRow(r.id, "vat_lieu", e.target.value)} />
                        </div>
                        <input type="number" min="0" className={"grp-input grp-sl" + (qtyOf(r) === 0 ? " qty-zero" : "")} aria-label="Số lượng"
                          title="Số lượng do AI đếm — có thể sửa tay"
                          value={r.soLuong != null ? r.soLuong : (r.instances.length || "")}
                          onClick={(e) => e.stopPropagation()}
                          onChange={(e) => updateRow(r.id, "soLuong", e.target.value)} />
                        <select className="grp-select" aria-label="Độ tin cậy" value={r.do_tin_cay} onClick={(e) => e.stopPropagation()} onChange={(e) => updateRow(r.id, "do_tin_cay", e.target.value)}>
                          {(r.do_tin_cay && !TINCAY_OPTS.includes(r.do_tin_cay) ? [r.do_tin_cay, ...TINCAY_OPTS] : TINCAY_OPTS).map((o) => <option key={o} value={o}>{o}</option>)}
                        </select>
                        <button className="icon-danger" aria-label="Xóa dòng" onClick={(e) => { e.stopPropagation(); deleteRow(r.id); }}><Trash2 size={14} /></button>
                      </div>
                    ))}
                  </div>
                ))
              ) : (
                <div className="empty"><div className="eyebrow">Chưa có dữ liệu</div><div className="msg">Tải một hoặc nhiều ảnh phối cảnh và bấm <b>Phân tích tất cả</b> để bắt đầu — hoặc <b>Thêm dòng</b> để nhập tay.</div></div>
              )}
              {hasRows && visibleCount === 0 && (
                <div className="empty"><div className="msg">Không có dòng nào khớp bộ lọc. <b onClick={() => { setSearch(""); setOnlyLow(false); setOnlyUnpinned(false); setImgFilter(null); setActiveSheet(ALL_SHEET); }} style={{ cursor: "pointer", color: "var(--ac2)" }}>Xoá bộ lọc</b></div></div>
              )}
            </div>

            <div className="toolbar">
              <button className="btn btn-ghost" onClick={addRow}><Plus size={15} /> Thêm dòng</button>
              <button className="btn btn-ghost" onClick={recode} disabled={!hasRows}><Hash size={15} /> Gộp trùng & đánh mã lại</button>
              <button className="btn btn-ghost" onClick={undo} disabled={!undoStack.length} title={undoStack.length ? ("Hoàn tác: " + undoStack[undoStack.length - 1].label + " (Ctrl/Cmd+Z)") : "Không có gì để hoàn tác"}><Undo2 size={15} /> Hoàn tác{undoStack.length ? " (" + undoStack.length + ")" : ""}</button>
              <div className="spacer" />
              <button className="btn btn-ghost" onClick={exportExcel} disabled={!hasRows}><Download size={15} /> Xuất Excel (bảng)</button>
            </div>
            {status && <div className="status"><span style={{ width: 6, height: 6, borderRadius: 6, background: "var(--ac2)", display: "inline-block" }} />{status}</div>}
          </div>
          </div>
        </section>
      </div>

      {/* C3 · lightbox soi crop */}
      {lightbox && (
        <div className="lb-backdrop" onClick={() => setLightbox(null)}>
          <div className="lb-card" onClick={(e) => e.stopPropagation()}>
            <div className="lb-head">
              <span className="lb-code">{lightbox.code}</span>
              <span className="lb-title">{lightbox.title}</span>
              <button className="lb-close" onClick={() => setLightbox(null)} aria-label="Đóng"><XIcon width={16} height={16} /></button>
            </div>
            <div className="lb-body">
              <img src={lightbox.src} alt={lightbox.title} />
              <div className="lb-meta">{lightbox.meta}</div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default InventoryExtractor;
