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
const ArrowLR = (p) => svg(p, P("m16 3 4 4-4 4"), P("M20 7H4"), P("m8 21-4-4 4-4"), P("M4 17h16"));

/* ================= UNCHANGED LOGIC (verbatim) ================= */
const NHOM_OPTS = ["Nội thất", "Đèn", "Vật liệu bề mặt", "Cửa & Vách kính", "Hardware", "Trang trí"];
const TINCAY_OPTS = ["Cao", "Thấp"];

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
- do_tin_cay: Cao / Thấp (chỉ 2 mức; không chắc chắn -> Thấp).
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
- Cá thể ĐỐI XỨNG GƯƠNG (lật trái–phải) hoặc XOAY sang hướng khác của CÙNG một thiết kế = CÙNG một loại: gộp về 1 dòng và tăng so_luong (vd 2 táp đầu giường đối xứng, 2 ghế cùng mẫu quay khác hướng).
- Vật liệu bề mặt / chi tiết dạng đường: 1 box đại diện (không dùng để đếm).
- Không bịa. Không dùng "|" hay ";" trong chữ.

KHÔNG BÓC (bỏ qua hoàn toàn, không tạo dòng): người/nhân viên/lễ tân (là người), màn hình & máy tính, laptop, bàn phím, và sách. (Lưu ý: quầy lễ tân, kệ/tủ sách vẫn là nội thất — vẫn bóc.)
TRANG TRÍ: chỉ bóc TRANH/ẢNH/đồ nghệ thuật (mã AW) và CÂY CẢNH/CHẬU (mã GW). KHÔNG bóc bình/lọ, khay/khay trà, gối tựa, giỏ/rổ mây đan, nến và đồ trang trí lặt vặt khác.

Ví dụ dòng:
PT|Vật liệu bề mặt|Sơn đen mờ|Tường|Vách tivi|Cao||0.05,0.10,0.40,0.85
MT|Vật liệu bề mặt|Nẹp đồng|Nẹp dọc trang trí|Vách tivi|Cao|Theo m dài|0.44,0.20,0.46,0.70
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
  "Nếu có dòng 'Gợi ý loại vật' bên dưới, hãy dùng nó để xác định đúng vật và đúng nhóm.\n" +
  "KHÔNG BÓC (nếu crop là các thứ này thì trả về RỖNG, không in dòng nào): người/nhân viên, màn hình/máy tính, laptop, bàn phím, sách.";

// Prompt cho PASS BỀ MẶT — ĐỘC LẬP (KHÔNG kế thừa PROMPT), chỉ lấy vật liệu bề mặt/hoàn thiện.
// Đồ rời (nội thất/đèn/thiết bị/gương/thảm) đã do tầng detect Gemini + đọc crop lo -> ở đây CẤM liệt kê
// để tránh TRÙNG LẶP và để model tập trung đúng phần bề mặt mà detector hay bỏ sót.
const SURFACE_PROMPT =
`Bạn là chuyên gia bóc tách VẬT LIỆU BỀ MẶT / HOÀN THIỆN từ ảnh phối cảnh (render) nội thất, phục vụ báo giá.
CHỈ liệt kê các LỚP HOÀN THIỆN CỐ ĐỊNH gắn với sàn – tường – trần – cửa:
- Sàn: gỗ (WF), gạch (TL), đá (ST), thảm trải sàn rời (RUG)
- Tường: ốp gỗ/lam (WD), veneer (VN), laminate (LM), MFC (MF), giấy dán tường (WP), sơn (PT), sơn hiệu ứng (PE), đá ốp (ST), kính/gương ốp mảng (GL)
- Trần: sơn/tấm trần, mảng đèn hắt (đặc điểm trần)
- Rèm/màn: rèm vải (CT), màn sáo (BL)
- Nẹp/phào (MD); vách kính/lan can kính (GL)

TUYỆT ĐỐI KHÔNG liệt kê đồ NỘI THẤT RỜI (giường, ghế, sofa, tủ, bàn, ottoman, kệ...), ĐÈN trang trí, THIẾT BỊ, GƯƠNG soi rời, cây/chậu, vật trang trí. Những thứ đó đã xử lý ở lượt khác — nếu liệt kê sẽ gây TRÙNG LẶP. (Ngoại lệ: THẢM trải sàn thì ĐƯỢC liệt kê ở đây với mã RUG.) Nếu ảnh không có mảng bề mặt nào rõ thì trả về rỗng.

CHỈ trả về các dòng phân tách bằng "|" — KHÔNG header, KHÔNG markdown, KHÔNG giải thích. Mỗi dòng gồm 9 cột theo đúng thứ tự:
ma|nhom|mon|vat_lieu_finish|so_luong|vi_tri|do_tin_cay|ghi_chu|boxes
- ma: tiền tố mã theo loại bề mặt (WF, TL, ST, WD, VN, LM, MF, WP, PT, PE, GL, SF, CT, BL, MD, RUG).
- nhom: LUÔN ghi đúng "Vật liệu bề mặt".
- mon: tên mảng bề mặt ngắn gọn (vd "Sàn gỗ", "Ốp tường lam gỗ", "Rèm cửa", "Trần thạch cao", "Sơn tường", "Thảm trải sàn").
- vat_lieu_finish: mô tả vật liệu/màu/vân ngắn gọn.
- so_luong: 1.
- vi_tri: khu vực (vd "Sàn phòng ngủ", "Tường đầu giường", "Cửa sổ").
- do_tin_cay: Cao / Thấp (chỉ 2 mức; không chắc chắn -> Thấp).
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

// Ép độ tin cậy về ĐÚNG 2 trạng thái: Cao / Thấp. "Trung bình" cũ và mọi giá trị lạ -> "Thấp" (để user review).
function normTinCay(v) {
  const s = stripVN(v || "").trim();
  if (s.includes("cao")) return "Cao";
  return "Thấp";
}

// --- Loại trừ vật thể KHÔNG bóc tách: người, màn hình/máy tính, laptop, sách ---
// Nhãn tiếng Anh do Gemini gán (mạnh nhất, chặn TRƯỚC khi crop).
const EXCLUDE_LABELS_EN = [
  "person", "people", "human", "man", "woman", "boy", "girl", "child", "kid",
  "staff", "worker", "employee", "receptionist", "guest", "figure", "silhouette",
  "monitor", "computer monitor", "computer", "desktop", "desktop computer", "pc", "imac",
  "laptop", "notebook computer", "keyboard",
  "book", "books",
];
function labelExcluded(label) {
  const s = " " + stripVN(label).replace(/[^a-z0-9\s]/g, " ").replace(/\s+/g, " ").trim() + " ";
  return EXCLUDE_LABELS_EN.some((w) => s.includes(" " + w + " "));
}
// Lưới an toàn theo TÊN món tiếng Việt (phòng khi lọt qua nhãn Gemini). Có chốt chặn để KHÔNG
// nhầm sang nội thất hợp lệ: "quầy lễ tân", "kệ sách", "tủ sách"… vẫn được giữ.
function monExcluded(mon) {
  const t = " " + normTokens(mon) + " "; // token đã bỏ dấu + sort
  const has = (w) => t.includes(" " + w + " ");
  const furn = has("quay") || has("ban") || has("ghe") || has("tu") || has("ke") || has("gia") || has("vach") || has("giuong");
  // người / nhân viên / lễ tân (là người, không phải quầy)
  if ((has("nguoi") || has("nhanvien") || (has("nhan") && has("vien")) || (has("le") && has("tan"))) && !furn) return true;
  // màn hình / máy tính / laptop / bàn phím
  if ((has("man") && has("hinh")) || (has("may") && has("tinh")) || has("laptop") || (has("ban") && has("phim"))) return true;
  // sách (không phải kệ/tủ/giá sách)
  if (has("sach") && !furn) return true;
  return false;
}

// --- Chuẩn hoá nhóm TRANG TRÍ ---
// Yêu cầu: tranh/ảnh/đồ nghệ thuật -> mã AW; chậu + cây cảnh -> mã GW.
// Loại bỏ các món trang trí lặt vặt: bình/lọ, khay (trà), gối, sách, giỏ/rổ mây đan, nến...
// Trả về: item (giữ/đổi mã) hoặc null (loại bỏ).
function normDecor(it) {
  // KHÔNG đụng các nhóm có ngữ nghĩa rõ ràng khác (tránh nhầm "tường" -> "tượng", "đèn cây" -> "cây"):
  const nhom = String(it.nhom || "");
  if (nhom === "Đèn" || nhom === "Vật liệu bề mặt" || nhom === "Cửa & Vách kính" || nhom === "Hardware") return it;
  const t = " " + normTokens(it.mon) + " ";
  const has = (w) => t.includes(" " + w + " ");
  const isArt = has("tranh") || has("poster") || has("canvas") || (has("nghe") && has("thuat"));
  const isPlant = has("chau") || has("bonsai") || has("kieng") || ((has("cay") || has("caycanh")) && !has("den"));
  // Chỉ xử lý trong phạm vi TRANG TRÍ (hoặc rõ ràng là tranh/cây, dù nhóm bị gán lệch).
  if (nhom !== "Trang trí" && !isArt && !isPlant) return it;
  if (isPlant) return { ...it, prefix: "GW", nhom: "Trang trí", ma: "" };   // chậu + cây cảnh -> GW
  if (isArt) return { ...it, prefix: "AW", nhom: "Trang trí", ma: "" };      // tranh/ảnh/nghệ thuật -> AW
  // Trong nhóm trang trí: loại món lặt vặt (bình/lọ, khay, gối, sách, giỏ/rổ mây, nến...)
  if (has("binh") || has("lo") || has("khay") || has("goi") || has("sach") || has("gio") || has("ro") || has("nen")) return null;
  return it;
}
const normalizeDecorList = (items) => (items || []).map(normDecor).filter(Boolean);

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

// Đếm SỐ CÁ THỂ THẬT trong 1 nhóm (cluster) đã gom bằng thị giác.
// Vấn đề: Gemini hay khoanh MỘT vật lớn thành nhiều box chồng nhau (giường = khung + đầu giường + chăn).
// Nếu cộng thẳng số box -> đếm dư (giường thành SL 3). Cách xử lý: các box CHỒNG nhau = CÙNG một vật
// -> gộp làm 1; chỉ box TÁCH RỜI mới tính là cá thể riêng. Trả về [{box(lớn nhất mỗi cụm), count}],
// đã sort theo diện tích giảm dần (phần tử đầu = box đại diện lớn nhất cả nhóm).
function spatialInstances(boxes) {
  const items = (boxes || []).map((b) => ({ box: b, count: (Number.isInteger(b.count) && b.count > 0) ? b.count : 1 }));
  items.sort((a, b) => _boxArea(b.box) - _boxArea(a.box)); // vật to trước -> làm "mỏ neo" của cụm
  const groups = [];
  for (const it of items) {
    const hit = groups.find((g) => boxesOverlap(g.box, it.box));
    if (hit) hit.count = Math.max(hit.count, it.count);   // chồng nhau = cùng 1 vật -> KHÔNG cộng
    else groups.push({ box: it.box, count: it.count });    // tách rời = cá thể mới
  }
  return groups;
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
    const do_tin_cay = normTinCay(parts[6]);
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
  const rank = { "Cao": 2, "Thấp": 1 };
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
  const rank = { "Cao": 2, "Thấp": 1 };
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

// Logo ARTIUS (chữ đen nền trong suốt) — nhúng sẵn để dùng trong Excel & header in PDF của file HTML (chạy offline).
const ARTIUS_LOGO_PNG = "data:image/png;base64,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";
const ARTIUS_LOGO_W = 576, ARTIUS_LOGO_H = 192;

/* ===================== HẠ TẦNG API: retry + backoff + đếm call/cost ===================== */
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const backoff = (attempt) => Math.round((600 * Math.pow(2, attempt)) * (0.8 + Math.random() * 0.4)); // ~0.6s,1.2s,2.4s + jitter
// Bộ đếm dùng chung cho MỌI lời gọi /api/* (module-level). Component đọc qua snapshot.
const API_STATS = { calls: 0, inTok: 0, outTok: 0, at: 0 };
let API_ON_UPDATE = null;   // component gắn callback để cập nhật UI realtime
// Giá ước tính (USD / 1 triệu token) — CHỈ để ước lượng, chỉnh nếu bảng giá đổi.
const PRICE_IN = 3, PRICE_OUT = 15;
function estCostUSD() { return API_STATS.inTok / 1e6 * PRICE_IN + API_STATS.outTok / 1e6 * PRICE_OUT; }
function accountUsage(data) { const u = data && data.usage; if (u) { API_STATS.inTok += u.input_tokens || 0; API_STATS.outTok += u.output_tokens || 0; } }
function resetApiStats() { API_STATS.calls = 0; API_STATS.inTok = 0; API_STATS.outTok = 0; API_STATS.at = 0; if (API_ON_UPDATE) API_ON_UPDATE(); }

// POST JSON tới /api/* với retry+backoff. Trả về JSON đã parse khi OK; NÉM lỗi rõ khi thất bại.
// Chỉ retry 429 / 5xx / lỗi mạng. 400 (hết credit) / 401 / 403 -> lỗi vĩnh viễn, không retry.
async function apiPost(url, body, opts = {}) {
  const tries = Math.max(1, opts.tries || 3);
  let lastErr = null;
  for (let a = 0; a < tries; a++) {
    let res;
    try {
      res = await fetch(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
    } catch (e) {
      lastErr = new Error("Lỗi mạng khi gọi API (fetch thất bại).");
      if (a < tries - 1) { await sleep(backoff(a)); continue; }
      throw lastErr;
    }
    if (res.ok) {
      const data = await res.json();
      API_STATS.calls++; accountUsage(data); API_STATS.at = Date.now();
      if (API_ON_UPDATE) API_ON_UPDATE();
      return data;
    }
    let detail = "";
    try { const t = await res.text(); try { const j = JSON.parse(t); detail = (j && j.error && j.error.message) || (j && j.message) || ""; } catch { detail = t.slice(0, 200); } } catch { /* bỏ qua */ }
    const err = new Error("HTTP " + res.status + (detail ? " · " + detail : ""));
    if ((res.status === 429 || res.status >= 500) && a < tries - 1) { lastErr = err; await sleep(backoff(a)); continue; }
    throw err;   // 4xx (kể cả 400 hết credit) -> dừng ngay
  }
  throw lastErr || new Error("Thất bại sau khi thử lại.");
}

/* ===================== CACHE KẾT QUẢ PHÂN TÍCH THEO HASH ẢNH ===================== */
// Hash nhanh (djb2, lấy mẫu thưa vì base64 rất dài) + độ dài -> khóa cache trong phiên.
function hashStr(s) {
  let h = 5381; const step = s.length > 20000 ? 251 : 1;
  for (let i = 0; i < s.length; i += step) h = ((h << 5) + h + s.charCodeAt(i)) | 0;
  return (h >>> 0).toString(36) + "_" + s.length;
}
const CACHE_VER = "v1-sonnet5";                 // đổi khi prompt/model đổi để vô hiệu cache cũ
const ANALYZE_CACHE = new Map();                // hash -> items (bản clone, KHÔNG giữ id); mất khi reload
function cacheClone(items) {
  return (items || []).map((it) => ({ ...it, id: 0, instances: (it.instances || []).map((b) => ({ ...b })), memberBoxes: (it.memberBoxes || []).map((b) => ({ ...b })) }));
}
function withFreshIds(items) {
  return (items || []).map((it) => ({ ...it, id: nextId(), instances: (it.instances || []).map((b) => ({ ...b })), memberBoxes: (it.memberBoxes || []).map((b) => ({ ...b })) }));
}

/* ===================== IndexedDB: lưu/khôi phục phiên ===================== */
function idbOpen() {
  return new Promise((res, rej) => {
    try {
      const r = indexedDB.open("speclens", 1);
      r.onupgradeneeded = () => { if (!r.result.objectStoreNames.contains("kv")) r.result.createObjectStore("kv"); };
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error || new Error("idb open error"));
    } catch (e) { rej(e); }
  });
}
async function idbSet(key, val) {
  const db = await idbOpen();
  return new Promise((res, rej) => { const tx = db.transaction("kv", "readwrite"); tx.objectStore("kv").put(val, key); tx.oncomplete = () => res(); tx.onerror = () => rej(tx.error); });
}
async function idbGet(key) {
  const db = await idbOpen();
  return new Promise((res, rej) => { const tx = db.transaction("kv", "readonly"); const rq = tx.objectStore("kv").get(key); rq.onsuccess = () => res(rq.result); rq.onerror = () => rej(rq.error); });
}
async function idbDel(key) {
  const db = await idbOpen();
  return new Promise((res, rej) => { const tx = db.transaction("kv", "readwrite"); tx.objectStore("kv").delete(key); tx.oncomplete = () => res(); tx.onerror = () => rej(tx.error); });
}

/* Nạp ExcelJS (bản UMD, expose window.ExcelJS) từ CDN — KHÔNG cần thêm dependency vào package.json.
   ExcelJS cần thiết để NHÚNG ẢNH crop vào ô Excel (SheetJS community không hỗ trợ ảnh).
   Thử lần lượt nhiều CDN; nếu offline/chặn mạng -> reject để caller rơi về bản xuất cơ bản (không ảnh). */
let _exceljsLoading = null;
function loadExcelJS() {
  if (typeof window !== "undefined" && window.ExcelJS) return Promise.resolve(window.ExcelJS);
  if (_exceljsLoading) return _exceljsLoading;
  const urls = [
    "https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js",
    "https://unpkg.com/exceljs@4.4.0/dist/exceljs.min.js",
    "https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js",
  ];
  _exceljsLoading = new Promise((resolve, reject) => {
    let i = 0;
    const tryNext = () => {
      if (window.ExcelJS) return resolve(window.ExcelJS);
      if (i >= urls.length) { _exceljsLoading = null; return reject(new Error("Không nạp được ExcelJS (mạng?)")); }
      const s = document.createElement("script");
      s.src = urls[i++]; s.async = true;
      s.onload = () => (window.ExcelJS ? resolve(window.ExcelJS) : tryNext());
      s.onerror = () => { s.remove(); tryNext(); };
      document.head.appendChild(s);
    };
    tryNext();
  });
  return _exceljsLoading;
}

/* ===================== KẾT XUẤT HTML TƯƠNG TÁC (hướng A) =====================
   Một file .html tự chứa: nhúng ảnh (dataURL) + dữ liệu bảng + crop từng món.
   Mở offline bằng trình duyệt: bấm/rê ký hiệu -> popover chi tiết; danh sách đồng bộ 2 chiều;
   tìm kiếm, lọc nhóm, bật/tắt nhãn, chuyển ảnh bằng tab, và In/PDF (kèm legend — nền tảng cho hướng B).
   payload = { meta, images:[{id,idx,name,src}], groups:[{nhom, items:[{id,no,ma,mon,vat_lieu,sl,tin,vi_tri,ghi_chu,crop,marks:[{img,x,y}]}]}] } */
const ARTIUS_HTML_SUPP = `
/* ===== Supplement: dùng CHUNG với CSS app để file HTML hiển thị Y HỆT webapp ===== */
html,body{margin:0}
.expbar{display:flex;align-items:center;gap:12px;padding:9px 16px;border-bottom:1px solid var(--line);background:var(--panel);position:sticky;top:0;z-index:20}
.expbar .expttl{font-weight:700;font-size:14px;color:var(--tx)}
.expbar .expmeta{font-size:11px;color:var(--mut);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.expbar .btn{background:#18202f;border:1px solid var(--line);color:var(--tx2);border-radius:8px;padding:7px 12px;font-size:12.5px;cursor:pointer;display:inline-flex;align-items:center;gap:6px}
.expbar .btn:hover{border-color:var(--ac)}
/* nút In/Xuất PDF nổi bật (sáng lên) — dùng accent gradient ARTIUS */
.expbar .btn.primary{background:linear-gradient(135deg,var(--ac2,#7ba3cf),var(--ac,#9dc0e6));border:1px solid var(--ac,#9dc0e6);color:var(--acink,#0c1524);font-weight:700;box-shadow:0 3px 12px rgba(123,163,207,0.35)}
.expbar .btn.primary:hover{filter:brightness(1.08)}
.expbar .sp{margin-left:auto}
#ax{display:flex;--splitpct:45%;height:calc(100vh - 47px);min-height:0;overflow:hidden}
.tabs-inner{display:contents}
#invbody input,#invbody select,#invbody .qty-btn,#invbody .axchk{pointer-events:none}
#invbody .grp-input:disabled,#invbody .grp-sl:disabled{opacity:1;-webkit-text-fill-color:currentColor;color:var(--tx2)}
#invbody .grp-select:disabled{opacity:1;color:var(--tx2)}
#invbody .grp-row{cursor:pointer}
.imgwrap.cropping .marker{display:none}
.imgstrip .imgtile{cursor:pointer}
/* B2: trạng thái xem bản phân trang (Paged.js) ngay trên màn hình trước khi in */
.print-only{display:none}
#pagedpreview{display:none;background:#3b3f46;padding:16px}
body.paged-on #screen{display:none}
body.paged-on #pagedpreview{display:block}
#pagedpreview .pagedjs_page{background:#fff;margin:0 auto 14px;box-shadow:0 4px 18px rgba(0,0,0,.4)}
/* ===== IN cơ bản (B1 · fallback khi không có Paged.js) — mỗi ảnh 1 trang: ảnh trên, legend dưới ===== */
@page{size:A4 landscape;margin:8mm}
@media print{
 html,body{background:#fff!important;color:#000!important}
 .expbar{display:none!important}
 body:not(.paged-on) #screen{display:none!important}
 body:not(.paged-on) .print-only{display:block!important}
 body.paged-on #screen,body.paged-on #printroot{display:none!important}
 body.paged-on #pagedpreview{display:block!important;background:#fff;padding:0}
 body.paged-on #pagedpreview .pagedjs_page{box-shadow:none;margin:0}
 .print-only .ppage{break-after:page;page-break-after:always}
 .print-only .ppage:last-child{break-after:auto;page-break-after:auto}
 .print-only .ptitle{display:flex;justify-content:space-between;align-items:flex-end;border-bottom:2px solid #222;padding-bottom:6px;margin-bottom:10px}
 .print-only .ptitle .pt-l{display:flex;align-items:center;gap:12px}
 .print-only .ptitle .plogo{height:34px;width:auto;display:block}
 .print-only .ptitle .pt-l b{font-size:15px}.print-only .ptitle .pt-l span{display:block;font-size:10px;color:#444;margin-top:2px}
 .print-only .ptitle .pt-r{font-size:11px;font-weight:700;color:#333;white-space:nowrap;padding-left:12px}
 /* Bố cục 2 cột mỗi trang: ảnh (trái) + bảng inventory (phải), vừa 1 trang A4 landscape */
 .print-only .prow{display:flex;gap:6mm;align-items:flex-start}
 .print-only .pcol-img{flex:0 0 54%;max-width:54%;min-width:0}
 .print-only .pcol-tbl{flex:1 1 auto;min-width:0}
 .print-only .pimg{text-align:center;margin:0}
 .print-only .pimg-inner{position:relative;display:inline-block;max-width:100%}
 .print-only .pimg-inner img{display:block;max-width:100%;max-height:175mm;height:auto;border:1px solid #ccc}
 .print-only .pmk{position:absolute;transform:translate(-50%,-50%);width:20px;height:20px;border-radius:50%;background:#fff;border:1.6px solid #c00;color:#c00;font-weight:700;font-size:10px;display:flex;align-items:center;justify-content:center;line-height:1}
 .print-only table.pleg{width:100%;border-collapse:collapse;font-size:10px}
 .print-only table.pleg th,.print-only table.pleg td{border:1px solid #999;padding:2px 5px;text-align:left;vertical-align:top}
 .print-only table.pleg thead th{background:#eee}.print-only table.pleg thead{display:table-header-group}
 .print-only table.pleg td.c,.print-only table.pleg th.c{text-align:center}
 .print-only table.pleg tr.lg td{background:#eef0f2;font-weight:700;text-transform:uppercase;font-size:9px}
 .print-only table.pleg tr{break-inside:avoid}
}`;
const ARTIUS_HTML_JS = `
var D=window.__DATA__||{groups:[],images:[],meta:{},colors:{}};
var COL=D.colors||{};
var items=[];D.groups.forEach(function(g){g.items.forEach(function(it){it._nhom=g.nhom;items.push(it);});});
var byId={};items.forEach(function(it){byId[it.id]=it;});
var imgs=D.images||[];
var st={cur:imgs.length?imgs[0].id:null,sel:null,hov:null,q:"",gf:"__all"};
function esc(s){s=(s==null?"":String(s));return s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");}
function col(n){return COL[n]||"#8a97a8";}
function separate(pts,minDist,w,h,iters){var p=pts.map(function(o){return{x:o.x,y:o.y};});var n=p.length,N=iters||70,it,i,j,k;for(it=0;it<N;it++){var moved=false;for(i=0;i<n;i++){for(j=i+1;j<n;j++){var dx=p[j].x-p[i].x,dy=p[j].y-p[i].y,d=Math.sqrt(dx*dx+dy*dy);if(d<minDist){if(d<0.001){var a=i*2.399963;dx=Math.cos(a);dy=Math.sin(a);d=1;}var push=(minDist-d)/2,ux=dx/d,uy=dy/d;p[i].x-=ux*push;p[i].y-=uy*push;p[j].x+=ux*push;p[j].y+=uy*push;moved=true;}}}for(k=0;k<n;k++){p[k].x=Math.max(minDist/2,Math.min(w-minDist/2,p[k].x));p[k].y=Math.max(minDist/2,Math.min(h-minDist/2,p[k].y));}if(!moved)break;}return p;}
function curImgObj(){for(var i=0;i<imgs.length;i++)if(imgs[i].id===st.cur)return imgs[i];return null;}
function marksOn(imgId){var a=[];items.forEach(function(it){it.marks.forEach(function(m){if(m.img===imgId)a.push({id:it.id,no:it.no,mon:it.mon,cx:(m.x1+m.x2)/2,cy:(m.y1+m.y2)/2});});});return a;}
function renderStrip(){var el=document.getElementById("imgstrip");if(!el)return;if(imgs.length<2){el.style.display="none";return;}var h="";imgs.forEach(function(im,i){var cnt=0;items.forEach(function(it){if(it.marks.some(function(m){return m.img===im.id;}))cnt++;});h+='<div class="imgtile'+(im.id===st.cur?" on":"")+'" data-img="'+im.id+'"><img src="'+im.src+'"><span class="idx">'+(i+1)+'</span><span class="st done"></span>'+(cnt>0?'<span class="cnt">'+cnt+' mã</span>':'')+'</div>';});el.innerHTML=h;}
function renderImage(){var im=curImgObj();var wrap=document.getElementById("imgwrap");var img=document.getElementById("baseimg");if(!im){wrap.style.display="none";return;}wrap.style.display="";img.onload=function(){renderMarkers();};if(img.getAttribute("src")!==im.src){img.setAttribute("src",im.src);}if(img.complete){renderMarkers();}setTimeout(renderMarkers,80);}
function renderMarkers(){var wrap=document.getElementById("imgwrap");var img=document.getElementById("baseimg");if(!wrap||!img)return;Array.prototype.slice.call(wrap.querySelectorAll(".marker")).forEach(function(n){n.parentNode.removeChild(n);});var W=img.clientWidth||wrap.clientWidth||1000,H=img.clientHeight||660;var mk=marksOn(st.cur);var pts=mk.map(function(m){return{x:m.cx*W,y:m.cy*H};});var sep=separate(pts,26,W,H,70);mk.forEach(function(m,i){var d=document.createElement("div");d.className="marker";d.setAttribute("data-id",m.id);d.style.left=(sep[i].x/W*100)+"%";d.style.top=(sep[i].y/H*100)+"%";d.title=m.no+". "+(m.mon||"—");d.textContent=m.no;wrap.appendChild(d);});applyStates();}
function applyStates(){var wrap=document.getElementById("imgwrap");if(wrap)Array.prototype.slice.call(wrap.querySelectorAll(".marker")).forEach(function(n){var id=n.getAttribute("data-id");n.classList.toggle("active",st.sel!=null&&id===st.sel);n.classList.toggle("dim",st.sel!=null&&id!==st.sel);n.classList.toggle("hl",id===st.hov);});Array.prototype.slice.call(document.querySelectorAll("#invbody .grp-row")).forEach(function(n){var id=n.getAttribute("data-id");n.classList.toggle("active-row",id===st.sel);n.classList.toggle("hl-row",id===st.hov);});}
function mtf(it){if(st.gf!=="__all"&&it._nhom!==st.gf)return false;if(!st.q)return true;var s=(it.no+" "+it.ma+" "+it.mon+" "+it.vat_lieu+" "+it.vi_tri).toLowerCase();return s.indexOf(st.q)>=0;}
function renderTabs(){var el=document.getElementById("invtabs");if(!el)return;var seen={},order=[];D.groups.forEach(function(g){if(!seen[g.nhom]){seen[g.nhom]=1;order.push({n:g.nhom,c:g.items.length});}});var h='<button class="sheet-tab'+(st.gf==="__all"?" on":"")+'" data-gf="__all">Tất cả <span class="tab-n">'+items.length+'</span></button>';order.forEach(function(o){h+='<button class="sheet-tab'+(st.gf===o.n?" on":"")+'" data-gf="'+esc(o.n)+'">'+esc(o.n)+' <span class="tab-n">'+o.c+'</span></button>';});el.innerHTML=h;}
function rowHtml(it){var thumb=it.crop?'<img class="grp-thumb" src="'+it.crop+'">':'<span class="grp-thumb-ph"></span>';return '<div class="grp-row'+(it.tin==="Thấp"?" row-low":"")+'" data-id="'+it.id+'" id="invrow-'+it.id+'" style="border-left-color:'+col(it._nhom)+'"><input type="checkbox" class="axchk" disabled><span class="grp-stt">'+it.no+'</span>'+thumb+'<input class="grp-input grp-code" value="'+esc(it.ma)+'" disabled><div class="grp-main"><input class="grp-input grp-mon" value="'+esc(it.mon)+'" disabled><input class="grp-input grp-vl" value="'+esc(it.vat_lieu)+'" disabled></div><div class="grp-qty"><button class="qty-btn" disabled>\u2212</button><input class="grp-input grp-sl'+(String(it.sl)==="0"?" qty-zero":"")+'" value="'+esc(it.sl)+'" disabled><button class="qty-btn" disabled>+</button></div><select class="grp-select" disabled><option>'+esc(it.tin||"Thấp")+'</option></select><div class="row-act"></div></div>';}
function renderRows(){var el=document.getElementById("invbody");if(!el)return;var h="";D.groups.forEach(function(g){var vis=g.items.filter(mtf);if(!vis.length)return;h+='<div><div class="grp-head" style="color:'+col(g.nhom)+'"><span class="grp-dot" style="background:'+col(g.nhom)+'"></span>'+esc(g.nhom)+' \u00b7 '+vis.length+'</div>';vis.forEach(function(it){h+=rowHtml(it);});h+='</div>';});el.innerHTML=h||'<div class="empty"><div class="msg">Không có dòng khớp bộ lọc.</div></div>';applyStates();}
function pickRow(id){var it=byId[id];if(!it)return;st.sel=id;var here=it.marks.some(function(m){return m.img===st.cur;});if(!here&&it.marks.length){st.cur=it.marks[0].img;renderStrip();renderImage();}else{applyStates();}var el=document.getElementById("invrow-"+id);if(el&&el.scrollIntoView)el.scrollIntoView({block:"center",behavior:"smooth"});renderCrop(it);}
function bindSearch(){var s=document.getElementById("invsearch");if(s)s.addEventListener("input",function(e){st.q=(e.target.value||"").toLowerCase().trim();renderRows();});}
function clearCrop(){var wrap=document.getElementById("imgwrap");if(!wrap)return;Array.prototype.slice.call(wrap.querySelectorAll(".cropbox")).forEach(function(n){n.parentNode.removeChild(n);});wrap.classList.remove("cropping");}
function renderCrop(it){var wrap=document.getElementById("imgwrap");if(!wrap||!it)return;clearCrop();var boxes=it.marks.filter(function(m){return m.img===st.cur;});if(!boxes.length)return;wrap.classList.add("cropping");boxes.forEach(function(b){var d=document.createElement("div");d.className="cropbox unlocked";d.style.left=(b.x1*100)+"%";d.style.top=(b.y1*100)+"%";d.style.width=((b.x2-b.x1)*100)+"%";d.style.height=((b.y2-b.y1)*100)+"%";var hs="";["nw","n","ne","e","se","s","sw","w"].forEach(function(dir){hs+='<span class="crop-h h-'+dir+'"></span>';});d.innerHTML=hs+'<div class="crop-tools"><button class="crop-del" title="Đóng" aria-label="Đóng">\u2715</button><button class="crop-tick" title="Xong" aria-label="Xong">\u2713</button></div>';d.addEventListener("click",function(e){e.stopPropagation();});var del=d.querySelector(".crop-del"),tick=d.querySelector(".crop-tick");if(del)del.addEventListener("click",function(e){e.stopPropagation();deselect();});if(tick)tick.addEventListener("click",function(e){e.stopPropagation();deselect();});wrap.appendChild(d);});}
function deselect(){st.sel=null;clearCrop();applyStates();}
function initEvents(){
 var strip=document.getElementById("imgstrip");if(strip)strip.addEventListener("click",function(e){var t=e.target.closest?e.target.closest(".imgtile"):null;if(t){st.cur=t.getAttribute("data-img");renderStrip();renderImage();}});
 var wrap=document.getElementById("imgwrap");
 if(wrap){wrap.addEventListener("click",function(e){var mk=e.target.closest?e.target.closest(".marker"):null;if(mk){pickRow(mk.getAttribute("data-id"));return;}if(e.target.closest&&e.target.closest(".cropbox"))return;deselect();});
 wrap.addEventListener("mouseover",function(e){var t=e.target.closest?e.target.closest(".marker"):null;if(t){st.hov=t.getAttribute("data-id");applyStates();}});
 wrap.addEventListener("mouseout",function(e){var t=e.target.closest?e.target.closest(".marker"):null;if(t){st.hov=null;applyStates();}});}
 var ib=document.getElementById("invbody");
 if(ib){ib.addEventListener("click",function(e){var t=e.target.closest?e.target.closest(".grp-row"):null;if(t)pickRow(t.getAttribute("data-id"));});
 ib.addEventListener("mouseover",function(e){var t=e.target.closest?e.target.closest(".grp-row"):null;if(t){st.hov=t.getAttribute("data-id");applyStates();}});
 ib.addEventListener("mouseout",function(e){var t=e.target.closest?e.target.closest(".grp-row"):null;if(t){st.hov=null;applyStates();}});}
 var tb=document.getElementById("invtabs");if(tb)tb.addEventListener("click",function(e){var b=e.target.closest?e.target.closest(".sheet-tab"):null;if(b){st.gf=b.getAttribute("data-gf");renderTabs();renderRows();}});
 bindSearch();
 var bp=document.getElementById("btnprint");if(bp)bp.addEventListener("click",function(){doPrint();});window.addEventListener("afterprint",function(){document.body.classList.remove("paged-on");var pp=document.getElementById("pagedpreview");if(pp)pp.innerHTML="";});
 window.addEventListener("keydown",function(e){if(e.key==="Escape")deselect();});
 var rt=null;window.addEventListener("resize",function(){clearTimeout(rt);rt=setTimeout(renderMarkers,120);});
}
var PAGED_CSS='@page{size:A4 landscape;margin:14mm 12mm 15mm;@top-left{content:string(rhtitle);font-size:9px;color:#666}@top-right{content:string(rhmeta);font-size:9px;color:#666}@bottom-center{content:"Trang " counter(page) " / " counter(pages);font-size:9px;color:#666}}.rh-title{string-set:rhtitle content(text)}.rh-meta{string-set:rhmeta content(text)}body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#000;background:#fff}.ppage{break-after:page}.ppage:last-child{break-after:auto}.ptitle{display:flex;justify-content:space-between;align-items:flex-end;border-bottom:2px solid #222;padding-bottom:6px;margin-bottom:10px}.ptitle .pt-l{display:flex;align-items:center;gap:12px}.ptitle .plogo{height:34px;width:auto;display:block}.ptitle .pt-l b{font-size:15px}.ptitle .pt-l span{display:block;font-size:10px;color:#444;margin-top:2px}.ptitle .pt-r{font-size:11px;font-weight:700;color:#333}.prow{display:flex;gap:6mm;align-items:flex-start}.pcol-img{flex:0 0 54%;max-width:54%;min-width:0}.pcol-tbl{flex:1 1 auto;min-width:0}.pimg{text-align:center;margin:0}.pimg-inner{position:relative;display:inline-block;max-width:100%}.pimg-inner img{display:block;max-width:100%;max-height:170mm;height:auto;border:1px solid #ccc}.pmk{position:absolute;transform:translate(-50%,-50%);width:20px;height:20px;border-radius:50%;background:#fff;border:1.6px solid #c00;color:#c00;font-weight:700;font-size:10px;display:flex;align-items:center;justify-content:center;line-height:1}table.pleg{width:100%;border-collapse:collapse;font-size:10px}table.pleg th,table.pleg td{border:1px solid #999;padding:2px 5px;text-align:left;vertical-align:top}table.pleg thead th{background:#eee}table.pleg thead{display:table-header-group}table.pleg td.c,table.pleg th.c{text-align:center}table.pleg tr.lg td{background:#eef0f2;font-weight:700;text-transform:uppercase;font-size:9px}table.pleg tr{break-inside:avoid}';
function printFlowHtml(){var m=D.meta||{};function head(sub){var meta=[m.location||"",m.client?("CĐT: "+m.client):"",m.date||""].filter(function(x){return x;}).map(esc).join("   \u00b7   ");var ttl=(m.code?esc(m.code)+" \u00b7 ":"")+esc(m.project||"Bảng bóc tách vật liệu");return '<div class="ptitle"><div class="pt-l">'+(m.logo?'<img class="plogo" src="'+m.logo+'">':'')+'<div class="pt-txt"><b class="rh-title">'+ttl+'</b><span class="rh-meta">'+meta+'</span></div></div><div class="pt-r">'+esc(sub)+'</div></div>';}
function legRows(pred){var h="";D.groups.forEach(function(g){var gi=g.items.filter(pred);if(!gi.length)return;h+='<tr class="lg"><td colspan="4">'+esc(g.nhom)+'</td></tr>';gi.forEach(function(it){h+='<tr><td class="c">'+it.no+'</td><td>'+esc(it.ma)+'</td><td>'+esc(it.mon)+'</td><td>'+esc(it.vat_lieu)+'</td></tr>';});});return h;}
var THEAD='<thead><tr><th class="c">#</th><th>Mã</th><th>Món</th><th>Vật liệu</th></tr></thead>';
var html="",total=imgs.length;
imgs.forEach(function(im,ix){var onImg=function(it){return it.marks.some(function(mk){return mk.img===im.id;});};var any=false;D.groups.forEach(function(g){if(g.items.some(onImg))any=true;});if(!any)return;var W=1000,H=650;var mk=marksOn(im.id);var pts=mk.map(function(x){return{x:x.cx*W,y:x.cy*H};});var sep=separate(pts,26,W,H,70);var marks="";mk.forEach(function(x,i){marks+='<div class="pmk" style="left:'+(sep[i].x/W*100)+'%;top:'+(sep[i].y/H*100)+'%">'+x.no+'</div>';});html+='<section class="ppage">'+head("Ảnh "+(ix+1)+"/"+total)+'<div class="prow"><div class="pcol-img"><div class="pimg"><div class="pimg-inner"><img src="'+im.src+'">'+marks+'</div></div></div><div class="pcol-tbl"><table class="pleg">'+THEAD+'<tbody>'+legRows(onImg)+'</tbody></table></div></div></section>';});
return html;}
function fillPrintRoot(){var r=document.getElementById("printroot");if(r)r.innerHTML=printFlowHtml();}
function nativePrint(){document.body.classList.remove("paged-on");fillPrintRoot();setTimeout(function(){window.print();},30);}
function runPaged(){var target=document.getElementById("pagedpreview");if(!target){nativePrint();return;}try{target.innerHTML="";document.body.classList.add("paged-on");var prev=new window.Paged.Previewer();prev.preview(printFlowHtml(),[{"artius-paged.css":PAGED_CSS}],target).then(function(){setTimeout(function(){window.print();},120);}).catch(function(){document.body.classList.remove("paged-on");nativePrint();});}catch(e){document.body.classList.remove("paged-on");nativePrint();}}
function doPrint(){if(window.__HAS_PAGED__&&window.Paged&&window.Paged.Previewer){runPaged();}else{nativePrint();}}
function init(){var m=D.meta||{};var t=document.getElementById("expttl");if(t)t.textContent=(m.code?m.code+" \u00b7 ":"")+(m.project||"Bảng bóc tách vật liệu");var mm=document.getElementById("expmeta");if(mm)mm.textContent=[m.location||"",m.client?("CĐT: "+m.client):"",m.date||""].filter(function(x){return x;}).join("   \u00b7   ");renderStrip();renderImage();renderTabs();renderRows();fillPrintRoot();initEvents();}
init();`;
function buildInteractiveHtml(payload, appCss, pagedB64) {
  const DATA = JSON.stringify(payload).replace(/</g, "\\u003c");
  const head = '<!doctype html><html lang="vi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>ARTIUS · Spec ảnh</title><style>' + (appCss || "") + ARTIUS_HTML_SUPP + '</style></head>';
  const shell = '<div id="screen"><div class="expbar"><div><div class="expttl" id="expttl"></div><div class="expmeta" id="expmeta"></div></div><span class="sp"></span><button class="btn primary" id="btnprint">In / Xuất PDF (A4)</button></div>'
    + '<div class="ax" id="ax">'
    + '<section class="pane pane-image"><div class="pane-body"><div>'
    + '<div class="block-head"><span class="section-label">02 · Ảnh phối cảnh</span></div>'
    + '<div class="panel"><div class="imgstrip" id="imgstrip"></div><div class="imgwrap" id="imgwrap"><img class="base" id="baseimg" alt="Ảnh phối cảnh"></div></div>'
    + '</div></div></section>'
    + '<section class="pane pane-table"><div class="pane-body"><div>'
    + '<div class="tbl-head"><div class="block-head"><span class="section-label">01 · Bảng Inventory & Mã vật liệu</span></div>'
    + '<div class="sheet-tabs"><div class="tabs-inner" id="invtabs"></div><div class="tabs-search"><div class="searchbox"><input id="invsearch" placeholder="Tìm mã, món, vật liệu, vị trí…"></div></div></div></div>'
    + '<div class="tbl-scroll"><div class="grp-caption"><span class="grp-cap-sel"></span><span class="grp-cap-stt">#</span><span class="grp-cap-thumb">Ảnh</span><span class="grp-cap-code">Mã</span><span class="grp-cap-main">Món / Vật liệu</span><span class="grp-cap-sl">SL</span><span class="grp-cap-select">Tin cậy</span><span class="grp-cap-act"></span></div>'
    + '<div class="grp-wrap" id="invbody"></div></div>'
    + '</div></div></section>'
    + '</div></div>'
    + '<div class="print-only" id="printroot"></div>'
    + '<div id="pagedpreview"></div>';
  const pagedLoader = pagedB64
    ? '<scr' + 'ipt>window.PagedConfig={auto:false};try{(0,eval)(decodeURIComponent(escape(atob("' + pagedB64 + '"))));window.__HAS_PAGED__=true;}catch(e){window.__HAS_PAGED__=false;}</scr' + 'ipt>'
    : '<scr' + 'ipt>window.__HAS_PAGED__=false;</scr' + 'ipt>';
  const body = '<body>' + shell
    + '<scr' + 'ipt>window.__DATA__=' + DATA + ';</scr' + 'ipt>'
    + pagedLoader
    + '<scr' + 'ipt>' + ARTIUS_HTML_JS + '</scr' + 'ipt></body></html>';
  return head + body;
}


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
Nhiệm vụ: gom những ô là CÙNG MỘT sản phẩm (cùng kiểu dáng/thiết kế/chất liệu; bỏ qua khác biệt nhỏ do góc nhìn, khoảng cách hay ánh sáng) vào chung 1 nhóm.
QUAN TRỌNG: coi bản ĐỐI XỨNG GƯƠNG (lật trái–phải) và bản XOAY sang hướng khác của cùng một thiết kế là CÙNG MỘT sản phẩm — ví dụ 2 táp/tủ đầu giường đặt đối xứng hai bên giường, 2 ghế cùng mẫu quay hướng khác nhau, 2 đèn tường đối xứng: tất cả phải nằm CHUNG 1 nhóm (rồi số lượng sẽ tự cộng). Chỉ tách nhóm khi thật sự KHÁC kiểu dáng/thiết kế.
Mỗi ô thuộc ĐÚNG 1 nhóm; mọi số 1..${n} xuất hiện đúng một lần.
CHỈ trả về JSON (không markdown, không giải thích), dạng:
{"groups":[{"members":[1,3,5],"loai":"ghế armchair","nhom":"Nội thất"},{"members":[2],"loai":"bàn trà tròn","nhom":"Nội thất"}]}
"nhom" chọn 1 trong ["Nội thất","Đèn","Vật liệu bề mặt","Cửa & Vách kính","Hardware","Trang trí"]. "loai" là tên loại ngắn gọn tiếng Việt.`;
  try {
    const data = await apiPost("/api/analyze", {
      model: "claude-sonnet-5", max_tokens: 1200,
      messages: [{ role: "user", content: [
        { type: "image", source: { type: "base64", media_type: "image/jpeg", data: montageB64 } },
        { type: "text", text: prompt },
      ] }],
    });
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

/* ===================== GỘP MÃ TƯƠNG TỰ (Hướng C: text lọc ứng viên -> Sonnet xác nhận) =====================
   Vấn đề: mergeKey gộp EXACT nên 2 vật giống nhau nhưng mô tả lệch 1 token vẫn ra 2 mã.
   Giải pháp: (1) lọc ứng viên bằng độ trùng token (rẻ, offline, recall cao),
              (2) dựng contact sheet crop đại diện rồi để Sonnet phán "có CÙNG sản phẩm không" (chính xác). */

// Tập token (bỏ dấu, hạ chữ) của 1 chuỗi.
function tokenSet(s) {
  return new Set(stripVN(s).replace(/[^a-z0-9\s]/g, " ").split(/\s+/).filter(Boolean));
}
// Jaccard giữa 2 tập token: |giao| / |hợp|. 0..1.
function jaccard(a, b) {
  if (!a.size && !b.size) return 1;
  if (!a.size || !b.size) return 0;
  let inter = 0; a.forEach((t) => { if (b.has(t)) inter++; });
  return inter / (a.size + b.size - inter);
}

// Ngưỡng lọc ứng viên (RỘNG có chủ đích — Sonnet mới là bộ lọc chính xác cuối).
const SIMILAR_MON = 0.5;     // độ trùng token TÊN MÓN tối thiểu
const SIMILAR_MAT = 0.34;    // hoặc: tên gần + vật liệu/finish trùng kha khá
// 2 dòng có "nghi giống" để đưa vào cụm ứng viên (đã cùng prefix trước khi gọi).
function pairCandidate(a, b) {
  const ma = tokenSet(a.mon), mb = tokenSet(b.mon);
  const sMon = jaccard(ma, mb);
  const sMat = jaccard(tokenSet(a.vat_lieu), tokenSet(b.vat_lieu));
  // Chốt an toàn: tên RẤT NGẮN (1 token) dễ gộp nhầm ("đèn thả" vs "đèn tường" đều có "đèn")
  // -> yêu cầu chặt hơn: gần như trùng tên VÀ vật liệu tương đồng.
  const shortName = ma.size <= 1 || mb.size <= 1;
  if (shortName) return sMon >= 0.75 && sMat >= 0.5;
  // Bình thường: tên trùng đủ cao, HOẶC tên khá gần + finish trùng kha khá.
  return sMon >= SIMILAR_MON || (sMon >= 0.34 && sMat >= SIMILAR_MAT);
}

// Gom các dòng CÙNG prefix mà pairCandidate thành cụm (union-find). Trả [[idx,...] (>=2 phần tử)].
function candidateClusters(rows) {
  const n = rows.length;
  const parent = Array.from({ length: n }, (_, i) => i);
  const find = (x) => { while (parent[x] !== x) { parent[x] = parent[parent[x]]; x = parent[x]; } return x; };
  const union = (x, y) => { const rx = find(x), ry = find(y); if (rx !== ry) parent[rx] = ry; };
  // Chỉ xét dòng đã có mã hợp lệ (đủ căn cứ) & cùng prefix.
  const pfxOf = (r) => String((r.ma && r.ma.split("-")[0]) || r.prefix || "").toUpperCase().replace(/[^A-Z]/g, "");
  for (let i = 0; i < n; i++) {
    if (!VALID_PREFIX.has(pfxOf(rows[i]))) continue;
    for (let j = i + 1; j < n; j++) {
      if (pfxOf(rows[i]) !== pfxOf(rows[j])) continue;
      if (pairCandidate(rows[i], rows[j])) union(i, j);
    }
  }
  const groups = new Map();
  for (let i = 0; i < n; i++) {
    if (!VALID_PREFIX.has(pfxOf(rows[i]))) continue;
    const r = find(i);
    if (!groups.has(r)) groups.set(r, []);
    groups.get(r).push(i);
  }
  return Array.from(groups.values()).filter((g) => g.length >= 2);
}

// Dựng contact sheet từ danh sách crop (dataURL) — xuyên ảnh. Async vì phải load ảnh. Trả dataURL jpeg.
function buildMontageFromCrops(cropUrls, cell = 240, pad = 8, maxEdge = 2200) {
  return new Promise((resolve) => {
    try {
      const n = cropUrls.length;
      if (!n) return resolve(null);
      const imgs = cropUrls.map((u) => { const im = new Image(); im.src = u; return im; });
      let pending = imgs.length;
      const done = () => {
        try {
          const cols = Math.ceil(Math.sqrt(n)), rowsN = Math.ceil(n / cols);
          let cW = cell, cH = cell;
          let W = cols * cW + (cols + 1) * pad, H = rowsN * cH + (rowsN + 1) * pad;
          const sc = Math.min(1, maxEdge / Math.max(W, H));
          if (sc < 1) { cW = Math.max(120, Math.floor(cW * sc)); cH = cW; W = cols * cW + (cols + 1) * pad; H = rowsN * cH + (rowsN + 1) * pad; }
          const c = document.createElement("canvas"); c.width = W; c.height = H;
          const ctx = c.getContext("2d");
          ctx.fillStyle = "#0d1119"; ctx.fillRect(0, 0, W, H);
          for (let i = 0; i < n; i++) {
            const col = i % cols, row = Math.floor(i / cols);
            const x = pad + col * (cW + pad), y = pad + row * (cH + pad);
            ctx.fillStyle = "#0d1119"; ctx.fillRect(x, y, cW, cH);
            const im = imgs[i];
            if (im && im.naturalWidth) {
              const scale = Math.min(cW / im.naturalWidth, cH / im.naturalHeight), dw = im.naturalWidth * scale, dh = im.naturalHeight * scale;
              ctx.drawImage(im, x + (cW - dw) / 2, y + (cH - dh) / 2, dw, dh);
            }
            const label = String(i + 1);
            ctx.font = "bold 18px sans-serif";
            const tw = ctx.measureText(label).width;
            ctx.fillStyle = "rgba(0,0,0,0.78)"; ctx.fillRect(x + 3, y + 3, tw + 12, 24);
            ctx.fillStyle = "#fff"; ctx.textBaseline = "top"; ctx.fillText(label, x + 9, y + 6);
            ctx.strokeStyle = "rgba(255,255,255,0.16)"; ctx.lineWidth = 1; ctx.strokeRect(x + 0.5, y + 0.5, cW - 1, cH - 1);
          }
          resolve(c.toDataURL("image/jpeg", 0.85));
        } catch (e) { resolve(null); }
      };
      imgs.forEach((im) => {
        const fin = () => { if (--pending <= 0) done(); };
        if (im.complete) fin(); else { im.onload = fin; im.onerror = fin; }
      });
    } catch (e) { resolve(null); }
  });
}

// Sonnet xác nhận (cross-image): các ô nào là CÙNG một sản phẩm. Dùng lại parseGroups.
async function confirmSameProduct(montageB64, n) {
  const prompt =
`Ảnh kèm theo là 1 lưới (contact sheet) gồm ${n} ô, mỗi ô đánh SỐ ở góc trên-trái (1..${n}). Các ô là crop của các vật thể LẤY TỪ NHIỀU ẢNH nội thất khác nhau, đã được sàng lọc là "trông na ná nhau".
Nhiệm vụ: gom những ô là CÙNG MỘT sản phẩm/vật liệu (cùng kiểu dáng, thiết kế, chất liệu, finish; bỏ qua khác biệt nhỏ do góc chụp, khoảng cách, ánh sáng, hay ảnh khác nhau) vào chung 1 nhóm.
Chỉ gom khi bạn TỰ TIN chúng là cùng một loại (tương đồng > ~50%). Nếu nghi ngờ khác kiểu/khác chất liệu -> ĐỂ RIÊNG (thà tách còn hơn gộp nhầm). Coi bản đối xứng gương / xoay hướng khác của cùng thiết kế là CÙNG một loại.
Mỗi ô thuộc ĐÚNG 1 nhóm; mọi số 1..${n} xuất hiện đúng một lần. Ô đứng một mình -> nhóm 1 phần tử.
CHỈ trả về JSON (không markdown, không giải thích), dạng:
{"groups":[{"members":[1,3],"loai":"ghế bành đơn"},{"members":[2],"loai":"ghế khác"}]}`;
  try {
    const data = await apiPost("/api/analyze", {
      model: "claude-sonnet-5", max_tokens: 1000,
      messages: [{ role: "user", content: [
        { type: "image", source: { type: "base64", media_type: "image/jpeg", data: montageB64 } },
        { type: "text", text: prompt },
      ] }],
    });
    if (data && data.type === "error") return null;
    const textOut = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n");
    return parseGroups(textOut, n);
  } catch (e) { return null; }
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
.marker { position:absolute; transform:translate(-50%,-50%); width:24px; height:24px; padding:0; border-radius:50%; background:rgba(255,255,255,0.10); color:#fff;
  font-size:11px; font-weight:700; display:flex; align-items:center; justify-content:center; border:1.5px solid rgba(255,255,255,0.35); backdrop-filter:blur(5px); -webkit-backdrop-filter:blur(5px); box-shadow:0 2px 8px rgba(0,0,0,.5); cursor:pointer; line-height:1; touch-action:none; opacity:.5; box-sizing:border-box; }
.marker.dim { opacity:.42; } .marker.active { background:rgba(255,255,255,0.12); color:#fff; border:1.5px solid rgba(127,216,171,0.5); transform:translate(-50%,-50%) scale(1.15); z-index:6; opacity:1; box-shadow:0 2px 8px rgba(0,0,0,.5); }
.marker.active::before, .marker.hl::before { content:""; position:absolute; inset:-3px; border-radius:50%; pointer-events:none; padding:3px;
  background:conic-gradient(from 0deg, rgba(127,216,171,0) 0deg, rgba(127,216,171,0.12) 140deg, #7fd8ab 275deg, #d8ffe9 330deg, #7fd8ab 360deg);
  -webkit-mask:linear-gradient(#000 0 0) content-box, linear-gradient(#000 0 0); -webkit-mask-composite:xor;
  mask:linear-gradient(#000 0 0) content-box, linear-gradient(#000 0 0); mask-composite:exclude;
  animation:mk-ring-spin 1.15s linear infinite; }
@keyframes mk-ring-spin { to { transform:rotate(360deg); } }
@media (prefers-reduced-motion: reduce) { .marker.active::before, .marker.hl::before { animation:none; background:#7fd8ab; } }
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
.sheet-tabs { display:flex; flex-wrap:wrap; align-items:center; gap:6px; margin:14px 0 12px; padding:0; }
.tabs-search { margin-left:auto; display:flex; align-items:center; gap:10px; }
.tabs-search .searchbox { flex:0 0 260px; }
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
.api-meter { display:inline-flex; align-items:center; gap:7px; margin-top:8px; font-size:11.5px; color:var(--faint); background:var(--input); border:1px solid rgba(255,255,255,0.07); border-radius:8px; padding:4px 10px; font-variant-numeric:tabular-nums; }
.api-meter .am-dot { width:6px; height:6px; border-radius:6px; background:var(--ac2); display:inline-block; }
.api-meter .am-reset { background:none; border:none; color:var(--faint); cursor:pointer; font-size:13px; padding:0 2px; line-height:1; }
.api-meter .am-reset:hover { color:var(--tx2); }

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
.marker.hl { opacity:1; transform:translate(-50%,-50%) scale(1.3); z-index:60; box-shadow:0 3px 12px rgba(0,0,0,0.55); }
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
.tb-field.tb-code { flex:0 1 130px; min-width:110px; }   /* Mã dự án: hẹp */
.tb-field.tb-date { flex:0 1 130px; min-width:112px; max-width:150px; }  /* Ngày lập: thu nhỏ vừa nội dung */
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
.pane-table .pane-body { overflow:hidden; padding:18px 22px 0; display:flex; flex-direction:column; }
.pane-table .pane-body > div { display:flex; flex-direction:column; min-height:0; flex:1 1 auto; }
.tbl-head { flex:0 0 auto; }
.tbl-scroll { flex:1 1 auto; min-height:0; overflow-y:auto; padding-bottom:44px; }
.tbl-actions { display:flex; align-items:center; gap:10px; flex-wrap:wrap; margin:2px 0 12px; padding-top:12px; border-top:1px solid var(--line); }
.tbl-actions .spacer { flex:1; }
.tbl-actions .btn, .tbl-actions .chip-toggle { height:30px; padding:0 12px; font-size:12px; }

@media (max-width: 1080px) {
  .ax2 { height:auto; overflow:visible; }
  .ax-split { flex-direction:column; }
  .pane-table, .pane-image { flex:1 1 auto; order:0; border-right:none; }
  .pane-table { border-bottom:1px solid var(--line); }
  .split-divider { display:none; }
  .pane-body { overflow:visible; }
  .pane-table .pane-body { overflow:visible; height:auto; display:block; }
  .tbl-scroll { overflow:visible; padding-bottom:0; }
  .grp-caption, .grp-head { position:static; }
}

/* Scrollbar đồng bộ theme (WebKit + Firefox) cho bảng Inventory & các vùng cuộn */
.pane-body, .tbl-scroll, .sched-scroll, .lb-card { scrollbar-width:thin; scrollbar-color:rgba(123,163,207,0.45) var(--input); }
.pane-body::-webkit-scrollbar, .tbl-scroll::-webkit-scrollbar, .sched-scroll::-webkit-scrollbar, .lb-card::-webkit-scrollbar { width:11px; height:11px; }
.pane-body::-webkit-scrollbar-track, .tbl-scroll::-webkit-scrollbar-track, .sched-scroll::-webkit-scrollbar-track, .lb-card::-webkit-scrollbar-track { background:var(--input); border-radius:8px; }
.pane-body::-webkit-scrollbar-thumb, .tbl-scroll::-webkit-scrollbar-thumb, .sched-scroll::-webkit-scrollbar-thumb, .lb-card::-webkit-scrollbar-thumb {
  background-color:rgba(123,163,207,0.4); border-radius:8px; border:2px solid var(--input); }
.pane-body::-webkit-scrollbar-thumb:hover, .tbl-scroll::-webkit-scrollbar-thumb:hover, .sched-scroll::-webkit-scrollbar-thumb:hover, .lb-card::-webkit-scrollbar-thumb:hover { background-color:var(--ac); }
.pane-body::-webkit-scrollbar-thumb:active, .tbl-scroll::-webkit-scrollbar-thumb:active, .sched-scroll::-webkit-scrollbar-thumb:active, .lb-card::-webkit-scrollbar-thumb:active { background-color:var(--ac2); }
.pane-body::-webkit-scrollbar-corner, .tbl-scroll::-webkit-scrollbar-corner, .sched-scroll::-webkit-scrollbar-corner, .lb-card::-webkit-scrollbar-corner { background:var(--input); }

/* Logo ARTIUS trong banner */
.tb-logo { height:44px; width:auto; display:block; }

/* ===== Bảng vật liệu GỘP THEO NHÓM (thay cho <table class="sched">) ===== */
.grp-wrap { border:none; border-radius:0; background:transparent; overflow:visible; }
.grp-head { padding:7px 12px; background:#0e1526; font-size:9px; font-weight:700; letter-spacing:.08em; text-transform:uppercase; border-bottom:1px solid var(--line2); border-top:1px solid var(--line2); display:flex; align-items:center; gap:7px; position:sticky; top:30px; z-index:2; }
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
.grp-sl { width:26px; flex:0 0 auto; text-align:center; font-size:11px; font-weight:700; color:var(--tx2); padding:4px 0; -moz-appearance:textfield; }
.grp-sl::-webkit-outer-spin-button, .grp-sl::-webkit-inner-spin-button { -webkit-appearance:none; margin:0; }
.grp-sl.qty-zero { color:var(--amber2); }
.grp-qty { width:72px; flex:0 0 auto; display:flex; align-items:center; justify-content:center; gap:3px; }
.qty-btn { width:18px; height:18px; flex:0 0 auto; padding:0; border-radius:5px; border:1px solid var(--line2); background:rgba(255,255,255,0.04); color:var(--tx2); font-size:13px; font-weight:700; line-height:1; cursor:pointer; display:flex; align-items:center; justify-content:center; }
.qty-btn:hover { background:rgba(127,216,171,0.14); color:#cfeede; border-color:rgba(127,216,171,0.45); }
.qty-btn:active { transform:scale(0.92); }
.grp-select { width:104px; flex:0 0 auto; background:transparent; border:none; font-family:var(--sans); color:var(--tx3); font-size:11px; padding:5px 0; cursor:pointer; }
.grp-select option { background:#101725; color:var(--tx2); }
.grp-note { width:110px; flex:0 0 auto; font-size:10.5px; color:var(--faint); padding:5px 4px; }
.grp-caption { display:flex; align-items:center; gap:9px; height:30px; box-sizing:border-box; padding:0 12px; background:#0e1526; border-bottom:1px solid var(--line2); font-size:9px; font-weight:700; letter-spacing:.06em; text-transform:uppercase; color:var(--mut2); position:sticky; top:0; z-index:4; }
.grp-cap-sel { width:15px; flex:0 0 auto; }
.grp-cap-thumb { width:72px; flex:0 0 auto; }
.grp-cap-code { width:60px; flex:0 0 auto; }
.grp-cap-main { flex:1; min-width:120px; }
.grp-cap-vitri { width:110px; flex:0 0 auto; }
.grp-cap-sl { width:72px; flex:0 0 auto; text-align:center; }
.grp-cap-select { width:104px; flex:0 0 auto; }
.grp-cap-act { width:60px; flex:0 0 auto; }
.grp-row .row-act { flex:0 0 60px; width:60px; display:flex; align-items:center; justify-content:flex-end; gap:1px; position:relative; }
.grp-row .row-act .icon-danger { flex:0 0 auto; width:auto; padding:7px; }
.icon-move { background:transparent; border:none; color:var(--mut2); cursor:pointer; padding:7px; border-radius:7px; display:inline-flex; }
.icon-move:hover { color:var(--ac2); background:rgba(123,163,207,0.12); }
.move-menu { position:fixed; z-index:80; min-width:170px; background:#101725; border:1px solid var(--line2); border-radius:10px; box-shadow:0 12px 34px rgba(0,0,0,0.55); padding:5px; }
.move-menu-cap { font-size:9px; letter-spacing:.08em; text-transform:uppercase; color:var(--faint); padding:5px 9px 6px; }
.move-menu-item { display:block; width:100%; text-align:left; background:transparent; border:none; color:var(--tx2); font-family:var(--sans); font-size:12px; padding:7px 9px; border-radius:7px; cursor:pointer; }
.move-menu-item:hover:not(:disabled) { background:rgba(127,216,171,0.12); color:#cfeede; }
.move-menu-item:disabled { color:var(--faint); cursor:default; }
.move-menu-item.cur { color:var(--ac3); }
.menu-backdrop { position:fixed; inset:0; z-index:79; background:transparent; }
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
  const rowsRef = useRef(rows);                 // luôn trỏ rows MỚI NHẤT (dùng trong các bước async sau setState)
  useEffect(() => { rowsRef.current = rows; }, [rows]);
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
  const [moveMenu, setMoveMenu] = useState(null);            // { id, x, y, up } — menu chuyển nhóm
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
  const [projectCode, setProjectCode] = useState("");
  const [dateStr, setDateStr] = useState(new Date().toLocaleDateString("vi-VN"));

  // B5: bộ đếm call + ước tính chi phí API (đọc từ API_STATS module-level).
  const [apiStats, setApiStats] = useState({ calls: 0, inTok: 0, outTok: 0, cost: 0 });
  const restoringRef = useRef(true);   // true trong lúc khôi phục phiên -> chặn autosave ghi đè
  useEffect(() => {
    API_ON_UPDATE = () => setApiStats({ calls: API_STATS.calls, inTok: API_STATS.inTok, outTok: API_STATS.outTok, cost: estCostUSD() });
    return () => { API_ON_UPDATE = null; };
  }, []);

  // B1: KHÔI PHỤC phiên từ IndexedDB khi mở app (dựng lại ảnh offscreen để crop/thumb hoạt động).
  useEffect(() => {
    (async () => {
      try {
        const s = await idbGet("session");
        if (s && ((s.rows && s.rows.length) || (s.images && s.images.length))) {
          (s.images || []).forEach((im) => {
            try {
              const el = new Image();
              el.onload = () => setReadyTick((t) => t + 1);
              el.src = im.preview || ("data:" + (im.mediaType || "image/jpeg") + ";base64," + (im.imgData || ""));
              imgElMap.current.set(im.id, el);
            } catch (e) { /* bỏ qua ảnh lỗi */ }
          });
          setImages(s.images || []);
          setRows(s.rows || []); rowsRef.current = s.rows || [];
          if (s.meta) {
            setProjectName(s.meta.projectName || ""); setProjectCode(s.meta.projectCode || "");
            setClient(s.meta.client || ""); setLocation(s.meta.location || "");
            if (s.meta.dateStr) setDateStr(s.meta.dateStr);
          }
          if (s.activeImgId) setActiveImgId(s.activeImgId);
          setStatus("Đã khôi phục phiên làm việc trước (" + ((s.images || []).length) + " ảnh · " + ((s.rows || []).length) + " dòng).");
        }
      } catch (e) { /* không có phiên / trình duyệt chặn -> bỏ qua */ }
      restoringRef.current = false;
    })();
  }, []);

  // B1: TỰ ĐỘNG LƯU phiên (debounce 800ms) mỗi khi bảng/ảnh/thông tin dự án đổi.
  useEffect(() => {
    if (restoringRef.current) return;
    const t = setTimeout(() => {
      try {
        if (!rows.length && !images.length) { idbDel("session").catch(() => {}); return; }
        const snapshot = {
          v: 1, rows, activeImgId,
          meta: { projectName, projectCode, client, location, dateStr },
          images: images.map((im) => ({ id: im.id, preview: im.preview, imgData: im.imgData, mediaType: im.mediaType, fileName: im.fileName, status: im.status })),
        };
        idbSet("session", snapshot).catch(() => {});
      } catch (e) { /* quota/lỗi -> bỏ qua, không chặn app */ }
    }, 800);
    return () => clearTimeout(t);
  }, [rows, images, projectName, projectCode, client, location, dateStr, activeImgId]);

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

  // Cuộn bảng inventory tới đúng hàng (khi bấm ký hiệu trên ảnh).
  function scrollToRow(rowId) {
    try {
      const node = typeof document !== "undefined" && document.getElementById("invrow-" + rowId);
      if (node && node.scrollIntoView) node.scrollIntoView({ behavior: "smooth", block: "center" });
    } catch (e) { /* no-op */ }
  }

  // Chọn 1 dòng trong bảng. Nếu có >1 ảnh và ảnh đang xem KHÔNG chứa ký hiệu của dòng này,
  // tự chuyển sang ảnh phối cảnh đầu tiên có ký hiệu của dòng đó.
  // Yêu cầu mới: bấm vào dòng -> VÀO LUÔN trạng thái crop ảnh của dòng đó (nếu dòng đã có ký hiệu).
  function selectRow(r) {
    setActiveId(r.id);
    setMarkerEdit(false); // crop unlock chỉ chạy khi KHÔNG ở chế độ Thêm ký hiệu
    const insts = r.instances || [];
    if (images.length > 1 && insts.length) {
      const imgIds = insts.map((b) => b.imgId);
      if (imgIds.indexOf(activeImgId) < 0) {
        const target = imgIds.find((id) => images.some((im) => im.id === id));
        if (target != null && target !== activeImgId) setActiveImgId(target);
      }
    }
    // Vào chế độ crop cho dòng vừa bấm (chỉ khi dòng có ít nhất 1 ký hiệu để crop).
    setCropRowId(insts.length ? r.id : null);
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
    let data;
    try {
      data = await apiPost("/api/analyze", {
        model: "claude-sonnet-5", max_tokens: 1000,
        messages: [{ role: "user", content: [
          { type: "image", source: { type: "base64", media_type: pic.media_type, data: pic.data } },
          { type: "text", text: PROMPT },
        ] }],
      });
    } catch (e) {
      throw new Error((e && e.message) || "Lỗi gọi API.");
    }
    if (data && data.type === "error") throw new Error(data.error ? data.error.message : "API trả về lỗi.");
    const textOut = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n");
    return parseItems(textOut).filter((it) => !monExcluded(it.mon));
  }

  // TẦNG 1 (grounding): gọi /api/detect (Gemini) -> danh sách vùng {label,count,x1,y1,x2,y2} trong [0..1].
  async function detectRegionsApi(pic) {
    const j = await apiPost("/api/detect", { image: { media_type: pic.media_type, data: pic.data } });
    return Array.isArray(j.regions) ? j.regions : [];
  }

  // TẦNG 2 (đọc): gửi 1 CROP cho Claude -> 1 item (dùng lại parseItems, lấy dòng đầu).
  // hintLabel = nhãn tiếng Anh do Gemini gán (vase/tray/basket/pillow...) -> giúp phân nhóm đúng.
  async function readCrop(cropB64, hintLabel) {
    const promptText = CROP_PROMPT + (hintLabel ? ("\n\nGợi ý loại vật (từ hệ phát hiện, tiếng Anh): " + hintLabel) : "");
    let data;
    try {
      data = await apiPost("/api/analyze", {
        model: "claude-sonnet-5", max_tokens: 400,
        messages: [{ role: "user", content: [
          { type: "image", source: { type: "base64", media_type: "image/jpeg", data: cropB64 } },
          { type: "text", text: promptText },
        ] }],
      });
    } catch (e) { return null; }
    if (data && data.type === "error") return null;
    const textOut = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n");
    const items = parseItems(textOut);
    return items[0] || null;
  }

  // PASS BỀ MẶT: đọc TOÀN ẢNH, chỉ lấy vật liệu bề mặt/hoàn thiện (sàn/tường/trần/rèm/sơn/đá ốp...).
  async function readSurfaces(pic) {
    let data;
    try {
      data = await apiPost("/api/analyze", {
        model: "claude-sonnet-5", max_tokens: 1000,
        messages: [{ role: "user", content: [
          { type: "image", source: { type: "base64", media_type: pic.media_type, data: pic.data } },
          { type: "text", text: SURFACE_PROMPT },
        ] }],
      });
    } catch (e) { return []; }
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
        // B: Gemini khoanh TỪNG cá thể; loại NGAY người/màn hình/máy tính/laptop/sách theo nhãn (trước khi crop).
        const regions = (await detectRegionsApi(pic)).filter((r) => !labelExcluded(r.label));
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
              // Đếm cá thể THẬT: box chồng nhau (giường bị khoanh 3 lần) gộp làm 1; chỉ box tách rời mới cộng.
              const inst = spatialInstances(idxs.map((i) => regions[i]));
              const rep = inst[0].box;   // đại diện = box lớn nhất cả nhóm
              const crop = makeExportCrop(el, rep, 1000);
              if (!crop || !crop.data) return null;
              const one = await readCrop(b64of(crop.data), g.loai || regions[idxs[0]].label);   // hint = loại do C gom được
              if (!one) return null;
              one.instances = [{ x1: rep.x1, y1: rep.y1, x2: rep.x2, y2: rep.y2 }];    // 1 ký hiệu = cá thể đại diện
              one.soLuong = inst.reduce((s, gg) => s + gg.count, 0);   // = số cá thể tách rời (đã bỏ phần chồng)
              if (g.nhom && !one.nhom) one.nhom = g.nhom;
              // box mọi cá thể (mỗi cụm 1 box, không hiển thị marker) — để dành cho tách nhóm / xuất sau này
              one.memberBoxes = inst.map((gg) => ({ x1: gg.box.x1, y1: gg.box.y1, x2: gg.box.x2, y2: gg.box.y2 }));
              return one;
            }),
            readSurfaces(pic).catch(() => []),
          ]);
          const objectItems = dedupeObjects(read.filter(Boolean));   // an toàn: gộp nốt nếu 2 đại diện lỡ chồng box
          // Lưới an toàn: bỏ nốt dòng bị loại trừ nếu lọt qua nhãn Gemini (vd Claude đặt tên "Người"/"Màn hình máy tính").
          const combined = [...objectItems, ...surfaceItems].filter((it) => !monExcluded(it.mon));
          if (combined.length) return combined;
        }
      }
    } catch (e) {
      // detect/crop lỗi -> rơi xuống fallback bên dưới (không chặn quy trình cũ)
    }
    return callAnalyzeSingle(imgId);
  }

  // Bọc callAnalyze bằng CACHE theo hash ảnh: cùng 1 ảnh (nội dung không đổi) sẽ KHÔNG gọi lại API
  // khi bấm "Phân tích tất cả" nhiều lần. Kết quả lưu bản clone (không giữ id); mỗi lần dùng cấp id mới.
  async function callAnalyzeCached(imgId, fresh) {
    const pic = apiImageFor(imgId);
    const key = pic && pic.data ? (CACHE_VER + "|" + hashStr(pic.data)) : null;
    if (!fresh && key && ANALYZE_CACHE.has(key)) return withFreshIds(ANALYZE_CACHE.get(key));   // cache HIT -> khỏi gọi API
    const items = await callAnalyze(imgId);
    if (key && items) ANALYZE_CACHE.set(key, cacheClone(items));   // luôn cập nhật cache (kể cả khi đọc tươi)
    return items;
  }

  // Phân tích 1 ảnh rồi GỘP kết quả VÀO bảng hiện có (không thay thế).
  // Đây là mấu chốt để món/vật liệu giống nhau ở NHIỀU ảnh gộp về 1 mã duy nhất:
  //   [dòng cũ (đã bỏ box của ảnh này)] + [item mới] -> mergeRows theo (prefix|món|vật liệu) -> codeItems.
  async function analyzeOne(imgId, fresh) {
    const im = images.find((x) => x.id === imgId);
    if (!im) return 0;
    setImages((prev) => prev.map((x) => (x.id === imgId ? { ...x, status: "analyzing", err: "" } : x)));
    try {
      let items = await callAnalyzeCached(imgId, fresh);
      items = normalizeDecorList(items);   // Trang trí: tranh->AW, cây/chậu->GW, loại bình/khay/gối/sách/giỏ/nến
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
        const finalRows = sortRows(combined);
        rowsRef.current = finalRows;    // cập nhật ĐỒNG BỘ để bước gộp mã đọc được rows mới nhất
        return finalRows;
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
  // ---- GỘP MÃ TƯƠNG TỰ: lọc ứng viên (text) -> Sonnet xác nhận thị giác -> gộp + gán lại mã ----
  // Chạy trên rowsRef.current (rows mới nhất). Trả số DÒNG bị gộp bớt (0 nếu không gộp gì).
  async function mergeSimilarCodes() {
    const cur = rowsRef.current || [];
    if (cur.length < 2) return 0;
    const clusters = candidateClusters(cur);              // các cụm nghi giống (cùng prefix, token gần)
    if (!clusters.length) return 0;

    const rank = { "Cao": 2, "Thấp": 1 };
    // Union-find trên TOÀN bảng để gộp các dòng được Sonnet xác nhận cùng nhau.
    const parent = cur.map((_, i) => i);
    const find = (x) => { while (parent[x] !== x) { parent[x] = parent[parent[x]]; x = parent[x]; } return x; };
    const union = (a, b) => { const ra = find(a), rb = find(b); if (ra !== rb) parent[ra] = rb; };

    for (const cl of clusters) {
      // Crop đại diện cho từng dòng trong cụm (bỏ dòng không tạo được crop -> không đủ căn cứ).
      const withCrop = cl.map((idx) => ({ idx, crop: rowExportCrop(cur[idx]) })).filter((o) => o.crop && o.crop.data);
      if (withCrop.length < 2) continue;                  // không đủ 2 ảnh để so -> bỏ qua cụm
      const montage = await buildMontageFromCrops(withCrop.map((o) => o.crop.data));
      if (!montage) continue;
      const b64 = montage.indexOf(",") >= 0 ? montage.slice(montage.indexOf(",") + 1) : montage;
      const groups = await confirmSameProduct(b64, withCrop.length);
      if (!groups) continue;
      // Mỗi group con (>=2 phần tử) do Sonnet xác nhận -> union các row index tương ứng.
      for (const g of groups) {
        if (!g.members || g.members.length < 2) continue;
        const rowIdxs = g.members.map((m) => withCrop[m - 1] && withCrop[m - 1].idx).filter((x) => x != null);
        for (let k = 1; k < rowIdxs.length; k++) union(rowIdxs[0], rowIdxs[k]);
      }
    }

    // Thực thi gộp theo các cụm union cuối cùng.
    const byRoot = new Map();
    cur.forEach((_, i) => { const r = find(i); if (!byRoot.has(r)) byRoot.set(r, []); byRoot.get(r).push(i); });
    let mergedCount = 0;
    const out = [];
    for (const idxs of byRoot.values()) {
      if (idxs.length === 1) { out.push(cur[idxs[0]]); continue; }
      // Chọn dòng ĐẠI DIỆN: do_tin_cay cao nhất, rồi mô tả dài nhất (nhiều thông tin nhất).
      idxs.sort((a, b) => {
        const ra = rank[cur[a].do_tin_cay] || 0, rb = rank[cur[b].do_tin_cay] || 0;
        if (rb !== ra) return rb - ra;
        return String(cur[b].vat_lieu || "").length - String(cur[a].vat_lieu || "").length;
      });
      const base = { ...cur[idxs[0]], instances: [...(cur[idxs[0]].instances || [])] };
      const locs = new Set(splitLocs(base.vi_tri));
      const srcImgs = new Set(base.srcImgs || []);
      let slMax = parseInt(base.soLuong, 10); slMax = Number.isFinite(slMax) ? slMax : 1;
      for (let k = 1; k < idxs.length; k++) {
        const it = cur[idxs[k]];
        (it.instances || []).forEach((b) => base.instances.push(b));   // giữ ĐỦ vị trí ký hiệu trên các ảnh
        splitLocs(it.vi_tri).forEach((s) => locs.add(s));
        (it.srcImgs || []).forEach((s) => srcImgs.add(s));
        const note = String(it.ghi_chu || "").trim();
        if (note && String(base.ghi_chu || "").indexOf(note) < 0) base.ghi_chu = [base.ghi_chu, note].filter(Boolean).join("; ");
        const sl = parseInt(it.soLuong, 10); if (Number.isFinite(sl)) slMax = Math.max(slMax, sl);   // xuyên ảnh: lấy MAX
        if ((rank[it.do_tin_cay] || 0) > (rank[base.do_tin_cay] || 0)) base.do_tin_cay = it.do_tin_cay;
      }
      base.vi_tri = Array.from(locs).join(", ");
      base.srcImgs = Array.from(srcImgs);
      base.soLuong = slMax || 1;
      out.push(base);
      mergedCount += idxs.length - 1;
    }
    if (!mergedCount) return 0;

    pushUndo("gộp mã tương tự");
    // Gán lại mã liên tục + đảm bảo thumb.
    let coded = codeItems(out, gpOf).map((r) => {
      if (r.thumb) return r;
      const b0 = (r.instances || [])[0];
      return { ...r, thumb: b0 && elReady(b0.imgId) ? makeThumb(getEl(b0.imgId), b0) : null };
    });
    setRows(sortRows(coded));
    rowsRef.current = sortRows(coded);
    return mergedCount;
  }

  async function analyzeActive() {
    if (!activeImage) { setError("Chưa có ảnh. Hãy tải ảnh phối cảnh lên trước."); return; }
    setLoading(true); setError(null); setStatus(null); setSelected(new Set());
    try {
      const n = await analyzeOne(activeImgId, true);   // đọc TƯƠI (bỏ cache) vì user chủ động phân tích lại 1 ảnh
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
    // Sau khi phân tích xong TẤT CẢ ảnh -> gộp mã tương tự (text lọc + Sonnet xác nhận).
    let mergedN = 0;
    if (ok) {
      try {
        setStatus("Đang gộp mã tương tự (Sonnet xác nhận)…");
        await new Promise((r) => setTimeout(r, 0));   // nhường 1 tick để rows + rowsRef flush xong
        mergedN = await mergeSimilarCodes();
      } catch (e) { /* gộp lỗi thì bỏ qua, giữ nguyên bảng đã phân tích */ }
    }
    setLoading(false);
    if (!ok) setError("Không phân tích được ảnh nào — " + (errs[0] || "kiểm tra kết nối rồi thử lại.") + (errs.length > 1 ? " (và " + (errs.length - 1) + " ảnh khác)" : ""));
    else {
      setStatus("Đã phân tích " + ok + "/" + list.length + " ảnh" + (fail ? " · " + fail + " ảnh lỗi" : "") + " · gộp trùng xuyên ảnh & gán mã" + (mergedN ? " · đã gộp thêm " + mergedN + " mã tương tự" : "") + ".");
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
  // Tăng/giảm số lượng bằng nút +/- (không xuống dưới 0).
  function bumpQty(r, d) {
    const cur = qtyOf(r);
    const base = Number.isFinite(cur) ? cur : 0;
    updateRow(r.id, "soLuong", String(Math.max(0, base + d)));
  }
  function deleteRow(id) { pushUndo("xoá dòng"); setRows((rs) => codeItems(rs.filter((r) => r.id !== id), gpOf)); if (activeId === id) setActiveId(null); setSelected((s) => { const n = new Set(s); n.delete(id); return n; }); }
  // Thêm dòng mới NGAY DƯỚI dòng đang chọn (cùng nhóm với dòng đó); nếu chưa chọn thì thêm cuối.
  function addRow() {
    const sel = rows.find((r) => r.id === activeId);
    const r = { id: nextId(), prefix: "", ma: "", nhom: sel ? sel.nhom : "Nội thất", mon: "", vat_lieu: "", soLuong: 1, vi_tri: "", do_tin_cay: "Thấp", ghi_chu: "", instances: [], thumb: null };
    setRows((rs) => {
      const idx = sel ? rs.findIndex((x) => x.id === sel.id) : -1;
      const next = idx >= 0 ? [...rs.slice(0, idx + 1), r, ...rs.slice(idx + 1)] : [...rs, r];
      return codeItems(next, gpOf);
    });
    setActiveId(r.id);
  }
  // Chuyển 1 dòng sang nhóm phân loại khác (thay đổi "nhom"). Mã prefix giữ nguyên — sửa tay ở ô Mã nếu cần.
  function moveRowToNhom(id, nhom) {
    pushUndo("chuyển nhóm");
    setRows((rs) => sortRows(rs.map((r) => (r.id === id ? { ...r, nhom } : r))));
    setMoveMenu(null);
  }
  function openMoveMenu(e, r) {
    e.stopPropagation();
    const rect = e.currentTarget.getBoundingClientRect();
    const mh = 40 + NHOM_OPTS.length * 34;   // ước lượng chiều cao menu
    const up = rect.bottom + mh > window.innerHeight;
    setMoveMenu({ id: r.id, x: rect.right, y: up ? rect.top : rect.bottom, up });
  }
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
    const rank = { "Cao": 2, "Thấp": 1 };
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
    aoa.push(["Dự án / Project:", (projectCode ? projectCode + " — " : "") + projectName]);
    aoa.push(["Địa điểm / Location:", location]);
    aoa.push(["Chủ đầu tư / Client:", client]);
    aoa.push(["Ngày / Date:", dateStr]);
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

  // Bản xuất Excel CƠ BẢN (SheetJS) — dùng làm FALLBACK khi không nạp được ExcelJS (offline).
  // Không nhúng được ảnh; giữ nguyên hành vi cũ để nút luôn có tác dụng.
  function exportExcelBasic() {
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
      setStatus("Đã xuất " + safeName() + ".xlsx (bản cơ bản, KHÔNG ảnh) — không nạp được ExcelJS.");
    } catch (e) {
      setStatus("Trình duyệt chặn tải file. Dùng “Sao chép bảng” rồi dán vào Excel.");
    }
  }

  // Lấy crop đại diện (dataURL) của 1 dòng để nhúng vào Excel — dùng box đầu tiên, ~520px cho nét khi in.
  function rowExportCrop(r) {
    const b0 = (r.instances || [])[0];
    if (b0 && elReady(b0.imgId)) { const c = makeExportCrop(getEl(b0.imgId), b0, 520); if (c && c.data) return c; }
    if (r.thumb) return { data: r.thumb, w: 72, h: 72 };
    return null;
  }

  /* Dựng 1 sheet nhóm theo ĐÚNG template ảnh đính kèm (ARTIUS spec) bằng ExcelJS:
     - khối tiêu đề (logo chữ + tên bảng), 4 dòng thông tin dự án, header hồng, dải nhóm xám
     - mỗi món = thẻ 6 dòng dọc: Mô tả / Vị trí / Kích thước / Mã(đỏ) / Nhãn hiệu / Ghi chú
     - cột E "HÌNH ẢNH MẪU 3D": tự động chèn ẢNH CROP của dòng (giữ tỉ lệ, canh giữa trong ô gộp) */
  function fillCatSheetXLSX(wb, ws, cat, list) {
    const thin = { style: "thin", color: { argb: "FF8A8A8A" } };
    const ALL = { top: thin, left: thin, bottom: thin, right: thin };
    const PINK = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE9CFD2" } };
    const GRAY = { type: "pattern", pattern: "solid", fgColor: { argb: "FFCFCFCF" } };
    const midC = { vertical: "middle", horizontal: "center", wrapText: true };
    const midL = { vertical: "middle", horizontal: "left", wrapText: true };

    // Bề rộng cột (xấp xỉ px để tính canh ảnh): A stt, B ký hiệu, C nhãn, D giá trị, E ảnh 3D, F ảnh duyệt
    ws.columns = [{ width: 6 }, { width: 13 }, { width: 24 }, { width: 40 }, { width: 30 }, { width: 30 }];

    // ----- Khối tiêu đề -----
    ws.mergeCells("A1:B1");
    ws.getCell("A1").alignment = midC;
    try {
      const logoId = wb.addImage({ base64: b64of(ARTIUS_LOGO_PNG), extension: "png" });
      const pxA = 6 * 7 + 5, pxB = 13 * 7 + 5, totalW = pxA + pxB, row1px = 46 * 96 / 72;
      const boxW = totalW - 12, boxH = row1px - 8;
      const sc = Math.min(boxW / ARTIUS_LOGO_W, boxH / ARTIUS_LOGO_H);
      const dw = Math.round(ARTIUS_LOGO_W * sc), dh = Math.round(ARTIUS_LOGO_H * sc);
      ws.addImage(logoId, {
        tl: { col: ((totalW - dw) / 2) / pxA, row: ((row1px - dh) / 2) / row1px },
        ext: { width: dw, height: dh }, editAs: "oneCell",
      });
    } catch (e) {
      ws.getCell("A1").value = "ARTIUS"; ws.getCell("A1").font = { bold: true, size: 14, name: "Arial" };
    }
    ws.mergeCells("C1:F1");
    ws.getCell("C1").value = "CHỈ DẪN KỸ THUẬT - TECHNICAL SPECIFICATION\nHẠNG MỤC - CATEGORY";
    ws.getCell("C1").font = { bold: true, size: 13, name: "Arial" };
    ws.getCell("C1").alignment = midC;
    ws.getRow(1).height = 46;

    const infos = [
      "Dự án - mã dự án / Project - project code: " + (projectCode ? projectCode + " — " : "") + (projectName || ""),
      "Địa điểm / Location: " + (location || ""),
      "Chủ đầu tư / Customer: " + (client || ""),
      "Hạng mục vật tư - thiết bị / Materials - equipment category: " + cat,
    ];
    infos.forEach((txt, k) => {
      const rr = 2 + k;
      ws.mergeCells("A" + rr + ":F" + rr);
      const c = ws.getCell("A" + rr);
      c.value = txt; c.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      c.font = { size: 11, name: "Arial", bold: k === 3 };
      ws.getRow(rr).height = 18;
    });

    // ----- Hàng header (gộp dọc 6-7), nền hồng -----
    ws.mergeCells("A6:A7"); ws.getCell("A6").value = "STT\nNO.";
    ws.mergeCells("B6:B7"); ws.getCell("B6").value = "KÝ HIỆU\nSYMBOL";
    ws.mergeCells("C6:D7"); ws.getCell("C6").value = "THÔNG TIN / INFORMATION";
    ws.mergeCells("E6:E7"); ws.getCell("E6").value = "HÌNH ẢNH MẪU 3D\nSAMPLE PICTURE 3D";
    ws.mergeCells("F6:F7"); ws.getCell("F6").value = "HÌNH ẢNH MẪU DUYỆT\nSAMPLE PICTURE APPROVED";
    ["A6", "B6", "C6", "E6", "F6"].forEach((a) => {
      const c = ws.getCell(a); c.fill = PINK; c.font = { bold: true, size: 9, name: "Arial" }; c.alignment = midC;
    });
    ws.getRow(6).height = 16; ws.getRow(7).height = 16;

    // ----- Dải nhóm (xám) -----
    ws.mergeCells("A8:F8");
    const band = ws.getCell("A8");
    band.value = String(cat || "").toUpperCase() + "/";
    band.fill = GRAY; band.font = { bold: true, size: 11, name: "Arial" }; band.alignment = { vertical: "middle", horizontal: "center" };
    ws.getRow(8).height = 18;

    // ----- Các thẻ món (6 dòng/thẻ) -----
    const CARD_H = 22;                 // pt/dòng
    const colPx = 30 * 7 + 5;          // ~ px cột E
    const rowPx = CARD_H * 96 / 72;    // ~ px/dòng
    let s = 9, stt = 0;
    list.forEach((r) => {
      stt++;
      const desc = [r.mon, r.vat_lieu].filter(Boolean).join(" — ");
      ws.mergeCells("A" + s + ":A" + (s + 5)); const ca = ws.getCell("A" + s); ca.value = stt; ca.font = { bold: true, size: 11, name: "Arial" }; ca.alignment = midC;
      ws.mergeCells("B" + s + ":B" + (s + 5)); const cb = ws.getCell("B" + s); cb.value = r.ma || ""; cb.font = { bold: true, size: 11, name: "Arial" }; cb.alignment = midC;
      ws.mergeCells("E" + s + ":E" + (s + 5));
      ws.mergeCells("F" + s + ":F" + (s + 5));
      const L = [
        ["Mô tả / Description:", desc],
        ["Vị trí sử dụng / Area:", r.vi_tri || ""],
        ["Kích thước / Dimension:", ""],
        ["Mã / Code:", ""],
        ["Nhãn hiệu / Brand:", ""],
        ["Ghi chú / Note:", r.ghi_chu || ""],
      ];
      for (let k = 0; k < 6; k++) {
        const rr = s + k;
        const lc = ws.getCell("C" + rr); lc.value = L[k][0]; lc.alignment = midL;
        lc.font = k === 3 ? { color: { argb: "FFC00000" }, name: "Arial", size: 10 } : { name: "Arial", size: 10 };
        const vc = ws.getCell("D" + rr); vc.value = L[k][1]; vc.alignment = midL; vc.font = { name: "Arial", size: 10 };
        ws.getRow(rr).height = CARD_H;
      }

      // Nhúng ảnh crop vào ô E (gộp 6 dòng) — giữ tỉ lệ, canh giữa.
      const crop = rowExportCrop(r);
      if (crop && crop.data) {
        try {
          const imgId = wb.addImage({ base64: b64of(crop.data), extension: "jpeg" });
          const boxW = colPx - 12, boxH = rowPx * 6 - 12;
          const cw = crop.w || 4, ch = crop.h || 3;
          const sc = Math.min(boxW / cw, boxH / ch, 1);
          const dw = Math.max(1, Math.round(cw * sc)), dh = Math.max(1, Math.round(ch * sc));
          ws.addImage(imgId, {
            tl: { col: 4 + ((colPx - dw) / 2) / colPx, row: (s - 1) + ((rowPx * 6 - dh) / 2) / rowPx },
            ext: { width: dw, height: dh },
            editAs: "oneCell",
          });
        } catch (e) { /* bỏ qua ảnh lỗi, vẫn giữ ô trống */ }
      }
      s += 6;
    });

    // ----- Viền lưới toàn bảng (ô gộp -> tự thành khung ngoài) -----
    const last = ws.rowCount;
    for (let rr = 1; rr <= last; rr++) for (let cc = 1; cc <= 6; cc++) ws.getCell(rr, cc).border = ALL;
  }

  // Dựng sheet TỔNG HỢP (ExcelJS) — giữ đủ cột gồm cả SL, có tiêu đề tô nền + viền.
  function fillSummarySheetXLSX(ws) {
    ws.columns = [{ width: 6 }, { width: 10 }, { width: 16 }, { width: 26 }, { width: 28 }, { width: 18 }, { width: 6 }, { width: 12 }, { width: 28 }];
    const thin = { style: "thin", color: { argb: "FF8A8A8A" } };
    const ALL = { top: thin, left: thin, bottom: thin, right: thin };
    const H = ["STT", "Mã", "Nhóm", "Món", "Vật liệu / Finish", "Vị trí", "SL", "Độ tin cậy", "Ghi chú"];
    const hr = ws.addRow(H);
    hr.eachCell((c) => {
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE9CFD2" } };
      c.font = { bold: true, size: 10, name: "Arial" };
      c.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      c.border = ALL;
    });
    hr.height = 20;
    rows.forEach((r, i) => {
      const row = ws.addRow([i + 1, r.ma, r.nhom, r.mon, r.vat_lieu, r.vi_tri, qtyOf(r), r.do_tin_cay, r.ghi_chu]);
      row.eachCell((c, col) => {
        c.font = { size: 10, name: "Arial" };
        c.alignment = { vertical: "middle", wrapText: true, horizontal: (col === 1 || col === 7) ? "center" : "left" };
        c.border = ALL;
      });
    });
    ws.views = [{ state: "frozen", ySplit: 1 }];
  }

  // Xuất Excel GIỐNG TEMPLATE + tự chèn ảnh crop vào cột "HÌNH ẢNH MẪU 3D" (ExcelJS).
  async function exportExcelRich() {
    const ExcelJS = await loadExcelJS();
    const wb = new ExcelJS.Workbook();
    wb.creator = "ARTIUS SpecLens";
    const used = new Set();
    const sheetName = (name) => {
      let n = String(name || "SHEET").replace(/[\\/?*\[\]:]/g, " ").trim().slice(0, 28) || "SHEET";
      let base = n, k = 2;
      while (used.has(n.toLowerCase())) { n = (base.slice(0, 24) + " " + k).trim(); k++; }
      used.add(n.toLowerCase()); return n;
    };
    // TỔNG HỢP
    fillSummarySheetXLSX(wb.addWorksheet(sheetName("TỔNG HỢP")));
    // Nhóm vật liệu -> mỗi nhóm 1 sheet theo template
    const groups = {}, order = [];
    rows.forEach((r) => {
      const pfx = (r.ma.split("-")[0] || "").toUpperCase();
      const c = sheetCategory(pfx);
      if (!groups[c]) { groups[c] = []; order.push(c); }
      groups[c].push(r);
    });
    order.forEach((c) => fillCatSheetXLSX(wb, wb.addWorksheet(sheetName(c)), c, groups[c]));

    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href = url; a.download = safeName() + ".xlsx"; a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
    const withImg = rows.filter((r) => !!rowExportCrop(r)).length;
    setStatus("Đã xuất " + safeName() + ".xlsx theo template — TỔNG HỢP + mỗi nhóm 1 sheet, đã chèn " + withImg + "/" + rows.length + " ảnh crop vào cột HÌNH ẢNH MẪU 3D.");
  }

  async function exportExcel() {
    if (!rows.length) return;
    setStatus("Đang dựng Excel theo template (nạp ExcelJS để nhúng ảnh)…");
    try {
      await exportExcelRich();
    } catch (e) {
      // Không nạp được ExcelJS (offline) -> rơi về bản cơ bản không ảnh, để nút vẫn dùng được.
      setStatus("Không nhúng được ảnh (thiếu ExcelJS/mạng) — xuất bản cơ bản không ảnh.");
      exportExcelBasic();
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
      ctx.beginPath(); ctx.arc(cx, cy, r, 0, Math.PI * 2);
      ctx.fillStyle = "#ffffff"; ctx.fill();                                // nền trắng
      ctx.lineWidth = Math.max(1.5, r * 0.16); ctx.strokeStyle = "#e01e1e"; ctx.stroke(); // viền đỏ
      ctx.fillStyle = "#e01e1e"; ctx.fillText(String(f.no), cx, cy);        // số đỏ
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
    const bundle = { meta: { project: projectName, code: projectCode, client, location, date: dateStr, images: images.length, generatedAt: new Date().toISOString() }, items };
    try {
      const blob = new Blob([JSON.stringify(bundle)], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a"); a.href = url; a.download = safeName() + "-bundle.json"; a.click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      const withImg = items.filter((x) => x.image).length;
      setStatus("Đã xuất gói .json (" + items.length + " dòng, " + withImg + " ảnh crop) — gửi file này vào chat để tạo .xlsx có ảnh nhúng.");
    } catch (e) { setStatus("Không tạo được gói .json trong môi trường này."); }
  }

  // Kết xuất HTML tương tác (hướng A): 1 file tự chứa, bấm ký hiệu -> chi tiết; chuyển ảnh bằng tab; In/PDF.
  async function exportHtml() {
    if (!rows.length) { setStatus("Chưa có dữ liệu để xuất HTML."); return; }
    const dispNo = buildDisplayNo(rows);
    const groups = groupRowsByNhom(rows).map((g) => ({
      nhom: g.nhom,
      items: g.rows.map((r) => {
        const b0 = (r.instances || [])[0];
        let crop = r.thumb || null;
        if (b0 && elReady(b0.imgId)) { const c = makeExportCrop(getEl(b0.imgId), b0, 300); if (c && c.data) crop = c.data; }
        return {
          id: String(r.id), no: dispNo.get(r.id) || 0, ma: r.ma || "", mon: r.mon || "", vat_lieu: r.vat_lieu || "",
          sl: qtyOf(r), tin: r.do_tin_cay || "", vi_tri: r.vi_tri || "", ghi_chu: r.ghi_chu || "", crop,
          marks: (r.instances || []).map((b) => ({ img: String(b.imgId), x1: b.x1, y1: b.y1, x2: b.x2, y2: b.y2 })),
        };
      }),
    }));
    const imgList = images.map((im, i) => {
      let src = im.preview;
      if (elReady(im.id)) {
        try {
          const el = getEl(im.id), nw = el.naturalWidth, nh = el.naturalHeight, sc = Math.min(1, 1600 / Math.max(nw, nh));
          const w = Math.max(1, Math.round(nw * sc)), h = Math.max(1, Math.round(nh * sc));
          const c = document.createElement("canvas"); c.width = w; c.height = h; c.getContext("2d").drawImage(el, 0, 0, w, h);
          src = c.toDataURL("image/jpeg", 0.82);
        } catch (e) { /* fallback preview */ }
      }
      return { id: String(im.id), idx: i + 1, name: im.fileName || ("Ảnh " + (i + 1)), src };
    });
    const payload = { meta: { project: projectName, code: projectCode, client, location, date: dateStr, logo: ARTIUS_LOGO_PNG, generatedAt: new Date().toISOString() }, colors: GROUP_COLOR, images: imgList, groups };
    // B2: tải Paged.js (polyfill CSS Paged Media) rồi NHÚNG vào file -> PDF phân trang chuẩn, chạy offline.
    setStatus("Đang chuẩn bị HTML + Paged.js cho bản in…");
    let pagedB64 = "";
    const cdns = [
      "https://unpkg.com/pagedjs@0.4.3/dist/paged.polyfill.min.js",
      "https://cdn.jsdelivr.net/npm/pagedjs@0.4.3/dist/paged.polyfill.min.js",
      "https://cdnjs.cloudflare.com/ajax/libs/pagedjs/0.4.3/paged.polyfill.min.js",
    ];
    for (const u of cdns) {
      try { const r = await fetch(u); if (r.ok) { const txt = await r.text(); if (txt && txt.length > 50000) { pagedB64 = btoa(unescape(encodeURIComponent(txt))); break; } } } catch (e) { /* thử CDN kế */ }
    }
    try {
      const html = buildInteractiveHtml(payload, css + cssExtra, pagedB64);
      const blob = new Blob([html], { type: "text/html;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a"); a.href = url; a.download = safeName() + "-spec-tuong-tac.html"; a.click();
      setTimeout(() => URL.revokeObjectURL(url), 1500);
      setStatus(pagedB64
        ? "Đã xuất HTML — bấm 'In / Xuất PDF' trong file để tạo PDF phân trang chuẩn (số trang, header lặp; Paged.js đã nhúng, chạy offline)."
        : "Đã xuất HTML — KHÔNG tải được Paged.js (mạng?), nút In/PDF dùng chế độ in cơ bản (B1, chưa có số trang). Xuất lại khi có mạng để nhúng Paged.js.");
    } catch (e) { setStatus("Không tạo được file HTML trong môi trường này."); }
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
            <div className="tb-field tb-code"><label>Mã dự án</label><input placeholder="Mã DA…" value={projectCode} onChange={(e) => setProjectCode(e.target.value)} /></div>
            <div className="tb-field"><label>Dự án</label><input placeholder="Tên dự án…" value={projectName} onChange={(e) => setProjectName(e.target.value)} /></div>
            <div className="tb-field"><label>Địa điểm</label><input placeholder="Địa điểm…" value={location} onChange={(e) => setLocation(e.target.value)} /></div>
            <div className="tb-field"><label>Chủ đầu tư</label><input placeholder="CĐT…" value={client} onChange={(e) => setClient(e.target.value)} /></div>
            <div className="tb-field tb-date"><label>Ngày lập</label><input placeholder="dd/mm/yyyy…" value={dateStr} onChange={(e) => setDateStr(e.target.value)} /></div>
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
            </div>
            <div className="panel">
              <input ref={fileRef} type="file" accept="image/*" multiple onChange={onPickFiles} style={{ display: "none" }} />

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
                      onPointerDown={(e) => startMarker(e, m.rowId, m.instIdx)} onClick={(e) => { e.stopPropagation(); setActiveId(m.rowId); scrollToRow(m.rowId); }}
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
                <div className="spacer" />
                {hasRows && (<button className={"btn btn-ghost" + (markerEdit ? " on" : "")} onClick={toggleEdit}><MapPin size={15} /> {markerEdit ? "Đang chỉnh ký hiệu" : "Thêm ký hiệu"}</button>)}
                <button className="btn btn-ghost" onClick={analyzeActive} disabled={loading || !activeImage}><MapPin size={15} /> Phân tích ảnh này</button>
                <button className="btn btn-primary" onClick={analyzeAll} disabled={loading || !hasImages}>{loading && <Loader2 size={15} className="spin" />}{loading ? "Đang phân tích…" : "Phân tích tất cả (" + images.length + ")"}</button>
              </div>
              {apiStats.calls > 0 && (
                <div className="api-meter" title="Số lần gọi API · token đã dùng · chi phí ƯỚC TÍNH (giá có thể đổi)">
                  <span className="am-dot" /> {apiStats.calls} calls · ~{Math.round((apiStats.inTok + apiStats.outTok) / 1000)}k tok · ~${apiStats.cost.toFixed(3)}
                  <button className="am-reset" onClick={() => { resetApiStats(); }} title="Đặt lại bộ đếm">↺</button>
                </div>
              )}
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
            <div className="tbl-head">
            <div className="block-head">
              <span className="section-label">01 · Bảng Inventory & Mã vật liệu</span>
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
                <div className="tabs-search">
                  <div className="searchbox">
                    <Search width={15} height={15} />
                    <input placeholder="Tìm theo mã, món, vật liệu, vị trí…" value={search} onChange={(e) => setSearch(e.target.value)} />
                    {search && <button className="clr" title="Xoá tìm" aria-label="Xoá tìm" onClick={() => setSearch("")}><XIcon width={14} height={14} /></button>}
                  </div>
                </div>
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

            <div className="tbl-actions">
              <button className="btn btn-ghost" onClick={addRow}><Plus size={15} /> Thêm dòng</button>
              <button className="btn btn-ghost" onClick={mergeSelected} disabled={selCount < 2} title={selCount < 2 ? "Tick chọn từ 2 dòng để gộp" : ("Gộp " + selCount + " dòng đã chọn thành 1")}><Combine size={15} /> Gộp dòng đã chọn{selCount >= 2 ? " (" + selCount + ")" : ""}</button>
              <button className="btn btn-ghost" onClick={undo} disabled={!undoStack.length} title={undoStack.length ? ("Hoàn tác: " + undoStack[undoStack.length - 1].label + " (Ctrl/Cmd+Z)") : "Không có gì để hoàn tác"}><Undo2 size={15} /> Hoàn tác{undoStack.length ? " (" + undoStack.length + ")" : ""}</button>
              <div className="spacer" />
              <button className={"chip-toggle" + (onlyLow ? " on" : "")} onClick={() => setOnlyLow((v) => !v)} title="Chỉ hiện dòng độ tin cậy Thấp">
                <span className="dotc" /> Chỉ tin cậy Thấp{lowN ? " (" + lowN + ")" : ""}
              </button>
              <button className={"chip-toggle" + (onlyUnpinned ? " on" : "")} onClick={() => setOnlyUnpinned((v) => !v)} title="Chỉ hiện dòng chưa gắn ký hiệu (SL=0, không có hình/chấm)">
                <span className="dotc" style={{ background: "var(--ac)" }} /> Chưa gắn KH{unpinnedN ? " (" + unpinnedN + ")" : ""}
              </button>
              <button className="btn btn-ghost" onClick={exportExcel} disabled={!hasRows} title="Xuất .xlsx theo template ARTIUS — mỗi nhóm 1 sheet, tự chèn ảnh crop vào cột HÌNH ẢNH MẪU 3D"><Download size={15} /> Xuất Excel (bảng)</button>
              <button className="btn btn-ghost" onClick={exportHtml} disabled={!hasRows || !hasImages} title="Xuất file HTML tương tác: bấm ký hiệu để xem chi tiết, chuyển ảnh bằng tab, In/PDF kèm legend"><ImageDown size={15} /> Xuất HTML</button>
            </div>

            </div>{/* /tbl-head */}

            <div className="tbl-scroll" onScroll={() => moveMenu && setMoveMenu(null)}>
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
            <div className="grp-wrap">
              {hasRows ? (
                groupRowsByNhom(rows.filter(passFilter)).map((g) => (
                  <div key={g.key}>
                    <div className="grp-head" style={{ color: GROUP_COLOR[g.key] || "var(--tx3)" }}>
                      <span className="grp-dot" style={{ background: GROUP_COLOR[g.key] || "var(--tx3)" }} />
                      {g.key} · {g.rows.length}
                    </div>
                    {g.rows.map((r) => (
                      <div key={r.id} id={"invrow-" + r.id}
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
                        <div className="grp-qty" onClick={(e) => e.stopPropagation()}>
                          <button className="qty-btn" aria-label="Giảm số lượng" title="Giảm" onClick={(e) => { e.stopPropagation(); bumpQty(r, -1); }}>−</button>
                          <input type="number" min="0" className={"grp-input grp-sl" + (qtyOf(r) === 0 ? " qty-zero" : "")} aria-label="Số lượng"
                            title="Số lượng do AI đếm — có thể sửa tay hoặc dùng nút +/−"
                            value={r.soLuong != null ? r.soLuong : (r.instances.length || "")}
                            onClick={(e) => e.stopPropagation()}
                            onChange={(e) => updateRow(r.id, "soLuong", e.target.value)} />
                          <button className="qty-btn" aria-label="Tăng số lượng" title="Tăng" onClick={(e) => { e.stopPropagation(); bumpQty(r, 1); }}>+</button>
                        </div>
                        <select className="grp-select" aria-label="Độ tin cậy" value={TINCAY_OPTS.includes(r.do_tin_cay) ? r.do_tin_cay : "Thấp"} onClick={(e) => e.stopPropagation()} onChange={(e) => updateRow(r.id, "do_tin_cay", e.target.value)}>
                          {TINCAY_OPTS.map((o) => <option key={o} value={o}>{o}</option>)}
                        </select>
                        <div className="row-act">
                          <button className="icon-move" aria-label="Chuyển sang nhóm khác" title="Chuyển sang nhóm phân loại khác" onClick={(e) => openMoveMenu(e, r)}><ArrowLR width={14} height={14} /></button>
                          <button className="icon-danger" aria-label="Xóa dòng" onClick={(e) => { e.stopPropagation(); deleteRow(r.id); }}><Trash2 size={14} /></button>
                        </div>
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
            </div>{/* /tbl-scroll */}

            {status && <div className="status"><span style={{ width: 6, height: 6, borderRadius: 6, background: "var(--ac2)", display: "inline-block" }} />{status}</div>}
          </div>
          </div>

          {moveMenu && (
            <>
              <div className="menu-backdrop" onClick={() => setMoveMenu(null)} />
              <div className="move-menu" style={{ left: moveMenu.x, top: moveMenu.y, transform: moveMenu.up ? "translate(-100%,-100%)" : "translate(-100%,0)" }} onClick={(e) => e.stopPropagation()}>
                <div className="move-menu-cap">Chuyển sang nhóm</div>
                {NHOM_OPTS.map((nh) => {
                  const cur = rows.find((r) => r.id === moveMenu.id);
                  const isCur = cur && cur.nhom === nh;
                  return <button key={nh} className={"move-menu-item" + (isCur ? " cur" : "")} disabled={isCur} onClick={() => moveRowToNhom(moveMenu.id, nh)}>{nh}</button>;
                })}
              </div>
            </>
          )}
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
