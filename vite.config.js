import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Dev thường ngày: `npm run dev` (vite) — nhanh, nhưng KHÔNG chạy được /api/analyze
// (đó là hàm Vercel, vite dev server không biết tới). Proxy /api sang `vercel dev`
// (cổng 3000 mặc định) để vẫn gọi AI thật được khi cần trong lúc phát triển UI.
// Muốn test đầy đủ (kể cả /api) mà không cần chạy 2 lệnh: dùng `npm run dev:vercel`.
export default defineConfig({
  plugins: [react()],
  server: {
    proxy: {
      "/api": "http://localhost:3000",
    },
  },
});
