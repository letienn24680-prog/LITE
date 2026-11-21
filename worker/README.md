# LITE R2 Backend API (`lite-r2-api`)

Cloudflare Worker không phụ thuộc thư viện runtime (`zero-dependency`, thuần Web Standards + `crypto.subtle`) cung cấp:
- Ký **AWS Signature V4 Presigned URL (`PUT`/`GET`)** trực tiếp cho **Cloudflare R2** (kèm nhánh dự phòng `/api/r2/direct/*` qua `R2_BUCKET` binding khi chưa gắn S3 API keys).
- Xác thực JWT người dùng **Supabase Auth** (`/auth/v1/user`).
- API tạo, mở khóa, kiểm tra và thu hồi liên kết chia sẻ tài liệu (`/api/shares/*`) đồng bộ với `public.document_shares` và RPC Supabase.

---

## Hướng dẫn triển khai trong 4 bước

### Bước 1: Tạo R2 Bucket và bật CORS cho trình duyệt
1. Đăng nhập [Cloudflare Dashboard](https://dash.cloudflare.com/) → **R2 Object Storage** → **Create bucket**.
2. Đặt tên bucket là `lite-user-files` (khớp với `bucket_name` trong [wrangler.toml](./wrangler.toml)).
3. Mở tab **Settings → CORS Policy** của bucket vừa tạo và thêm chính sách cho phép trình duyệt upload/xem PDF trực tiếp qua Presigned URL:
   ```json
   [
     {
       "AllowedOrigins": ["http://localhost:5173", "https://your-lite-domain.pages.dev"],
       "AllowedMethods": ["GET", "PUT", "HEAD"],
       "AllowedHeaders": ["*"],
       "ExposeHeaders": ["ETag", "Content-Length"],
       "MaxAgeSeconds": 3600
     }
   ]
   ```

### Bước 2: Cấp S3 API Token cho Cloudflare R2
1. Tại trang tổng quan **R2 Object Storage**, ghi lại **Account ID** (chuỗi hex 32 ký tự).
2. Chọn **Manage R2 API Tokens → Create API token**:
   - **Permissions:** `Object Read & Write`.
   - **Specify bucket(s):** Chỉ chọn `lite-user-files`.
3. Lưu lại `Access Key ID` và `Secret Access Key` vừa được cấp.

### Bước 3: Cấu hình Wrangler, Secrets và Deploy Worker
1. Cài đặt devDependencies và điền biến môi trường vào [wrangler.toml](./wrangler.toml) (`ALLOWED_ORIGIN`, `SUPABASE_URL`, `R2_ACCOUNT_ID`, `R2_BUCKET_NAME`):
   ```bash
   cd worker
   npm install
   ```
2. Nạp 3 bí mật (secrets) bằng Wrangler CLI:
   ```bash
   npx wrangler secret put SUPABASE_ANON_KEY
   npx wrangler secret put R2_ACCESS_KEY_ID
   npx wrangler secret put R2_SECRET_ACCESS_KEY
   ```
3. Kiểm tra TypeScript, chạy thử cục bộ hoặc deploy lên Cloudflare:
   ```bash
   npm run typecheck
   npm run dev
   npm run deploy
   ```
   Sau khi deploy, kiểm tra nhanh endpoint: `GET https://lite-r2-api.<subdomain>.workers.dev/api/health`.

### Bước 4: Gắn URL Worker vào Frontend LITE (`.env`)
1. Mở (hoặc tạo) tệp `.env` ở thư mục gốc dự án LITE và thêm URL Worker vừa deploy:
   ```env
   VITE_SUPABASE_URL=https://your-project-ref.supabase.co
   VITE_SUPABASE_ANON_KEY=your-supabase-anon-key
   VITE_R2_API_URL=https://lite-r2-api.<subdomain>.workers.dev
   ```
2. Khởi động lại frontend (`npm run dev` hoặc build lại `npm run build`) để LITE tự động ưu tiên upload tệp qua R2 Presigned URL và mở liên kết chia sẻ công khai không cần đăng nhập.
