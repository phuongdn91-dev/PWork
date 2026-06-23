# PWork — Hướng dẫn Triển khai

## Cấu trúc file
```
pwork-app/
├── index.html          ← App chính (toàn bộ UI + logic)
├── manifest.json       ← PWA manifest (cài lên điện thoại)
├── sw.js               ← Service Worker (offline support)
├── icon-192.png        ← Icon app 192x192 (tự thêm)
├── icon-512.png        ← Icon app 512x512 (tự thêm)
└── gas-backend.gs      ← Google Apps Script (copy vào GAS)
```

---

## Bước 1: Host lên web server

### Cách đơn giản nhất — GitHub Pages (miễn phí)
1. Tạo repo mới trên GitHub (ví dụ: `pwork-app`)
2. Upload tất cả file vào repo (trừ `gas-backend.gs`)
3. Vào **Settings > Pages > Branch: main / root** → Save
4. URL app: `https://[username].github.io/pwork-app/`

### Alternatives
- **Netlify** (kéo thả folder): netlify.com
- **Vercel**: `npx vercel` trong thư mục
- **Serve cục bộ**: `npx serve .` (cần Node.js)

---

## Bước 2: Cài lên điện thoại (PWA)

### Android (Chrome)
1. Mở URL app trong Chrome
2. Nhấn menu **⋮ > Thêm vào màn hình chính**
3. Nhấn **Thêm** → App xuất hiện như app thật

### iPhone (Safari)
1. Mở URL app trong Safari (bắt buộc dùng Safari)
2. Nhấn nút **Chia sẻ** (⬆️)
3. Chọn **Thêm vào màn hình chính**
4. Nhấn **Thêm**

---

## Bước 3: Kết nối Google Sheet (tuỳ chọn)

### 3.1. Tạo Google Sheet
1. Vào [sheets.google.com](https://sheets.google.com) → Tạo spreadsheet mới
2. Đặt tên: `PWork Database`

### 3.2. Tạo Google Apps Script
1. Trong Google Sheet: **Extensions > Apps Script**
2. Xóa code mặc định, paste toàn bộ nội dung file `gas-backend.gs`
3. Nhấn **Save** (Ctrl+S)

### 3.3. Deploy GAS
1. Click **Deploy > New deployment**
2. Chọn type: **Web app**
3. Description: `PWork API v1`
4. Execute as: **Me**
5. Who has access: **Anyone** (cần thiết để app gọi được)
6. Click **Deploy**
7. Copy **Web app URL** (dạng `https://script.google.com/macros/s/XXX/exec`)

### 3.4. Cấu hình trong app
Mở `index.html`, tìm dòng:
```javascript
const GAS_URL = ''; // 'https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec'
```
Thay bằng URL vừa copy:
```javascript
const GAS_URL = 'https://script.google.com/macros/s/AKfycb.../exec';
```

---

## Bước 4: Thêm Icon app

Tạo 2 file icon PNG:
- `icon-192.png` — 192×192px
- `icon-512.png` — 512×512px

Gợi ý: Dùng [favicon.io](https://favicon.io) hoặc [realfavicongenerator.net]

---

## Tính năng đã có

| Tính năng | Trạng thái |
|-----------|------------|
| Thêm công việc (đầy đủ trường) | ✅ |
| Upload file đính kèm (lưu local) | ✅ |
| Danh sách công việc + tìm kiếm | ✅ |
| Lọc theo trạng thái | ✅ |
| Cập nhật tiến độ % (slider) | ✅ |
| Nhật ký triển khai / kế hoạch | ✅ |
| Báo cáo kết quả + đính kèm file | ✅ |
| Tự lưu ngày hoàn thành | ✅ |
| Báo cáo tổng hợp theo kỳ | ✅ |
| Xuất báo cáo (.txt) | ✅ |
| Đồng bộ Google Sheet | ✅ (cần cấu hình GAS_URL) |
| Cài được lên điện thoại (PWA) | ✅ |
| Offline support | ✅ |

---

## Roadmap mở rộng AI (Phase 2)

```javascript
// Tích hợp Gemini API vào app
// Thêm floating chat button
// Prompt template:
const AI_PROMPT = `
Dữ liệu công việc VNPT-NET:
${JSON.stringify(tasks, null, 2)}

Câu hỏi: ${userQuestion}

Trả lời bằng tiếng Việt, ngắn gọn, đúng trọng tâm.
`;
```

Các usecase AI:
- "Tóm tắt tiến độ tuần này"  
- "Công việc nào sắp đến hạn?"
- "Ai đang chủ trì nhiều việc nhất?"
- "Tổng hợp báo cáo dạng văn bản"
