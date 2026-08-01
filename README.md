# Yasuda Yen Xa — Bilingual Daily Report / Daily Plan Tool

App nhập liệu song ngữ (Anh đứng / Việt nghiêng) cho **Daily Report** (kèm trang ảnh)
và **Daily Plan** (không ảnh) của dự án Yen Xa Sewerage System — Package 4.

## Cấu trúc file

| File | Vai trò |
|---|---|
| `daily_report_app.html` | Giao diện nhập liệu (mở bằng trình duyệt, hoặc host qua GitHub Pages) |
| `local_server.py` | Server chạy TRÊN MÁY BẠN (không cần cài Flask) |
| `app.py` | Server để **DEPLOY ONLINE** (dùng Flask, chạy trên Render/Railway...) |
| `json_to_excel.py` | Ánh xạ dữ liệu vào đúng toạ độ ô trong template |
| `report_engine.py` | Engine lõi sinh file Excel |
| `requirements.txt` | Danh sách thư viện cần cài khi deploy online |
| `20260731-DR1.xlsx` / `DR2.xlsx` | 2 file mẫu gốc |

## Cách 1 — Chạy trên máy (như trước giờ, không cần internet)

```bash
pip install openpyxl pillow
python local_server.py
```
Mở `daily_report_app.html`, dùng bình thường.

## Cách 2 — Chạy ONLINE (ai cũng mở link là xuất được, không cần cài Python)

**Lưu ý:** GitHub/GitHub Pages chỉ host được file tĩnh (HTML), không chạy được
Python. Cần thêm 1 dịch vụ host Python — dùng **Render.com** (miễn phí):

### Bước 1 — Đưa toàn bộ thư mục này lên GitHub
(xem hướng dẫn upload ở phần trước — kéo-thả hoặc GitHub Desktop)

### Bước 2 — Deploy server lên Render
1. Vào **https://render.com** → đăng ký (có thể dùng tài khoản GitHub để đăng nhập luôn)
2. Bấm **New +** → **Web Service**
3. Chọn **Build and deploy from a Git repository** → kết nối GitHub → chọn repo `YasudaYenxaDatabase`
4. Điền:
   - **Name**: tuỳ ý, vd `yasuda-yenxa-export`
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
   - **Instance Type**: Free
5. Bấm **Create Web Service** → đợi build xong (vài phút)
6. Sau khi deploy xong, Render cho 1 địa chỉ dạng:
   `https://yasuda-yenxa-export.onrender.com`

### Bước 3 — Nối app HTML với server online
Mở `daily_report_app.html` bằng Notepad, tìm dòng:
```js
const LOCAL_SERVER_BASE = "http://localhost:8765";
```
Đổi thành URL Render vừa có (không có dấu `/` ở cuối):
```js
const LOCAL_SERVER_BASE = "https://yasuda-yenxa-export.onrender.com";
```
Lưu file, upload lại lên GitHub (ghi đè bản cũ).

### Bước 4 — Bật GitHub Pages để có link truy cập online
1. Vào repo trên GitHub → **Settings → Pages**
2. Ở mục **Source**, chọn nhánh `main`, thư mục `/ (root)` → **Save**
3. Sau vài phút, GitHub cho link dạng:
   `https://xubi-ysd.github.io/YasudaYenxaDatabase/daily_report_app.html`

Từ giờ, ai mở link đó cũng nhập liệu và xuất Excel trực tiếp được — không cần cài Python.

**Lưu ý về gói miễn phí của Render:** server sẽ "ngủ" sau ~15 phút không ai
dùng, lần truy cập đầu tiên sau đó có thể mất 30-60 giây để "thức dậy" —
bình thường, không phải lỗi.

## Ghi chú kỹ thuật

- Toạ độ ô nhập liệu được xác minh bằng cách so sánh trực tiếp file mẫu đã
  điền với bản trống tương ứng.
- Ô mô tả có sẵn (dòng 8–11) giữ nguyên nếu không nhập gì mới; ô số liệu để
  trống thật nếu không nhập.
- Cột "Loại nhân lực/thiết bị" chỉ đọc khi đã có sẵn nhãn; tự cho phép nhập
  mới ở dòng còn trống.
