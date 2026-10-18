# 🚀 Qr Grandstream by CLG - Cloud Google Sheets & GitHub Pages

Hệ thống quản lý quét mã thiết bị Grandstream tự động hóa hoàn toàn trên nền tảng Cloud:
- **Lưu trữ dữ liệu**: Trực tiếp vào **Google Sheets** (thay thế file Excel nội bộ).
- **Xử lý Webhook**: Sử dụng **Google Apps Script** (chạy serverless 24/7, không cần mở máy tính chạy Python Flask).
- **Giao diện Dashboard**: Triển khai trên **GitHub Pages** (giao diện hiện đại, Dark/Light Mode, đồng bộ dữ liệu thời gian thực).

---

## 🌟 Ưu Điểm So Với Bản Chạy Excel Trước Đây
1. **Không cần chạy server máy tính**: Bạn không cần phải mở máy tính bật `python appv3.py`.
2. **Quét mọi lúc mọi nơi**: Điện thoại dùng 4G/5G hoặc bất kỳ mạng WiFi nào cũng có thể quét và gửi dữ liệu thành công.
3. **Cộng tác thời gian thực**: Toàn bộ nhân viên có thể xem và chia sẻ file Google Sheet cùng lúc.
4. **Cảnh báo trùng lặp thông minh**: Tự động phát hiện và chặn trùng địa chỉ MAC, đồng thời cảnh báo lên Dashboard.

---

## 🛠 Hướng Dẫn Cài Đặt Chi Tiết Từ A - Z

### Bước 1: Tạo Google Sheet & Dán mã Script
1. Truy cập [Google Sheets](https://sheets.new) để tạo một bảng tính mới. Đặt tên bảng tính, ví dụ: `Danh Sách Thiết Bị Grandstream`.
2. Trên thanh menu, chọn **Tiện ích mở rộng** (Extensions) > **Apps Script**.
3. Xóa toàn bộ mã mặc định trong tệp `Mã.gs` (hoặc `Code.gs`).
4. Mở tệp [`google_apps_script.js`](./google_apps_script.js) trong thư mục này, sao chép toàn bộ nội dung và dán vào Apps Script.
5. Nhấn biểu tượng 💾 **Lưu dự án** (Ctrl + S).

---

### Bước 2: Triển Khai (Deploy) Thành Web App
1. Ở góc trên bên phải của trang Apps Script, nhấn nút **Triển khai** (Deploy) > chọn **Tùy chọn triển khai mới...** (New deployment).
2. Nhấn vào biểu tượng bánh răng ⚙️ bên cạnh "Chọn loại" > chọn **Ứng dụng web** (Web app).
3. Điền các thông tin:
   - **Mô tả** (Description): `Webhook barKoder Grandstream`
   - **Thực thi dưới dạng** (Execute as): `Tôi` (Me - địa chỉ email của bạn)
   - **Ai có quyền truy cập** (Who has access): Chọn **Bất kỳ ai** (Anyone) *(Lưu ý: Bắt buộc chọn "Anyone" để ứng dụng barKoder và Dashboard gửi/nhận dữ liệu không bị chặn).*
4. Nhấn **Triển khai** (Deploy).
5. Google sẽ yêu cầu cấp quyền:
   - Chọn tài khoản Google của bạn.
   - Nhấn **Nâng cao** (Advanced) > Nhấn tiếp **Đi tới (Không an toàn)** / *Go to project (unsafe)*.
   - Nhấn **Cho phép** (Allow).
6. Sau khi hoàn tất, sao chép đường dẫn **URL ứng dụng web** (có dạng `https://script.google.com/macros/s/.../exec`).

---

### Bước 3: Cấu Hình Ứng Dụng Quét Mã (barKoder) Trên Điện Thoại
1. Mở ứng dụng **barKoder** trên điện thoại.
2. Vào **Settings** > Tìm mục **Webhook** (hoặc Batch Upload / Webhook URL).
3. Dán link **URL ứng dụng web** bạn vừa copy ở Bước 2.
4. Đảm bảo phương thức là **POST**.
5. Bây giờ, khi bạn quét đủ 3 mã (MAC, Serial, Password), barKoder sẽ tự động đẩy lên Google Sheet ngay lập tức!

---

### Bước 4: Đưa Lên GitHub & Bật GitHub Pages
1. Tạo một Repository mới trên tài khoản [GitHub](https://github.com/new) của bạn (ví dụ đặt tên: `qr-grandstream`).
2. Mở terminal tại thư mục này và chạy các lệnh sau để đẩy code lên GitHub:
   ```bash
   git init
   git add .
   git commit -m "Khoi tao he thong Grandstream Cloud"
   git branch -M main
   git remote add origin https://github.com/<tai-khoan-github-cua-ban>/qr-grandstream.git
   git push -u origin main
   ```
3. Sau khi đẩy code lên, truy cập vào trang repository trên GitHub:
   - Vào mục **Settings** > chọn mục **Pages** ở thanh menu bên trái.
   - Tại mục **Build and deployment** > **Branch**: Chọn nhánh `main` (hoặc `master`) và thư mục `/ (root)`.
   - Nhấn nút **Save**.
4. Chờ khoảng 1-2 phút, GitHub Pages sẽ cung cấp cho bạn một đường link website trực tuyến miễn phí (dạng `https://<tai-khoan>.github.io/qr-grandstream/`).

---

### Bước 5: Kết Nối Dashboard Với Google Sheet
1. Mở link GitHub Pages bạn vừa tạo ở Bước 4.
2. Nhấn nút **⚙️ Cấu hình** ở góc trên bên phải Dashboard:
   - Dán **URL ứng dụng web** (Google Apps Script Web App URL).
   - Dán đường link bảng tính **Google Sheet** của bạn (để tiện mở nhanh).
   - Nhấn **Lưu cấu hình**.
3. Bạn có thể nhấn nút **🧪 Test Scan** trên Dashboard để kiểm tra: hệ thống sẽ tự động thêm 1 thiết bị thử nghiệm vào Google Sheet và hiển thị ngay trên bảng Dashboard!

---

## 📁 Cấu Trúc Thư Mục
```text
├── index.html              # Trang Dashboard trực tuyến (chạy trên GitHub Pages)
├── google_apps_script.js   # Mã nguồn Apps Script để dán vào Google Sheets
├── logo.png                # Logo thương hiệu CLG
├── icon.ico                # Favicon cho trang web
├── .gitignore              # Tệp loại trừ khi đưa lên Git
├── README.md               # Hướng dẫn chi tiết
└── appv3.py                # File Python cũ (được giữ lại để tham khảo nếu cần)
```

---

## 🔒 Bảo Mật & Lưu Ý
- Link Web App Apps Script không công khai dữ liệu đăng nhập nhạy cảm ngoài các trường thiết bị quét.
- Cấu hình URL được lưu ngay trong trình duyệt của bạn (`localStorage`), không lưu hardcode lộ thông tin trong mã nguồn công khai.
"# qr-grandstream" 
