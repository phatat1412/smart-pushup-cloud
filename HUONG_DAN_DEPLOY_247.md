# 🚀 HƯỚNG DẪN DEPLOY SMART PUSH-UP LÊN CLOUD 24/7 (KHÔNG CẦN BẬT LAPTOP)

Hệ thống đã được thiết kế hoàn chỉnh theo kiến trúc Cloud IoT toàn cầu:
- Cloud MQTT Broker: broker.emqx.io (Port 1883 TCP cho ESP32, Port 8084 WSS cho Web/App)
- Topic cá nhân: pushup_ntp113/device01/... (Không bao giờ trùng lặp)
- Cloud Server: Chạy trên Render.com hoặc Glitch.com (Miễn phí 100%, có HTTPS)

---

## CÁCH 1: DEPLOY LÊN RENDER.COM (KHUYÊN DÙNG CHO ĐỒ ÁN - 3 BƯỚC)

### Bước 1: Đẩy mã nguồn lên GitHub của bạn
Mở CMD hoặc PowerShell tại thư mục này (D:\IOT\Smart Push Up\cloud_server) và chạy:
git remote add origin https://github.com/<TEN_GITHUB_CUA_BAN>/smart-pushup-cloud.git
git push -u origin main

### Bước 2: Kết nối với Render.com
1. Truy cập: https://render.com -> Đăng nhập bằng tài khoản GitHub.
2. Bấm nút New + (màu xanh góc phải) -> Chọn Web Service.
3. Chọn repository smart-pushup-cloud vừa đẩy lên.
4. Render sẽ tự động điền cấu hình:
   - Name: smart-pushup-ntp113
   - Runtime: Node
   - Build Command: npm install
   - Start Command: npm start
   - Plan: Free
5. Bấm Deploy Web Service.

Trong 1 phút, Render sẽ cấp cho bạn đường link cố định vĩnh viễn:
👉 https://smart-pushup-ntp113.onrender.com

---

## CÁCH 2: DEPLOY LÊN GLITCH.COM (SIÊU NHANH 30 GIÂY - KHÔNG CẦN GITHUB)
1. Truy cập: https://glitch.com -> Đăng nhập (Google / Facebook).
2. Bấm New Project -> Chọn glitch-hello-node.
3. Kéo toàn bộ các file trong thư mục này (cloud_server) thả vào cửa sổ Glitch.
4. Glitch sẽ tự động chạy ngay lập tức và cấp cho bạn link web 24/7:
👉 https://smart-pushup-ntp113.glitch.me

---

## BƯỚC CUỐI: NẠP CODE VÀO ESP32
1. Mở file D:\IOT\CODE_TEST\CODE_TEST.ino trong Arduino IDE.
2. Bấm nút Upload (Mũi tên sang phải) để nạp vào ESP32.
3. Cắm ESP32 vào sạc dự phòng.
4. TẮT LAPTOP HOÀN TOÀN! Mở điện thoại vào link Cloud ở trên để tập luyện mọi lúc mọi nơi!
