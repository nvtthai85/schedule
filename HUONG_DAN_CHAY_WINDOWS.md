# HƯỚNG DẪN CHẠY ỨNG DỤNG TRÊN MÁY TÍNH WINDOWS (1-CLICK CHẠY NGAY)

Tài liệu này hướng dẫn chi tiết cách chạy ứng dụng **Hệ Thống Lập Kế Hoạch Kiểm Soát FECT** trên bất kỳ máy tính Windows nào theo tiêu chí: **Chỉ cần click là chạy, không cần cài đặt hay cấu hình thủ công.**

---

## 🌟 CÁCH 1: DÙNG FILE BATCH TỰ ĐỘNG `CHAY_UNG_DUNG.bat` (ĐƠN GIẢN & TIỆN LỢI NHẤT)

Trong thư mục mã nguồn đã tạo sẵn file **`CHAY_UNG_DUNG.bat`**.

### Cách thực hiện:
1. Copy toàn bộ thư mục ứng dụng sang máy Windows.
2. **Click đúp (double-click)** chuột vào file **`CHAY_UNG_DUNG.bat`**.

### Cơ chế thông minh của file:
- **Nếu máy đã cài Python**: File sẽ tự động kích hoạt môi trường, tự cài thư viện (chỉ 30 giây lần đầu) và tự bật trình duyệt web `http://localhost:8000`.
- **Nếu máy HOÀN TOÀN CHƯA CÓ PYTHON**:
  - File sẽ tự động tải bản **Python Portable chính chủ siêu nhẹ (~15MB)** từ python.org qua PowerShell.
  - Tự động giải nén vào thư mục `python_windows/` và cài đặt sẵn các thư viện cần thiết.
  - **Không yêu cầu quyền Admin (quyền quản trị viên)** của máy tính công ty.
  - Sau khi xong, trình duyệt Chrome / Edge sẽ tự động bật lên và sẵn sàng sử dụng.
  - *Từ các lần sau, bạn click đúp là ứng dụng chạy ngay lập tức trong 1 giây.*

---

## 🚀 CÁCH 2: ĐÓNG GÓI THÀNH FILE `.EXE` (CHẠY ĐỘC LẬP 100% NHƯ PHẦN MỀM WINDOWS)

Nếu bạn muốn tạo ra một file `.exe` duy nhất để gửi cho đồng nghiệp (họ chỉ cần click đúp vào file `.exe` là ứng dụng chạy và tự mở trình duyệt):

### Bước 1: Đóng gói (chỉ làm 1 lần trên máy tính có Python)
- Trong thư mục dự án, click đúp vào file **`DONG_GOI_EXE.bat`** (hoặc mở cmd gõ `python build_exe.py`).
- Quá trình đóng gói sẽ tự động chạy trong khoảng 1 phút.

### Bước 2: Nhận kết quả
- Sau khi đóng gói xong, trong thư mục dự án sẽ xuất hiện thư mục:
  `dist\Lap_Ke_Hoach_FECT\`
- Trong thư mục này có sẵn file:
  👉 **`Lap_Ke_Hoach_FECT.exe`**

### Bước 3: Phân phối & Sử dụng
- Bạn chỉ cần nén thư mục `Lap_Ke_Hoach_FECT` thành file `.zip` và gửi qua Zalo / Google Drive / Email cho đồng nghiệp.
- Người nhận chỉ cần giải nén ra và **click đúp vào `Lap_Ke_Hoach_FECT.exe`**:
  - Ứng dụng tự khởi động.
  - Tự động mở trình duyệt web `http://localhost:8000`.
  - **100% không cần cài Python, không cần cài thư viện, không cần gõ bất kỳ dòng lệnh nào.**

---

## 📦 CÁCH 3: TẠO GÓI PORTABLE OFFLINE (CHO MÁY KHÔNG CÓ MẠNG INTERNET)

Nếu máy tính đích ở phòng ban bảo mật, bị chặn mạng hoặc không kết nối Internet:

1. Trên một máy tính có kết nối mạng, chạy file `CHAY_UNG_DUNG.bat` lần đầu tiên để nó hoàn tất việc tải thư mục `python_windows/`.
2. Copy toàn bộ thư mục dự án (đã chứa thư mục `python_windows/`) vào USB.
3. Cắm USB sang máy tính không có mạng:
   - Click đúp vào **`CHAY_UNG_DUNG.bat`**.
   - Ứng dụng sẽ sử dụng trực tiếp Python nhúng trong thư mục `python_windows/` và khởi chạy ngay lập tức mà không cần kết nối mạng.

---

## 🛑 CÁCH TẮT ỨNG DỤNG KHI DÙNG XONG
- Khi không sử dụng nữa, bạn chỉ cần **đóng cửa sổ màu đen (Command Prompt)** hoặc nhấn tổ hợp phím **`Ctrl + C`** trên cửa sổ đó.
