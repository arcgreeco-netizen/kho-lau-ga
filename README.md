# Kho Lẩu gà – Web app quản lý kho (Google Apps Script)

Trang web chạy trên điện thoại, gắn vào một Google Sheet quản lý kho nhà hàng. Nhân viên đăng nhập bằng Gmail, chỉ thấy các mục được chủ quán cấp quyền.

## Chức năng

| Mục | Làm gì | Ghi vào tab |
|---|---|---|
| 📋 Kiểm kho | Đếm tồn cuối ngày: cột NGUYÊN (thùng / gói / cây • kg) và LẺ (lon / viên / cái), có nhãn ✅ Khớp / ⚠ Thiếu / ℹ Dư so với máy | Phiếu kiểm kho ngày |
| 🛒 Order | Bếp ghi SL đặt / SL nhận cho Rau – Thịt và Hàng khô | Order Rau - Thịt, Order Hàng khô |
| 📦 Xuất/Nhập | Ghi xuất kho (ra quầy / bếp) và nhập hàng NCC (giá, SL, kg gà) | Xuất kho, Nhập hàng |
| 📊 Đối chiếu | Xem dùng thực tế vs theo bán / định lượng, và kiểm kê tuần; lọc nhanh món lệch | Đối chiếu NL, Kiểm kê tuần (chỉ đọc) |
| ⚙ Phân quyền | Chủ quán tick mục cho từng email; lưu xong tự chia sẻ Sheet cho email mới | Phân quyền |

## File

- `WebApp.gs` – xử lý dữ liệu (đọc / ghi Sheet), kiểm tra quyền theo email.
- `Index.html` – giao diện điện thoại (mỗi món 1 hàng, menu 1 hàng có icon).
- `Ma.gs` – menu "Order Bếp" trong Google Sheet.

## Cài đặt

1. Mở Google Sheet → **Tiện ích → Apps Script**.
2. Tạo các file `WebApp.gs`, `Index.html` (loại HTML), dán nội dung tương ứng; `Ma.gs` dán vào file `Mã.gs`.
3. **Triển khai → Tùy chọn triển khai mới → Ứng dụng web**:
   - Thực thi bằng tên: **Người dùng truy cập vào ứng dụng web**
   - Người có quyền truy cập: **Bất cứ ai có Tài khoản Google**
4. Mở link `/exec` bằng tài khoản chủ Sheet → bấm **⚙** để thêm email nhân viên và tick quyền.
5. Mỗi nhân viên mở link lần đầu phải bấm **Cho phép** quyền truy cập (script tự tạo nên Google báo "chưa xác minh").

Sheet cần có các tab: Danh mục, Phiếu kiểm kho ngày, Order Rau - Thịt, Order Hàng khô, Xuất kho, Nhập hàng, Đối chiếu NL, Kiểm kê tuần. Tab Phân quyền tự tạo.

## Cập nhật code

Sửa file trong Apps Script → Lưu → **Triển khai → Quản lý các tùy chọn triển khai → ✏ → Phiên bản: Phiên bản mới → Triển khai** (link giữ nguyên).
