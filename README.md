# OPPO Experience & Service Store - KPI & Sales Dashboard

Trang web theo dõi trực quan các chỉ số hoạt động, điểm KPI theo danh mục, doanh thu và năng suất nhân viên của cửa hàng.

## Tính năng chính
1. **Hiển thị trực quan đa chiều**:
   - **Tổng quan KPI**: Tiến độ hoàn thành điểm thực tế so với Target tháng từng nhóm hàng (`Phone`, `Tai Nghe`, `Pad`, `Ốp Lưng`, `OPPO Care`, `MDCL`, `Sạc dự phòng`, `Watch`).
   - **Doanh số & Sản phẩm**: Biểu đồ xu hướng điểm & doanh thu theo ngày, cơ cấu đóng góp điểm, Top sản phẩm bán chạy, tỷ trọng Khách lẻ / Đại lý.
   - **Năng suất & Nhân viên**: Biểu đồ xếp chồng Volume kỹ thuật viên và bảng điểm đóng góp từng nhân sự.
   - **Tra cứu đơn hàng**: Lọc nhanh giao dịch từ cả hai nguồn `Center SO` và `GCSM SO`.
2. **Upload file Excel trực tiếp trên web**:
   - Nhấn nút **Upload File Excel (.xlsx)** hoặc **kéo thả file** báo cáo Tổng Hợp Sale vào trang web.
   - Hệ thống tự động bóc tách dữ liệu từ các sheet (`Tổng hợp`, `Center SO`, `GCSM SO`) ngay trên trình duyệt.
3. **Tự động lưu trữ lần upload gần nhất**:
   - Dữ liệu của lần upload gần nhất được tự động lưu vào `localStorage` của trình duyệt. Khi tải lại trang hoặc mở lại trình duyệt, dữ liệu mới nhất vẫn được giữ nguyên.
   - Hỗ trợ nút **Xuất JSON** (`default-data.js`) để bạn cập nhật dữ liệu mặc định lên kho lưu trữ GitHub bất cứ lúc nào.

## Hướng dẫn đưa lên GitHub Pages (Chạy miễn phí 100%)
1. Giải nén file `.zip` này.
2. Tạo một Repository mới trên GitHub (ví dụ: `oppo-shop-dashboard`).
3. Upload toàn bộ các file và thư mục bên trong (`index.html`, `README.md`, `assets/`, `data/`) lên nhánh `main` của Repository.
4. Vào mục **Settings** -> **Pages** trên GitHub:
   - Ở mục **Source**, chọn **Deploy from a branch**.
   - Chọn nhánh `main` và thư mục `/ (root)` -> nhấn **Save**.
5. Chờ khoảng 1–2 phút, GitHub sẽ cung cấp đường link trang web (dạng `https://<ten-tai-khoan>.github.io/oppo-shop-dashboard/`) để xem trên mọi thiết bị máy tính và điện thoại.
