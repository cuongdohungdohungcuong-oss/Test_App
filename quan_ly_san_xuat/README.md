# Quản lý sản xuất thức ăn chăn nuôi (Streamlit + Excel)

App độc lập, cùng cấu trúc với app quản lý kho ở thư mục gốc. Dữ liệu lưu trong `data/san_xuat.xlsx`.

## Chạy

```bash
cd quan_ly_san_xuat
pip install -r requirements.txt
streamlit run app.py
```

Tài khoản mặc định: `admin` / `admin123`. Hãy đổi mật khẩu sau khi triển khai.

Trong GitHub Codespaces / Dev Container, app kho chạy ở cổng **8501** và app sản xuất tự chạy ở cổng **8502**.

## Chức năng

| Tab | Nội dung |
|---|---|
| Lệnh sản xuất | Danh sách, lọc theo ngày/trạng thái; tạo/sửa lệnh (xem trước nhu cầu NL theo BOM và tồn kho); chuyển trạng thái Kế hoạch → Đang sản xuất → Hoàn thành / Hủy — **kiểm tra đủ NL tại kho xuất trước khi bắt đầu** (có thể bỏ qua có chủ đích); **in phiếu lệnh** (HTML, Ctrl+P để in/PDF) |
| Ghi nhận thực tế | Nhập sản lượng & tiêu hao NL thực tế; hoàn thành lệnh sẽ sinh phiếu `<MaLenh>-X` (xuất NL) và `<MaLenh>-N` (nhập TP) |
| Công thức (BOM) | Định mức kg NL / 1.000 kg thành phẩm, kiểm tra tổng = 1.000 |
| Báo cáo sản xuất | KPI, sản lượng KH vs TT theo ngày/TP/dây chuyền/ca, hao hụt NL; tải Excel |
| Tồn kho | Tồn hiện tại theo kho và mã hàng; phiếu **Nhập kho** (`PN-…`) và **Điều chỉnh** tồn (`DC-…`, số âm để giảm) |
| Phiếu kho | Toàn bộ phiếu kho: sinh từ sản xuất và nhập tay (cùng cột với sheet `GiaoDich` của app kho) |
| Danh mục | Thành phẩm, nguyên liệu, dây chuyền, kho, người dùng (quản trị) |

## Vai trò

- `QuanTri`: toàn quyền, cấp tài khoản
- `QuanLySanXuat`: lập kế hoạch, công thức, danh mục, ghi nhận
- `ToTruongSX`: bắt đầu lệnh, ghi nhận thực tế
- `XemBaoCao`: chỉ xem

## Cấu trúc

| File | Vai trò |
|---|---|
| `app.py` | Giao diện Streamlit |
| `excel_io.py` | Đọc/ghi Excel, tự bổ sung sheet còn thiếu |
| `production.py` | Nghiệp vụ (BOM, lệnh, hoàn thành, báo cáo), không phụ thuộc Streamlit |
| `sample_sx.py` | Sinh lệnh sản xuất mẫu (`python sample_sx.py`) |
| `auth.py` | Băm mật khẩu PBKDF2 |
| `tests/` | `pytest tests` |
