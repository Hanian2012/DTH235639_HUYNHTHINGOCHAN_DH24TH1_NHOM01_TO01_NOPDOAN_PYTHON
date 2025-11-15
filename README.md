# DTH235639_HUYNHTHINGOCHAN_DH24TH1_NHOM01_TO01_NOPDOAN_PYTHON
# Ứng dụng Quản lý Giáo viên Trường THPT 🧑‍🏫

## Giới thiệu
Đây là **Ứng dụng Quản lý Giáo viên** được phát triển bằng ngôn ngữ lập trình **Python** và thư viện **Tkinter** để xây dựng giao diện người dùng đồ họa (GUI). Ứng dụng này sử dụng **SQL Server** làm hệ quản trị cơ sở dữ liệu backend.

Mục tiêu của dự án là số hóa và tối ưu hóa công tác quản lý nhân sự tại các trường THPT, bao gồm hồ sơ, hợp đồng, nghỉ phép, và đặc biệt là quản lý tiền lương một cách hiệu quả và chính xác.

## 🎯 Tính năng nổi bật
Ứng dụng cung cấp một bộ công cụ quản lý toàn diện:

* **Quản lý Hồ sơ Giáo viên:** Thao tác CRUD (Thêm, Sửa, Xóa) và tìm kiếm thông tin cá nhân/chuyên môn của giáo viên.
* **Quản lý Hợp đồng Lao động:**
    * Lưu trữ, tra cứu chi tiết hợp đồng.
    * **Tạo và Xuất Hợp đồng tự động** ra file **Word** (`.docx`).
    * **Xuất Danh sách Hợp đồng** ra file **Excel** (`.xlsx`) để báo cáo.
* **✨ Quản lý Tiền Lương:** Quản lý lương cơ bản, phụ cấp, và hỗ trợ xuất **Bảng Lương** ra file **Excel**.
* **Quản lý Nghỉ phép:** Ghi nhận và theo dõi các đơn xin nghỉ phép.
* **Hệ thống Đăng nhập:** Đảm bảo tính bảo mật và phân quyền truy cập hệ thống.

## 🛠️ Công nghệ & Thư viện
| Thành phần | Công nghệ/Thư viện | Mục đích |
| :--- | :--- | :--- |
| Ngôn ngữ chính | **Python 3.x** | Ngôn ngữ lập trình chính. |
| Giao diện (GUI) | **Tkinter** | Xây dựng giao diện người dùng đồ họa. |
| Cơ sở dữ liệu | **SQL Server** | Lưu trữ và quản lý dữ liệu. |
| Kết nối DB | `pyodbc` | Kết nối Python với SQL Server. |
| Xử lý Excel | `openpyxl` | Tạo và ghi dữ liệu ra file Excel. |
| Xử lý Word | `python-docx` | Tạo file Word (Hợp đồng) tự động. |
| Calendar | `tkcalendar` | Hỗ trợ chọn ngày tháng trên giao diện. |

## 🚀 Hướng dẫn Cài đặt & Khởi chạy

### 1. Yêu cầu Hệ thống
* Đã cài đặt **Python 3.x**.
* Đã cài đặt và cấu hình **SQL Server** (cùng với **ODBC Driver**).

### 2. Cài đặt Thư viện Python
Mở Terminal hoặc Command Prompt và chạy lệnh sau để cài đặt các thư viện cần thiết:
```bash
pip install pyodbc tkcalendar openpyxl python-docx
