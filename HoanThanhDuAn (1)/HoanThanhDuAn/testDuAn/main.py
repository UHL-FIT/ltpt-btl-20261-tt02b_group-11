import tkinter as tk

# Import hàm tạo giao diện từ file form.py
from views.form import tao_form

# Import tất cả hàm chức năng:
# thêm, sửa, xóa,...
from controllers.chucnang import *

# Import hàm thống kê từ file thongke.py
from controllers.thongke import thong_ke

# Tạo cửa sổ chính
root = tk.Tk()

# Đặt tiêu đề cửa sổ
root.title("Quản Lý Công thức Nấu Ăn")

# Kích thước cửa sổ
# ngang 900 - cao 600
root.geometry("900x600")

# Đổi màu nền giao diện
root.configure(bg="#F5F5F5")

# Gọi hàm tạo form giao diện
tao_form(

    # Cửa sổ chính
    root,

    # Hàm thêm món
    them_mon,

    # Hàm sửa món
    sua_mon,

    # Hàm xóa món
    xoa_mon,

    # Hàm thống kê
    thong_ke
)

# Chạy chương trình
root.mainloop()