import tkinter as tk
from tkinter import messagebox

# ================= THÊM MÓN =================
def them_mon(tree, entry_ten, entry_loai, entry_nguyenlieu, entry_thoigian):

    # Lấy dữ liệu từ ô nhập
    ten = entry_ten.get()
    loai = entry_loai.get()
    nguyenlieu = entry_nguyenlieu.get()
    thoigian = entry_thoigian.get()

    # ================= KIỂM TRA RỖNG =================
    if (
        ten == "" or
        loai == "" or
        nguyenlieu == "" or
        thoigian == ""
    ):

        # Hiện cảnh báo nếu thiếu dữ liệu
        messagebox.showwarning(
            "Thông báo",
            "Vui lòng nhập đầy đủ thông tin"
        )

        return

    # ================= THÊM DỮ LIỆU VÀO BẢNG =================
    tree.insert("", tk.END, values=(
        ten,
        loai,
        nguyenlieu,
        thoigian
    ))

    # ================= XÓA NỘI DUNG Ô NHẬP =================
    entry_ten.delete(0, tk.END)
    entry_loai.delete(0, tk.END)
    entry_nguyenlieu.delete(0, tk.END)
    entry_thoigian.delete(0, tk.END)

# ================= XÓA MÓN =================
def xoa_mon(tree):

    # Lấy dòng đang chọn
    selected = tree.selection()

    # Nếu chưa chọn dòng
    if not selected:
        messagebox.showwarning(
            "Thông báo",
            "Chọn món cần xóa"
        )
        return

    # Xóa dòng khỏi bảng
    tree.delete(selected)

# ================= SỬA MÓN =================
def sua_mon(tree, entry_ten, entry_loai, entry_nguyenlieu, entry_thoigian):

    # Lấy dòng đang chọn
    selected = tree.selection()

    # Kiểm tra đã chọn dòng chưa
    if not selected:
        messagebox.showwarning(
            "Thông báo",
            "Chọn món cần sửa"
        )
        return

    # ================= CẬP NHẬT DỮ LIỆU =================
    tree.item(selected, values=(
        entry_ten.get(),
        entry_loai.get(),
        entry_nguyenlieu.get(),
        entry_thoigian.get()
    ))

# ================= CHỌN MÓN =================
def chon_mon(
    event,
    tree,
    entry_ten,
    entry_loai,
    entry_nguyenlieu,
    entry_thoigian
):

    # Lấy dòng đang chọn
    selected = tree.selection()

    # Nếu có dòng được chọn
    if selected:

        # Lấy dữ liệu của dòng
        values = tree.item(selected[0], "values")

        # ================= HIỆN DỮ LIỆU LÊN Ô NHẬP =================

        # Hiện tên món
        entry_ten.delete(0, tk.END)
        entry_ten.insert(0, values[0])

        # Hiện loại món
        entry_loai.delete(0, tk.END)
        entry_loai.insert(0, values[1])

        # Hiện nguyên liệu
        entry_nguyenlieu.delete(0, tk.END)
        entry_nguyenlieu.insert(0, values[2])

        # Hiện thời gian
        entry_thoigian.delete(0, tk.END)
        entry_thoigian.insert(0, values[3])