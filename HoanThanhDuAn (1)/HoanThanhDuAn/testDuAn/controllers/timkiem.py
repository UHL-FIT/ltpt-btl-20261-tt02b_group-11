import tkinter as tk
from tkinter import messagebox


def tim_kiem_mon(tree, entry_ten, cb_loaimon, cb_thoigian):
    # Lấy từ khóa từ các ô nhập
    ten_val = entry_ten.get().lower().strip()
    loaimon_val = cb_loaimon.get().strip()
    thoigian_val = cb_thoigian.get().strip()

    # Khởi tạo danh sách các item bị ẩn nếu chưa có
    if not hasattr(tree, "hidden_items"):
        tree.hidden_items = []

    # 1. Khôi phục lại toàn bộ item đã bị ẩn từ lần tìm kiếm trước
    for item in tree.hidden_items:
        # Kiểm tra xem item còn tồn tại không (phòng trường hợp đã bị xóa)
        if tree.exists(item):
            tree.reattach(item, "", tk.END)
    tree.hidden_items.clear()

    # Nếu tất cả các ô tìm kiếm đều trống hoặc ở trạng thái mặc định thì dừng lại
    is_ten_empty = (ten_val == "")
    is_loaimon_empty = (loaimon_val == "" or loaimon_val == "Tất cả")
    is_thoigian_empty = (thoigian_val == "" or thoigian_val == "Tất cả")

    if is_ten_empty and is_loaimon_empty and is_thoigian_empty:
        return

    tim_thay = False
    items_to_hide = []

    # 2. Duyệt qua các item hiện có trên bảng
    for item in tree.get_children():
        values = tree.item(item, "values")
        if not values or len(values) < 7:
            continue

        ten_mon = str(values[1]).lower()
        loai_mon = str(values[2]).strip()
        thoi_gian_str = str(values[5]).strip()

        # Kiểm tra Tên món
        match_ten = True
        if not is_ten_empty:
            if ten_val not in ten_mon:
                match_ten = False

        # Kiểm tra Loại món
        match_loai = True
        if not is_loaimon_empty:
            if loaimon_val.lower() != loai_mon.lower():
                match_loai = False

        # Kiểm tra Thời gian
        match_thoigian = True
        if not is_thoigian_empty:
            try:
                t = int(thoi_gian_str)
                if thoigian_val == "Nhanh (< 15 phút)":
                    match_thoigian = (t < 15)
                elif thoigian_val == "Trung bình (15 - 30 phút)":
                    match_thoigian = (15 <= t <= 30)
                elif thoigian_val == "Hơi lâu (30 - 45 phút)":
                    match_thoigian = (30 < t <= 45)
                elif thoigian_val == "Lâu (> 45 phút)":
                    match_thoigian = (t > 45)
            except ValueError:
                match_thoigian = False

        # Kết hợp điều kiện (AND)
        if match_ten and match_loai and match_thoigian:
            tim_thay = True
        else:
            # Lưu lại các item không khớp để ẩn đi
            items_to_hide.append(item)

    # 3. Xử lý kết quả
    if tim_thay:
        # Nếu có món khớp, tiến hành ẩn các món không khớp
        for item in items_to_hide:
            tree.detach(item)
            tree.hidden_items.append(item)
    else:
        # Nếu không tìm thấy, thông báo cho người dùng
        messagebox.showinfo(
            "Thông báo",
            "Không tìm thấy món ăn nào phù hợp với bộ lọc!"
        )