import tkinter as tk
from tkinter import messagebox
from collections import Counter, defaultdict

def lay_thong_ke_nhanh(data):
    if len(data) == 0:
        return "0", "Chưa có", "Chưa có", "Chưa có"

    tong_mon = str(len(data))

    # Tìm món có thời gian nấu lâu nhất (cột index 5 - ban đầu là index 4)
    mon_max = "Không rõ"
    max_thoigian = -1.0
    for row in data:
        try:
            val = float(row[5])
            if val > max_thoigian:
                max_thoigian = val
                mon_max = row[1]
        except Exception:
            pass

    # Tìm món có đánh giá cao nhất (cột index 6 - ban đầu là index 5)
    top_mon = "Không rõ"
    max_danhgia = ""
    for row in data:
        try:
            val = str(row[6])
            if val > max_danhgia:
                max_danhgia = val
                top_mon = row[1]
        except Exception:
            pass

    # Tìm nguyên liệu được sử dụng phổ biến nhất (cột index 3 - ban đầu là index 2)
    ds_nguyenlieu = []
    for row in data:
        try:
            nl = row[3]
            tach = str(nl).split(",")
            for item in tach:
                clean_item = item.strip().lower()
                if clean_item:
                    ds_nguyenlieu.append(clean_item)
        except Exception:
            pass

    nl_phobien = "Không rõ"
    if ds_nguyenlieu:
        try:
            counts = Counter(ds_nguyenlieu)
            nl_phobien = counts.most_common(1)[0][0].title()
        except Exception:
            pass

    return tong_mon, mon_max, top_mon, nl_phobien


def thong_ke(tree):
    # ===== LẤY DỮ LIỆU TRÊN LUỒNG CHÍNH =====
    data = []
    for item in tree.get_children():
        values = tree.item(item, "values")
        data.append(values)

    # ===== KIỂM TRA =====
    if len(data) == 0:
        messagebox.showinfo(
            "Thông báo",
            "Chưa có dữ liệu"
        )
        return

    # Hiển thị trạng thái chờ bằng con trỏ chuột
    root = tree.winfo_toplevel()
    root.config(cursor="watch")

    def worker():
        try:
            # ===== THỜI GIAN TRUNG BÌNH =====
            loai_mon_times = defaultdict(list)
            ds_nguyenlieu = []
            
            mon_max_name = "Không rõ"
            mon_max_time = -1.0
            
            for row in data:
                try:
                    loai = row[2]
                    thoigian = float(row[5])
                    loai_mon_times[loai].append(thoigian)
                    
                    if thoigian > mon_max_time:
                        mon_max_time = thoigian
                        mon_max_name = row[1]
                except Exception:
                    pass
                
                try:
                    nl = row[3]
                    tach = str(nl).split(",")
                    for item in tach:
                        clean_item = item.strip().lower()
                        if clean_item:
                            ds_nguyenlieu.append(clean_item)
                except Exception:
                    pass

            # Tính thời gian trung bình của từng nhóm
            tb_thoigian = {}
            for loai, times in loai_mon_times.items():
                if times:
                    tb_thoigian[loai] = sum(times) / len(times)

            # Thống kê nguyên liệu (top 5)
            counts = Counter(ds_nguyenlieu)
            thongke_nguyenlieu = counts.most_common(5)

            if mon_max_time == -1.0:
                mon_max_time = 0.0

            # Gọi lại luồng chính để hiển thị giao diện popup an toàn
            root.after(0, lambda: hien_thi_popup(tb_thoigian, thongke_nguyenlieu, mon_max_name, mon_max_time))
        except Exception as e:
            root.after(0, lambda err=e: error_callback(str(err)))

    def error_callback(err_msg):
        root.config(cursor="")
        messagebox.showerror("Lỗi", f"Không thể tính toán thống kê: {err_msg}")

    def hien_thi_popup(tb_thoigian, thongke_nguyenlieu, mon_max_name, mon_max_time):
        root.config(cursor="")

        # TẠO CỬA SỔ THỐNG KÊ TRÊN LUỒNG CHÍNH
        win = tk.Toplevel()
        win.title("Thống kê món ăn")
        win.geometry("700x500")
        win.configure(bg="#F8FAFC")

        # ===== TITLE =====
        lbl_title = tk.Label(
            win,
            text="📊 THỐNG KÊ MÓN ĂN",
            bg="#F8FAFC",
            fg="#0F172A",
            font=("Segoe UI", 18, "bold")
        )
        lbl_title.pack(pady=15)

        # ==================================================
        # FRAME THỜI GIAN
        # ==================================================
        frame_tb = tk.LabelFrame(
            win,
            text="⏱ Thời gian trung bình",
            font=("Segoe UI", 11, "bold"),
            bg="white",
            fg="#2563EB",
            padx=10,
            pady=10
        )
        frame_tb.pack(fill="x", padx=20, pady=10)

        for loai, tb in tb_thoigian.items():
            lbl = tk.Label(
                frame_tb,
                text=f"{loai}: {tb:.1f} phút",
                bg="white",
                font=("Segoe UI", 10)
            )
            lbl.pack(anchor="w")

        # ==================================================
        # FRAME NGUYÊN LIỆU
        # ==================================================
        frame_nl = tk.LabelFrame(
            win,
            text="🧂 Nguyên liệu dùng nhiều",
            font=("Segoe UI", 11, "bold"),
            bg="white",
            fg="#2563EB",
            padx=10,
            pady=10
        )
        frame_nl.pack(fill="x", padx=20, pady=10)

        for nl, sl in thongke_nguyenlieu:
            lbl = tk.Label(
                frame_nl,
                text=f"{nl}: {sl} lần",
                bg="white",
                font=("Segoe UI", 10)
            )
            lbl.pack(anchor="w")

        # ==================================================
        # FRAME MÓN LÂU NHẤT
        # ==================================================
        frame_max = tk.LabelFrame(
            win,
            text="🔥 Món chuẩn bị lâu nhất",
            font=("Segoe UI", 11, "bold"),
            bg="white",
            fg="#2563EB",
            padx=10,
            pady=10
        )
        frame_max.pack(fill="x", padx=20, pady=10)

        lbl_max = tk.Label(
            frame_max,
            text=f"{mon_max_name} - {int(mon_max_time) if mon_max_time == int(mon_max_time) else mon_max_time} phút",
            bg="white",
            font=("Segoe UI", 11, "bold"),
            fg="#EF4444"
        )
        lbl_max.pack(anchor="w")

        # ==================================================
        # NÚT ĐÓNG
        # ==================================================
        btn_close = tk.Button(
            win,
            text="Đóng",
            bg="#64748B",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            width=12,
            relief="flat",
            cursor="hand2",
            command=win.destroy
        )
        btn_close.bind("<Enter>", lambda e: btn_close.config(bg="#475569"))
        btn_close.bind("<Leave>", lambda e: btn_close.config(bg="#64748B"))
        btn_close.pack(pady=20)

    import threading
    threading.Thread(target=worker, daemon=True).start()