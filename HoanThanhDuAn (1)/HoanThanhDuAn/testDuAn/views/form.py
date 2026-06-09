import tkinter as tk
from tkinter import ttk

from controllers.chucnang import chon_mon, xuat_word
from controllers.timkiem import tim_kiem_mon
from controllers.thongke import lay_thong_ke_nhanh

from views.style import cai_dat_style
from models.dulieu_mau import them_du_lieu_mau


def hien_thi_chi_tiet_mon(parent, values):
    # values: (Chọn, Tên món, Loại món, Nguyên liệu, Định lượng, Thời gian, Đánh giá)
    if len(values) < 7:
        return
    
    win = tk.Toplevel(parent)
    win.title(f"Chi tiết công thức: {values[1]}")
    win.geometry("650x400")
    win.configure(bg="#F8FAFC")
    win.transient(parent)
    win.grab_set()
    
    # Title Frame
    title_frame = tk.Frame(win, bg="#1E293B", pady=20)
    title_frame.pack(fill="x")
    
    tk.Label(title_frame, text=values[1].upper(), bg="#1E293B", fg="#F8FAFC", font=("Segoe UI", 20, "bold")).pack()
    tk.Label(title_frame, text=f"Loại món: {values[2]}  •  Đánh giá: {values[6]}", bg="#1E293B", fg="#94A3B8", font=("Segoe UI", 11, "italic")).pack(pady=(5, 0))
    
    # Content Frame
    content = tk.Frame(win, bg="#F8FAFC", padx=25, pady=20)
    content.pack(fill="both", expand=True)
    
    # Grid of quick specs
    specs_frame = tk.Frame(content, bg="white", bd=1, relief="flat", highlightbackground="#E2E8F0", highlightthickness=1)
    specs_frame.pack(fill="x", pady=(0, 15))
    
    def add_spec(parent_f, icon, label, val, row, col):
        f = tk.Frame(parent_f, bg="white", padx=15, pady=10)
        f.grid(row=row, column=col, sticky="nsew")
        tk.Label(f, text=f"{icon} {label}", bg="white", fg="#64748B", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        tk.Label(f, text=val, bg="white", fg="#1E293B", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(2, 0))
        
    specs_frame.columnconfigure(0, weight=1)
    specs_frame.columnconfigure(1, weight=1)
    specs_frame.columnconfigure(2, weight=1)
    
    add_spec(specs_frame, "⚖", "Định lượng", values[4], 0, 0)
    add_spec(specs_frame, "⏱", "Thời gian", f"{values[5]} phút", 0, 1)
    add_spec(specs_frame, "⭐", "Đánh giá", values[6], 0, 2)
    
    # Ingredients card
    ing_frame = tk.LabelFrame(content, text="🧂 NGUYÊN LIỆU CHÍNH", font=("Segoe UI", 11, "bold"), bg="white", fg="#2563EB", bd=1, relief="flat", highlightbackground="#E2E8F0", highlightthickness=1, padx=15, pady=10)
    ing_frame.pack(fill="both", expand=True, pady=(0, 15))
    
    lbl_ing = tk.Label(ing_frame, text=values[3], bg="white", fg="#334155", font=("Segoe UI", 11), wraplength=550, justify="left")
    lbl_ing.pack(anchor="w")
    
    # Bottom Close Button
    btn_close = tk.Button(win, text="Đóng cửa sổ", bg="#64748B", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2", command=win.destroy)
    btn_close.pack(pady=15, ipady=8, ipadx=20)
    btn_close.bind("<Enter>", lambda e: btn_close.config(bg="#475569"))
    btn_close.bind("<Leave>", lambda e: btn_close.config(bg="#64748B"))


def tao_form(root, them_mon, sua_mon, xoa_mon, thong_ke_func):
    dang_them = False

    root.configure(bg="#F1F5F9")
    root.geometry("1200x700")
    cai_dat_style()

    # ================= HIỆU ỨNG =================
    def add_hover(btn, bg_normal, bg_hover):
        btn.bind("<Enter>", lambda e: btn.config(bg=bg_hover))
        btn.bind("<Leave>", lambda e: btn.config(bg=bg_normal))

    # ================= LAYOUT CHÍNH =================
    nav_frame = tk.Frame(root, bg="#1E293B", width=220)
    nav_frame.pack(side="left", fill="y")
    nav_frame.pack_propagate(False)

    content_frame = tk.Frame(root, bg="#F1F5F9")
    content_frame.pack(side="right", fill="both", expand=True)

    page_danh_sach = tk.Frame(content_frame, bg="#F1F5F9")
    page_nhap_lieu = tk.Frame(content_frame, bg="#F1F5F9")
    page_thong_ke = tk.Frame(content_frame, bg="#F1F5F9")

    # ================= KHỞI TẠO CÁC BIẾN DỮ LIỆU =================
    columns = ("Chọn", "Tên món", "Loại món", "Nguyên liệu", "Định lượng", "Thời gian", "Đánh giá")
    frame_table = tk.Frame(page_danh_sach, bg="white", bd=1, highlightbackground="#E2E8F0", highlightthickness=1, relief="flat")
    tree = ttk.Treeview(frame_table, columns=columns, show="headings", height=15)
    
    frame_input = tk.Frame(page_nhap_lieu, bg="white", bd=1, highlightbackground="#E2E8F0", highlightthickness=1, relief="flat")
    
    entry_ten = tk.Entry(frame_input)
    entry_loai = tk.Entry(frame_input)
    entry_nguyenlieu = tk.Entry(frame_input)
    entry_dinhluong = tk.Entry(frame_input)
    entry_thoigian = tk.Entry(frame_input)
    entry_danhgia = tk.Entry(frame_input)

    frame_info = tk.Frame(page_thong_ke, bg="white", bd=1, highlightbackground="#E2E8F0", highlightthickness=1, relief="flat")
    lbl_tongmon = tk.Label(frame_info, bg="white", fg="#334155", font=("Segoe UI", 16))
    lbl_launhat = tk.Label(frame_info, bg="white", fg="#334155", font=("Segoe UI", 16))
    lbl_top = tk.Label(frame_info, bg="white", fg="#334155", font=("Segoe UI", 16))
    lbl_nguyenlieu = tk.Label(frame_info, bg="white", fg="#334155", font=("Segoe UI", 16))

    # ================= HÀM CHUYỂN TRANG =================
    def show_page(page):
        page_danh_sach.pack_forget()
        page_nhap_lieu.pack_forget()
        page_thong_ke.pack_forget()
        page.pack(fill="both", expand=True)
        if page == page_thong_ke:
            update_quick_stats()

    # ================= HÀM WRAPPER =================
    def update_quick_stats():
        # Lấy dữ liệu trên luồng chính để đảm bảo thread-safe
        data = []
        for item in tree.get_children():
            data.append(tree.item(item, "values"))

        if hasattr(tree, "hidden_items"):
            for item in tree.hidden_items:
                if tree.exists(item):
                    data.append(tree.item(item, "values"))

        # Hiển thị đang tải trên UI
        lbl_tongmon.config(text="🍽 Tổng món: Đang nạp...")
        lbl_launhat.config(text="⏰ Lâu nhất: Đang nạp...")
        lbl_top.config(text="⭐ Đánh giá cao: Đang nạp...")
        lbl_nguyenlieu.config(text="🥩 Phổ biến: Đang nạp...")

        import threading
        def worker():
            try:
                tong_mon, mon_max, top_mon, nl_phobien = lay_thong_ke_nhanh(data)
                root.after(0, lambda: lbl_tongmon.config(text=f"🍽 Tổng món: {tong_mon}"))
                root.after(0, lambda: lbl_launhat.config(text=f"⏰ Lâu nhất: {mon_max}"))
                root.after(0, lambda: lbl_top.config(text=f"⭐ Đánh giá cao: {top_mon}"))
                root.after(0, lambda: lbl_nguyenlieu.config(text=f"🥩 Phổ biến: {nl_phobien}"))
            except Exception:
                root.after(0, lambda: lbl_tongmon.config(text="🍽 Tổng món: Lỗi nạp"))

        threading.Thread(target=worker, daemon=True).start()

    def them_wrapper():
        nonlocal dang_them
        dang_them = True
        them_mon(tree, entry_ten, entry_loai, entry_nguyenlieu, entry_dinhluong, entry_thoigian, entry_danhgia)
        update_quick_stats()
        update_header_checkbox()
        show_page(page_danh_sach)

    def sua_wrapper():
        sua_mon(tree, entry_ten, entry_loai, entry_nguyenlieu, entry_dinhluong, entry_thoigian, entry_danhgia)
        update_quick_stats()
        update_header_checkbox()
        show_page(page_danh_sach)

    def xoa_wrapper():
        xoa_mon(tree, entry_ten, entry_loai, entry_nguyenlieu, entry_dinhluong, entry_thoigian, entry_danhgia)
        update_quick_stats()
        update_header_checkbox()

    # ================= XÂY DỰNG NAV BAR =================
    tk.Label(nav_frame, text="COOKING\nMANAGER", bg="#1E293B", fg="#F8FAFC", font=("Segoe UI", 20, "bold")).pack(pady=40)
    
    def tao_nut_nav(text, page):
        btn = tk.Button(
            nav_frame, text=text, bg="#334155", fg="white", font=("Segoe UI", 12, "bold"),
            relief="flat", cursor="hand2", activebackground="#475569", activeforeground="white",
            command=lambda: show_page(page)
        )
        btn.pack(fill="x", padx=15, pady=5, ipady=12)
        add_hover(btn, "#334155", "#475569")

    tao_nut_nav("📋 Danh Sách", page_danh_sach)
    tao_nut_nav("✍️ Nhập Liệu", page_nhap_lieu)
    tao_nut_nav("📊 Thống Kê", page_thong_ke)

    btn_chitiet = tk.Button(
        nav_frame, text="📈 Bảng Chi Tiết", bg="#F59E0B", fg="white", font=("Segoe UI", 12, "bold"),
        relief="flat", cursor="hand2", command=lambda: thong_ke_func(tree)
    )
    btn_chitiet.pack(fill="x", padx=15, pady=40, ipady=12)
    add_hover(btn_chitiet, "#F59E0B", "#D97706")

    # ================= TRANG 1: DANH SÁCH =================
    top_danhsach = tk.Frame(page_danh_sach, bg="white", height=70, highlightbackground="#E2E8F0", highlightthickness=1)
    top_danhsach.pack(fill="x")
    tk.Label(top_danhsach, text="📋 DANH SÁCH CÔNG THỨC", bg="white", fg="#0F172A", font=("Segoe UI", 20, "bold")).pack(side="left", padx=20, pady=15)

    # Thanh bộ lọc tìm kiếm kết hợp mới (Căn giữa)
    frame_filter = tk.Frame(page_danh_sach, bg="white", highlightbackground="#E2E8F0", highlightthickness=1)
    frame_filter.pack(fill="x", padx=25, pady=(15, 5))

    # Container bên trong giúp căn giữa toàn bộ thanh tìm kiếm & bộ lọc
    container_filter = tk.Frame(frame_filter, bg="white")
    container_filter.pack(anchor="center", pady=10)

    # 1. Nhóm Tìm kiếm theo Tên món
    lbl_ten = tk.Label(container_filter, text="🍜 Tên món:", bg="white", fg="#475569", font=("Segoe UI", 10, "bold"))
    lbl_ten.pack(side="left", padx=(5, 2))
    
    entry_tim_ten = tk.Entry(container_filter, font=("Segoe UI", 10), width=18, relief="flat", bd=0, highlightbackground="#CBD5E1", highlightcolor="#3B82F6", highlightthickness=1)
    entry_tim_ten.pack(side="left", padx=5, ipady=4)

    def thuc_hien_tim_nhanh(event=None):
        # Đặt lại các bộ lọc nâng cao khi thực hiện tìm nhanh theo tên
        cb_tim_loai.set("Tất cả")
        cb_tim_tg.set("Tất cả")
        tim_kiem_mon(tree, entry_tim_ten, cb_tim_loai, cb_tim_tg)

    # Bind phím Enter để tìm nhanh
    entry_tim_ten.bind("<Return>", thuc_hien_tim_nhanh)

    btn_tim_ten = tk.Button(
        container_filter, text="Tìm", bg="#3B82F6", fg="white", font=("Segoe UI", 10, "bold"),
        relief="flat", cursor="hand2", command=thuc_hien_tim_nhanh
    )
    btn_tim_ten.pack(side="left", padx=(2, 15), ipadx=10, ipady=4)
    add_hover(btn_tim_ten, "#3B82F6", "#2563EB")

    # Vạch ngăn cách dọc tinh tế giữa Tìm kiếm và Lọc nâng cao
    lbl_divider = tk.Label(container_filter, text="|", bg="white", fg="#CBD5E1", font=("Segoe UI", 12))
    lbl_divider.pack(side="left", padx=10)
    
    # 2. Nhóm Lọc theo Loại và Thời gian
    lbl_loai = tk.Label(container_filter, text="🥗 Loại món:", bg="white", fg="#475569", font=("Segoe UI", 10, "bold"))
    lbl_loai.pack(side="left", padx=(5, 2))
    
    cb_tim_loai = ttk.Combobox(container_filter, font=("Segoe UI", 10), width=12, state="readonly")
    cb_tim_loai.pack(side="left", padx=5, ipady=2)
    cb_tim_loai.set("Tất cả")
    
    def cap_nhat_loai_mon():
        types = set()
        for item in tree.get_children():
            vals = tree.item(item, "values")
            if vals and len(vals) > 2:
                types.add(vals[2].strip())
        if hasattr(tree, "hidden_items"):
            for item in tree.hidden_items:
                if tree.exists(item):
                    vals = tree.item(item, "values")
                    if vals and len(vals) > 2:
                        types.add(vals[2].strip())
        cb_tim_loai['values'] = ["Tất cả"] + sorted(list(types))

    cb_tim_loai.config(postcommand=cap_nhat_loai_mon)
    
    lbl_tg = tk.Label(container_filter, text="⏱ Thời gian:", bg="white", fg="#475569", font=("Segoe UI", 10, "bold"))
    lbl_tg.pack(side="left", padx=(10, 2))
    
    cb_tim_tg = ttk.Combobox(container_filter, font=("Segoe UI", 10), width=16, state="readonly")
    cb_tim_tg.pack(side="left", padx=5, ipady=2)
    cb_tim_tg['values'] = ("Tất cả", "Nhanh (< 15 phút)", "Trung bình (15 - 30 phút)", "Hơi lâu (30 - 45 phút)", "Lâu (> 45 phút)")
    cb_tim_tg.set("Tất cả")

    btn_tim = tk.Button(
        container_filter, text="🔍 Lọc", bg="#10B981", fg="white", font=("Segoe UI", 10, "bold"), 
        relief="flat", cursor="hand2", command=lambda: tim_kiem_mon(tree, entry_tim_ten, cb_tim_loai, cb_tim_tg)
    )
    btn_tim.pack(side="left", padx=(15, 5), ipadx=12, ipady=4)
    add_hover(btn_tim, "#10B981", "#059669")

    def reset_bo_loc():
        entry_tim_ten.delete(0, tk.END)
        cb_tim_loai.set("Tất cả")
        cb_tim_tg.set("Tất cả")
        tim_kiem_mon(tree, entry_tim_ten, cb_tim_loai, cb_tim_tg)

    btn_reset = tk.Button(
        container_filter, text="🔄 Reset", bg="#64748B", fg="white", font=("Segoe UI", 10, "bold"),
        relief="flat", cursor="hand2", command=reset_bo_loc
    )
    btn_reset.pack(side="left", padx=5, ipadx=10, ipady=4)
    add_hover(btn_reset, "#64748B", "#475569")

    # Khung chứa các nút Chỉnh sửa, Xóa, Xuất Word ở dưới cùng
    frame_btn_ds = tk.Frame(page_danh_sach, bg="#F1F5F9")
    frame_btn_ds.pack(side="bottom", fill="x", padx=25, pady=(10, 20))
    
    center_btn_ds = tk.Frame(frame_btn_ds, bg="#F1F5F9")
    center_btn_ds.pack(anchor="center")

    btn_edit = tk.Button(center_btn_ds, text="✏ Chỉnh sửa", bg="#F59E0B", fg="white", font=("Segoe UI", 11, "bold"), relief="flat", cursor="hand2", command=lambda: show_page(page_nhap_lieu))
    btn_edit.pack(side="left", padx=10, ipady=8, ipadx=20)
    add_hover(btn_edit, "#F59E0B", "#D97706")
    
    btn_del = tk.Button(center_btn_ds, text="🗑 Xóa món", bg="#EF4444", fg="white", font=("Segoe UI", 11, "bold"), relief="flat", cursor="hand2", command=xoa_wrapper)
    btn_del.pack(side="left", padx=10, ipady=8, ipadx=20)
    add_hover(btn_del, "#EF4444", "#DC2626")

    btn_export = tk.Button(center_btn_ds, text="📝 Xuất Word", bg="#2563EB", fg="white", font=("Segoe UI", 11, "bold"), relief="flat", cursor="hand2", command=lambda: xuat_word(tree))
    btn_export.pack(side="left", padx=10, ipady=8, ipadx=20)
    add_hover(btn_export, "#2563EB", "#1D4ED8")

    frame_table.pack(side="top", fill="both", expand=True, padx=25, pady=(10, 10))
    
    for col in columns:
        if col == "Chọn":
            tree.heading(col, text="☐")
            tree.column(col, width=50, minwidth=50, anchor="center", stretch=tk.NO)
        else:
            tree.heading(col, text=col)
            if col == "Tên món": tree.column(col, minwidth=150, anchor="w", stretch=tk.YES)
            elif col == "Nguyên liệu": tree.column(col, minwidth=180, anchor="w", stretch=tk.YES)
            else: tree.column(col, minwidth=80, anchor="center", stretch=tk.YES)

    def auto_resize_columns(event):
        total = event.width
        if total > 600:
            tree.column("Chọn", width=50)
            remaining = total - 50
            tree.column("Tên món", width=int(remaining * 0.25))
            tree.column("Loại món", width=int(remaining * 0.15))
            tree.column("Nguyên liệu", width=int(remaining * 0.30))
            tree.column("Định lượng", width=int(remaining * 0.10))
            tree.column("Thời gian", width=int(remaining * 0.10))
            tree.column("Đánh giá", width=int(remaining * 0.10))
            
    tree.bind("<Configure>", auto_resize_columns)

    scroll_y = ttk.Scrollbar(frame_table, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scroll_y.set)
    scroll_y.pack(side="right", fill="y")
    tree.pack(fill="both", expand=True)

    # ================= LOGIC CHECKBOX CHỌN TẤT CẢ / CHỌN TỪNG DÒNG =================
    global_select_all = False

    def update_header_checkbox():
        nonlocal global_select_all
        all_items = tree.get_children()
        if not all_items:
            global_select_all = False
            tree.heading("Chọn", text="☐")
            return
        
        all_checked = True
        for item in all_items:
            vals = tree.item(item, "values")
            if vals and vals[0] == "☐":
                all_checked = False
                break
        global_select_all = all_checked
        tree.heading("Chọn", text="☑" if all_checked else "☐")

    def toggle_all_checkboxes():
        nonlocal global_select_all
        global_select_all = not global_select_all
        new_val = "☑" if global_select_all else "☐"
        tree.heading("Chọn", text=new_val)
        for item in tree.get_children():
            vals = list(tree.item(item, "values"))
            if vals:
                vals[0] = new_val
                tree.item(item, values=vals)

    def toggle_row_checkbox(item):
        vals = list(tree.item(item, "values"))
        if vals:
            vals[0] = "☑" if vals[0] == "☐" else "☐"
            tree.item(item, values=vals)
            update_header_checkbox()

    def check_click(event):
        region = tree.identify("region", event.x, event.y)
        if region == "heading":
            col = tree.identify_column(event.x)
            if col == "#1": # Cột Chọn
                toggle_all_checkboxes()
                return "break"
        elif region == "cell":
            col = tree.identify_column(event.x)
            if col == "#1": # Cột Chọn
                item = tree.identify_row(event.y)
                if item:
                    toggle_row_checkbox(item)
                return "break"

    tree.bind("<ButtonRelease-1>", check_click)

    def chon_mon_wrapper(event):
        nonlocal dang_them
        if dang_them:
            dang_them = False
            return
        chon_mon(event, tree, entry_ten, entry_loai, entry_nguyenlieu, entry_dinhluong, entry_thoigian, entry_danhgia)
    tree.bind("<<TreeviewSelect>>", chon_mon_wrapper)

    def xem_chi_tiet(event):
        item = tree.identify_row(event.y)
        if not item:
            return
        vals = tree.item(item, "values")
        if not vals:
            return
        hien_thi_chi_tiet_mon(root, vals)
    
    tree.bind("<Double-1>", xem_chi_tiet)



    # ================= TRANG 2: NHẬP LIỆU =================
    tk.Label(page_nhap_lieu, text="✍️ QUẢN LÝ CÔNG THỨC", bg="#F1F5F9", fg="#0F172A", font=("Segoe UI", 24, "bold")).pack(pady=(40, 10))

    frame_input.pack(padx=150, pady=20, ipadx=20, ipady=20, fill="x")
    frame_input.grid_columnconfigure(0, weight=1, minsize=180)
    frame_input.grid_columnconfigure(1, weight=3)

    def config_entry(text, row, entry_obj):
        tk.Label(frame_input, text=text, bg="white", font=("Segoe UI", 12, "bold"), fg="#334155").grid(row=row, column=0, padx=20, pady=12, sticky="e")
        entry_obj.config(
            font=("Segoe UI", 13), bg="#F8FAFC", fg="#1E293B", relief="flat", bd=1,
            highlightthickness=1, highlightbackground="#CBD5E1", highlightcolor="#3B82F6"
        )
        entry_obj.grid(row=row, column=1, padx=20, pady=12, sticky="we", ipady=8)

    config_entry("🍜 Tên món:", 1, entry_ten)
    config_entry("🥗 Loại món:", 2, entry_loai)
    config_entry("🧂 Nguyên liệu:", 3, entry_nguyenlieu)
    config_entry("⚖ Định lượng:", 4, entry_dinhluong)
    config_entry("⏱ Thời gian:", 5, entry_thoigian)
    config_entry("⭐ Đánh giá:", 6, entry_danhgia)


    frame_btn_nhap = tk.Frame(page_nhap_lieu, bg="#F1F5F9")
    frame_btn_nhap.pack(pady=(10, 40))
    
    btn_add = tk.Button(frame_btn_nhap, text="➕ Thêm Mới", bg="#10B981", fg="white", font=("Segoe UI", 12, "bold"), relief="flat", cursor="hand2", command=them_wrapper)
    btn_add.pack(side="left", padx=20, ipadx=30, ipady=12)
    add_hover(btn_add, "#10B981", "#059669")
    
    btn_save = tk.Button(frame_btn_nhap, text="💾 Lưu Cập Nhật", bg="#3B82F6", fg="white", font=("Segoe UI", 12, "bold"), relief="flat", cursor="hand2", command=sua_wrapper)
    btn_save.pack(side="left", padx=20, ipadx=30, ipady=12)
    add_hover(btn_save, "#3B82F6", "#2563EB")

    # ================= TRANG 3: THỐNG KÊ =================
    tk.Label(page_thong_ke, text="📊 BÁO CÁO THỐNG KÊ", bg="#F1F5F9", fg="#0F172A", font=("Segoe UI", 24, "bold")).pack(pady=(40, 20))
    
    frame_info.pack(padx=200, pady=20, ipadx=20, ipady=20, fill="x")
    tk.Label(frame_info, text="💡 THỐNG KÊ TỔNG QUAN", bg="white", fg="#0F172A", font=("Segoe UI", 18, "bold")).pack(pady=20)
    tk.Frame(frame_info, bg="#E2E8F0", height=2).pack(fill="x", padx=100, pady=10)
    
    lbl_tongmon.pack(pady=10)
    lbl_launhat.pack(pady=10)
    lbl_top.pack(pady=10)
    lbl_nguyenlieu.pack(pady=10)
    
    tk.Label(frame_info, text="🍜  🍔  🍕  🍣", bg="white", font=("Segoe UI Emoji", 36)).pack(pady=30)

    # ================= KHỞI ĐỘNG =================
    them_du_lieu_mau(tree)
    update_quick_stats()
    show_page(page_danh_sach)

    return tree