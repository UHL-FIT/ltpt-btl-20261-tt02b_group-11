# ===== FILE: form.py =====

# Import thư viện tkinter
import tkinter as tk

# Import ttk để dùng Treeview
from tkinter import ttk

# Import hàm chọn món từ file chucnang.py
from chucnang import chon_mon

# ================= TẠO FORM =================
def tao_form(
    root,
    them_mon,
    sua_mon,
    xoa_mon,
    thong_ke
):

    # Biến kiểm tra trạng thái đang thêm món
    dang_them = False

    # ================= HÀM ĐỔI TRẠNG THÁI THÊM =================
    def set_dang_them():

        # Cho phép sửa biến bên ngoài hàm
        nonlocal dang_them

        # Đánh dấu đang thêm món
        dang_them = True

    # ================= STYLE TREEVIEW =================

    # Tạo style cho bảng
    style = ttk.Style()

    # Chọn giao diện mặc định
    style.theme_use("default")

    # Style cho bảng dữ liệu
    style.configure(
        "Treeview",

        # Màu nền bảng
        background="white",

        # Màu chữ
        foreground="black",

        # Chiều cao mỗi dòng
        rowheight=30,

        # Màu nền dữ liệu
        fieldbackground="white",

        # Font chữ
        font=("Arial", 10)
    )

    # Style cho tiêu đề cột
    style.configure(
        "Treeview.Heading",

        # Font tiêu đề
        font=("Arial", 11, "bold")
    )

    # ================= TIÊU ĐỀ =================

    title = tk.Label(

        # Hiển thị trên cửa sổ root
        root,

        # Nội dung tiêu đề
        text="QUẢN LÝ CÔNG THỨC NẤU ĂN",

        # Font chữ
        font=("Arial", 24, "bold"),

        # Màu nền
        bg="#F5F5F5",

        # Màu chữ
        fg="#2E7D32"
    )

    # Hiển thị tiêu đề
    title.pack(pady=20)

    # ================= FRAME NHẬP DỮ LIỆU =================

    frame_input = tk.Frame(

        # Hiển thị trên root
        root,

        # Màu nền
        bg="white",

        # Độ dày viền
        bd=2,

        # Kiểu viền
        relief="groove"
    )

    # Hiển thị frame
    frame_input.pack(
        pady=10
    )

    # ================= FONT LABEL =================

    font_label = ("Arial", 11)

    # ================= TÊN MÓN =================

    tk.Label(
        frame_input,

        # Nội dung label
        text="Tên món",

        # Màu nền
        bg="white",

        # Font
        font=font_label

    ).grid(
        row=0,
        column=0,
        padx=5,
        pady=10
    )

    # Ô nhập tên món
    entry_ten = tk.Entry(
        frame_input,

        # Độ rộng
        width=18,

        # Font
        font=("Arial", 11)
    )

    entry_ten.grid(
        row=0,
        column=1,
        padx=5
    )

    # ================= LOẠI MÓN =================

    tk.Label(
        frame_input,
        text="Loại món",
        bg="white",
        font=font_label

    ).grid(
        row=1,
        column=0,
        padx=5,
        pady=10
    )

    entry_loai = tk.Entry(
        frame_input,
        width=18,
        font=("Arial", 11)
    )

    entry_loai.grid(
        row=1,
        column=1,
        padx=5
    )

    # ================= NGUYÊN LIỆU =================

    tk.Label(
        frame_input,
        text="Nguyên liệu",
        bg="white",
        font=font_label

    ).grid(
        row=0,
        column=2,
        padx=5,
        pady=10
    )

    entry_nguyenlieu = tk.Entry(
        frame_input,
        width=18,
        font=("Arial", 11)
    )

    entry_nguyenlieu.grid(
        row=0,
        column=3,
        padx=5
    )

    # ================= THỜI GIAN =================

    tk.Label(
        frame_input,
        text="Thời gian",
        bg="white",
        font=font_label

    ).grid(
        row=1,
        column=2,
        padx=5,
        pady=10
    )

    entry_thoigian = tk.Entry(
        frame_input,
        width=18,
        font=("Arial", 11)
    )

    entry_thoigian.grid(
        row=1,
        column=3,
        padx=5
    )

    # ================= FRAME NÚT =================

    frame_btn = tk.Frame(
        root,
        bg="#F5F5F5"
    )

    frame_btn.pack(pady=15)

    # ================= NÚT THÊM =================

    btn_them = tk.Button(
        frame_btn,

        # Nội dung nút
        text="Thêm",

        # Độ rộng
        width=12,

        # Màu nền
        bg="#4CAF50",

        # Màu chữ
        fg="black",

        # Font chữ
        font=("Arial", 10, "bold"),

        # Kiểu viền
        relief="flat",

        # Chức năng nút
        command=lambda:

            [
                # Đánh dấu đang thêm
                set_dang_them(),

                # Gọi hàm thêm món
                them_mon(
                    tree,
                    entry_ten,
                    entry_loai,
                    entry_nguyenlieu,
                    entry_thoigian
                )
            ]
    )

    btn_them.grid(
        row=0,
        column=0,
        padx=8
    )

    # ================= NÚT SỬA =================

    btn_sua = tk.Button(
        frame_btn,
        text="Sửa",
        width=12,
        bg="#FF9800",
        fg="black",
        font=("Arial", 10, "bold"),
        relief="flat",

        command=lambda: sua_mon(
            tree,
            entry_ten,
            entry_loai,
            entry_nguyenlieu,
            entry_thoigian
        )
    )

    btn_sua.grid(
        row=0,
        column=1,
        padx=8
    )

    # ================= NÚT XÓA =================

    btn_xoa = tk.Button(
        frame_btn,
        text="Xóa",
        width=12,
        bg="#C65850",
        fg="black",
        font=("Arial", 10, "bold"),
        relief="flat",

        command=lambda: xoa_mon(tree)
    )

    btn_xoa.grid(
        row=0,
        column=2,
        padx=8
    )

    # ================= NÚT THỐNG KÊ =================

    btn_thongke = tk.Button(
        frame_btn,
        text="Thống kê",
        width=12,
        bg="#2196F3",
        fg="black",
        font=("Arial", 10, "bold"),
        relief="flat",

        command=lambda: thong_ke(tree)
    )

    btn_thongke.grid(
        row=0,
        column=3,
        padx=8
    )

    # ================= FRAME BẢNG =================

    frame_table = tk.Frame(
        root,
        bg="white",
        bd=2,
        relief="groove"
    )

    frame_table.pack(
        padx=15,
        pady=15,
        fill="both",
        expand=True
    )

    # ================= TÊN CỘT =================

    columns = (
        "Tên món",
        "Loại món",
        "Nguyên liệu",
        "Thời gian"
    )

    # ================= TẠO BẢNG =================

    tree = ttk.Treeview(
        frame_table,

        # Tên cột
        columns=columns,

        # Chỉ hiện tiêu đề
        show="headings",

        # Số dòng hiển thị
        height=10
    )

    # ================= TẠO CỘT =================

    for col in columns:

        # Tạo tiêu đề cột
        tree.heading(col, text=col)

        # Chỉnh kích thước cột
        tree.column(
            col,
            width=180,
            anchor="center"
        )

    # Hiển thị bảng
    tree.pack(fill="both", expand=True)

    # ================= WRAPPER CHỌN MÓN =================

    def chon_mon_wrapper(event):

        # Cho phép sửa biến bên ngoài
        nonlocal dang_them

        # Nếu vừa thêm món
        if dang_them:

            # Reset trạng thái
            dang_them = False

            return

        # Gọi hàm chọn món
        chon_mon(
            event,
            tree,
            entry_ten,
            entry_loai,
            entry_nguyenlieu,
            entry_thoigian
        )

    # ================= BẮT SỰ KIỆN CLICK =================

    tree.bind(
        "<<TreeviewSelect>>",
        chon_mon_wrapper
    )

    # Trả về bảng
    return tree