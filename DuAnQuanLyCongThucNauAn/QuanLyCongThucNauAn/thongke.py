# Import thư viện pandas
# Dùng để xử lý và thống kê dữ liệu
import pandas as pd

# Import hộp thoại thông báo
from tkinter import messagebox

# ================= HÀM THỐNG KÊ =================
def thong_ke(tree):

    # Tạo danh sách rỗng để lưu dữ liệu
    data = []

    # ================= LẤY DỮ LIỆU TỪ BẢNG =================

    # Duyệt từng dòng trong Treeview
    for item in tree.get_children():

        # Lấy dữ liệu của dòng
        values = tree.item(item, "values")

        # Thêm dữ liệu vào danh sách
        data.append(values)

    # ================= KIỂM TRA DỮ LIỆU =================

    # Nếu bảng chưa có dữ liệu
    if len(data) == 0:

        # Hiện thông báo
        messagebox.showinfo(
            "Thống kê",
            "Chưa có dữ liệu"
        )

        return

    # ================= TẠO DATAFRAME =================

    # Chuyển dữ liệu sang DataFrame
    df = pd.DataFrame(

        # Dữ liệu
        data,

        # Tên cột
        columns=[
            "Tên món",
            "Loại món",
            "Nguyên liệu",
            "Thời gian"
        ]
    )

    # ================= THỐNG KÊ =================

    # Đếm tổng số món ăn
    tong_mon = len(df)

    # Đếm số lượng từng loại món
    thongke_loai = df["Loại món"].value_counts()

    # ================= TẠO CHUỖI KẾT QUẢ =================

    # Tạo chuỗi ban đầu
    ketqua = f"Tổng số món ăn: {tong_mon}\n\n"

    # Thêm tiêu đề thống kê loại món
    ketqua += "Thống kê loại món:\n"

    # Duyệt từng loại món
    for loai, soluong in thongke_loai.items():

        # Thêm vào chuỗi kết quả
        ketqua += f"- {loai}: {soluong}\n"

    # ================= HIỂN THỊ KẾT QUẢ =================

    # Hiện hộp thoại thống kê
    messagebox.showinfo(
        "Thống kê",
        ketqua
    )