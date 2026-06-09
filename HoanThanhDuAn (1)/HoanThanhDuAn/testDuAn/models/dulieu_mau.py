# ===== FILE: dulieu_mau.py =====

import tkinter as tk


def them_du_lieu_mau(tree):

    tree.insert(
        "",
        tk.END,
        values=(
            "☐",
            "Phở bò",
            "Món nước",
            "Bò, bánh phở, hành",
            "2 người",
            "45",
            "5 sao"
        )
    )

    tree.insert(
        "",
        tk.END,
        values=(
            "☐",
            "Cơm chiên hải sản",
            "Món chiên",
            "Cơm, tôm, mực",
            "3 người",
            "30",
            "4 sao"
        )
    )

    tree.insert(
        "",
        tk.END,
        values=(
            "☐",
            "Trà sữa",
            "Đồ uống",
            "Trà, sữa, trân châu",
            "1 ly",
            "10",
            "5 sao"
        )
    )