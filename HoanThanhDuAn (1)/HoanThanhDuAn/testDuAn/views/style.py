from tkinter import ttk

def cai_dat_style():

    style = ttk.Style()

    style.theme_use("default")

    style.configure(
        "Treeview",
        background="white",
        foreground="#333333",
        rowheight=40,
        fieldbackground="white",
        font=("Segoe UI", 11),
        borderwidth=0
    )

    style.configure(
        "Treeview.Heading",
        background="#F3F4F6",
        foreground="#374151",
        font=("Segoe UI", 11, "bold"),
        borderwidth=1,
        relief="flat"
    )

    style.map(
        "Treeview",
        background=[("selected", "#DBEAFE")],
        foreground=[("selected", "#1E3A8A")]
    )