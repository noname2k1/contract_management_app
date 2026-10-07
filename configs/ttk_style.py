from tkinter import ttk


def style_config():
    style = ttk.Style()
    try:
        style.theme_use("vista")
    except Exception:  # noqa: BLE001, S110
        pass
    # ------------------------------
    # NỀN CHUNG
    # ------------------------------
    style.configure(
        "TFrame",
        background="#f2f2f2",
    )
    style.configure(
        "TLabelframe",
        background="#f2f2f2",
    )
    style.configure(
        "TLabelframe.Label",
        background="#f2f2f2",
        font=("Arial", 10),
    )
    # ------------------------------
    # LABEL
    # ------------------------------
    style.configure(
        "TLabel",
        background="#f2f2f2",
        font=("Arial", 10),
    )
    style.configure(
        "AppTitle.TLabel",
        background="#f2f2f2",
        font=("Arial", 22, "bold"),
    )
    style.configure(
        "PageTitle.TLabel",
        background="#f2f2f2",
        font=("Arial", 18, "bold"),
    )
    style.configure(
        "Section.TLabel",
        background="#f2f2f2",
        font=("Arial", 11, "bold"),
    )
    style.configure(
        "Header.TLabel",
        background="#f2f2f2",
        font=("Arial", 10, "bold"),
    )
    # ------------------------------
    # BUTTON
    # ------------------------------
    style.configure(
        "TButton",
        font=("Arial", 10),
        padding=(8, 4),
    )
    style.configure(
        "Menu.TButton",
        font=("Arial", 10),
        padding=(10, 8),
    )
    style.configure(
        "Action.TButton",
        font=("Arial", 10, "bold"),
        padding=(10, 6),
    )
    # ------------------------------
    # RADIOBUTTON
    # ------------------------------
    style.configure(
        "TRadiobutton",
        background="#f2f2f2",
        font=("Arial", 10),
    )
    # ------------------------------
    # ENTRY
    # ------------------------------
    style.configure(
        "TEntry",
        padding=(4, 3),
    )
    # ------------------------------
    # TREEVIEW
    # ------------------------------
    style.configure(
        "Treeview",
        font=("Arial", 9),
        rowheight=24,
    )
    style.configure(
        "Treeview.Heading",
        font=("Arial", 9, "bold"),
        padding=(4, 4),
    )
