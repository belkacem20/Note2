import math
import os
import sys
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, simpledialog

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# ملف الإكسل يُحفظ بجانب البرنامج
if getattr(sys, "frozen", False):
    BASE = os.path.dirname(sys.executable)  # عند التشغيل كملف exe
else:
    BASE = os.path.dirname(os.path.abspath(__file__))
FILE = os.path.join(BASE, "ملاحظاتي.xlsx")

# الفئات الجاهزة (يمكن إضافة فئات جديدة من داخل البرنامج بزر «فئة جديدة»)
DEFAULT_CATEGORIES = ["Croissance 1", "Croissance 2", "VL", "Démarrage"]
HEADERS = ["التاريخ", "الوقت", "الفئة", "الملاحظة"]
COLORS = ["#1565c0", "#2e7d32", "#ef6c00", "#6a1b9a", "#00838f", "#ad1457"]
categories = list(DEFAULT_CATEGORIES)


def style_header(ws):
    fill = PatternFill("solid", fgColor="1F4E78")
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF", size=13)
        c.fill = fill
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26
    ws.freeze_panes = "A2"


def style_row(ws, r):
    thin = Side(style="thin", color="BBBBBB")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for i, cell in enumerate(ws[r]):
        cell.border = border
        cell.font = Font(size=12)
        if i < 3:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        else:
            # الملاحظة تبقى أسطراً كما كتبتها، والاتجاه يتبع لغة النص
            cell.alignment = Alignment(horizontal="general", vertical="top", wrap_text=True)
            cell.data_type = "s"  # يبقى النص نصاً حتى لو بدأ بـ = أو كان رقماً
    text = str(ws.cell(r, 4).value or "")
    lines = sum(max(1, math.ceil(len(line) / 75)) for line in text.split("\n"))
    ws.row_dimensions[r].height = max(22, 18 * lines + 4)


def load_book():
    if os.path.exists(FILE):
        return load_workbook(FILE)
    wb = Workbook()
    ws = wb.active
    ws.title = "الملاحظات"
    ws.sheet_view.rightToLeft = True
    ws.append(HEADERS)
    for col, width in zip("ABCD", (14, 10, 14, 80)):
        ws.column_dimensions[col].width = width
    style_header(ws)
    return wb


def read_rows():
    """كل الملاحظات من الملف: (التاريخ، الوقت، الفئة، النص)"""
    if not os.path.exists(FILE):
        return []
    ws = load_workbook(FILE).active
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        d, t, c, n = (list(r) + [None] * 4)[:4]
        if n is None:
            continue
        rows.append((str(d or ""), str(t or ""), str(c or ""), str(n)))
    return rows


def save_note():
    text = box.get("1.0", tk.END).strip()
    if not text:
        messagebox.showwarning("تنبيه", "اكتب ملاحظة أولاً")
        return
    cat = cat_var.get()
    now = datetime.now()
    try:
        wb = load_book()
        ws = wb.active
        ws.append([now.strftime("%d/%m/%Y"), now.strftime("%H:%M:%S"), cat, text])
        style_row(ws, ws.max_row)
        ws.auto_filter.ref = "A1:D%d" % ws.max_row
        wb.save(FILE)
    except PermissionError:
        messagebox.showerror("خطأ", "أغلق ملف الإكسل أولاً ثم أعد المحاولة")
        return
    box.delete("1.0", tk.END)
    refresh_list()
    rebuild_category_widgets()
    status.config(text="تم الحفظ ✓  %s  —  %s" % (cat, now.strftime("%H:%M:%S")))


def refresh_list():
    listbox.delete(0, tk.END)
    try:
        rows = read_rows()
    except PermissionError:
        return
    for d, t, c, n in reversed(rows[-20:]):
        short = n.replace("\n", " | ")[:50]
        listbox.insert(tk.END, "%s  %s  [%s]  —  %s" % (d, t, c, short))


def show_notes(cat):
    """عرض الملاحظات بالترتيب الذي كُتبت به، مع تاريخ كل ملاحظة."""
    try:
        rows = read_rows()
    except PermissionError:
        messagebox.showerror("خطأ", "أغلق ملف الإكسل أولاً ثم أعد المحاولة")
        return
    if cat:
        rows = [r for r in rows if r[2] == cat]
    title = cat if cat else "كل الملاحظات"
    if not rows:
        messagebox.showinfo("ملاحظاتي", "لا توجد ملاحظات في: " + title)
        return

    win = tk.Toplevel(root)
    win.title("%s  (%d)" % (title, len(rows)))
    win.geometry("640x520")
    frame = tk.Frame(win)
    frame.pack(fill="both", expand=True, padx=10, pady=10)
    scroll = tk.Scrollbar(frame)
    scroll.pack(side="left", fill="y")
    txt = tk.Text(frame, font=("Arial", 12), wrap="word", yscrollcommand=scroll.set)
    txt.pack(side="right", fill="both", expand=True)
    scroll.config(command=txt.yview)
    txt.tag_configure("head", font=("Arial", 12, "bold"), foreground="#1F4E78",
                      background="#E3EEF9", spacing1=10, spacing3=4)

    for d, t, c, n in reversed(rows):  # الأحدث أولاً
        head = "%s   %s" % (d, t)
        if not cat:
            head += "   —   " + c
        txt.insert(tk.END, head + "\n", "head")
        txt.insert(tk.END, n + "\n\n")
    txt.config(state="disabled")


def add_category():
    name = simpledialog.askstring("فئة جديدة", "اكتب اسم الفئة:", parent=root)
    if not name or not name.strip():
        return
    name = name.strip()
    if name not in categories:
        categories.append(name)
    cat_var.set(name)
    rebuild_category_widgets()


def rebuild_category_widgets():
    # الفئات الموجودة في الملف تُضاف تلقائياً
    try:
        for _, _, c, _ in read_rows():
            if c and c not in categories:
                categories.append(c)
    except PermissionError:
        pass

    for w in pick_frame.winfo_children():
        w.destroy()
    for w in view_frame.winfo_children():
        w.destroy()

    per_row = 4
    for i, c in enumerate(categories):
        row, col = divmod(i, per_row)
        col = per_row - 1 - col  # الأولى على اليمين
        tk.Radiobutton(pick_frame, text=c, variable=cat_var, value=c,
                       font=("Arial", 11)).grid(row=row, column=col, padx=4, sticky="e")
        tk.Button(view_frame, text=c, font=("Arial", 11, "bold"),
                  bg=COLORS[i % len(COLORS)], fg="white", width=11,
                  command=lambda c=c: show_notes(c)).grid(row=row, column=col, padx=3, pady=3)

    row, col = divmod(len(categories), per_row)
    tk.Button(view_frame, text="عرض الكل", font=("Arial", 11, "bold"), bg="#444444",
              fg="white", width=11,
              command=lambda: show_notes(None)).grid(row=row, column=per_row - 1 - col,
                                                     padx=3, pady=3)


root = tk.Tk()
root.title("ملاحظاتي اليومية")
root.geometry("580x720")
cat_var = tk.StringVar(value=DEFAULT_CATEGORIES[0])

tk.Label(root, text="اكتب ملاحظتك: oussama", font=("Arial", 13)).pack(anchor="e", padx=12, pady=(12, 4))
box = tk.Text(root, height=9, font=("Arial", 13), wrap="word")
box.pack(fill="x", padx=12)

tk.Label(root, text="الفئة:", font=("Arial", 12)).pack(anchor="e", padx=12, pady=(10, 0))
pick_frame = tk.Frame(root)
pick_frame.pack(anchor="e", padx=12)

buttons = tk.Frame(root)
buttons.pack(pady=8)
tk.Button(buttons, text="حفظ الملاحظة", font=("Arial", 13, "bold"), bg="#2e7d32", fg="white",
          command=save_note).pack(side="right", padx=6)
tk.Button(buttons, text="+ فئة جديدة", font=("Arial", 12), bg="#757575", fg="white",
          command=add_category).pack(side="right", padx=6)
status = tk.Label(root, text="", fg="#2e7d32")
status.pack()

tk.Label(root, text="عرض الملاحظات حسب الفئة:", font=("Arial", 12)).pack(anchor="e", padx=12, pady=(8, 2))
view_frame = tk.Frame(root)
view_frame.pack(anchor="e", padx=12)

tk.Label(root, text="آخر الملاحظات:", font=("Arial", 12)).pack(anchor="e", padx=12, pady=(10, 2))
listbox = tk.Listbox(root, font=("Arial", 11), height=6)
listbox.pack(fill="both", expand=True, padx=12, pady=(0, 12))

rebuild_category_widgets()
refresh_list()
root.mainloop()
