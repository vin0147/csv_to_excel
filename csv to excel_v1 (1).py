import os
import re
import sys
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment


# ==========================================
# 1. 核心功能：讀取 CSV、套用樣式與自動欄寬
# ==========================================
def read_csv_auto_encoding(csv_path):
    """依序嘗試常見編碼讀取 CSV，並全部以字串讀入保留開頭的 0"""
    encodings = ["utf-8-sig", "utf-8", "cp950", "big5"]
    for enc in encodings:
        try:
            return pd.read_csv(csv_path, encoding=enc, dtype=str)
        except Exception:
            continue
    raise ValueError(f"無法辨識檔案編碼：{os.path.basename(csv_path)}")


def format_worksheet(ws):
    """套用表頭樣式、凍結首列、篩選器，並自動適配欄位寬度"""
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    header_fill = PatternFill(start_color="1D4ED8", end_color="1D4ED8", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for col_idx, col_cells in enumerate(ws.columns, start=1):
        max_width = 0
        for cell in col_cells:
            val = str(cell.value) if cell.value is not None else ""
            width = sum(2 if ord(ch) > 127 else 1 for ch in val)
            if width > max_width:
                max_width = width

        final_width = max(10, min(max_width + 4, 60))
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = final_width


def make_safe_sheet_name(raw_name, used_names):
    """清理 Excel 工作表不允許的字元，限制 31 字元內並避免重複"""
    clean_name = re.sub(r'[\\/*?:\[\]]', "_", raw_name).strip()
    if not clean_name:
        clean_name = "Sheet"
    base_name = clean_name[:31]

    final_name = base_name
    counter = 1
    while final_name in used_names:
        suffix = f"_{counter}"
        final_name = base_name[: 31 - len(suffix)] + suffix
        counter += 1

    used_names.add(final_name)
    return final_name


def convert_single_csv(csv_path, excel_path):
    """獨立模式：單一 CSV 轉成單一 Excel"""
    df = read_csv_auto_encoding(csv_path)
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="資料明細")
        format_worksheet(writer.sheets["資料明細"])


def convert_merged_csvs(csv_paths, excel_path):
    """合併模式：多個 CSV 合併成同一個 Excel 的不同工作表"""
    used_sheet_names = set()
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        for path in csv_paths:
            df = read_csv_auto_encoding(path)
            raw_name = os.path.splitext(os.path.basename(path))[0]
            sheet_name = make_safe_sheet_name(raw_name, used_sheet_names)

            df.to_excel(writer, index=False, sheet_name=sheet_name)
            format_worksheet(writer.sheets[sheet_name])


# ==========================================
# 2. 介面互動功能
# ==========================================
def get_output_dir():
    """取得並自動建立執行檔所在目錄下的 output 資料夾"""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(base_dir, "output")
    os.makedirs(output_dir, exist_ok=True)
    return output_dir


def open_output_folder():
    """開啟 output 資料夾"""
    output_dir = get_output_dir()
    if sys.platform == "win32":
        os.startfile(output_dir)
    elif sys.platform == "darwin":
        subprocess.run(["open", output_dir])
    else:
        subprocess.run(["xdg-open", output_dir])


def update_selected_count():
    """更新目前勾選檔案數量提示"""
    total = len(file_vars)
    checked = sum(1 for _, var in file_vars if var.get())
    count_label.config(text=f"已勾選 {checked} / {total} 個 CSV")


def select_folder():
    folder = filedialog.askdirectory(title="請選擇包含 CSV 的資料夾")
    if not folder:
        return

    folder_var.set(folder)
    refresh_file_list(folder)


def refresh_file_list(folder):
    for widget in list_frame.winfo_children():
        widget.destroy()
    file_vars.clear()

    csv_files = sorted([f for f in os.listdir(folder) if f.lower().endswith(".csv")])

    if not csv_files:
        tk.Label(
            list_frame,
            text="此資料夾內沒有找到任何 .csv 檔案",
            fg="#64748B",
            bg="#FFFFFF",
            font=("微軟正黑體", 10)
        ).pack(anchor="w", padx=12, pady=10)
        update_selected_count()
        return

    for idx, filename in enumerate(csv_files):
        row_bg = "#FFFFFF" if idx % 2 == 0 else "#F8FAFC"
        var = tk.BooleanVar(value=True)
        chk = tk.Checkbutton(
            list_frame,
            text=f"📄 {filename}",
            variable=var,
            command=update_selected_count,
            font=("微軟正黑體", 10),
            bg=row_bg,
            activebackground="#EEF2FF",
            fg="#1E293B",
            selectcolor="#DBEAFE",
            anchor="w",
            padx=8,
            pady=3
        )
        chk.pack(fill="x", padx=4, pady=1)
        file_vars.append((filename, var))

    update_selected_count()


def toggle_all(state):
    for _, var in file_vars:
        var.set(state)
    update_selected_count()


def start_conversion():
    folder = folder_var.get()
    if not folder:
        messagebox.showwarning("提醒", "請先選擇 CSV 所在的目錄！")
        return

    selected_files = [name for name, var in file_vars if var.get()]
    if not selected_files:
        messagebox.showwarning("提醒", "請至少勾選一個 CSV 檔案！")
        return

    output_dir = get_output_dir()
    mode = mode_var.get()

    if mode == "merge":
        csv_paths = [os.path.join(folder, name) for name in selected_files]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_path = os.path.join(output_dir, f"合併輸出_{timestamp}.xlsx")

        try:
            convert_merged_csvs(csv_paths, excel_path)
            messagebox.showinfo(
                "合併轉換完成",
                f"已成功將 {len(selected_files)} 個檔案合併至同一個 Excel！\n\n輸出檔案：\n{excel_path}"
            )
        except Exception as e:
            messagebox.showerror("轉換失敗", f"合併過程發生錯誤：\n{e}")

    else:
        success_count = 0
        error_messages = []

        for filename in selected_files:
            csv_path = os.path.join(folder, filename)
            excel_name = os.path.splitext(filename)[0] + ".xlsx"
            excel_path = os.path.join(output_dir, excel_name)

            try:
                convert_single_csv(csv_path, excel_path)
                success_count += 1
            except Exception as e:
                error_messages.append(f"{filename}：{e}")

        if error_messages:
            err_text = "\n".join(error_messages)
            messagebox.showwarning(
                "完成（部分失敗）",
                f"成功轉換 {success_count} 個檔案至 output 資料夾。\n\n失敗清單：\n{err_text}"
            )
        else:
            messagebox.showinfo(
                "轉換完成",
                f"已成功將 {success_count} 個檔案轉換為獨立 Excel！\n\n輸出資料夾：\n{output_dir}"
            )


# ==========================================
# 3. 建立視窗畫面（色彩分區設計）
# ==========================================
BG_MAIN = "#F4F7FB"
BG_CARD = "#FFFFFF"
BORDER_COLOR = "#D6E0EC"

root = tk.Tk()
root.title("CSV 批次轉 Excel 工具")
root.geometry("620x580")
root.configure(bg=BG_MAIN)

folder_var = tk.StringVar()
mode_var = tk.StringVar(value="merge")  # 預設為合併模式
file_vars = []

# --- 頂部標題橫幅 ---
banner = tk.Frame(root, bg="#1E3A8A", padx=16, pady=12)
banner.pack(fill="x")
tk.Label(
    banner,
    text="📊 CSV 批次轉 Excel 小工具",
    font=("微軟正黑體", 14, "bold"),
    bg="#1E3A8A",
    fg="#FFFFFF"
).pack(side="left")
tk.Label(
    banner,
    text="自動編碼辨識・保留前導零・自動欄寬",
    font=("微軟正黑體", 9),
    bg="#1E3A8A",
    fg="#BFDBFE"
).pack(side="right", anchor="s")

# --- 區塊 1：目錄選擇卡片 ---
top_card = tk.Frame(root, bg=BG_CARD, highlightbackground=BORDER_COLOR, highlightthickness=1, padx=12, pady=10)
top_card.pack(fill="x", padx=14, pady=(12, 6))

tk.Button(
    top_card,
    text="📁 選擇 CSV 目錄",
    command=select_folder,
    font=("微軟正黑體", 10, "bold"),
    bg="#2563EB",
    fg="#FFFFFF",
    activebackground="#1D4ED8",
    activeforeground="#FFFFFF",
    relief="flat",
    padx=12,
    pady=4,
    cursor="hand2"
).pack(side="left")

tk.Entry(
    top_card,
    textvariable=folder_var,
    state="readonly",
    font=("微軟正黑體", 10),
    readonlybackground="#F8FAFC",
    fg="#334155",
    relief="solid",
    bd=1
).pack(side="left", fill="x", expand=True, padx=(10, 0), ipady=3)

# --- 區塊 2：輸出模式與批次控制列 ---
ctrl_card = tk.Frame(root, bg="#EEF2FF", highlightbackground="#C7D2FE", highlightthickness=1, padx=12, pady=8)
ctrl_card.pack(fill="x", padx=14, pady=6)

mode_row = tk.Frame(ctrl_card, bg="#EEF2FF")
mode_row.pack(fill="x")

tk.Label(
    mode_row,
    text="輸出模式：",
    font=("微軟正黑體", 10, "bold"),
    bg="#EEF2FF",
    fg="#3730A3"
).pack(side="left")

tk.Radiobutton(
    mode_row,
    text="合併為單一 Excel（多分頁）",
    variable=mode_var,
    value="merge",
    font=("微軟正黑體", 10, "bold"),
    bg="#EEF2FF",
    activebackground="#EEF2FF",
    fg="#1E293B",
    selectcolor="#FFFFFF"
).pack(side="left", padx=(4, 12))

tk.Radiobutton(
    mode_row,
    text="各自轉為獨立 Excel",
    variable=mode_var,
    value="single",
    font=("微軟正黑體", 10),
    bg="#EEF2FF",
    activebackground="#EEF2FF",
    fg="#334155",
    selectcolor="#FFFFFF"
).pack(side="left")

action_row = tk.Frame(ctrl_card, bg="#EEF2FF")
action_row.pack(fill="x", pady=(6, 0))

tk.Button(
    action_row,
    text="全選",
    command=lambda: toggle_all(True),
    font=("微軟正黑體", 9, "bold"),
    bg="#4F46E5",
    fg="#FFFFFF",
    activebackground="#4338CA",
    activeforeground="#FFFFFF",
    relief="flat",
    width=7,
    cursor="hand2"
).pack(side="left", padx=(0, 6))

tk.Button(
    action_row,
    text="取消全選",
    command=lambda: toggle_all(False),
    font=("微軟正黑體", 9),
    bg="#E0E7FF",
    fg="#3730A3",
    activebackground="#C7D2FE",
    relief="flat",
    width=8,
    cursor="hand2"
).pack(side="left")

count_label = tk.Label(
    action_row,
    text="已勾選 0 / 0 個 CSV",
    font=("微軟正黑體", 9, "bold"),
    bg="#EEF2FF",
    fg="#4338CA"
)
count_label.pack(side="right")

# --- 區塊 3：可捲動的檔案勾選清單區 ---
list_card = tk.Frame(root, bg=BG_CARD, highlightbackground=BORDER_COLOR, highlightthickness=1, padx=6, pady=6)
list_card.pack(fill="both", expand=True, padx=14, pady=6)

canvas = tk.Canvas(list_card, borderwidth=0, highlightthickness=0, bg=BG_CARD)
scrollbar = tk.Scrollbar(list_card, orient="vertical", command=canvas.yview)
list_frame = tk.Frame(canvas, bg=BG_CARD)

list_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
canvas_window = canvas.create_window((0, 0), window=list_frame, anchor="nw")
canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas_window, width=e.width))
canvas.configure(yscrollcommand=scrollbar.set)

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# --- 區塊 4：底部操作按鈕區 ---
bottom_frame = tk.Frame(root, bg=BG_MAIN, padx=14, pady=8)
bottom_frame.pack(fill="x", pady=(0, 8))

tk.Button(
    bottom_frame,
    text="🚀 一鍵轉換為 Excel（存至 output）",
    command=start_conversion,
    font=("微軟正黑體", 11, "bold"),
    bg="#16A34A",
    fg="#FFFFFF",
    activebackground="#15803D",
    activeforeground="#FFFFFF",
    relief="flat",
    pady=8,
    cursor="hand2"
).pack(side="left", fill="x", expand=True, padx=(0, 8))

tk.Button(
    bottom_frame,
    text="📂 開啟 output",
    command=open_output_folder,
    font=("微軟正黑體", 10, "bold"),
    bg="#0284C7",
    fg="#FFFFFF",
    activebackground="#0369A1",
    activeforeground="#FFFFFF",
    relief="flat",
    padx=12,
    pady=8,
    cursor="hand2"
).pack(side="right")

root.mainloop()