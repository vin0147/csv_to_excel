import os
import re
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

    header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # 自動適配欄位寬度（中文字算 2 格、英數字算 1 格）
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
        tk.Label(list_frame, text="此資料夾內沒有找到任何 CSV 檔案", fg="gray", bg="white").pack(anchor="w", pady=5)
        return

    for filename in csv_files:
        var = tk.BooleanVar(value=True)
        chk = tk.Checkbutton(list_frame, text=filename, variable=var, font=("微軟正黑體", 10), bg="white")
        chk.pack(anchor="w", padx=10, pady=2)
        file_vars.append((filename, var))


def toggle_all(state):
    for _, var in file_vars:
        var.set(state)


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

    # 模式一：合併成同一個 Excel（多工作表）
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

    # 模式二：各自轉成獨立 Excel
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
# 3. 建立視窗畫面
# ==========================================
root = tk.Tk()
root.title("CSV 批次轉 Excel 工具")
root.geometry("580x540")

folder_var = tk.StringVar()
mode_var = tk.StringVar(value="merge")  # 預設為合併模式
file_vars = []

# --- 上方：目錄選擇區 ---
top_frame = tk.Frame(root, pady=10, padx=15)
top_frame.pack(fill="x")

tk.Button(top_frame, text="選擇 CSV 目錄", command=select_folder, font=("微軟正黑體", 10, "bold"), bg="#E1E8F0").pack(side="left")
tk.Entry(top_frame, textvariable=folder_var, state="readonly", font=("微軟正黑體", 10)).pack(side="left", fill="x", expand=True, padx=10)

# --- 輸出模式切換區 ---
mode_frame = tk.Frame(root, padx=15, pady=5)
mode_frame.pack(fill="x")

tk.Label(mode_frame, text="輸出模式：", font=("微軟正黑體", 10, "bold")).pack(side="left")
tk.Radiobutton(mode_frame, text="合併為單一 Excel（多分頁）", variable=mode_var, value="merge", font=("微軟正黑體", 10)).pack(side="left", padx=5)
tk.Radiobutton(mode_frame, text="各自轉為獨立 Excel", variable=mode_var, value="single", font=("微軟正黑體", 10)).pack(side="left", padx=5)

# --- 全選 / 取消全選按鈕 ---
btn_frame = tk.Frame(root, padx=15, pady=5)
btn_frame.pack(fill="x")

tk.Button(btn_frame, text="全選", command=lambda: toggle_all(True), width=8).pack(side="left", padx=(0, 5))
tk.Button(btn_frame, text="取消全選", command=lambda: toggle_all(False), width=8).pack(side="left")

# --- 可捲動的檔案勾選清單區 ---
container = tk.Frame(root, padx=15, pady=5)
container.pack(fill="both", expand=True)

canvas = tk.Canvas(container, borderwidth=1, relief="sunken", bg="white")
scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
list_frame = tk.Frame(canvas, bg="white")

list_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
canvas.create_window((0, 0), window=list_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# --- 下方：一鍵轉換按鈕 ---
bottom_frame = tk.Frame(root, pady=15, padx=15)
bottom_frame.pack(fill="x")

tk.Button(
    bottom_frame,
    text="一鍵轉換為 Excel（輸出至 output 資料夾）",
    command=start_conversion,
    font=("微軟正黑體", 12, "bold"),
    bg="#2E7D32",
    fg="white",
    height=2
).pack(fill="x")

root.mainloop()