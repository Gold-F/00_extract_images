from tkinter import messagebox
import csv
import os
import sys

# pyinstallerでバンドルせず、後から配置
def collect_pdf_files(folderRef):
    pdf_files = [f for f in os.listdir(folderRef) if f.endswith('.pdf')]
    if len(pdf_files) < 1:
        messagebox.showerror("終了",f"'{folderRef}'  にPDFファイルを入れてください。")
        sys.exit()
    return pdf_files

# pyinstallerでバンドルせず、後から配置
def load_size_file(file_path):
    if not os.path.exists(file_path):
        messagebox.showerror("終了",f"'{file_path}' が見つかりません。")
        sys.exit()
    try:
        with open(file_path, newline="", encoding="utf-8-sig") as f: # BOM付き無し両方のUTF-8に対応
            reader = csv.reader(f)
            next(reader)  # Skip header row
            return [(int(a), int(b)) for a, b in reader]
    except UnicodeDecodeError:
        messagebox.showerror(
            "文字コードエラー",
            "CSVファイルの文字コードが UTF-8 ではありません。\n"
        )
        sys.exit()
        
def load_trim_file(file_path):
    if not os.path.exists(file_path):
        messagebox.showerror("終了",f"'{file_path}' が見つかりません。")
        sys.exit()
    try:
        with open(file_path, newline="", encoding="utf-8-sig") as f: # BOM付き無し両方のUTF-8に対応
            reader = csv.reader(f)
            next(reader)  # Skip header row
            row = next(reader)  # 1行だけ読む
            return tuple(map(int, row))
    except UnicodeDecodeError:
        messagebox.showerror(
            "文字コードエラー",
            "CSVファイルの文字コードが UTF-8 ではありません。\n"
        )
        sys.exit()
        
# pyinstallerで「--add-data で同梱したリソース（font, image, csv）」の実体場所を取る
def resource_path(relative_path):
    if hasattr(sys, "_MEIPASS"): #.exe化されたときの一時フォルダパス取得
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__)) # .py実行時のフォルダパス取得

    full_path = os.path.join(base_path, relative_path)

    if not os.path.exists(full_path):
        messagebox.showerror("エラー", f"フォントファイルが見つかりません: {full_path}")
        sys.exit()

    return full_path

