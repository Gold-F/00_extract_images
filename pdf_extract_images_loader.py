from tkinter import messagebox
import csv
import os  # フォルダ操作に使用
import sys  # プログラム終了に使用


def collect_pdf_files(folderRef):
    pdf_files = [f for f in os.listdir(folderRef) if f.endswith('.pdf')]
    if len(pdf_files) < 1:
        messagebox.showerror("終了",f"'{folderRef}'  にPDFファイルを入れてください。")
        sys.exit()
    return pdf_files

def load_csv_files(file_path):
    if not os.path.exists(file_path):
        messagebox.showerror("終了",f"'{file_path}' が見つかりません。")
        sys.exit()
    with open(file_path, newline="", encoding="utf-8-sig") as f: # BOM付き無し両方のUTF-8に対応
        reader = csv.reader(f)
        next(reader)  # Skip header row
        return [(int(a), int(b)) for a, b in reader]

    
