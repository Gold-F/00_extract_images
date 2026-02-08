# main.py
import tkinter as tk
import traceback
import sys
import os
from pdf_extract_images_detector import AdDetector
from pdf_extract_images_ui import get_ad_text, get_sub_text, get_teikei_flag, get_custom_size
from pdf_extract_images_loader import collect_pdf_files, load_size_file, load_trim_file, resource_path


def main():
    try:
        # コンソールにメッセージを表示
        print("START [ PDF_extract_images ]")  

        base_dir = (
            os.path.dirname(sys.executable) # sys.executable + sys.frozen → exe 本体の場所
            if getattr(sys, 'frozen', False)
            else os.path.dirname(os.path.abspath(__file__))
        )
        storage_dir = os.path.join(base_dir, "storage_")
        input_dir = os.path.join(base_dir, "input_files")
        font_path = resource_path(os.path.join("fonts", "NotoSansJP-Regular.ttf")) #ビルド時にバンドルするinternalフォルダの中に入る

        # データ読み込み
        pdf_files = collect_pdf_files(storage_dir)
        ad_sizes = load_size_file(os.path.join(input_dir, "ad_size_list.csv"))
        trim_sizes = load_trim_file(os.path.join(input_dir, "trim_size_list.csv"))

        root = tk.Tk()
        root.withdraw()  # メインウィンドウを非表示にする

        # ユーザー入力を取得
        ad_text = get_ad_text()
        sub_text = get_sub_text()
        teikei_flag = get_teikei_flag()

        # 定型外サイズがある場合は追加登録
        if not teikei_flag:
            custom_size = get_custom_size()
            ad_sizes.append(custom_size)

        detector = AdDetector(
            storage_dir = storage_dir,
            font_path = font_path,
            pdf_files = pdf_files,
            ad_text = ad_text,
            sub_text = sub_text,
            ad_sizes = ad_sizes,
            trim_sizes = trim_sizes
        )
        detector.run()
       
        
    except Exception as e:
        print("エラーが発生しました:", file=sys.stderr)
        traceback.print_exc()  # スタックトレース（エラーの詳細）を表示
        input("\n[Enter] を押すと終了します")  # 閉じないように待機
        sys.exit(1)  # エラー終了コードで終了
    tk.messagebox.showinfo("完了報告", "Finish!")
    sys.exit()

if __name__ == "__main__":
     main()
    
    