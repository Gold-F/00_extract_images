# PDFを読み込み、各ページを画像として保存し、指定サイズの長方形を検出して切り出す
# 矩形を検出して定型サイズに合う広告を切り出す（定型外サイズも登録可）
# A4サイズ縦向き(横幅380㎜のみA3横向き)の中央に画像を配置する
# 下部に日付と企画名を日本語で描画する
# エラー表示用として黒い画面を出す。
# 新聞対応
import cv2
import numpy as np
from PIL import Image
from PIL import ImageDraw
import pypdfium2 as pdfium
from pypdf import PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import tkinter as tk
from tkinter import simpledialog, messagebox
import os  # フォルダ操作に使用
import sys  # プログラム終了に使用
import traceback  # エラー表示用

folderRef = os.path.join(os.path.dirname(os.path.abspath(__file__)), "storage_")
dpi = 150  # 解像度（DPI）

# 日本語対応フォントの登録
def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):  # PyInstaller実行時
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)
font_path = resource_path(r"fonts/NotoSansJP-Regular.ttf")  # TTF or OTF推奨
if not os.path.exists(font_path):
    messagebox.showerror("エラー", f"フォントファイルが見つかりません: {font_path}")
    sys.exit()
pdfmetrics.registerFont(TTFont("NotoJP", font_path))

# サイズリスト（幅×高さ mm）
sizes = [
    (92, 32),
    (61, 32),
    (380, 66),
    (188, 66),
    (124, 66),
    (92, 66),
    (73, 66),
    (61, 66),
    (45, 66),
    (30, 66),
    (380, 100),
    (188, 100),
    (380, 135),
    (188, 135),
    (380, 169),
    (188, 169),
    (73, 32),
    (45, 32)
]
root = tk.Tk()
root.withdraw()  # メインウィンドウを非表示にする

def get_multiline_input(title="", prompt="", initial_text="", maxlen=25, maxlines=2, height=2):
    def on_ok():
        nonlocal result
        result = text_widget.get("1.0", "end-1c")  #1行目の0文字目　から すべてのテキスト（末尾の改行を除く）を取得
        dialog.destroy()

    def on_cancel():
        nonlocal result
        result = None
        dialog.destroy()
    
    def clear_initial_text(event):
        if text_widget.get("1.0", "end-1c") == initial_text:
            text_widget.delete("1.0", "end")
        # イベントバインドを1回だけで解除する
        text_widget.unbind("<FocusIn>")

    def on_text_change(event):
        text_widget.edit_modified(False)  # まずフラグをリセット
        text = text_widget.get("1.0", "end-1c")
        lines = text.splitlines()
        # チェック処理
        text_widget.bind("<Return>", on_return)
        if len(lines) > maxlines or any(len(line) > maxlen for line in lines):
            # 修正処理：前の有効状態に戻す
            text_widget.delete("1.0", "end")
            text_widget.insert("1.0", on_text_change.last_valid_text)
        else:
            # 有効なら最新状態を保存
            on_text_change.last_valid_text = text

    def on_return(event):
        text = text_widget.get("1.0", "end-1c")
        lines = text.splitlines()
        if len(lines) >= maxlines:
            return "break"  # Enter を無効化（Tkinterの特殊戻り値）
        return None  # 通常通り改行
    
    result = None
    dialog = tk.Toplevel()
    dialog.title(title)
    dialog.grab_set()  # モーダルにする

    tk.Label(dialog, text=prompt).pack(padx=10, pady=5)
    text_widget = tk.Text(dialog, width=maxlen+25, height=height)
    text_widget.insert("1.0", initial_text)
    text_widget.pack(padx=10)
    button_frame = tk.Frame(dialog)
    button_frame.pack(pady=10)
    tk.Button(button_frame, text="   O K   ", command=on_ok).pack(side="left", padx=10)
    tk.Button(button_frame, text="Cancel", command=on_cancel).pack(side="left", padx=10)
    text_widget.bind("<FocusIn>", clear_initial_text)
    text_widget.bind("<<Modified>>", on_text_change)
    text_widget.bind("<Return>", on_return)
    on_text_change.last_valid_text = initial_text
    # ×ボタンでon_cancel()が呼ばれるようにする
    dialog.protocol("WM_DELETE_WINDOW", on_cancel)
    dialog.wait_window()
 
    if result is None:
        return None  # Cancelボタンか×ボタンが押された場合
    return result

def get_custom_size():    
    # 幅の入力ループ
    while True:
        width = simpledialog.askstring("幅", "定型外サイズの「横幅」を\n半角数字で入力してください \n単位：mm")
        if width is None:
            sys.exit()
        if width.strip() == "" or not width.isdigit():
            messagebox.showerror("注意", "半角数字を入力してください\n(20mm以上)")
            continue
        width = int(width.strip())
        if width >= 20:
            break
        else:
            messagebox.showerror("注意", "半角数字を入力してください\n(20mm以上)")
            continue

    # 高さの入力ループ
    while True:
        height = simpledialog.askstring("高さ", "定型外サイズの「高さ」を\n半角数字で入力してください \n単位：mm")
        if height is None:
            sys.exit()
        if height.strip() == "" or not height.isdigit():
            messagebox.showerror("注意", "半角数字を入力してください\n(20mm以上)")
            continue
        height = int(height.strip())
        if height >= 20:
            break
        else:
            messagebox.showerror("注意", "半角数字を入力してください\n(20mm以上)")
            continue

    sizes.append((width, height))    
    messagebox.showinfo("定型外サイズ登録完了", f"定型外 (幅: {width} mm  高さ: {height} mm)\nが追加されました。")
    

def sanitize_filename(text):
    invalid_chars = '\\/*?:"<>|'
    return text.translate(str.maketrans({ch: '_' for ch in invalid_chars}))

def main():
    try:
        # コンソールにメッセージを表示
        print("START [ PDF_blanket ]")  
        if not os.path.exists(folderRef):
            os.makedirs(folderRef, exist_ok=True)

        pdf_files = [f for f in os.listdir(folderRef) if f.endswith('.pdf')]
        if len(pdf_files) < 1:
            messagebox.showerror("終了",f"'{folderRef}'  にPDFファイルを入れてください。")
            sys.exit()

        # 企画名（下部テキスト）入力ウィンドウを表示
        max_length = 25
        max_lines = 2
        initial_text = f"下部に掲載されるテキストです\n1行の最大文字数は{max_length}文字、{max_lines}行まで入力可能です"
        prompt = (
            "『広告掲載年月日 企画名』を入力してください。1行目がファイル名になります。\n"
            "『\\/*?:\"<>|』は使えません。\n"
        )
        while True:
            ad_text = get_multiline_input(title="企画名入力（下部テキスト）", prompt=prompt, initial_text=initial_text)
            # 終了（テキスト入力で「終了」ボタンまたは×ボタンが押された場合）
            if ad_text is None:
                messagebox.showinfo("キャンセル", "処理を中止しました。")
                sys.exit()
            confirm = messagebox.askyesno("テキスト確認", f"以下の内容でよろしいですか？\n\n{ad_text}")
            if confirm:
                break  # 入力確定
            else:            
                initial_text = ad_text # 再入力のため、初期テキストを更新してループ継続

        # 企画名（上部テキスト）入力ウィンドウを表示
        initial_text = ""
        prompt = f"上部に入れるテキストがある場合は入力してください。\n1行の最大文字数は{max_length}文字、{max_lines}行まで入力可能です。\n"
        while True:
            sub_text = get_multiline_input(title="上部テキスト入力", prompt=prompt, initial_text=initial_text, maxlen=max_length, maxlines=max_lines)
            # 終了（テキスト入力で「終了」ボタンまたは×ボタンが押された場合）
            if sub_text is None:
                messagebox.showinfo("キャンセル", "処理を中止しました。")
                sys.exit()
            confirm = messagebox.askyesno("テキスト確認", f"以下の内容でよろしいですか？\n\n{sub_text}")
            if confirm:
                break  # 入力確定
            else:            
                initial_text = sub_text # 再入力のため、初期テキストを更新してループ継続

        # 定型外サイズの登録、無い場合はスルー
        teikei_flag = True
        response = messagebox.askyesnocancel("定形外の確認", "定型サイズは「はい」、定型外サイズを含む場合は「いいえ」を押してください")
        if response is False:  # 「いいえ」が押された場合
            teikei_flag = False
            get_custom_size() # 定型外サイズをsizesに追加
        elif response is None:  # 「キャンセル」または×が押された場合
            messagebox.showinfo("終了", "処理を終了します。")
            sys.exit()

        page_counter = 0
        for pdf_file in pdf_files:
            input_pdf_path = os.path.join(folderRef, pdf_file)
            print(f"Processing file: {os.path.basename(input_pdf_path)}")

            with pdfium.PdfDocument(input_pdf_path) as pdf:
                for page_num in range(len(pdf)):
                    page_counter += 1
                    doc_name = os.path.basename(input_pdf_path)
                    doc_name = os.path.splitext(doc_name)[0] # 拡張子.pdf を除去　
                    counter_folder = os.path.join(folderRef, f"{page_counter:03d}_{doc_name}")
                    os.makedirs(counter_folder, exist_ok=True)
                    parts_folder = os.path.join(counter_folder, "parts")
                    os.makedirs(parts_folder, exist_ok=True)
                    page = pdf[page_num]
                    pil_image = page.render(
                        scale=dpi / 72
                    ).to_pil()
                    input_png_path = os.path.join(counter_folder, f"page_{page_counter:03d}.png")
                    pil_image.save(input_png_path)                   
                    print(f"    Extracted page  page_{page_counter:03d}.png")

                    # withブロックを抜けたら自動的にcloseされる！

                    #[pillow] 画像を開く
                    image_pil = Image.open(input_png_path)
                    # [pillow] 外側の長方形があると内部の長方形が取得できないのでトリミングする
                    left, right, top, bottom = int(62), int(62), 0, 0 # 左右のみトリミング
                    width, height = image_pil.size
                    cropped_image = image_pil.crop((left, top, width - right, height - bottom))

                    # [pillow] ImageDrawでcheck用の長方形を描くオブジェクトを作成 
                    draw_check = ImageDraw.Draw(cropped_image)
                    # [pillow] グレースケールに変換 *通常はグレースケールで行なうが、今回はカラーの方が検出精度がよい
                    # cropped_image_gray = cropped_image.convert("L")
                    # [pillow] 明示的にRGBに変換
                    cropped_image_rgb = cropped_image.convert("RGB")
                    # [numpy] 画像をNumPy配列に変換し、OpenCV形式に変更
                    image_cv2 = np.array(cropped_image_rgb) 
                    #[opencv] 画像をガウシアンぼかし エッジがなめらかになる
                    image_cv2 = cv2.GaussianBlur(image_cv2, (5, 5), 0)
                    # [opencv] B, G, Rチャンネル別に取り出し
                    b_channel = image_cv2[:, :, 0]
                    g_channel = image_cv2[:, :, 1]
                    r_channel = image_cv2[:, :, 2]

                    # [opencv] 各チャンネルでCannyエッジ検出 threshold1 を少し高く設定して、弱いエッジを取り除きます。threshold2 を高く設定することで、より強いエッジを検出します。
                    edges_b = cv2.Canny(b_channel, 30, 250)
                    edges_g = cv2.Canny(g_channel, 30, 250)
                    edges_r = cv2.Canny(r_channel, 30, 250)

                    # [opencv] 各エッジ画像を合成（最大値を取る）
                    edges_combined = cv2.max(edges_b, cv2.max(edges_g, edges_r))

                    # [opencv] 輪郭検出（外側の枠を取得）   
                    contours, _ = cv2.findContours(edges_combined, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
                    # contours, _ = cv2.findContours(edges_combined, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
                    contours_sorted = sorted(contours, key=lambda cnt: cv2.boundingRect(cnt)[2], reverse=True)
                    # 結果を表示（デバッグ用）
                    # output_image_path = os.path.join(parts_folder, "edges_1.jpg")
                    # cv2.imwrite(output_image_path, edges_combined)

                    single_num = 0 # 個別PDFの連番       
                    single_pdf_paths = [] # 個別PDFの保存パス一覧(マージ用)
                    # [opencv] 指定サイズ内の枠を探す
                    for cnt in contours_sorted:
                        if cv2.contourArea(cnt) < 13000:  # 150dpiで2cm四方以下はスルー
                            continue
                        x, y, w, h = cv2.boundingRect(cnt) # intで取得される
                        tolerance = dpi*3/25.4 # 3mmの許容誤差
                        if w < dpi*30/25.4 - tolerance and teikei_flag:
                            break
                        if h < dpi*32/25.4 - tolerance and teikei_flag:
                            continue
                        for size in sizes:
                            size_width_mm, size_height_mm = size
                            scaled_width_px = size_width_mm * dpi / 25.4
                            scaled_height_px = size_height_mm * dpi / 25.4

                            if abs(w - scaled_width_px) <= tolerance and abs(h - scaled_height_px) <= tolerance:
                                single_num += 1 
                                # [pillow] ROIを切り出し
                                roi_pil = cropped_image.crop((x, y, x + w, y + h))                                      
                                output_path_pil = os.path.join(parts_folder, f"ad_{single_num:03d}.jpg")
                                roi_pil.save(output_path_pil, quality=90, dpi=(dpi, dpi)) # 解像度設定

                                # [reportlab] A4縦またはA3横のPDFを作成
                                if size_width_mm > 190:
                                    page_size = (1190.54, 841.89)  # A3横向き
                                else:
                                    page_size = A4  # A4縦向き
                                page_width_pt, page_height_pt = page_size
                                output_path_pdf = os.path.join(parts_folder, f"ad_{single_num:03d}.pdf")
                                single_pdf_paths.append(output_path_pdf)  # 後で結合用に追加しておく
                                c = canvas.Canvas(output_path_pdf, pagesize=page_size)

                                # ピクセル → ポイントに変換
                                w_pt = w / dpi * 72
                                h_pt = h / dpi * 72

                                # ページ中央に配置するためのオフセット
                                offset_x_pt = (page_width_pt - w_pt) / 2
                                offset_y_pt = (page_height_pt - h_pt) / 2
                                
                                # [reportlab] 画像挿入
                                c.drawImage(output_path_pil, offset_x_pt, offset_y_pt, width=w_pt, height=h_pt)
                                font_name = "NotoJP"
                                font_size = 18
                                # [reportlab] 日本語テキスト描画（フォント設定必要）
                                c.setFont(font_name, font_size)

                                # テキスト位置の初期値（左下が原点）、右揃え
                                ad_y = 150
                                # reportlabのdrawString() は改行されないので1行ずつ描画
                                if ad_text:
                                    for line in ad_text.splitlines():
                                        text_width = pdfmetrics.stringWidth(line, font_name, font_size)
                                        ad_x = (page_width_pt - text_width) - 80
                                        c.drawString(ad_x, ad_y, line)
                                        ad_y -= 28  # 行間

                                # 上部テキストの描画
                                sub_y = page_height_pt - 100
                                # センター揃え
                                if sub_text:
                                    for line in sub_text.splitlines():
                                        text_width = pdfmetrics.stringWidth(line, font_name, font_size)
                                        sub_x = (page_width_pt - text_width) / 2
                                        c.drawString(sub_x, sub_y, line)
                                        sub_y -= 28  # 行間

                                # # A4の縦横センターに横180mm縦170mmの赤い長方形を描画
                                # if page_size == A4:
                                #     rect_width_mm = 180
                                #     rect_height_mm = 170
                                #     rect_width_pt = rect_width_mm / 25.4 * 72
                                #     rect_height_pt = rect_height_mm / 25.4 * 72
                                #     rect_x = (page_width_pt - rect_width_pt) / 2
                                #     rect_y = (page_height_pt - rect_height_pt) / 2
                                #     c.setStrokeColorRGB(1, 0, 0)  # 赤
                                #     c.setLineWidth(3)
                                #     c.rect(rect_x, rect_y, rect_width_pt, rect_height_pt, stroke=1, fill=0)
                                c.save()                

                                # [pillow] check用画像に長方形と対角線を描画
                                bright_green = (0, 255, 100)  # 明るく鮮やかな緑
                                draw_check.rectangle([x, y, x + w, y + h], outline=bright_green, width=4)
                                draw_check.line([x, y, x + w, y + h], fill=bright_green, width=4)
                                draw_check.line([x + w, y, x, y + h], fill=bright_green, width=4)
                                # 黒い画面に出力
                                print(f"        ad_{single_num:03d}  width : {int(w*25.4 / dpi)} , height : {int(h* 25.4 / dpi)}  (mm)")
                                # [pillow] check用画像を画質を落としてJPGで保存
                                output_path_check = os.path.join(counter_folder, f"check_{page_counter:03d}.jpg")
                                cropped_image.save(output_path_check, format="JPEG", quality=50, dpi=(dpi, dpi))
            
                    # [fitz] A4・A3のPDFをマージ   
                    if single_pdf_paths:
                        merged_pdf_writer = PdfWriter()
                        for path in single_pdf_paths:
                            merged_pdf_writer.append(path)
                        first_line = ad_text.splitlines()[0] if ad_text.splitlines() else "binder"
                        output_path_merged = os.path.join(counter_folder, f"{sanitize_filename(first_line)}_{page_counter:03d}.pdf")
                        merged_pdf_writer.write(output_path_merged)
                        print( f"    The file {first_line}_{page_counter:03d}.pdf has been generated.") 
    except Exception as e:
        print("エラーが発生しました:", file=sys.stderr)
        traceback.print_exc()  # スタックトレース（エラーの詳細）を表示
        input("\n[Enter] を押すと終了します")  # 閉じないように待機
        sys.exit(1)  # エラー終了コードで終了
    messagebox.showinfo("完了報告", "Finish!")
    sys.exit()
if __name__ == "__main__":
    main()