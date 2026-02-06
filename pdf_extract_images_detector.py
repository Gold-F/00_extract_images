# PDFを読み込み、各ページを画像として保存し、指定サイズの長方形を検出して切り出す
# 矩形を検出して定型サイズに合う広告を切り出す（定型外サイズも登録可）
# A4サイズ縦向き(横幅380㎜のみA3横向き)の中央に画像を配置する
# 下部に日付と企画名を日本語で描画する
# エラー表示用として黒い画面を出す。

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
from tkinter import messagebox
import os  # フォルダ操作に使用
import sys  # プログラム終了に使用

class AdDetector:
    def __init__(self,storage_dir, font_path, pdf_files, ad_text, sub_text, ad_sizes, trim_sizes):
        self.storage_dir = storage_dir
        self.font_path = font_path
        self.pdf_files = pdf_files
        self.ad_text = ad_text
        self.sub_text = sub_text
        self.ad_sizes = ad_sizes
        self.trim_sizes = trim_sizes
        self.dpi = 150
        self.tolerance_px = self.dpi * 3 / 25.4
        self.min_area = 13000
        self.tolerance_mm = 3
        # フォント設定
        self.set_font(self.font_path)


        
    def set_font(self, font_path):
    # 日本語対応フォントの登録 # TTF or OTF推奨
        if not os.path.exists(font_path):
            messagebox.showerror("エラー", f"フォントファイルが見つかりません: {font_path}")
            sys.exit()
        pdfmetrics.registerFont(TTFont("NotoJP", font_path))

    def sanitize_filename(self, text):
        invalid_chars = '\\/*?:"<>|'
        return text.translate(str.maketrans({ch: '_' for ch in invalid_chars}))

    def run(self):
        dpi = self.dpi
        storage_dir = self.storage_dir
        pdf_files = self.pdf_files
        ad_sizes = self.ad_sizes
        page_counter = 0
        for pdf_file in pdf_files:
            input_pdf_path = os.path.join(storage_dir, pdf_file)
            print(f"Processing file: {os.path.basename(input_pdf_path)}")

            # withブロックを抜けたら自動的にcloseされる！
            with pdfium.PdfDocument(input_pdf_path) as pdf:
                for page_num in range(len(pdf)):
                    page_counter += 1
                    doc_name = os.path.basename(input_pdf_path)
                    doc_name = os.path.splitext(doc_name)[0] # 拡張子.pdf を除去　
                    counter_folder = os.path.join(storage_dir, f"{page_counter:03d}_{doc_name}")
                    os.makedirs(counter_folder, exist_ok=True)
                    parts_folder = os.path.join(counter_folder, "parts")
                    os.makedirs(parts_folder, exist_ok=True)
                    page = pdf[page_num]
                    pil_image = page.render(
                        scale=dpi / 72
                    ).to_pil()
                    input_png_path = os.path.join(counter_folder, f"page_{page_counter:03d}.png")
                    pil_image.save(input_png_path)                   

                    
                    #[pillow] 画像を開く
                    image_pil = Image.open(input_png_path)
                    # [pillow] 外側の長方形があると内部の長方形が取得できないのでトリミングする
                    left, right = self.trim_sizes
                    top, bottom = 0, 0
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
                    # 輪郭を幅でソート（大きい順）
                    contours_sorted = sorted(contours, key=lambda cnt: cv2.boundingRect(cnt)[2], reverse=True)
                    # 結果を表示（デバッグ用）
                    # output_image_path = os.path.join(parts_folder, "edges_1.jpg")
                    # cv2.imwrite(output_image_path, edges_combined)
                    print(f"        Detected {len(contours_sorted)} contours")
                    single_num = 0 # 個別PDFの連番       
                    single_pdf_paths = [] # 個別PDFの保存パス一覧(マージ用)
                    # [opencv] 指定サイズ内の枠を探す
                    for cnt in contours_sorted:
                        if cv2.contourArea(cnt) < self.min_area:  # 150dpiで2cm四方以下はスルー
                            continue
                        x, y, w, h = cv2.boundingRect(cnt) # intで取得される
                        tolerance = dpi*self.tolerance_mm/25.4 # 3mmの許容誤差
                        # if w < dpi*30/25.4 - tolerance and teikei_flag:
                        #     break
                        # if h < dpi*32/25.4 - tolerance and teikei_flag:
                        #     continue
                        for size in ad_sizes:
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
                                if self.ad_text:
                                    for line in self.ad_text.splitlines():
                                        text_width = pdfmetrics.stringWidth(line, font_name, font_size)
                                        ad_x = (page_width_pt - text_width) - 80
                                        c.drawString(ad_x, ad_y, line)
                                        ad_y -= 28  # 行間

                                # 上部テキストの描画
                                sub_y = page_height_pt - 100
                                # センター揃え
                                if self.sub_text:
                                    for line in self.sub_text.splitlines():
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
                        first_line = self.ad_text.splitlines()[0] if self.ad_text.splitlines() else "binder"
                        output_path_merged = os.path.join(counter_folder, f"{self.sanitize_filename(first_line)}_{page_counter:03d}.pdf")
                        merged_pdf_writer.write(output_path_merged)
                        print( f"    The file {first_line}_{page_counter:03d}.pdf has been generated.") 

