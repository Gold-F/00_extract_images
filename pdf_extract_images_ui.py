# ユーザーの入力から企画名や定型外サイズを取得するUIモジュール

import tkinter as tk
from tkinter import simpledialog, messagebox
import os  # フォルダ操作に使用
import sys  # プログラム終了に使用


def get_ad_text():
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
    return ad_text

def get_sub_text():
    max_length = 25
    max_lines = 2
    # 企画名（上部テキスト）入力ウィンドウを表示
    initial_text = ""
    prompt = f"上部に入れるテキストがある場合は入力してください。\n1行の最大文字数は{max_length}文字、{max_lines}行まで入力可能です。\n"
    while True:
        sub_text = get_multiline_input(title="上部テキスト入力", prompt=prompt, initial_text=initial_text)
        # 終了（テキスト入力で「終了」ボタンまたは×ボタンが押された場合）
        if sub_text is None:
            messagebox.showinfo("キャンセル", "処理を中止しました。")
            sys.exit()
        confirm = messagebox.askyesno("テキスト確認", f"以下の内容でよろしいですか？\n\n{sub_text}")
        if confirm:
            break  # 入力確定
        else:            
            initial_text = sub_text # 再入力のため、初期テキストを更新してループ継続
    return sub_text

def get_teikei_flag_and_custom_size():
    # 定型外サイズの登録、無い場合はスルー
    teikei_flag = True
    response = messagebox.askyesnocancel("定形外の確認", "定型サイズは「はい」、定型外サイズを含む場合は「いいえ」を押してください")
    if response is False:  # 「いいえ」が押された場合
        teikei_flag = False
        custom_w, custom_h = get_custom_size() # 定型外サイズをsizesに追加
    elif response is None:  # 「キャンセル」または×が押された場合
        messagebox.showinfo("終了", "処理を終了します。")
        sys.exit()
    return teikei_flag, (custom_w, custom_h)

def get_teikei_flag():
    # 定型外サイズの登録、無い場合はスルー
    response = messagebox.askyesnocancel("定形外の確認", "定型サイズは「はい」、定型外サイズを含む場合は「いいえ」を押してください")
    if response is None:
        messagebox.showinfo("終了", "処理を終了します。")
        sys.exit()
    return response  # True or False


def get_multiline_input(title="", prompt="", initial_text=""):
    max_length = 25
    max_lines = 2
    height=2
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
        if len(lines) > max_lines or any(len(line) > max_length for line in lines):
            # 修正処理：前の有効状態に戻す
            text_widget.delete("1.0", "end")
            text_widget.insert("1.0", on_text_change.last_valid_text)
        else:
            # 有効なら最新状態を保存
            on_text_change.last_valid_text = text

    def on_return(event):
        text = text_widget.get("1.0", "end-1c")
        lines = text.splitlines()
        if len(lines) >= max_lines:
            return "break"  # Enter を無効化（Tkinterの特殊戻り値）
        return None  # 通常通り改行

    result = None
    dialog = tk.Toplevel()
    dialog.title(title)
    dialog.grab_set()  # モーダルにする

    tk.Label(dialog, text=prompt).pack(padx=10, pady=5)
    text_widget = tk.Text(dialog, width=max_length+25, height=height)
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

    messagebox.showinfo("定型外サイズ登録完了", f"定型外 (幅: {width} mm  高さ: {height} mm)\nが追加されました。")
    return (width, height)
    



