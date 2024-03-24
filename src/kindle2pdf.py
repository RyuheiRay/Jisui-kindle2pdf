import os
import re
import sys
import time
from pathlib import Path
import pyautogui
import pygetwindow as gw
from PIL import Image, ImageEnhance, ImageGrab
from screeninfo import get_monitors

from PyPDF2 import PdfWriter, PdfReader
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# スクリーンショット画面検索文字列を定義
KINDLE_NAME = 'Kindle for PC'  # Kindle for PC

# 出力パス、ファイル名を定義
OUTPUT_PATH = Path(os.path.dirname(sys.argv[0])) / 'result'
OUTPUT_FILE_NAME = 'output.pdf'
TEMP_FILE_NAME = 'screenshot.png'

# PDF出力のターゲットデバイス定義
OUTPUT_TARGET = {'target': 'Remarkable2', 'height': 1872, 'width': 1404}
OUTPUT_H = OUTPUT_TARGET['height']
OUTPUT_W = OUTPUT_TARGET['width']

# PDF変換モードを定義
PDF_CONVERT_MODE_PIL = 1  # PILを使用してPDFに変換
PDF_CONVERT_MODE_PYPDF = 2  # PyPDF2を使用してPDFに変換
PDF_CONVERT_MODE_REPORTLAB = 3  # reportlabを使用してPDFに変換

# デフォルトのPDF変換モードを設定
PDF_CONVERT_MODE = PDF_CONVERT_MODE_PIL

# コントラスト調整定数
CONTRAST_FACTOR = 1.7


def kindle2pdf():
    # Kindleウィンドウを見つける
    try:
        kindle_window = gw.getWindowsWithTitle(KINDLE_NAME)[0]
    except IndexError:
        print(f"Please open '{KINDLE_NAME}' and prepare for capture.")
        return

    print(f"Start...for {OUTPUT_TARGET['target']}({OUTPUT_H}x{OUTPUT_W})")

    # 必要なパスを作成
    OUTPUT_PATH.mkdir(parents=False, exist_ok=True)

    # スクリーンの幅と高さを取得
    screen_width, screen_height = get_display_resolution(KINDLE_NAME)

    print("Capturing... Don't touch PC")

    # ウィンドウを前面に移動
    kindle_window.activate()

    # ウインドウを全画面表示
    is_full_screen = (kindle_window.width == screen_width and kindle_window.height == screen_height)
    if not is_full_screen:
        pyautogui.press('f11')
        time.sleep(4)
        is_full_screen = True

    # Kindleの表紙に移動
    pyautogui.hotkey('ctrl', 'g')
    time.sleep(0.5)
    pyautogui.press('1')
    time.sleep(0.1)
    pyautogui.press('enter')
    time.sleep(0.1)

    previous_screenshot = None
    while True:
        current_screenshot = capture_kindle_screenshot()
        if previous_screenshot:
            if previous_screenshot == current_screenshot:
                break
        previous_screenshot = current_screenshot
        pyautogui.press('pageup')
        time.sleep(0.1)

    # 全ページスクリーンショットを保存
    idx = 1
    scale_factor = OUTPUT_H / screen_height
    previous_screenshot = None
    while True:
        current_screenshot = capture_kindle_screenshot()
        if previous_screenshot:
            if previous_screenshot == current_screenshot:
                # 最後のスクリーンショット保存
                save_screenshot_png(current_screenshot, idx, scale_factor, CONTRAST_FACTOR)
                break
        previous_screenshot = current_screenshot
        pyautogui.press('pagedown')
        # スクリーンショット保存
        save_screenshot_png(current_screenshot, idx, scale_factor, CONTRAST_FACTOR)
        idx = idx + 1
    print("Capturing done...you can now touch PC")

    # Kindle全画面表示解除
    if is_full_screen:
        pyautogui.press('f11')

    # PDFに変換
    print("Converting to pdf...")
    output_file_name = get_next_output_filename(OUTPUT_PATH, convert_to_valid_filename(get_kindle_title(kindle_window) + '.pdf'))
    print(output_file_name)
    if PDF_CONVERT_MODE == PDF_CONVERT_MODE_PIL:
        convert_png_to_pdf_pil(OUTPUT_PATH, output_file_name)
    elif PDF_CONVERT_MODE == PDF_CONVERT_MODE_PYPDF:
        convert_png_to_pdf_pyfpdf(OUTPUT_PATH, output_file_name)
    elif PDF_CONVERT_MODE == PDF_CONVERT_MODE_REPORTLAB:
        convert_png_to_pdf_reportlab(OUTPUT_PATH, output_file_name)
    else:
        convert_png_to_pdf_pil(OUTPUT_PATH, output_file_name)

    # 中間ファイルを削除
    delete_tmp_files()

    print(f"\nConversion complete!")
    print(f"{idx} pages converted to {output_file_name}.")


def get_kindle_region():
    # Kindleのウィンドウ領域を取得
    window_rect = gw.getWindowsWithTitle(KINDLE_NAME)[0]
    # クロップ領域を設定
    scale_factor = window_rect.height / OUTPUT_H
    region_left = window_rect.left + int(window_rect.width / 2 - OUTPUT_W * scale_factor / 2)
    region_right = region_left + int(OUTPUT_W * scale_factor)
    kindle_region = (region_left, window_rect.top, region_right, window_rect.bottom)
    return kindle_region


def capture_kindle_screenshot():
    return ImageGrab.grab(bbox=get_kindle_region(), all_screens=True)


def get_display_resolution(window_title):
    # Get the window with the specified title
    window = gw.getWindowsWithTitle(window_title)
    if len(window) == 0:
        print("Window with the specified title was not found.")
        return None

    # Get the display resolution where the window is displayed
    for monitor in get_monitors():
        if monitor.x <= window[0].left <= monitor.x + monitor.width and \
                monitor.y <= window[0].top <= monitor.y + monitor.height:
            return monitor.width, monitor.height

    print("Could not retrieve the display resolution where the specified window is displayed.")
    return None


def convert_png_to_pdf_pil(folder_path, output_file):
    images = []
    # 指定したフォルダ内のPNG画像を取得し、ファイル名順にソートする
    for filename in sorted(os.listdir(folder_path)):
        if filename.endswith('.png'):
            img_path = os.path.join(folder_path, filename)
            images.append(Image.open(img_path).convert('RGB'))
    # PNG画像を1枚のPDFに結合する
    if images:
        images[0].save(folder_path / output_file, save_all=True, append_images=images[1:])


def convert_png_to_pdf_pyfpdf(folder_path, output_file):
    images = []
    for filename in sorted(os.listdir(folder_path)):
        if filename.endswith('.png'):
            img_path = os.path.join(folder_path, filename)
            images.append(img_path)

    if images:
        output_pdf = PdfWriter()
        for img_path in images:
            img = Image.open(img_path)
            img_pdf = img.convert('RGB')
            # 画像のサイズに合わせてPDFファイルを作成
            img_pdf.save('temp.pdf')
            img_pdf_file = open('temp.pdf', 'rb')
            img_pdf_reader = PdfReader(img_pdf_file)
            output_pdf.add_page(img_pdf_reader.pages[0])
            img_pdf_file.close()
            os.remove('temp.pdf')

        with open(output_file, 'wb') as f:
            output_pdf.write(f)


def convert_to_valid_filename(input_str):
    # 不適切な文字を空白に置換する正規表現パターン
    invalid_chars_regex = r'[\\/:*?"<>|]'

    # 不適切な文字を空白に置換して、空白をアンダースコアに置換する
    valid_filename = re.sub(invalid_chars_regex, '', input_str)
    valid_filename = valid_filename.replace(' ', '_')

    return valid_filename


def get_kindle_title(kindle_windows):
    # Kindle for PCのウィンドウを取得
    if not kindle_windows:
        return OUTPUT_FILE_NAME

    # 最初のKindle for PCウィンドウのタイトルを取得
    kindle_window_title = kindle_windows.title

    # "Kindle for PC -"の後の文字列を取得
    pattern = re.compile(r"Kindle for PC - (.*)")
    match = pattern.search(kindle_window_title)
    if match:
        return match.group(1)
    else:
        return None


def convert_png_to_pdf_reportlab(folder_path, output_file):
    c = canvas.Canvas(str(output_file), pagesize=letter)
    for filename in sorted(os.listdir(folder_path)):
        if filename.endswith('.png'):
            img_path = os.path.join(folder_path, filename)
            c.drawImage(img_path, 0, 0, width=letter[0], height=letter[1])
            c.showPage()
    c.save()


def delete_tmp_files():
    # フォルダが存在するかどうかを確認
    if not os.path.exists(OUTPUT_PATH):
        print(f"The specified folder '{OUTPUT_PATH}' does not exist.")
        return

    # フォルダ内のファイルを走査し、"screenshot" で始まり ".png" で終わるファイルを削除
    for filename in os.listdir(OUTPUT_PATH):
        if filename.startswith("screenshot") and filename.endswith(".png"):
            file_path = os.path.join(OUTPUT_PATH, filename)
            try:
                os.remove(file_path)  # ファイルを削除
            except Exception as e:
                print(f"Error: Failed to delete {OUTPUT_PATH}. Reason: {e}")


def save_screenshot_png(screenshot, idx, scale_factor, contrast_factor):
    tmp_path = rename_with_index(OUTPUT_PATH / TEMP_FILE_NAME, idx)
    screenshot.save(tmp_path)
    upscale_screenshot_with_contrast(tmp_path, tmp_path, scale_factor, contrast_factor)


def rename_with_postfix(filename, postfix):
    base, ext = os.path.splitext(filename)
    return f"{base}_{postfix}{ext}"


def rename_with_index(filename, idx):
    base, ext = os.path.splitext(filename)
    return f"{base}_{idx:05d}{ext}"


def upscale_screenshot_with_contrast(input_path, output_path, scale_factor, contrast_factor):
    # 画像読み込み
    image = Image.open(input_path)
    # コントラストの調整
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(contrast_factor)
    # アップスケール
    width, height = image.size
    new_width = int(width * scale_factor)
    new_height = int(height * scale_factor)
    scaled_image = image.resize((new_width, new_height), Image.LANCZOS)
    # 保存
    scaled_image.save(output_path)


def get_next_output_filename(folder_path, base_filename):
    # フォルダ内のファイルを取得
    files = os.listdir(folder_path)
    # base_filename にマッチするファイル名のリストを作成
    matching_files = [f for f in files if f.startswith(base_filename.split('.')[0])]
    # マッチするファイルがなければ、そのまま base_filename を返す
    if not matching_files:
        return base_filename
    # 数字を抽出してリストに格納
    numbers = []
    for f in matching_files:
        # ファイル名が '_' を含んでいる場合のみ数字を抽出する
        if '_' in f:
            number_str = f.split('_')[-1].split('.')[0]
            if number_str.isdigit():  # 数字のみで構成されているか確認
                numbers.append(int(number_str))
    # 数字が見つからない場合は、0をセット
    if not numbers:
        max_number = 0
    else:
        # 最大の数字を取得
        max_number = max(numbers)
    # 次の数字を計算
    next_number = max_number + 1
    # 出力ファイル名を生成して返す
    return f"{base_filename.split('.')[0]}_{str(next_number).zfill(3)}.pdf"
