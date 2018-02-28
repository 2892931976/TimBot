import os
import win32clipboard
from cStringIO import StringIO
from PIL import Image


def copy_to_clipboard(file_path):
    image = Image.open(file_path)

    output = StringIO()
    image.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()

    send_to_clipboard(win32clipboard.CF_DIB, data)


def copy_text_to_clipboard(text):
    send_to_clipboard(win32clipboard.CF_TEXT, text)


def send_to_clipboard(clip_type, data):
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(clip_type, data)
        win32clipboard.CloseClipboard()
    except BaseException as e:
        print(e)


def is_64windows():
  return 'PROGRAMFILES(X86)' in os.environ

def get_tim_path():
    if is_64windows():
        return 'C:\Program Files (x86)\Tencent\TIM\Bin\TIM.exe'
    else:
        return 'C:\Program Files\Tencent\TIM\Bin\TIM.exe'