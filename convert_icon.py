"""將 PNG 圖示轉換為 ICO 格式"""
from PIL import Image

# 讀取 PNG 圖片
img = Image.open('icon.png')

# 轉換為 RGBA (如果不是的話)
if img.mode != 'RGBA':
    img = img.convert('RGBA')

# 建立多種大小的圖示 (Windows 常用的圖示大小)
icon_sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]

# 儲存為 ICO 格式
img.save('icon.ico', format='ICO', sizes=icon_sizes)
print('圖示轉換完成: icon.ico')
