"""
生成申通风格图标 - 蓝色方框 + 白色快递文字 + 橙色装饰
"""
from PIL import Image, ImageDraw, ImageFont
import os

# 创建输出目录
icons_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "icons")
os.makedirs(icons_dir, exist_ok=True)

# 生成 64x64 图标
size = 128
img = Image.new("RGB", (size, size), "#0B3D91")  # 申通蓝
draw = ImageDraw.Draw(img)

# 画橙色装饰条
draw.rectangle([0, size*0.6, size, size], fill="#FF6B00")  # 橙色

# 画白色"S"字母
try:
    font = ImageFont.truetype("arial.ttf", 70)
except:
    font = ImageFont.load_default()

draw.text((size*0.32, size*0.02), "S", fill="white", font=font)

# 保存
icon_path = os.path.join(icons_dir, "sto_icon.png")
img.save(icon_path)

# 也保存一张小图 32x32
small = img.resize((32, 32), Image.LANCZOS)
small_path = os.path.join(icons_dir, "sto_icon_32.png")
small.save(small_path)

# 生成一张用于标题栏的图
title_bg = Image.new("RGB", (400, 60), "white")
title_draw = ImageDraw.Draw(title_bg)
try:
    title_font = ImageFont.truetype("arial.ttf", 28)
    small_font = ImageFont.truetype("arial.ttf", 14)
except:
    title_font = ImageFont.load_default()
    small_font = ImageFont.load_default()

title_draw.text((70, 0), "申通派费计算系统", fill="#FF6B00", font=title_font)

print(f"✅ 图标已生成：{icon_path}")
print(f"✅ 小图标已生成：{small_path}")
print(f"   颜色：申通蓝 #0B3D91 + 橙色 #FF6B00")
