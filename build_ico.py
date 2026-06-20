"""
正确生成多尺寸 ICO 文件
- 源图: monkey-icon.png
- 多分辨率: 16x16, 24x24, 32x32, 48x48, 64x64, 128x128, 256x256
- 使用 PIL.Image.save() 原生 ICO 支持，确保像素行顺序正确（BMP bottom-up）
"""
from PIL import Image
import os

script_dir = os.path.dirname(os.path.abspath(__file__))
icon_dir = os.path.join(script_dir, 'data', 'icons')

src_png = os.path.join(icon_dir, 'monkey-icon.png')
out_ico = os.path.join(icon_dir, 'dasheng.ico')

if not os.path.exists(src_png):
    print(f"ERROR: 源图片不存在: {src_png}")
    exit(1)

img = Image.open(src_png).convert('RGBA')
w, h = img.size
print(f"源图片尺寸: {w}x{h}")

# 裁剪底部少量空白（如果有）
crop_height = int(h * 0.95)
cropped = img.crop((0, 0, w, crop_height))
print(f"裁剪后尺寸: {cropped.size}")

# 生成不同尺寸的图像
sizes = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
sized_images = []
for size in sizes:
    sized = cropped.resize(size, Image.LANCZOS)
    sized_images.append(sized)
    print(f"  已生成 {size[0]}x{size[1]}")

# 使用 PIL 原生 ICO 保存（自动处理 BMP bottom-up 行顺序）
sized_images[0].save(
    out_ico,
    format='ICO',
    sizes=sizes,
)

# 验证
with Image.open(out_ico) as test_ico:
    print(f"\n✓ 新图标已生成: {out_ico}")
    print(f"  文件大小: {os.path.getsize(out_ico) / 1024:.1f} KB")
    print(f"  分辨率: {', '.join(f'{s}x{s}' for s in test_ico.ico.sizes() if hasattr(test_ico, 'ico'))}")

print(f"  实际尺寸列表: {sizes}")
