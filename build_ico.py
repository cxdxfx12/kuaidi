"""
正确生成多尺寸 ICO 文件（使用标准 BMP/DIB 格式）
- 源图: monkey-icon.png (459x510)
- 多分辨率: 16x16, 24x24, 32x32, 48x48, 64x64, 128x128, 256x256
"""
from PIL import Image
import os
import struct

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

crop_height = int(h * 0.92)
cropped = img.crop((0, 0, w, crop_height))
print(f"裁剪后尺寸: {cropped.size}")

sizes = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]

images_data = []
for size in sizes:
    resized = cropped.resize(size, Image.LANCZOS)
    
    raw = resized.tobytes('raw', 'BGRA')
    
    bmp_info = struct.pack(
        '<IIIHHIIIIII',
        40,           # biSize
        size[0],      # biWidth
        size[1] * 2,  # biHeight (height * 2 for ICO format)
        1,            # biPlanes
        32,           # biBitCount
        0,            # biCompression
        len(raw),     # biSizeImage
        0,            # biXPelsPerMeter
        0,            # biYPelsPerMeter
        0,            # biClrUsed
        0             # biClrImportant
    )
    
    images_data.append((size[0], size[1], bmp_info + raw))

ico_data = []

ico_header = struct.pack('<HHH', 0, 1, len(images_data))
ico_data.append(ico_header)

offset = 6 + len(images_data) * 16

directory_entries = []
for w, h, img_data in images_data:
    actual_w = w if w != 256 else 0
    actual_h = h if h != 256 else 0
    img_size = len(img_data)
    
    entry = struct.pack(
        '<BBBBHHII',
        actual_w,
        actual_h,
        0,
        0,
        1,
        32,
        img_size,
        offset
    )
    directory_entries.append(entry)
    offset += img_size

ico_data.extend(directory_entries)

for _, _, img_data in images_data:
    ico_data.append(img_data)

with open(out_ico, 'wb') as f:
    f.write(b''.join(ico_data))

print(f"OK 新图标已生成：{out_ico}")
print(f"  分辨率: {', '.join(f'{s[0]}x{s[1]}' for s in sizes)}")

with open(out_ico, 'rb') as f:
    data = f.read()
    num_images = struct.unpack('<H', data[4:6])[0]
    print(f"  ICO 实际图像数: {num_images}")
    
    offset = 6
    for i in range(num_images):
        w = data[offset]
        h = data[offset + 1]
        w = 256 if w == 0 else w
        h = 256 if h == 0 else h
        img_size = struct.unpack('<I', data[offset + 8:offset + 12])[0]
        print(f"    图像 {i}: {w}x{h}, 大小={img_size}字节")
        offset += 16
