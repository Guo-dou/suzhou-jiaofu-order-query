# -*- coding: utf-8 -*-
"""生成多尺寸高清图标: 大尺寸=「苏州交复」完整版, 小尺寸(<=48)=「交」字简化版
小尺寸若照搬四字会糊成一团, 故独立绘制单字版, 保证任务栏/标题栏清晰。"""
import io
import os
import struct
from PIL import Image, ImageDraw, ImageFont

OUT = r"C:\Users\Administrator\Doubao\chats\2026-09-09\new-chat"
FONT = r"C:\Windows\Fonts\msyhbd.ttc"
TOP = (0, 105, 227)       # #0069E3
BOTTOM = (10, 132, 255)   # #0A84FF


def base_icon(size):
    """渐变蓝圆角底"""
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    for y in range(size):
        t = y / size
        c = tuple(int(TOP[i] + (BOTTOM[i] - TOP[i]) * t) for i in range(3))
        d.line([(0, y), (size, y)], fill=c + (255,))
    # 顶部柔和高光
    hl = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    dh = ImageDraw.Draw(hl)
    for y in range(size // 2):
        alpha = int(70 * (1 - y / (size // 2)))
        dh.line([(0, y), (size, y)], fill=(255, 255, 255, alpha))
    img = Image.alpha_composite(img, hl)
    mask = Image.new("L", (size, size), 0)
    dm = ImageDraw.Draw(mask)
    dm.rounded_rectangle([0, 0, size - 1, size - 1], radius=size * 0.226, fill=255)
    img.putalpha(mask)
    return img


def full_icon(size):
    """完整版: 「苏州交复」+ 副文字"""
    img = base_icon(size)
    d = ImageDraw.Draw(img)
    text = "苏州交复"
    f_main = ImageFont.truetype(FONT, int(size * 0.222))
    bb = d.textbbox((0, 0), text, font=f_main)
    tw, th = bb[2] - bb[0], bb[3] - bb[1]
    tx = (size - tw) / 2 - bb[0]
    ty = size / 2 - th / 2 - bb[1] - size * 0.03
    d.text((tx, ty), text, font=f_main, fill=(255, 255, 255, 255))
    if size >= 128:
        sub = "SUZHOU  JIAOFU"
        f_sub = ImageFont.truetype(FONT, max(12, int(size * 0.043)))
        bb2 = d.textbbox((0, 0), sub, font=f_sub)
        tw2, th2 = bb2[2] - bb2[0], bb2[3] - bb2[1]
        sx = (size - tw2) / 2 - bb2[0]
        sy = ty + th + size * 0.094 - bb2[1]
        d.text((sx, sy), sub, font=f_sub, fill=(255, 255, 255, 215))
    return img


def simple_icon(size):
    """简化版: 渐变底 + 单个「交」字, 小尺寸依然清晰"""
    img = base_icon(size)
    d = ImageDraw.Draw(img)
    f = ImageFont.truetype(FONT, int(size * 0.60))
    bb = d.textbbox((0, 0), "交", font=f)
    tw, th = bb[2] - bb[0], bb[3] - bb[1]
    tx = (size - tw) / 2 - bb[0]
    ty = (size - th) / 2 - bb[1]
    d.text((tx, ty), "交", font=f, fill=(255, 255, 255, 255))
    return img


def bmp_bytes(im, size):
    """32bit BGRA + AND mask 的 DIB 数据（ICO 小尺寸用）"""
    ih = struct.pack("<IiiHHIIiiII", 40, size, size * 2, 1, 32, 0, 0, 0, 0, 0, 0)
    px = im.convert("RGBA").load()
    rows = []
    for y in range(size - 1, -1, -1):
        row = bytearray()
        for x in range(size):
            r, g, b, a = px[x, y]
            row += bytes((b, g, r, a))
        rows.append(bytes(row))
    xor = b"".join(rows)
    mask_row_len = ((size + 31) // 32) * 4
    and_mask = b"\x00" * (mask_row_len * size)
    return ih + xor + and_mask


def build_ico(path, entries):
    """entries: [(请求尺寸, 已绘制好的同尺寸 RGBA 图)]"""
    blobs = []
    for size, im in entries:
        if size >= 256:
            data = im.resize((256, 256), Image.LANCZOS)
            buf = io.BytesIO()
            data.save(buf, format="PNG")
            blobs.append((size, buf.getvalue()))
        else:
            blobs.append((size, bmp_bytes(im, size)))
    header = struct.pack("<HHH", 0, 1, len(blobs))
    offset = 6 + 16 * len(blobs)
    out = bytearray(header)
    for (size, data) in blobs:
        w = 256 if size >= 256 else size
        out += struct.pack("<BBBBHHII", w & 0xFF, w & 0xFF, 0, 0, 1, 32,
                           len(data), offset)
        offset += len(data)
    for (_size, data) in blobs:
        out += data
    with open(path, "wb") as f:
        f.write(out)
    print(f"ICO 已生成: {path} ({len(blobs)} 个尺寸)")


# 完整版: 64/128/256/512（512 在 ICO 中按 256 存储）
full = [(64, full_icon(64)), (128, full_icon(128)), (256, full_icon(256))]
# 简化版: 16/24/32/48
simple = [(16, simple_icon(16)), (24, simple_icon(24)),
          (32, simple_icon(32)), (48, simple_icon(48))]
build_ico(os.path.join(OUT, "icon.ico"), simple + full)

# 源 PNG: 完整版 1024
full_icon(1024).save(os.path.join(OUT, "icon.png"))
print("icon.png 已生成 (1024 完整版)")
