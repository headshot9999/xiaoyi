from PIL import Image, ImageDraw, ImageFont
import math

CJK_FONT  = "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"
EN_FONT   = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
FRAMES    = 24  # frames per loop

def blink(i, lo=0.3, hi=1.0):
    raw = (math.sin(i / FRAMES * 2 * math.pi - math.pi / 2) + 1) / 2
    return lo + (hi - lo) * raw

def sc(rgb, f):
    return tuple(min(255, int(c * f)) for c in rgb)

def save_gif(frames, path):
    frames[0].save(path, save_all=True, append_images=frames[1:],
                   duration=[70] * len(frames), loop=0, disposal=2)
    print(f"已生成: {path}  ({frames[0].size[0]}x{frames[0].size[1]}px, {len(frames)}帧)")

def draw_star(draw, cx, cy, ro, ri, fill):
    pts = [(cx + (ro if k % 2 == 0 else ri) * math.cos(math.radians(k * 36 - 90)),
            cy + (ro if k % 2 == 0 else ri) * math.sin(math.radians(k * 36 - 90)))
           for k in range(10)]
    draw.polygon(pts, fill=fill)

def draw_flame(draw, cx, cy, w, h, fill):
    pts = [
        (cx,          cy - h * 0.95),
        (cx + w*0.28, cy - h * 0.52),
        (cx + w*0.48, cy - h * 0.05),
        (cx + w*0.32, cy - h * 0.30),
        (cx + w*0.12, cy + h * 0.08),
        (cx,          cy + h * 0.50),
        (cx - w*0.12, cy + h * 0.08),
        (cx - w*0.32, cy - h * 0.30),
        (cx - w*0.48, cy - h * 0.05),
        (cx - w*0.28, cy - h * 0.52),
    ]
    draw.polygon(pts, fill=fill)
    inner = [
        (cx,          cy - h * 0.50),
        (cx + w*0.13, cy - h * 0.18),
        (cx + w*0.08, cy + h * 0.12),
        (cx,          cy + h * 0.32),
        (cx - w*0.08, cy + h * 0.12),
        (cx - w*0.13, cy - h * 0.18),
    ]
    r, g, b = fill[0], fill[1], fill[2]
    draw.polygon(inner, fill=(min(255, r+40), min(255, g+80), min(255, b+20), fill[3] if len(fill)==4 else 255))

def draw_bulb(draw, cx, cy, size, fill):
    r = size * 0.36
    oy = cy - size * 0.05  # bulb center offset upward
    draw.ellipse([cx - r, oy - r * 1.15, cx + r, oy + r * 0.55], fill=fill)
    bw, bh = r * 0.62, r * 0.32
    draw.rectangle([cx - bw, oy + r*0.45, cx + bw, oy + r*0.45 + bh], fill=fill)
    draw.rectangle([cx - bw*0.8, oy + r*0.45 + bh, cx + bw*0.8, oy + r*0.45 + bh*2], fill=fill)
    for deg in range(0, 360, 45):
        rad = math.radians(deg - 90)
        x1 = cx + (r + 3) * math.cos(rad)
        y1 = oy + (r + 3) * math.sin(rad)
        x2 = cx + (r + size * 0.20) * math.cos(rad)
        y2 = oy + (r + size * 0.20) * math.sin(rad)
        draw.line([(x1, y1), (x2, y2)], fill=fill, width=max(2, int(size * 0.07)))


# ─── GIF 1: NEW! ────────────────────────────────────────────────────────────
def gen_new_gif():
    W, H = 200, 76
    RED    = (228, 45, 35)
    YELLOW = (255, 218, 0)
    font = ImageFont.truetype(EN_FONT, 50)
    frames = []
    for i in range(FRAMES):
        f = blink(i)
        img  = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        skew = 13
        bg_pts = [(skew, 0), (W - 4, 0), (W - 4 - skew, H - 1), (4, H - 1)]
        draw.polygon(bg_pts, fill=sc(RED, f) + (255,))
        # shadow
        bbox = draw.textbbox((0, 0), "NEW!", font=font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
        x = (W - tw) // 2 - bbox[0]
        y = (H - th) // 2 - bbox[1] - 2
        draw.text((x + 2, y + 2), "NEW!", font=font, fill=(120, 20, 10, int(180 * f)))
        draw.text((x, y), "NEW!", font=font, fill=sc(YELLOW, f) + (255,))
        frames.append(img)
    save_gif(frames, "/workspace/icon1_new.gif")


# ─── GIF 2: 紧急 (方波急促闪烁 + 橙色边框 + 超大黄色火焰) ──────────────────
def gen_urgent_gif():
    W, H    = 214, 68
    RED_HI  = (255, 12, 12)   # 极亮红（亮态）
    RED_LO  = (80,  4,  4)    # 极暗红（暗态）
    YELLOW  = (255, 220, 0)   # 荧光黄火焰
    ORANGE  = (255, 110, 0)   # 橙色边框
    WHITE   = (255, 255, 255)
    font    = ImageFont.truetype("/usr/share/fonts/opentype/noto/NotoSerifCJK-Bold.ttc", 30, index=0)

    FAST = 14  # 更少帧 + 更短帧时 = 更急促
    frames = []

    for i in range(FAST):
        t = i / FAST
        # 方波闪烁：快速亮→暗切换，在暗态停留更短（制造紧张感）
        if t < 0.38:
            f = 1.0
        elif t < 0.50:
            f = max(0.0, 1.0 - (t - 0.38) / 0.12)
        elif t < 0.72:
            f = 0.0
        else:
            f = min(1.0, (t - 0.72) / 0.28)

        img  = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        r    = H // 2

        # 背景在亮红与极暗红之间剧烈切换
        bg = tuple(int(RED_LO[j] + (RED_HI[j] - RED_LO[j]) * f) for j in range(3))
        draw.rounded_rectangle([2, 2, W - 2, H - 2], radius=r, fill=bg + (255,))

        # 橙色脉冲边框（亮态加粗）
        border_w = 3 if f > 0.5 else 1
        draw.rounded_rectangle([2, 2, W - 2, H - 2], radius=r,
                                outline=tuple(int(c * max(0.2, f)) for c in ORANGE) + (255,),
                                width=border_w)

        # 超大黄色火焰
        flame_col = tuple(int(c * max(0.15, f)) for c in YELLOW) + (255,)
        draw_flame(draw, 43, H // 2 + 1, 22, 30, flame_col)

        # 分割线
        draw.line([(74, 12), (74, H - 12)],
                  fill=(255, 255, 255, int(150 * f)), width=1)

        # 粗体白色"紧急！"
        bbox = draw.textbbox((0, 0), "紧急！", font=font)
        tw   = bbox[2] - bbox[0]
        th   = bbox[3] - bbox[1]
        tx   = 80 + (W - 88 - tw) // 2 - bbox[0]
        ty   = (H - th) // 2 - bbox[1]
        draw.text((tx, ty), "紧急！", font=font,
                  fill=tuple(int(c * max(0.2, f)) for c in WHITE) + (255,))
        frames.append(img)

    frames[0].save("/workspace/icon2_urgent.gif", save_all=True,
                   append_images=frames[1:], duration=[42] * FAST, loop=0, disposal=2)
    print(f"已生成: /workspace/icon2_urgent.gif  ({W}x{H}px, {FAST}帧, 42ms/帧)")


# ─── GIF 3: 重要 (浅粉胶囊 + 灯泡图标) ──────────────────────────────────────
def gen_important1_gif():
    W, H   = 216, 64
    BG     = (225, 80, 40)
    WHITE  = (255, 255, 255)
    YELLOW = (255, 240, 0)
    font   = ImageFont.truetype("/usr/share/fonts/opentype/noto/NotoSerifCJK-Bold.ttc", 28, index=0)
    frames = []
    for i in range(FRAMES):
        f    = blink(i)
        img  = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        r    = H // 2
        # orange-red pill background
        draw.rounded_rectangle([2, 2, W - 2, H - 2], radius=r,
                                fill=sc(BG, f) + (255,))
        # bulb icon in vivid yellow
        draw_bulb(draw, 42, H // 2, 40, sc(YELLOW, f) + (255,))
        # separator line
        draw.line([(70, 12), (70, H - 12)], fill=(255, 255, 255, int(120 * f)), width=1)
        # text in bold white
        bbox = draw.textbbox((0, 0), "重要", font=font)
        tw   = bbox[2] - bbox[0]
        th   = bbox[3] - bbox[1]
        tx   = 76 + (W - 84 - tw) // 2 - bbox[0]
        ty   = (H - th) // 2 - bbox[1]
        draw.text((tx, ty), "重要", font=font, fill=sc(WHITE, f) + (255,))
        frames.append(img)
    save_gif(frames, "/workspace/icon3_important1.gif")


# ─── GIF 4: 重要 (红色横幅 + 白色波浪 + 星形图标) ─────────────────────────────
def gen_important2_gif():
    W, H  = 196, 64
    RED   = (238, 55, 48)
    WHITE = (255, 255, 255)
    PINK  = (255, 185, 180)
    font  = ImageFont.truetype(CJK_FONT, 28, index=0)
    frames = []
    for i in range(FRAMES):
        f    = blink(i)
        img  = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        # red rounded rectangle
        draw.rounded_rectangle([0, 0, W, H], radius=14,
                                fill=sc(RED, f) + (255,))
        # white wave swoosh at bottom-left
        wave = [(0, H * 0.42), (0, H), (W * 0.22, H),
                (W * 0.14, H * 0.72), (W * 0.07, H * 0.54)]
        draw.polygon(wave, fill=(255, 255, 255, int(210 * f)))
        # star: outer white, inner pink
        draw_star(draw, 44, H // 2 - 1, 17, 8,
                  sc(WHITE, f) + (255,))
        draw_star(draw, 44, H // 2 - 1, 11, 5,
                  sc(PINK, f) + (255,))
        # text
        bbox = draw.textbbox((0, 0), "重要", font=font)
        tw   = bbox[2] - bbox[0]
        th   = bbox[3] - bbox[1]
        tx   = 70 + (W - 78 - tw) // 2 - bbox[0]
        ty   = (H - th) // 2 - bbox[1]
        draw.text((tx, ty), "重要", font=font, fill=sc(WHITE, f) + (255,))
        frames.append(img)
    save_gif(frames, "/workspace/icon4_important2.gif")


if __name__ == "__main__":
    gen_new_gif()
    gen_urgent_gif()
    gen_important1_gif()
    gen_important2_gif()
    print("\n全部生成完毕！")
