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


# ─── GIF 2: 紧急 (白色胶囊 + 红边框 + 火焰图标) ─────────────────────────────
def gen_urgent_gif():
    W, H = 196, 64
    RED   = (255, 75, 65)
    WHITE = (255, 255, 255)
    font  = ImageFont.truetype(CJK_FONT, 28, index=0)
    frames = []
    for i in range(FRAMES):
        f    = blink(i)
        img  = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        r    = H // 2
        # white pill background
        draw.rounded_rectangle([2, 2, W - 2, H - 2], radius=r,
                                fill=(255, 255, 255, 255))
        # red border
        draw.rounded_rectangle([2, 2, W - 2, H - 2], radius=r,
                                outline=sc(RED, f) + (255,), width=2)
        # flame icon
        draw_flame(draw, 38, H // 2 + 1, 13, 19, sc(RED, f) + (255,))
        # separator
        draw.line([(60, 12), (60, H - 12)], fill=sc(RED, f) + (150,), width=1)
        # text
        bbox = draw.textbbox((0, 0), "紧急", font=font)
        tw   = bbox[2] - bbox[0]
        th   = bbox[3] - bbox[1]
        tx   = 66 + (W - 74 - tw) // 2 - bbox[0]
        ty   = (H - th) // 2 - bbox[1]
        draw.text((tx, ty), "紧急", font=font, fill=sc(RED, f) + (255,))
        frames.append(img)
    save_gif(frames, "/workspace/icon2_urgent.gif")


# ─── GIF 3: 重要 (浅粉胶囊 + 灯泡图标) ──────────────────────────────────────
def gen_important1_gif():
    W, H   = 216, 64
    CORAL  = (255, 95, 85)
    BG     = (255, 236, 235)
    font   = ImageFont.truetype(CJK_FONT, 28, index=0)
    frames = []
    for i in range(FRAMES):
        f    = blink(i, lo=0.4)
        img  = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        r    = H // 2
        # pink pill background (brightness pulses slightly)
        bg_f = 0.85 + 0.15 * f
        draw.rounded_rectangle([2, 2, W - 2, H - 2], radius=r,
                                fill=sc(BG, bg_f) + (255,))
        # bulb icon
        draw_bulb(draw, 42, H // 2, 40, sc(CORAL, f) + (255,))
        # text
        bbox = draw.textbbox((0, 0), "重要", font=font)
        tw   = bbox[2] - bbox[0]
        th   = bbox[3] - bbox[1]
        tx   = 72 + (W - 80 - tw) // 2 - bbox[0]
        ty   = (H - th) // 2 - bbox[1]
        draw.text((tx, ty), "重要", font=font, fill=sc(CORAL, f) + (255,))
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
