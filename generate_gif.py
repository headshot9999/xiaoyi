from PIL import Image, ImageDraw, ImageFont
import math

FONT_PATH = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
SIZE = 128
TOTAL_FRAMES = 20
TEXT = "New"

def make_frame(alpha: float) -> Image.Image:
    """alpha: 0.0 = 完全透明, 1.0 = 完全不透明"""
    img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # 背景圆形 #00b5ef
    bg_alpha = int(255 * alpha)
    draw.ellipse([4, 4, SIZE - 4, SIZE - 4], fill=(0, 181, 239, bg_alpha))

    # 白色描边
    stroke_alpha = int(200 * alpha)
    draw.ellipse([4, 4, SIZE - 4, SIZE - 4], outline=(255, 255, 255, stroke_alpha), width=3)

    # "重要"文字
    font_size = 44
    font = ImageFont.truetype(FONT_PATH, font_size)
    text_alpha = int(255 * alpha)

    bbox = draw.textbbox((0, 0), TEXT, font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]
    x = (SIZE - text_w) // 2 - bbox[0]
    y = (SIZE - text_h) // 2 - bbox[1]

    draw.text((x, y), TEXT, font=font, fill=(255, 255, 255, text_alpha))
    return img


def rgba_to_p(img: Image.Image) -> Image.Image:
    """转换 RGBA 图像为带调色板的 P 模式（支持透明）"""
    return img.convert("RGBA")


frames = []
durations = []

for i in range(TOTAL_FRAMES):
    t = i / TOTAL_FRAMES
    # 平滑正弦闪烁：0→1→0 一个周期
    alpha = (math.sin(t * 2 * math.pi - math.pi / 2) + 1) / 2
    frame = make_frame(alpha)
    frames.append(frame)
    durations.append(60)  # 每帧60ms，整体约1.2秒一轮

output_path = "/workspace/important.gif"
frames[0].save(
    output_path,
    save_all=True,
    append_images=frames[1:],
    duration=durations,
    loop=0,
    disposal=2,
)

print(f"GIF 已生成：{output_path}")
print(f"尺寸：{SIZE}x{SIZE}px，帧数：{TOTAL_FRAMES}，循环：无限")
