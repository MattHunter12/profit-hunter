from PIL import Image, ImageDraw, ImageFont, ImageFilter
from pathlib import Path

OUT = Path("landing/assets")
OUT.mkdir(parents=True, exist_ok=True)

W, H = 1600, 1000
img = Image.new("RGB", (W, H), "#f3efe6")
draw = ImageDraw.Draw(img)

for y in range(H):
    r = int(242 - y * 0.015)
    g = int(239 - y * 0.010)
    b = int(230 - y * 0.002)
    draw.line([(0, y), (W, y)], fill=(r, g, b))

try:
    font_big = ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial Bold.ttf", 64)
    font_med = ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial Bold.ttf", 34)
    font = ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial.ttf", 28)
    font_small = ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial.ttf", 22)
except Exception:
    font_big = font_med = font = font_small = ImageFont.load_default()

# Soft desk shadow
shadow = Image.new("RGBA", (W, H), (0, 0, 0, 0))
sdraw = ImageDraw.Draw(shadow)
sdraw.rounded_rectangle((255, 200, 1345, 760), radius=42, fill=(42, 48, 56, 80))
shadow = shadow.filter(ImageFilter.GaussianBlur(36))
img = Image.alpha_composite(img.convert("RGBA"), shadow)
draw = ImageDraw.Draw(img)

# Laptop/dashboard panel
draw.rounded_rectangle((240, 160, 1320, 720), radius=36, fill="#fbfaf5", outline="#ded7ca", width=3)
draw.rounded_rectangle((240, 160, 1320, 250), radius=36, fill="#19352f")
draw.rectangle((240, 215, 1320, 250), fill="#19352f")
draw.text((300, 192), "Money Leak Finder", fill="#fbfaf5", font=font_med)
draw.text((300, 272), "Map the paycheck before it disappears", fill="#685f52", font=font)

# KPI cards
cards = [
    ("Take-home pay", "$4,850", "#e8f0d9"),
    ("Fixed bills", "$2,940", "#dce8ef"),
    ("Leak found", "$418", "#f4dccf"),
]
for i, (label, value, color) in enumerate(cards):
    x = 300 + i * 315
    draw.rounded_rectangle((x, 330, x + 270, 455), radius=18, fill=color)
    draw.text((x + 24, 352), label, fill="#51483f", font=font_small)
    draw.text((x + 24, 385), value, fill="#17231f", font=font_big)

# Four buckets
buckets = [
    ("Bills", "#c85f42", 0.61),
    ("Debt", "#28715e", 0.16),
    ("Safety", "#d6a63d", 0.09),
    ("Life", "#4b6f9f", 0.14),
]
x0, y0 = 300, 520
for i, (name, color, pct) in enumerate(buckets):
    y = y0 + i * 44
    draw.text((x0, y), name, fill="#28251f", font=font_small)
    draw.rounded_rectangle((x0 + 160, y + 4, x0 + 790, y + 28), radius=12, fill="#ebe4d8")
    draw.rounded_rectangle((x0 + 160, y + 4, x0 + 160 + int(630 * pct), y + 28), radius=12, fill=color)
    draw.text((x0 + 815, y), f"{int(pct * 100)}%", fill="#28251f", font=font_small)

# Notebook and calculator details
draw.rounded_rectangle((1050, 495, 1245, 655), radius=18, fill="#ffffff", outline="#ddd4c6", width=2)
draw.text((1080, 525), "Next move", fill="#1b322c", font=font_small)
draw.line((1080, 570, 1215, 570), fill="#d0c7ba", width=3)
draw.line((1080, 605, 1185, 605), fill="#d0c7ba", width=3)

draw.rounded_rectangle((980, 760, 1210, 930), radius=24, fill="#25312d")
for row in range(3):
    for col in range(4):
        x = 1010 + col * 45
        y = 800 + row * 38
        draw.rounded_rectangle((x, y, x + 30, y + 24), radius=6, fill="#f6f0e5")

img = img.convert("RGB")
img.save(OUT / "profit-hunter-money-leak-dashboard.png", quality=92)
