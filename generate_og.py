"""Generate a 1200x630 Open Graph preview image for the resume site."""
from PIL import Image, ImageDraw, ImageFont

W, H = 1200, 630
NAVY = (31, 42, 68)
NAVY2 = (43, 59, 94)
ACCENT = (110, 160, 210)
LIGHT = (215, 226, 239)
WHITE = (255, 255, 255)

img = Image.new("RGB", (W, H), NAVY)
draw = ImageDraw.Draw(img)

# Vertical gradient navy -> lighter navy
for y in range(H):
    t = y / H
    r = int(NAVY[0] + (NAVY2[0] - NAVY[0]) * t)
    g = int(NAVY[1] + (NAVY2[1] - NAVY[1]) * t)
    b = int(NAVY[2] + (NAVY2[2] - NAVY[2]) * t)
    draw.line([(0, y), (W, y)], fill=(r, g, b))


def load_font(size, bold=False):
    candidates = (
        ["seguisb.ttf", "segoeuib.ttf", "arialbd.ttf"] if bold
        else ["segoeui.ttf", "arial.ttf"]
    )
    for name in candidates:
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    return ImageFont.load_default()


f_name = load_font(88, bold=True)
f_role = load_font(38, bold=False)
f_stat = load_font(30, bold=True)
f_mark = load_font(46, bold=True)

MARGIN = 80

# Initials badge (top-left)
badge_x, badge_y, badge_s = MARGIN, 70, 92
draw.rounded_rectangle(
    [badge_x, badge_y, badge_x + badge_s, badge_y + badge_s],
    radius=18, outline=ACCENT, width=3,
)
mb = draw.textbbox((0, 0), "AV", font=f_mark)
draw.text(
    (badge_x + (badge_s - (mb[2] - mb[0])) / 2 - mb[0],
     badge_y + (badge_s - (mb[3] - mb[1])) / 2 - mb[1]),
    "AV", font=f_mark, fill=WHITE,
)

# Name
draw.text((MARGIN, 250), "Art Van de Riet", font=f_name, fill=WHITE)

# Role
draw.text((MARGIN, 356), "Systems Engineer  |  Automation & Change Management",
          font=f_role, fill=ACCENT)

# Accent divider
draw.line([(MARGIN, 430), (MARGIN + 300, 430)], fill=ACCENT, width=4)

# Highlight stats row (auto-fit width, left-aligned to name)
stats = ["10+ yrs experience", "1,000+ changes/yr", "0 error rate", "4 tools shipped"]
y = 472
avail = W - 2 * MARGIN
bullet = "  \u2022  "


def row_width(font):
    total = 0
    for i, s in enumerate(stats):
        if i > 0:
            total += draw.textlength(bullet, font=font)
        total += draw.textlength(s, font=font)
    return total


# Shrink gap/font until the row fits within the available width
stat_size = 30
while stat_size > 20:
    f_stat = load_font(stat_size, bold=True)
    if row_width(f_stat) <= avail:
        break
    stat_size -= 1

x = MARGIN
for i, s in enumerate(stats):
    if i > 0:
        draw.text((x, y), bullet, font=f_stat, fill=ACCENT)
        x += draw.textlength(bullet, font=f_stat)
    draw.text((x, y), s, font=f_stat, fill=LIGHT)
    x += draw.textlength(s, font=f_stat)

img.save("og-image.png", "PNG")
print("Saved og-image.png", img.size)
