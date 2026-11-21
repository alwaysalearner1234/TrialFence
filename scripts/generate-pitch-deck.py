"""Generate the TrialFence 7-slide live pitch deck.

Dependencies:
    python -m pip install python-pptx qrcode[pil] reportlab

The slides are rendered at 1920x1080 so fonts and layout remain identical on
any presentation laptop. The same slide images are embedded into PowerPoint
and PDF; the generator remains the editable source of truth.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Iterable

import qrcode
from PIL import Image, ImageDraw, ImageFilter, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas


ROOT = Path(__file__).resolve().parents[1]
OUT_PPTX = ROOT / "TrialFence_Phase2_Pitch_Deck.pptx"
OUT_PDF = ROOT / "TrialFence_Phase2_Pitch_Deck.pdf"
OUT_NOTES = ROOT / "PITCH-DECK-SPEAKER-NOTES.md"
TEMP = Path(os.environ.get("LOCALAPPDATA", str(ROOT))) / "Temp" / "opencode" / "trialfence-pitch"
SLIDES_DIR = TEMP / "slides"
SLIDES_DIR.mkdir(parents=True, exist_ok=True)

W, H = 1920, 1080
TOTAL = 7
LIVE_URL = "https://trialfence-frontend.onrender.com/"
REPO_URL = "https://github.com/alwaysalearner1234/TrialFence"

# Restrained healthcare palette.
BG = "#F7FAFC"
WHITE = "#FFFFFF"
NAVY = "#123746"
INK = "#21323A"
SLATE = "#60717A"
TEAL = "#109E92"
TEAL_DARK = "#087A72"
MINT = "#DDF3EF"
SKY = "#E4F0F7"
BLUE = "#3F7F9B"
AMBER = "#F4A261"
AMBER_LIGHT = "#FFF0DE"
CORAL = "#E76F51"
CORAL_LIGHT = "#FDE9E4"
GREEN = "#2F9D67"
GREEN_LIGHT = "#E4F5EC"
LINE = "#D6E1E5"
SOFT = "#EEF3F5"


def _font_path(*names: str) -> str:
    fonts = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts"
    for name in names:
        path = fonts / name
        if path.exists():
            return str(path)
    raise FileNotFoundError(f"None of these fonts were found: {names}")


FONT_REGULAR = _font_path("segoeui.ttf", "arial.ttf")
FONT_BOLD = _font_path("segoeuib.ttf", "arialbd.ttf")
FONT_SEMIBOLD = _font_path("seguisb.ttf", "segoeuib.ttf", "arialbd.ttf")


def font(size: int, weight: str = "regular") -> ImageFont.FreeTypeFont:
    path = FONT_BOLD if weight == "bold" else FONT_SEMIBOLD if weight == "semibold" else FONT_REGULAR
    return ImageFont.truetype(path, size=size)


def rgb(value: str) -> tuple[int, int, int]:
    value = value.lstrip("#")
    return tuple(int(value[i : i + 2], 16) for i in (0, 2, 4))


def gradient_bg(top: str = BG, bottom: str = "#EDF8F5") -> Image.Image:
    a, b = rgb(top), rgb(bottom)
    image = Image.new("RGB", (W, H), a)
    px = image.load()
    for y in range(H):
        t = y / (H - 1)
        c = tuple(round(a[i] * (1 - t) + b[i] * t) for i in range(3))
        for x in range(W):
            px[x, y] = c
    return image


def add_soft_blobs(image: Image.Image) -> None:
    layer = Image.new("RGBA", image.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    d.ellipse((1450, -220, 2100, 430), fill=(*rgb(MINT), 150))
    d.ellipse((-260, 740, 360, 1360), fill=(*rgb(SKY), 125))
    layer = layer.filter(ImageFilter.GaussianBlur(48))
    image.paste(layer, (0, 0), layer)


def text_box(
    draw: ImageDraw.ImageDraw,
    xy: tuple[int, int],
    text: str,
    fnt: ImageFont.FreeTypeFont,
    fill: str = INK,
    max_width: int | None = None,
    line_gap: int = 8,
    anchor: str | None = None,
    align: str = "left",
) -> tuple[int, int, int, int]:
    if max_width is None or "\n" in text:
        lines = text.split("\n")
    else:
        words = text.split()
        lines: list[str] = []
        current = ""
        for word in words:
            candidate = word if not current else f"{current} {word}"
            if draw.textlength(candidate, font=fnt) <= max_width:
                current = candidate
            else:
                if current:
                    lines.append(current)
                current = word
        if current:
            lines.append(current)
    line_height = fnt.getbbox("Ag")[3] - fnt.getbbox("Ag")[1]
    height = len(lines) * line_height + max(0, len(lines) - 1) * line_gap
    x, y = xy
    if anchor == "mm":
        y -= height // 2
    widest = 0
    for line in lines:
        width = int(draw.textlength(line, font=fnt))
        widest = max(widest, width)
        lx = x
        if anchor == "mm" or align == "center":
            lx = x - width // 2
        elif align == "right":
            lx = x - width
        draw.text((lx, y), line, font=fnt, fill=fill)
        y += line_height + line_gap
    return (x, y - height, x + widest, y)


def card(
    image: Image.Image,
    box: tuple[int, int, int, int],
    fill: str = WHITE,
    outline: str = LINE,
    radius: int = 28,
    shadow: bool = True,
    shadow_alpha: int = 26,
) -> None:
    x1, y1, x2, y2 = box
    if shadow:
        layer = Image.new("RGBA", image.size, (0, 0, 0, 0))
        d = ImageDraw.Draw(layer)
        d.rounded_rectangle((x1 + 8, y1 + 12, x2 + 8, y2 + 12), radius=radius, fill=(18, 55, 70, shadow_alpha))
        layer = layer.filter(ImageFilter.GaussianBlur(10))
        image.paste(layer, (0, 0), layer)
    d = ImageDraw.Draw(image)
    d.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=2)


def pill(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], label: str, fill: str, color: str = NAVY) -> None:
    draw.rounded_rectangle(box, radius=(box[3] - box[1]) // 2, fill=fill)
    text_box(draw, ((box[0] + box[2]) // 2, (box[1] + box[3]) // 2), label, font(21, "semibold"), color, anchor="mm", align="center")


def arrow(draw: ImageDraw.ImageDraw, start: tuple[int, int], end: tuple[int, int], color: str = TEAL, width: int = 8) -> None:
    draw.line((start, end), fill=color, width=width)
    x2, y2 = end
    x1, y1 = start
    if abs(x2 - x1) >= abs(y2 - y1):
        direction = 1 if x2 > x1 else -1
        points = [(x2, y2), (x2 - direction * 22, y2 - 14), (x2 - direction * 22, y2 + 14)]
    else:
        direction = 1 if y2 > y1 else -1
        points = [(x2, y2), (x2 - 14, y2 - direction * 22), (x2 + 14, y2 - direction * 22)]
    draw.polygon(points, fill=color)


def check(draw: ImageDraw.ImageDraw, center: tuple[int, int], color: str = WHITE, size: int = 28, width: int = 7) -> None:
    x, y = center
    draw.line((x - size // 2, y, x - 4, y + size // 3, x + size // 2, y - size // 2), fill=color, width=width, joint="curve")


def cross(draw: ImageDraw.ImageDraw, center: tuple[int, int], color: str = WHITE, size: int = 22, width: int = 6) -> None:
    x, y = center
    draw.line((x - size // 2, y - size // 2, x + size // 2, y + size // 2), fill=color, width=width)
    draw.line((x + size // 2, y - size // 2, x - size // 2, y + size // 2), fill=color, width=width)


def person_icon(draw: ImageDraw.ImageDraw, center: tuple[int, int], color: str = NAVY, scale: float = 1.0) -> None:
    x, y = center
    r = int(22 * scale)
    draw.ellipse((x - r, y - int(52 * scale), x + r, y - int(8 * scale)), fill=color)
    draw.rounded_rectangle((x - int(40 * scale), y, x + int(40 * scale), y + int(58 * scale)), radius=int(22 * scale), fill=color)


def clinic_icon(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], color: str = TEAL) -> None:
    x1, y1, x2, y2 = box
    draw.rounded_rectangle(box, radius=14, outline=color, width=6, fill=WHITE)
    mid = (x1 + x2) // 2
    draw.rectangle((mid - 10, y1 + 18, mid + 10, y1 + 68), fill=color)
    draw.rectangle((mid - 26, y1 + 35, mid + 26, y1 + 53), fill=color)
    for x in (x1 + 22, x2 - 40):
        draw.rectangle((x, y2 - 40, x + 18, y2 - 20), fill=color)


def shield_icon(draw: ImageDraw.ImageDraw, center: tuple[int, int], color: str = TEAL, scale: float = 1.0) -> None:
    x, y = center
    w, h = int(100 * scale), int(120 * scale)
    pts = [(x, y - h // 2), (x + w // 2, y - h // 3), (x + int(w * 0.42), y + h // 4), (x, y + h // 2), (x - int(w * 0.42), y + h // 4), (x - w // 2, y - h // 3)]
    draw.polygon(pts, fill=color)
    check(draw, (x, y + 2), WHITE, int(45 * scale), max(4, int(8 * scale)))


def database_icon(draw: ImageDraw.ImageDraw, center: tuple[int, int], color: str = CORAL, scale: float = 1.0) -> None:
    x, y = center
    w, h = int(125 * scale), int(115 * scale)
    draw.rectangle((x - w // 2, y - h // 3, x + w // 2, y + h // 3), fill=color)
    draw.ellipse((x - w // 2, y - h // 2, x + w // 2, y - h // 6), fill=color)
    draw.ellipse((x - w // 2, y + h // 6, x + w // 2, y + h // 2), fill=color)
    draw.arc((x - w // 2, y - h // 8, x + w // 2, y + h // 4), 0, 180, fill=WHITE, width=4)


def browser_chrome(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int]) -> tuple[int, int, int, int]:
    x1, y1, x2, y2 = box
    draw.rounded_rectangle(box, radius=26, fill=WHITE, outline=LINE, width=2)
    draw.rounded_rectangle((x1, y1, x2, y1 + 58), radius=26, fill="#E9F0F2")
    draw.rectangle((x1, y1 + 30, x2, y1 + 58), fill="#E9F0F2")
    for i, color in enumerate((CORAL, AMBER, GREEN)):
        draw.ellipse((x1 + 22 + i * 30, y1 + 20, x1 + 38 + i * 30, y1 + 36), fill=color)
    draw.rounded_rectangle((x1 + 145, y1 + 14, x2 - 28, y1 + 43), radius=14, fill=WHITE)
    text_box(draw, (x1 + 170, y1 + 18), "trialfence-frontend.onrender.com", font(15), SLATE)
    return (x1, y1 + 58, x2, y2)


def base_slide(number: int, kicker: str, title: str, subtitle: str | None = None) -> tuple[Image.Image, ImageDraw.ImageDraw]:
    image = gradient_bg(BG, "#F3F9F7")
    add_soft_blobs(image)
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((104, 72, 158, 126), radius=14, fill=TEAL)
    text_box(draw, (131, 99), "TF", font(20, "bold"), WHITE, anchor="mm", align="center")
    text_box(draw, (177, 83), "TRIALFENCE", font(22, "bold"), NAVY)
    text_box(draw, (177, 111), kicker.upper(), font(17, "semibold"), TEAL_DARK)
    text_box(draw, (110, 166), title, font(52, "bold"), NAVY, max_width=1650, line_gap=4)
    if subtitle:
        text_box(draw, (112, 242), subtitle, font(27), SLATE, max_width=1600, line_gap=7)
    draw.line((110, 1012, 1810, 1012), fill=LINE, width=2)
    text_box(draw, (112, 1027), "COMPSPHERE / HACKSPHERE 2026 · PHASE 2", font(16, "semibold"), SLATE)
    text_box(draw, (1808, 1027), f"{number} / {TOTAL}", font(17, "semibold"), SLATE, align="right")
    return image, draw


def make_slide_1() -> Image.Image:
    image = gradient_bg("#FBFDFE", "#EAF7F4")
    add_soft_blobs(image)
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((110, 80, 170, 140), radius=16, fill=TEAL)
    text_box(draw, (140, 110), "TF", font(22, "bold"), WHITE, anchor="mm", align="center")
    text_box(draw, (192, 92), "TRIALFENCE", font(25, "bold"), NAVY)
    text_box(draw, (192, 125), "PRIVACY-PRESERVING TRIAL ENROLLMENT", font(17, "semibold"), TEAL_DARK)

    text_box(draw, (112, 255), "One trial.\nOne anonymous enrollment.", font(72, "bold"), NAVY, line_gap=6)
    text_box(
        draw,
        (118, 470),
        "Stop duplicate enrollment across clinical trial sites — without sharing participant identity.",
        font(31),
        INK,
        max_width=850,
        line_gap=10,
    )
    pill(draw, (118, 625, 390, 682), "NO PII ON-CHAIN", MINT, TEAL_DARK)
    pill(draw, (410, 625, 717, 682), "ANONYMOUS PROOF", SKY, NAVY)
    pill(draw, (737, 625, 1010, 682), "REPEAT BLOCKED", AMBER_LIGHT, "#9A570E")

    # Hero diagram.
    card(image, (1080, 120, 1800, 895), fill=WHITE, radius=38)
    text_box(draw, (1440, 175), "A shared gate, not a shared identity database", font(25, "semibold"), NAVY, anchor="mm", align="center")
    person_icon(draw, (1440, 285), NAVY, 0.95)
    text_box(draw, (1440, 365), "ANONYMOUS VOLUNTEER", font(18, "bold"), SLATE, anchor="mm", align="center")

    card(image, (1145, 420, 1390, 585), fill="#F5FBFA", radius=22, shadow=False)
    clinic_icon(draw, (1180, 452, 1268, 548), TEAL)
    text_box(draw, (1290, 465), "SITE A", font(24, "bold"), NAVY)
    text_box(draw, (1290, 505), "first proof", font(20), SLATE)
    card(image, (1490, 420, 1735, 585), fill="#F5FBFA", radius=22, shadow=False)
    clinic_icon(draw, (1525, 452, 1613, 548), BLUE)
    text_box(draw, (1635, 465), "SITE B", font(24, "bold"), NAVY)
    text_box(draw, (1635, 505), "repeat proof", font(20), SLATE)
    arrow(draw, (1385, 350), (1270, 420), TEAL, 6)
    arrow(draw, (1495, 350), (1615, 420), BLUE, 6)

    draw.rounded_rectangle((1250, 660, 1630, 815), radius=30, fill=NAVY)
    shield_icon(draw, (1335, 738), TEAL, 0.62)
    text_box(draw, (1420, 690), "FIRST: ACCEPT", font(24, "bold"), WHITE)
    text_box(draw, (1420, 738), "REPEAT: BLOCK", font(24, "bold"), "#FFD2C7")
    text_box(draw, (1440, 855), "Same protocol · same private tag", font(18, "semibold"), SLATE, anchor="mm", align="center")

    draw.line((110, 957, 1810, 957), fill=LINE, width=2)
    text_box(draw, (115, 981), "Lidiya · Syed Saad · Krishna", font(19, "semibold"), NAVY)
    text_box(draw, (960, 981), "LIVE", font(16, "bold"), TEAL, align="right")
    text_box(draw, (980, 981), "trialfence-frontend.onrender.com", font(18), INK)
    return image


def make_slide_2() -> Image.Image:
    image, draw = base_slide(2, "The problem", "Clinical trials face a privacy trade-off")
    card(image, (105, 275, 1025, 900), fill=WHITE, radius=32)
    text_box(draw, (150, 315), "The same person can enroll at two sites", font(31, "bold"), NAVY)

    # Two sites and one repeated participant.
    for x, label, color in ((205, "SITE A", TEAL), (665, "SITE B", BLUE)):
        card(image, (x, 430, x + 300, 640), fill="#F7FBFC", radius=24, shadow=False)
        clinic_icon(draw, (x + 35, 472, x + 137, 585), color)
        text_box(draw, (x + 165, 480), label, font(26, "bold"), NAVY)
        person_icon(draw, (x + 210, 570), NAVY, 0.55)
    text_box(draw, (515, 390), "SAME PERSON", font(21, "bold"), CORAL, anchor="mm", align="center")
    arrow(draw, (515, 422), (355, 485), CORAL, 6)
    arrow(draw, (515, 422), (815, 485), CORAL, 6)
    draw.rounded_rectangle((230, 715, 900, 825), radius=24, fill=CORAL_LIGHT)
    draw.ellipse((265, 744, 323, 802), fill=CORAL)
    cross(draw, (294, 773), WHITE, 24, 6)
    text_box(draw, (350, 741), "Duplicate data can distort safety and efficacy results.", font(25, "semibold"), INK, max_width=505)

    # Evidence cards.
    card(image, (1080, 275, 1810, 515), fill=MINT, outline="#B8E2DA", radius=30)
    text_box(draw, (1130, 315), "48%", font(76, "bold"), TEAL_DARK)
    text_box(draw, (1390, 325), "hid participation\nin another study", font(29, "semibold"), NAVY, line_gap=5)
    text_box(draw, (1135, 455), "Peer-reviewed participant deception study", font(18), SLATE)
    card(image, (1080, 550, 1810, 790), fill=AMBER_LIGHT, outline="#F3D4AC", radius=30)
    text_box(draw, (1130, 590), "22%", font(76, "bold"), "#A75B09")
    text_box(draw, (1390, 600), "used changed identity details\nto enroll again", font(29, "semibold"), NAVY, line_gap=5)
    text_box(draw, (1135, 730), "Among contaminated studies in the same research", font(18), SLATE)
    draw.rounded_rectangle((1080, 825, 1810, 915), radius=22, fill=NAVY)
    text_box(draw, (1445, 870), "Sharing raw IDs may stop duplicates — but creates a privacy risk.", font(24, "semibold"), WHITE, anchor="mm", align="center")
    text_box(draw, (112, 972), "Source: NIH / PMC7876535", font(17), SLATE)
    return image


def make_slide_3() -> Image.Image:
    image, draw = base_slide(3, "The solution", "TrialFence proves uniqueness — not identity")
    draw.rounded_rectangle((560, 260, 1360, 332), radius=34, fill=MINT)
    text_box(draw, (960, 296), "An anonymous proof says “I am eligible” without revealing who I am.", font(25, "semibold"), TEAL_DARK, anchor="mm", align="center")

    card(image, (105, 375, 830, 820), fill=WHITE, radius=30)
    text_box(draw, (155, 415), "Traditional approach", font(29, "bold"), NAVY)
    text_box(draw, (155, 465), "Share a central identity database", font(24), SLATE)
    database_icon(draw, (465, 615), CORAL, 1.0)
    clinic_icon(draw, (185, 565, 300, 690), CORAL)
    clinic_icon(draw, (630, 565, 745, 690), CORAL)
    arrow(draw, (310, 625), (390, 625), CORAL, 6)
    arrow(draw, (620, 625), (540, 625), CORAL, 6)
    pill(draw, (205, 742, 730, 798), "RAW ID DATA MOVES BETWEEN SITES", CORAL_LIGHT, "#A93E2B")

    arrow(draw, (875, 600), (1038, 600), TEAL, 10)
    card(image, (1090, 375, 1815, 820), fill=WHITE, radius=30)
    text_box(draw, (1140, 415), "TrialFence", font(29, "bold"), TEAL_DARK)
    text_box(draw, (1140, 465), "Share only an anonymous proof", font(24), SLATE)
    clinic_icon(draw, (1160, 565, 1275, 690), TEAL)
    clinic_icon(draw, (1635, 565, 1750, 690), BLUE)
    shield_icon(draw, (1450, 622), TEAL, 0.85)
    arrow(draw, (1290, 625), (1370, 625), TEAL, 6)
    arrow(draw, (1618, 625), (1535, 625), BLUE, 6)
    pill(draw, (1190, 742, 1715, 798), "NO CROSS-SITE ID DATABASE", MINT, TEAL_DARK)

    outcomes = [
        (110, "01", "No PII on-chain"),
        (675, "02", "Repeat enrollment blocked"),
        (1240, "03", "Different studies stay unlinkable"),
    ]
    for x, n, label in outcomes:
        draw.ellipse((x, 868, x + 56, 924), fill=TEAL)
        text_box(draw, (x + 28, 896), n, font(17, "bold"), WHITE, anchor="mm", align="center")
        text_box(draw, (x + 76, 878), label, font(24, "semibold"), NAVY, max_width=420)
    return image


def make_slide_4() -> Image.Image:
    image, draw = base_slide(4, "How it works", "Four simple steps from ID check to safe enrollment")
    steps = [
        ("1", "Create a private credential", "Generated in the volunteer’s browser.", TEAL),
        ("2", "Check identity once", "A coordinator adds an anonymous commitment.", BLUE),
        ("3", "Build a proof", "Bound to one study protocol and one site.", AMBER),
        ("4", "Gate enrollment", "Accept the first proof. Block the repeat.", GREEN),
    ]
    xs = [105, 555, 1005, 1455]
    for i, ((n, title, body, color), x) in enumerate(zip(steps, xs)):
        card(image, (x, 305, x + 360, 700), fill=WHITE, radius=28)
        draw.ellipse((x + 32, 335, x + 102, 405), fill=color)
        text_box(draw, (x + 67, 370), n, font(29, "bold"), WHITE, anchor="mm", align="center")
        if i == 0:
            person_icon(draw, (x + 180, 500), NAVY, 0.72)
        elif i == 1:
            clinic_icon(draw, (x + 125, 448, x + 235, 570), BLUE)
        elif i == 2:
            shield_icon(draw, (x + 180, 505), AMBER, 0.72)
        else:
            draw.ellipse((x + 125, 448, x + 235, 558), fill=GREEN)
            check(draw, (x + 180, 503), WHITE, 52, 9)
        text_box(draw, (x + 28, 585), title, font(26, "bold"), NAVY, max_width=304, line_gap=4)
        text_box(draw, (x + 28, 650), body, font(20), SLATE, max_width=304, line_gap=5)
        if i < 3:
            arrow(draw, (x + 375, 500), (x + 430, 500), TEAL, 7)

    draw.rounded_rectangle((105, 770, 1815, 925), radius=28, fill=NAVY)
    text_box(draw, (155, 805), "STAYS PRIVATE", font(18, "bold"), "#8BE0D5")
    text_box(draw, (155, 845), "Name · document number · browser identity secret", font(27, "semibold"), WHITE)
    draw.line((960, 795, 960, 900), fill="#55717D", width=2)
    text_box(draw, (1010, 805), "SHARED WITH THE GATE", font(18, "bold"), "#A9CDE0")
    text_box(draw, (1010, 845), "Anonymous commitment · proof · one-time private tag", font(27, "semibold"), WHITE)
    return image


def make_slide_5() -> Image.Image:
    image, draw = base_slide(5, "Working product", "A working prototype — not a mock-up")
    # Product UI mock.
    inner = browser_chrome(draw, (95, 255, 1185, 920))
    x1, y1, x2, y2 = inner
    draw.rectangle((x1, y1, x1 + 235, y2), fill="#15323E")
    draw.rounded_rectangle((x1 + 30, y1 + 32, x1 + 85, y1 + 87), radius=14, fill=TEAL)
    text_box(draw, (x1 + 58, y1 + 60), "TF", font(18, "bold"), WHITE, anchor="mm", align="center")
    text_box(draw, (x1 + 103, y1 + 42), "TrialFence", font(22, "bold"), WHITE)
    for i, (label, active) in enumerate((("Volunteer", False), ("Coordinator", False), ("Sponsor / IRB", True))):
        y = y1 + 155 + i * 72
        if active:
            draw.rounded_rectangle((x1 + 20, y - 12, x1 + 215, y + 45), radius=16, fill="#214C5B")
        text_box(draw, (x1 + 48, y), label, font(19, "semibold"), "#AEE5DE" if active else "#D2E0E5")

    mx = x1 + 275
    text_box(draw, (mx, y1 + 35), "Sponsor / IRB dashboard", font(28, "bold"), NAVY)
    text_box(draw, (mx, y1 + 78), "Live on-chain events and signed denials", font(19), SLATE)
    stats = [("2", "ACCEPTED", GREEN_LIGHT, GREEN), ("1", "DUPLICATE", CORAL_LIGHT, CORAL), ("2", "PROTOCOLS", MINT, TEAL_DARK)]
    for i, (num, label, bg, color) in enumerate(stats):
        sx = mx + i * 245
        draw.rounded_rectangle((sx, y1 + 130, sx + 215, y1 + 245), radius=18, fill=bg)
        text_box(draw, (sx + 22, y1 + 150), num, font(44, "bold"), color)
        text_box(draw, (sx + 22, y1 + 207), label, font(16, "bold"), SLATE)
    text_box(draw, (mx, y1 + 300), "ENROLLMENT AUDIT", font(17, "bold"), SLATE)
    rows = [
        ("Protocol A", "Site North", "ACCEPTED", GREEN_LIGHT, GREEN),
        ("Protocol A", "Site South", "DUPLICATE", CORAL_LIGHT, CORAL),
        ("Protocol B", "Site East", "ACCEPTED", GREEN_LIGHT, GREEN),
    ]
    for i, (protocol, site, status, bg, color) in enumerate(rows):
        y = y1 + 340 + i * 76
        draw.rounded_rectangle((mx, y, x2 - 35, y + 58), radius=14, fill="#F7FAFB", outline=LINE, width=1)
        text_box(draw, (mx + 22, y + 16), protocol, font(19, "semibold"), NAVY)
        text_box(draw, (mx + 210, y + 16), site, font(18), SLATE)
        draw.rounded_rectangle((x2 - 230, y + 10, x2 - 55, y + 48), radius=18, fill=bg)
        text_box(draw, (x2 - 142, y + 29), status, font(15, "bold"), color, anchor="mm", align="center")

    text_box(draw, (1245, 265), "Live browser result", font(27, "bold"), NAVY)
    result_cards = [
        (GREEN, GREEN_LIGHT, "1", "First site", "ACCEPTED"),
        (CORAL, CORAL_LIGHT, "2", "Second site", "DUPLICATE BLOCKED"),
        (TEAL, MINT, "3", "New protocol", "ACCEPTED · UNLINKABLE"),
    ]
    for i, (color, bg, n, title, result) in enumerate(result_cards):
        y = 330 + i * 185
        card(image, (1235, y, 1815, y + 150), fill=bg, outline=bg, radius=24, shadow=False)
        draw.ellipse((1270, y + 36, 1345, y + 111), fill=color)
        if i == 1:
            cross(draw, (1308, y + 73), WHITE, 27, 6)
        else:
            check(draw, (1308, y + 73), WHITE, 34, 7)
        text_box(draw, (1380, y + 28), title, font(23, "semibold"), NAVY)
        text_box(draw, (1380, y + 78), result, font(21, "bold"), color)

    for i, (value, label) in enumerate((("22", "CONTRACT TESTS"), ("100%", "INTEGRATION FLOW"), ("LIVE", "ON RENDER"))):
        x = 1240 + i * 190
        text_box(draw, (x, 895), value, font(28, "bold"), TEAL_DARK)
        text_box(draw, (x, 935), label, font(13, "bold"), SLATE)
    return image


def make_slide_6() -> Image.Image:
    image, draw = base_slide(6, "Impact & pilot", "Built for the people who protect trial integrity", "One privacy-safe rule can support every participating site.")
    audiences = [
        (105, TEAL, "SPONSORS & CROs", "Cleaner safety and efficacy data", "Reduce duplicate enrollment without creating a shared patient database."),
        (665, BLUE, "RESEARCH SITES", "A simple accept-or-block answer", "Sites keep their own identity workflow and share only anonymous proof."),
        (1225, AMBER, "IRBs & REGULATORS", "An audit trail without raw identity", "Review accepted events and signed denials without cross-site PII."),
    ]
    for x, color, label, title, body in audiences:
        card(image, (x, 325, x + 500, 725), fill=WHITE, radius=30)
        draw.ellipse((x + 38, 365, x + 120, 447), fill=color)
        if label.startswith("SPONSORS"):
            shield_icon(draw, (x + 79, 405), WHITE, 0.35)
        elif label.startswith("RESEARCH"):
            clinic_icon(draw, (x + 50, 380, x + 108, 434), WHITE)
        else:
            draw.rectangle((x + 60, 382, x + 100, 425), outline=WHITE, width=5)
            check(draw, (x + 80, 405), WHITE, 22, 5)
        text_box(draw, (x + 145, 378), label, font(18, "bold"), color)
        text_box(draw, (x + 38, 485), title, font(29, "bold"), NAVY, max_width=420, line_gap=5)
        text_box(draw, (x + 38, 590), body, font(22), SLATE, max_width=420, line_gap=7)

    draw.rounded_rectangle((105, 790, 1815, 930), radius=30, fill=NAVY)
    text_box(draw, (155, 820), "BEST NEXT STEP", font(17, "bold"), "#8BE0D5")
    text_box(draw, (155, 858), "Pilot one protocol across 2–3 sites", font(34, "bold"), WHITE)
    pill(draw, (980, 825, 1200, 885), "PRIVACY REVIEW", "#274B59", WHITE)
    pill(draw, (1220, 825, 1445, 885), "SITE ONBOARDING", "#274B59", WHITE)
    pill(draw, (1465, 825, 1745, 885), "TESTNET VALIDATION", "#274B59", WHITE)
    return image


def make_qr(url: str, path: Path, color: str = NAVY) -> Image.Image:
    qr = qrcode.QRCode(version=None, error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=12, border=2)
    qr.add_data(url)
    qr.make(fit=True)
    image = qr.make_image(fill_color=color, back_color="white").convert("RGB")
    image.save(path)
    return image


def make_slide_7() -> Image.Image:
    image = gradient_bg("#F8FCFC", "#DFF3EF")
    add_soft_blobs(image)
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((105, 72, 163, 130), radius=15, fill=TEAL)
    text_box(draw, (134, 101), "TF", font(21, "bold"), WHITE, anchor="mm", align="center")
    text_box(draw, (185, 84), "TRIALFENCE", font(23, "bold"), NAVY)
    text_box(draw, (185, 114), "LIVE DEMO", font(17, "semibold"), TEAL_DARK)

    text_box(draw, (110, 205), "Ready to prove it live", font(66, "bold"), NAVY)
    text_box(draw, (115, 300), "One identity. One protocol. One accepted enrollment.", font(29, "semibold"), SLATE)

    demo_steps = [
        ("1", "SET UP", "Register one study protocol and two sites."),
        ("2", "VERIFY", "Check a fictional ID and issue anonymous membership."),
        ("3", "PROVE", "Accept Site A · block Site B · accept a new protocol."),
    ]
    for i, (n, title, body) in enumerate(demo_steps):
        y = 405 + i * 155
        draw.ellipse((120, y, 200, y + 80), fill=TEAL if i < 2 else NAVY)
        text_box(draw, (160, y + 40), n, font(30, "bold"), WHITE, anchor="mm", align="center")
        text_box(draw, (240, y + 5), title, font(18, "bold"), TEAL_DARK)
        text_box(draw, (240, y + 42), body, font(27, "semibold"), NAVY, max_width=750, line_gap=5)

    qr_live_path = TEMP / "qr-live.png"
    qr_repo_path = TEMP / "qr-repo.png"
    qr_live = make_qr(LIVE_URL, qr_live_path, NAVY)
    qr_repo = make_qr(REPO_URL, qr_repo_path, NAVY)
    qr_live = qr_live.resize((250, 250), Image.Resampling.LANCZOS)
    qr_repo = qr_repo.resize((250, 250), Image.Resampling.LANCZOS)

    for x, label, url, qr_img, accent in (
        (1120, "LIVE DEMO", "trialfence-frontend.onrender.com", qr_live, TEAL),
        (1490, "GITHUB", "github.com/alwaysalearner1234/TrialFence", qr_repo, BLUE),
    ):
        card(image, (x, 285, x + 320, 730), fill=WHITE, radius=30)
        image.paste(qr_img, (x + 35, 330))
        draw.rounded_rectangle((x + 35, 600, x + 285, 655), radius=25, fill=accent)
        text_box(draw, (x + 160, 628), label, font(18, "bold"), WHITE, anchor="mm", align="center")
        text_box(draw, (x + 160, 690), url, font(14, "semibold"), SLATE, max_width=270, line_gap=3, anchor="mm", align="center")

    draw.rounded_rectangle((105, 860, 1815, 965), radius=28, fill=NAVY)
    text_box(draw, (960, 912), "Protect trial integrity — without building a participant identity database.", font(30, "bold"), WHITE, anchor="mm", align="center")
    text_box(draw, (110, 1005), "Lidiya · Syed Saad · Krishna", font(18, "semibold"), NAVY)
    text_box(draw, (1810, 1005), "COMPSPHERE / HACKSPHERE 2026 · PHASE 2", font(16, "semibold"), SLATE, align="right")
    return image


def make_contact_sheet(slides: Iterable[Image.Image], path: Path) -> None:
    slides = list(slides)
    thumb_w, thumb_h = 768, 432
    gap = 28
    rows = (len(slides) + 1) // 2
    sheet = Image.new("RGB", (thumb_w * 2 + gap * 3, rows * thumb_h + gap * (rows + 1)), "#CBD6DA")
    for i, slide in enumerate(slides):
        thumb = slide.resize((thumb_w, thumb_h), Image.Resampling.LANCZOS)
        x = gap + (i % 2) * (thumb_w + gap)
        y = gap + (i // 2) * (thumb_h + gap)
        sheet.paste(thumb, (x, y))
    sheet.save(path, quality=95)


def save_pptx(slide_paths: list[Path]) -> None:
    prs = Presentation()
    prs.slide_width = Inches(13.333333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    for path in slide_paths:
        slide = prs.slides.add_slide(blank)
        slide.shapes.add_picture(str(path), 0, 0, width=prs.slide_width, height=prs.slide_height)
    # Remove the default first slide if a template ever adds one (the default
    # Presentation currently starts empty, so this is normally a no-op).
    while len(prs.slides) > len(slide_paths):
        slide_id = prs.slides._sldIdLst[0]
        prs.part.drop_rel(slide_id.rId)
        del prs.slides._sldIdLst[0]

    # Transparent clickable areas over the two QR cards on slide 7.
    last = prs.slides[-1]
    for x, url in ((1120, LIVE_URL), (1490, REPO_URL)):
        shape = last.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x / 144), Inches(285 / 144), Inches(320 / 144), Inches(445 / 144))
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
        shape.fill.transparency = 100
        shape.line.fill.background()
        shape.click_action.hyperlink.address = url

    prs.core_properties.title = "TrialFence — Privacy-Preserving Clinical Trial Enrollment"
    prs.core_properties.subject = "CompSphere / HackSphere 2026 Phase 2 live pitch"
    prs.core_properties.author = "TrialFence — Lidiya, Syed Saad, Krishna"
    prs.core_properties.keywords = "clinical trials, privacy, zero knowledge, duplicate enrollment"
    prs.save(OUT_PPTX)


def save_pdf(slide_paths: list[Path]) -> None:
    page_w, page_h = 960, 540  # 16:9 in PDF points
    pdf = canvas.Canvas(str(OUT_PDF), pagesize=(page_w, page_h), pageCompression=1)
    pdf.setTitle("TrialFence Phase 2 Pitch Deck")
    pdf.setAuthor("TrialFence — Lidiya, Syed Saad, Krishna")
    for i, path in enumerate(slide_paths):
        pdf.drawImage(ImageReader(str(path)), 0, 0, page_w, page_h, preserveAspectRatio=True, anchor="c")
        if i == 6:
            pdf.linkURL(LIVE_URL, (560, 175, 720, 365), relative=0)
            pdf.linkURL(REPO_URL, (745, 175, 905, 365), relative=0)
        pdf.showPage()
    pdf.save()


NOTES = """# TrialFence Pitch Deck — Speaker Notes

Keep the deck on screen; do not read it word-for-word. Aim for 4–5 minutes before the live demo.

## Slide 1 — One trial. One anonymous enrollment. (25 seconds)
TrialFence stops the same person from joining the same clinical trial protocol at multiple sites, without creating a shared identity database. The volunteer proves eligibility anonymously; the shared gate accepts the first enrollment and blocks the repeat.

## Slide 2 — The problem (35 seconds)
Clinical trials need clean enrollment data, but sites do not want to exchange raw participant IDs. Research shows that hidden participation and changed identity details are real sources of study contamination. Today, teams often choose between weak duplicate checks and risky data sharing.

## Slide 3 — The solution (35 seconds)
TrialFence changes what sites share. The coordinator still checks identity locally, but sites share only an anonymous cryptographic proof. The proof shows that the volunteer is eligible and has not already enrolled in this protocol. It does not reveal the volunteer’s identity.

## Slide 4 — How it works (45 seconds)
The browser creates a private credential. A coordinator verifies a fictional demo ID and adds only an anonymous commitment. The browser then makes a proof tied to the selected protocol and site. The smart contract accepts the first proof and blocks the same protocol-specific private tag at another site.

## Slide 5 — Working product (40 seconds)
This is a running full-stack prototype with volunteer, coordinator, and sponsor workspaces. In the tested flow, Site North was accepted, Site South was blocked as a duplicate with a signed denial record, and the same identity was accepted in a different protocol with an unrelated tag. The repository includes 22 contract tests and a full integration suite.

## Slide 6 — Impact and pilot (30 seconds)
Sponsors get cleaner trial data, sites get an instant accept-or-block answer, and IRBs get an audit trail without raw cross-site identity. The next practical step is a small pilot: one protocol, two or three sites, privacy review, onboarding, and testnet validation.

## Slide 7 — Live demo and close (20 seconds, then demo)
We will now register a protocol and two sites, issue anonymous membership, accept the first enrollment, block the duplicate, and show that a new protocol remains unlinkable. Close with: “TrialFence protects trial integrity without building a participant identity database.”
"""


def main() -> None:
    makers = [make_slide_1, make_slide_2, make_slide_3, make_slide_4, make_slide_5, make_slide_6, make_slide_7]
    slides = [maker() for maker in makers]
    slide_paths: list[Path] = []
    for i, slide in enumerate(slides, 1):
        path = SLIDES_DIR / f"slide-{i:02d}.png"
        slide.save(path, optimize=True)
        slide_paths.append(path)
    save_pptx(slide_paths)
    save_pdf(slide_paths)
    OUT_NOTES.write_text(NOTES, encoding="utf-8")
    preview = TEMP / "TrialFence_Pitch_Deck_Contact_Sheet.png"
    make_contact_sheet(slides, preview)
    print(f"PPTX: {OUT_PPTX}")
    print(f"PDF:  {OUT_PDF}")
    print(f"Notes:{OUT_NOTES}")
    print(f"Preview: {preview}")


if __name__ == "__main__":
    main()
