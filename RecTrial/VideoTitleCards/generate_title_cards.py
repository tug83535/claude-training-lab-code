"""
Video Title Card Generator
==========================
Regenerates all 4 video title cards with consistent iPipeline branding.

Output: C:\\Users\\connor.atlee\\RecTrial\\VideoTitleCards_v2\\
  V1_Title_Card.png
  V2_Title_Card.png
  V3_Title_Card.png
  V4_Title_Card.png

Run: python generate_title_cards.py
"""

from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

# ---------- OUTPUT ----------
OUT_DIR = Path(r"C:\Users\connor.atlee\RecTrial\VideoTitleCards_v2")
OUT_DIR.mkdir(parents=True, exist_ok=True)

# ---------- DIMENSIONS ----------
W, H = 1920, 1080

# ---------- COLORS (iPipeline brand — matches originals) ----------
BG_MAIN         = (17,  46,  81)   # #112E51 iPipeline Navy — main background
BG_PANEL_INNER  = (11,  33,  60)   # Darker blue for angled right panel
BG_PANEL_OUTER  = (14,  40,  72)   # Slightly lighter panel farther right
AQUA            = (43,  204, 211)  # #2BCCD3 — bright accent / outline
INNOVATION_BLUE = (75,  155, 203)  # #4B9BCB — subtitle + pill
PILL_BORDER     = (55,  115, 165)  # A touch darker than Innovation Blue
ARCTIC_WHITE    = (249, 249, 249)  # #F9F9F9 — titles, pill text
FOOTER_GRAY     = (130, 148, 168)  # Medium gray for footer
BODY_TEXT       = (225, 225, 225)  # Body paragraph text (disclaimer)

# ---------- FONTS ----------
FONT_BOLD    = r"C:\Windows\Fonts\arialbd.ttf"
FONT_REGULAR = r"C:\Windows\Fonts\arial.ttf"

def F(bold: bool, size: int):
    return ImageFont.truetype(FONT_BOLD if bold else FONT_REGULAR, size)

# ---------- CARD CONTENT ----------
CARDS = [
    {
        "video_num": 1,
        "total": 4,
        "title": "What's Possible",
        "subtitle": "A quick look at what Excel automation can do",
    },
    {
        "video_num": 2,
        "total": 4,
        "title": "Full Demo Walkthrough",
        "subtitle": "End-to-end demo of all 62 automated actions",
    },
    {
        "video_num": 3,
        "total": 4,
        "title": "Universal Tools",
        "subtitle": "VBA tools that work on any Excel file",
    },
    {
        "video_num": 4,
        "total": 4,
        "title": "Python Automation for Finance",
        "subtitle": "Real Python scripts you can run from Command Prompt",
    },
]

# ---------- LAYOUT ----------
MARGIN_LEFT = 100
PILL_X       = 100
PILL_Y       = 275
PILL_W       = 215
PILL_H       = 50
PILL_RADIUS  = 8

TITLE_Y           = 345
TITLE_FONT_SIZE   = 92   # Arial Bold
TITLE_UNDERLINE_GAP = 14 # pixels below title baseline
UNDERLINE_THICKNESS = 4

SUBTITLE_GAP      = 36   # pixels below underline
SUBTITLE_FONT_SIZE = 32   # Arial Bold

FOOTER_Y          = 1035
FOOTER_FONT_SIZE  = 16

# ---------- RIGHT-SIDE DECORATION ----------
# Two prominent aqua-outlined angled panels + bottom-right aqua trapezoid.
# Matches the original iPipeline video title card layout.
def draw_right_side_decoration(draw: ImageDraw.ImageDraw):
    # OUTER panel (rightmost) — lighter navy fill
    draw.polygon([
        (1530, -5),
        (W + 5, -5),
        (W + 5, H + 5),
        (1730, H + 5),
    ], fill=BG_PANEL_OUTER)

    # INNER thin vertical strip (darker) — between the two aqua lines
    draw.polygon([
        (1530, -5),
        (1580, -5),
        (1780, H + 5),
        (1730, H + 5),
    ], fill=BG_PANEL_INNER)

    # Left aqua accent line — thick + bright
    draw.line([(1530, 0), (1730, H)], fill=AQUA, width=5)

    # Right aqua accent line (parallel, shifted right) — thick + bright
    draw.line([(1580, 0), (1780, H)], fill=AQUA, width=5)

    # Top horizontal aqua cap across both panels
    draw.line([(1530, 4), (W, 4)], fill=AQUA, width=4)

    # Bottom-right aqua trapezoid accent — larger + more visible
    draw.polygon([
        (1735, H),
        (W, 820),
        (W, H),
    ], fill=AQUA)

# ---------- PILL ----------
def draw_pill(draw: ImageDraw.ImageDraw, text: str):
    x0, y0, x1, y1 = PILL_X, PILL_Y, PILL_X + PILL_W, PILL_Y + PILL_H
    draw.rounded_rectangle(
        [x0, y0, x1, y1],
        radius=PILL_RADIUS,
        fill=INNOVATION_BLUE,
        outline=PILL_BORDER,
        width=2,
    )
    font = F(bold=True, size=18)
    bbox = draw.textbbox((0, 0), text, font=font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    tx = x0 + (PILL_W - tw) / 2
    ty = y0 + (PILL_H - th) / 2 - 2
    draw.text((tx, ty), text, font=font, fill=ARCTIC_WHITE)

# ---------- TITLE + UNDERLINE + SUBTITLE ----------
def draw_title_block(draw: ImageDraw.ImageDraw, title: str, subtitle: str):
    # Title
    title_font = F(bold=True, size=TITLE_FONT_SIZE)
    draw.text((MARGIN_LEFT, TITLE_Y), title, font=title_font, fill=ARCTIC_WHITE)

    # Measure title for underline width
    bbox = draw.textbbox((MARGIN_LEFT, TITLE_Y), title, font=title_font)
    title_bottom = bbox[3]
    title_right = bbox[2]

    # Underline (width = title width + small extension)
    underline_y = title_bottom + TITLE_UNDERLINE_GAP
    draw.rectangle(
        [MARGIN_LEFT, underline_y, title_right + 10, underline_y + UNDERLINE_THICKNESS],
        fill=AQUA,
    )

    # Subtitle
    subtitle_font = F(bold=True, size=SUBTITLE_FONT_SIZE)
    subtitle_y = underline_y + UNDERLINE_THICKNESS + SUBTITLE_GAP
    draw.text((MARGIN_LEFT, subtitle_y), subtitle, font=subtitle_font, fill=INNOVATION_BLUE)

# ---------- FOOTER ----------
def draw_footer(draw: ImageDraw.ImageDraw):
    font = F(bold=True, size=FOOTER_FONT_SIZE)
    draw.text(
        (MARGIN_LEFT + 10, FOOTER_Y),
        "Excel Automation | iPipeline",
        font=font,
        fill=FOOTER_GRAY,
    )

# ---------- MAIN GENERATE ----------
def generate_card(card: dict) -> Path:
    img = Image.new("RGB", (W, H), BG_MAIN)
    draw = ImageDraw.Draw(img)

    draw_right_side_decoration(draw)

    pill_text = f"VIDEO {card['video_num']} OF {card['total']}"
    draw_pill(draw, pill_text)

    draw_title_block(draw, card["title"], card["subtitle"])

    draw_footer(draw)

    out_path = OUT_DIR / f"V{card['video_num']}_Title_Card.png"
    img.save(out_path, "PNG", optimize=True)
    return out_path


def generate_disclaimer_card() -> Path:
    """Generate the Demonstration Data Notice disclaimer card using the
    same branding as the 4 video title cards."""
    img = Image.new("RGB", (W, H), BG_MAIN)
    draw = ImageDraw.Draw(img)

    draw_right_side_decoration(draw)

    # DISCLAIMER pill (narrower than video-card pill)
    pill_text = "DISCLAIMER"
    disc_pill_w = 150
    disc_pill_h = 40
    disc_pill_x = PILL_X
    disc_pill_y = 170
    draw.rounded_rectangle(
        [disc_pill_x, disc_pill_y, disc_pill_x + disc_pill_w, disc_pill_y + disc_pill_h],
        radius=PILL_RADIUS,
        fill=INNOVATION_BLUE,
        outline=PILL_BORDER,
        width=2,
    )
    pill_font = F(bold=True, size=16)
    bbox = draw.textbbox((0, 0), pill_text, font=pill_font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    draw.text(
        (disc_pill_x + (disc_pill_w - tw) / 2, disc_pill_y + (disc_pill_h - th) / 2 - 2),
        pill_text, font=pill_font, fill=ARCTIC_WHITE,
    )

    # Title
    title_font = F(bold=True, size=72)
    title_x = MARGIN_LEFT
    title_y = 225
    draw.text((title_x, title_y), "Demonstration Data Notice", font=title_font, fill=ARCTIC_WHITE)

    # Measure for underline
    tbbox = draw.textbbox((title_x, title_y), "Demonstration Data Notice", font=title_font)
    title_right = tbbox[2]
    title_bottom = tbbox[3]
    underline_y = title_bottom + 10
    draw.rectangle(
        [MARGIN_LEFT, underline_y, title_right + 10, underline_y + UNDERLINE_THICKNESS],
        fill=AQUA,
    )

    # Subtitle (in Innovation Blue)
    subtitle_font = F(bold=True, size=28)
    subtitle_y = underline_y + UNDERLINE_THICKNESS + 28
    draw.text(
        (MARGIN_LEFT, subtitle_y),
        "All data in this workbook is purely fictitious",
        font=subtitle_font, fill=INNOVATION_BLUE,
    )

    # Body paragraphs
    body_font = F(bold=False, size=22)
    body_y = subtitle_y + 65
    paragraph_1 = [
        "All company names, financial figures, account balances, revenue",
        "numbers, cost allocations, and any other data in this file are",
        "entirely made up for demonstration and training purposes only.",
    ]
    for line in paragraph_1:
        draw.text((MARGIN_LEFT, body_y), line, font=body_font, fill=BODY_TEXT)
        body_y += 32
    body_y += 14
    paragraph_2 = [
        "None of the financials are remotely accurate or based on any",
        "real company's actual financial data.",
    ]
    for line in paragraph_2:
        draw.text((MARGIN_LEFT, body_y), line, font=body_font, fill=BODY_TEXT)
        body_y += 32

    # Divider line
    body_y += 18
    draw.line(
        [(MARGIN_LEFT, body_y), (MARGIN_LEFT + 520, body_y)],
        fill=INNOVATION_BLUE, width=2,
    )

    # Purpose header
    body_y += 30
    purpose_font = F(bold=True, size=26)
    draw.text((MARGIN_LEFT, body_y), "Purpose", font=purpose_font, fill=AQUA)

    # Purpose body
    body_y += 48
    purpose_body = [
        "This file demonstrates VBA macros, SQL queries, and Python",
        "scripts for demo purposes only.",
    ]
    for line in purpose_body:
        draw.text((MARGIN_LEFT, body_y), line, font=body_font, fill=BODY_TEXT)
        body_y += 32

    draw_footer(draw)

    out_path = OUT_DIR / "disclaimer_card.png"
    img.save(out_path, "PNG", optimize=True)
    return out_path


def main():
    print(f"Output folder: {OUT_DIR}\n")
    for card in CARDS:
        path = generate_card(card)
        print(f"  Generated: {path.name}  ({card['title']})")
    disc_path = generate_disclaimer_card()
    print(f"  Generated: {disc_path.name}  (Demonstration Data Notice)")
    print("\nDone.")


if __name__ == "__main__":
    main()
