from __future__ import annotations

import json
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont, ImageFilter


ROOT = Path(__file__).resolve().parent.parent
DATA_FILE = ROOT / "data" / "guida_multimedia.json"
STATIC_DIR = ROOT / "static" / "guide-media"
IMAGES_DIR = STATIC_DIR / "images"
LOGO_FILE = ROOT / "static" / "logo-ca.ico"

CANVAS = (1600, 900)
BACKGROUND_TOP = (255, 246, 238)
BACKGROUND_BOTTOM = (245, 233, 219)
ORANGE = (214, 122, 51)
ORANGE_SOFT = (238, 184, 144)
INK = (70, 42, 20)
MUTED = (125, 90, 62)
CARD = (255, 252, 248)


def load_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    candidates = [
        "C:/Windows/Fonts/seguiemj.ttf",
        "C:/Windows/Fonts/segoeuib.ttf" if bold else "C:/Windows/Fonts/segoeui.ttf",
        "C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf",
    ]
    for candidate in candidates:
        path = Path(candidate)
        if path.exists():
            return ImageFont.truetype(str(path), size=size)
    return ImageFont.load_default()


def rounded_box(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], fill, outline=None, radius: int = 28, width: int = 2):
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def wrap_text(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.FreeTypeFont, max_width: int) -> list[str]:
    words = text.split()
    if not words:
        return [""]
    lines: list[str] = []
    current = words[0]
    for word in words[1:]:
        attempt = f"{current} {word}"
        bbox = draw.textbbox((0, 0), attempt, font=font)
        if bbox[2] - bbox[0] <= max_width:
            current = attempt
        else:
            lines.append(current)
            current = word
    lines.append(current)
    return lines


def add_text_block(draw: ImageDraw.ImageDraw, position: tuple[int, int], text: str, font, fill, max_width: int, line_gap: int = 8) -> int:
    x, y = position
    lines = wrap_text(draw, text, font, max_width)
    cursor_y = y
    for line in lines:
        draw.text((x, cursor_y), line, font=font, fill=fill)
        bbox = draw.textbbox((x, cursor_y), line, font=font)
        cursor_y += (bbox[3] - bbox[1]) + line_gap
    return cursor_y


def make_background() -> Image.Image:
    image = Image.new("RGB", CANVAS, BACKGROUND_TOP)
    draw = ImageDraw.Draw(image)
    width, height = CANVAS
    for y in range(height):
        ratio = y / max(height - 1, 1)
        color = tuple(
            int(BACKGROUND_TOP[i] * (1 - ratio) + BACKGROUND_BOTTOM[i] * ratio)
            for i in range(3)
        )
        draw.line((0, y, width, y), fill=color)
    draw.ellipse((width - 460, -120, width + 120, 380), fill=(255, 233, 215))
    draw.ellipse((-180, height - 320, 360, height + 180), fill=(250, 228, 208))
    return image.filter(ImageFilter.GaussianBlur(radius=0.5))


def load_logo() -> Image.Image | None:
    if not LOGO_FILE.exists():
        return None
    logo = Image.open(LOGO_FILE).convert("RGBA")
    logo.thumbnail((180, 180))
    return logo


def render_slide(course_key: str, slide_index: int, slide: dict, total_slides: int, logo: Image.Image | None) -> None:
    image = make_background()
    draw = ImageDraw.Draw(image)
    width, height = CANVAS

    title_font = load_font(42, bold=True)
    subtitle_font = load_font(24, bold=False)
    bullet_font = load_font(28, bold=False)
    kicker_font = load_font(22, bold=True)
    small_font = load_font(18, bold=False)

    rounded_box(draw, (70, 70, width - 70, height - 70), fill=(255, 255, 255), outline=(245, 222, 202), radius=36, width=3)
    rounded_box(draw, (70, 70, width - 70, 190), fill=(255, 245, 236), outline=None, radius=36)
    draw.rounded_rectangle((92, 92, 280, 142), radius=22, fill=ORANGE)
    draw.text((118, 104), "GUIDA IN LINEA", font=kicker_font, fill=(255, 255, 255))

    draw.text((330, 98), slide["title"], font=title_font, fill=INK)
    draw.text((332, 152), slide["subtitle"], font=subtitle_font, fill=MUTED)

    if logo:
        image.paste(logo, (width - 250, 96), logo)

    left_panel = (96, 230, 920, 784)
    right_panel = (962, 230, width - 96, 784)
    rounded_box(draw, left_panel, fill=CARD, outline=(242, 225, 210), radius=28, width=2)
    rounded_box(draw, right_panel, fill=(255, 247, 241), outline=(242, 225, 210), radius=28, width=2)

    draw.text((126, 258), "Cosa puoi fare", font=load_font(30, bold=True), fill=INK)
    y = 320
    for bullet in slide["bullets"]:
        draw.ellipse((132, y + 8, 150, y + 26), fill=ORANGE)
        y = add_text_block(draw, (172, y), bullet, bullet_font, INK, 680, line_gap=8) + 16

    draw.text((992, 258), "Contesto operativo", font=load_font(30, bold=True), fill=INK)
    context_lines = [
        "Interfaccia pensata per operatori e segreteria",
        "Flussi coerenti con tesseramento, iscrizioni e pagamenti",
        "Dati, ricevute e report sempre collegati fra loro"
    ]
    y = 320
    for line in context_lines:
        y = add_text_block(draw, (992, y), line, load_font(24), MUTED, 430, line_gap=6) + 18

    rounded_box(draw, (992, 520, width - 128, 710), fill=(255, 255, 255), outline=(244, 225, 206), radius=24, width=2)
    draw.text((1024, 548), "Focus slide", font=load_font(28, bold=True), fill=INK)
    add_text_block(draw, (1024, 602), slide["audio_text"], load_font(21), MUTED, 360, line_gap=7)

    footer = f"{slide_index + 1}/{total_slides}  •  Oratorio Carlo Acutis"
    draw.text((126, height - 122), footer, font=small_font, fill=MUTED)
    draw.text((width - 350, height - 122), course_key.replace('_', ' ').title(), font=small_font, fill=ORANGE)

    output_dir = IMAGES_DIR / course_key
    output_dir.mkdir(parents=True, exist_ok=True)
    image.save(output_dir / f"{slide_index + 1:02d}.png", quality=95)


def main() -> None:
    payload = json.loads(DATA_FILE.read_text(encoding="utf-8"))
    logo = load_logo()
    for course_key, course in payload.items():
        slides = course.get("slides", [])
        for index, slide in enumerate(slides):
            render_slide(course_key, index, slide, len(slides), logo)


if __name__ == "__main__":
    main()
