#!/usr/bin/env python3
"""
build_og.py -- 링크 미리보기 카드(assets/img/og-card.jpg)를 만든다.

카카오톡·슬랙·트위터·링크드인에 홈페이지 주소를 붙였을 때 뜨는 그림이다.
없으면 회색 네모만 나오고, 있으면 사이트 헤더와 같은 모습이 나온다.

크기는 1200 x 630 (Open Graph 표준 비율 1.91:1).

**로컬(Windows/WSL)에서만 돌린다.** Times New Roman 이 필요하기 때문이다.
GitHub Actions(리눅스)에는 그 폰트가 없으므로, 여기서 만든 jpg 를 저장소에 함께 올린다.
build.py 가 자동으로 부르되, 폰트가 없으면 조용히 건너뛴다.

    python build_og.py        # 카드만 다시 만든다
"""
from pathlib import Path

from PIL import Image, ImageDraw, ImageEnhance, ImageFilter

ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "assets"
OUT = ASSETS / "img" / "og-card.jpg"

W, H = 1200, 630

INK_DEEP = (23, 20, 36)       # #171424  짙은 남보라
BURNT = (191, 87, 0)          # #bf5700  UT 주황
WHITE = (255, 255, 255)

# 윈도우 기본 폰트 위치. WSL 에서도 이 경로로 읽힌다.
FONT_DIR = Path("/mnt/c/Windows/Fonts")
F_TIMES_ITALIC = FONT_DIR / "timesi.ttf"
F_ARIAL = FONT_DIR / "arial.ttf"
F_ARIAL_BOLD = FONT_DIR / "arialbd.ttf"


def fonts_available() -> bool:
    return all(f.exists() for f in (F_TIMES_ITALIC, F_ARIAL, F_ARIAL_BOLD))


def load(path: Path, size: int):
    from PIL import ImageFont
    return ImageFont.truetype(str(path), size)


def draw_tracked(draw, xy, text, font, fill, tracking=0.0, anchor_center=None):
    """
    글자 사이를 벌려서(letter-spacing) 그린다. Pillow 에는 이 기능이 없어서 직접 만든다.
    tracking 은 글자 크기에 대한 비율이다 (0.1 이면 글자 높이의 10% 만큼 벌린다).
    anchor_center 에 x 를 주면 그 x 를 중심으로 가운데 맞춘다.
    """
    gap = font.size * tracking
    widths = [draw.textlength(ch, font=font) for ch in text]
    total = sum(widths) + gap * max(len(text) - 1, 0)

    x, y = xy
    if anchor_center is not None:
        x = anchor_center - total / 2
    for ch, w in zip(text, widths):
        draw.text((x, y), ch, font=font, fill=fill)
        x += w + gap
    return total


def make_background() -> Image.Image:
    """사이트 헤더와 같은 배경. header.jpg 를 채우고 짙은 남보라를 덮는다."""
    bg = Image.new("RGB", (W, H), INK_DEEP)
    src = ASSETS / "img" / "backgrounds" / "header.jpg"
    if src.exists():
        im = Image.open(src).convert("RGB")
        # cover: 짧은 쪽을 기준으로 키운 뒤 가운데를 잘라낸다
        scale = max(W / im.width, H / im.height)
        im = im.resize((round(im.width * scale), round(im.height * scale)), Image.LANCZOS)
        left = (im.width - W) // 2
        top = (im.height - H) // 2
        im = im.crop((left, top, left + W, top + H))
        im = ImageEnhance.Brightness(im).enhance(0.55)
        # 배경 무늬가 글자와 경쟁하지 않게 살짝 흐린다
        im = im.filter(ImageFilter.GaussianBlur(1.2))
        bg = Image.blend(bg, im, 0.62)
    return bg


# portrait-ut.jpg (1440 x 1112) 안에서 사람이 있는 자리.
# 사진 가운데가 아니라 오른쪽에 앉아 있어서, 가운데로 자르면 몸이 잘린다.
# (중심 x, 중심 y, 한 변) 단위는 원본 픽셀. 사진을 바꾸면 이 값도 다시 잡는다.
PORTRAIT_BOX = (905, 560, 540)


def circular_portrait(size: int) -> Image.Image | None:
    """초상 사진을 정사각으로 자르고 흰 테두리를 두른 원형으로 만든다."""
    src = ASSETS / "img" / "people" / "portrait-ut.jpg"
    if not src.exists():
        return None
    im = Image.open(src).convert("RGB")

    cx, cy, side = PORTRAIT_BOX
    side = min(side, im.width, im.height)
    # 자르는 네모가 사진 밖으로 나가지 않게 안쪽으로 밀어 넣는다
    left = min(max(cx - side // 2, 0), im.width - side)
    top = min(max(cy - side // 2, 0), im.height - side)
    im = im.crop((left, top, left + side, top + side))

    ss = 4  # 테두리를 매끄럽게 하려고 4배로 그린 뒤 줄인다
    big = size * ss
    im = im.resize((big, big), Image.LANCZOS)

    mask = Image.new("L", (big, big), 0)
    ImageDraw.Draw(mask).ellipse((0, 0, big - 1, big - 1), fill=255)

    ring = Image.new("RGBA", (big, big), (0, 0, 0, 0))
    ring.paste(im, (0, 0), mask)
    d = ImageDraw.Draw(ring)
    d.ellipse((0, 0, big - 1, big - 1), outline=WHITE + (235,), width=ss * 3)

    return ring.resize((size, size), Image.LANCZOS)


def build():
    if not fonts_available():
        return False

    card = make_background()
    d = ImageDraw.Draw(card)

    # ---- 오른쪽: 초상 사진 --------------------------------------------------
    p_size = 310
    p_cx = W - 60 - p_size // 2          # 오른쪽 여백 60px
    portrait = circular_portrait(p_size)
    if portrait is not None:
        card.paste(portrait, (p_cx - p_size // 2, H // 2 - p_size // 2 - 18), portrait)

    # ---- 왼쪽: 이름과 소속 ---------------------------------------------------
    left = 72
    text_right = p_cx - p_size // 2 - 48      # 사진에 닿지 않는 선
    box_w = text_right - left

    f_name = load(F_TIMES_ITALIC, 84)
    f_role = load(F_ARIAL_BOLD, 26)
    f_line = load(F_ARIAL, 24)

    y = 190
    d.text((left, y), "Hyunje Yang", font=f_name, fill=WHITE)
    y += 108

    # 이름 아래 주황색 짧은 선
    d.rectangle((left, y, left + 96, y + 5), fill=BURNT)
    y += 34

    draw_tracked(d, (left, y), "PhD Candidate", f_role, WHITE, tracking=0.09)
    y += 42
    draw_tracked(d, (left, y), "The University of Texas at Austin", f_line,
                 (222, 219, 231), tracking=0.06)
    y += 38

    f_topic = load(F_ARIAL, 22)
    d.text((left, y), "Machine learning for coastal flood prediction",
           font=f_topic, fill=(190, 186, 206))

    # ---- 아래: 주소 띠 -------------------------------------------------------
    bar_h = 62
    strip = Image.new("RGB", (W, bar_h), INK_DEEP)
    card.paste(strip, (0, H - bar_h))
    d = ImageDraw.Draw(card)
    d.rectangle((0, H - bar_h, W, H - bar_h + 3), fill=BURNT)

    f_url = load(F_ARIAL_BOLD, 22)
    ty = H - bar_h + (bar_h - 22) // 2 - 3
    draw_tracked(d, (0, ty), "HYUNJEYANG.COM", f_url, WHITE,
                 tracking=0.16, anchor_center=W / 2)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    card.save(OUT, "JPEG", quality=88, optimize=True, progressive=True)
    return True


def main():
    if build():
        kb = OUT.stat().st_size / 1024
        print(f"  만듦  assets/img/og-card.jpg  ({W}x{H}, {kb:.0f} KB)")
    else:
        print("  건너뜀  og-card.jpg (Times New Roman 폰트를 찾지 못했습니다)")


if __name__ == "__main__":
    main()
