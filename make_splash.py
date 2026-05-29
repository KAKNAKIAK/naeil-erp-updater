# -*- coding: utf-8 -*-
"""스플래시 이미지 생성기 (앱용 / 설치용). 버전 변경 시 VERSION만 고치고 재실행."""
from PIL import Image, ImageDraw, ImageFont

VERSION = "v1.1.19"
W, H = 480, 270

BG = (21, 22, 28)          # #15161c
CARD = (33, 35, 44)        # #21232c
ACCENT = (59, 130, 246)    # #3b82f6
GREEN = (34, 197, 94)      # #22c55e
FG = (243, 244, 246)       # #f3f4f6
MUTED = (150, 154, 166)    # #969aa6
TROUGH = (14, 15, 20)      # #0e0f14

MALGUN = "C:/Windows/Fonts/malgun.ttf"
MALGUN_BD = "C:/Windows/Fonts/malgunbd.ttf"


def font(path, size):
    return ImageFont.truetype(path, size)


def make(out_path, status_text):
    img = Image.new("RGB", (W, H), BG)
    d = ImageDraw.Draw(img)

    # 상단 액센트 바
    d.rectangle([0, 0, W, 6], fill=ACCENT)

    # 앱 아이콘 썸네일
    try:
        icon = Image.open("app_icon.ico").convert("RGBA")
        icon = icon.resize((72, 72), Image.LANCZOS)
        img.paste(icon, (40, 56), icon)
    except Exception:
        pass

    f_brand = font(MALGUN, 14)
    f_title = font(MALGUN_BD, 26)
    f_ver = font(MALGUN_BD, 13)
    f_status = font(MALGUN, 14)

    tx = 132
    d.text((tx, 60), "Naeil Tour", font=f_brand, fill=MUTED)
    d.text((tx, 82), "ERP 항공요금 업데이트", font=f_title, fill=FG)

    # 버전 배지
    vb = d.textbbox((0, 0), VERSION, font=f_ver)
    vw, vh = vb[2] - vb[0], vb[3] - vb[1]
    bx, by = tx, 126
    d.rounded_rectangle([bx, by, bx + vw + 16, by + vh + 10], radius=6, fill=CARD)
    d.text((bx + 8, by + 4), VERSION, font=f_ver, fill=ACCENT)

    # 상태 텍스트
    d.text((40, 198), status_text, font=f_status, fill=MUTED)

    # 진행 막대(트로프 + 초록 청크)
    py0, py1 = 228, 240
    d.rounded_rectangle([40, py0, W - 40, py1], radius=6, fill=TROUGH)
    d.rounded_rectangle([40, py0, 40 + int((W - 80) * 0.4), py1], radius=6, fill=GREEN)

    img.save(out_path)
    print("saved", out_path, img.size)


if __name__ == "__main__":
    make("splash.png", "프로그램을 준비하고 있습니다…")
    make("install_splash.png", "설치 프로그램을 준비하고 있습니다…")
