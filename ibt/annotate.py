# -*- coding: utf-8 -*-
"""액티비티 예시 이미지에 '문제/지문/작성가이드/루브릭' 매핑을 오버레이."""
from PIL import Image, ImageDraw, ImageFont

SRC = r"C:\Users\happy\.claude\image-cache\ed9409fd-121b-48b1-8034-4476f9a2dcab\2.png"
OUT = r"J:\claude\ibt\액티비티-매핑.png"
FONT = "C:/Windows/Fonts/malgun.ttf"
FONTB = "C:/Windows/Fonts/malgunbd.ttf"

base = Image.open(SRC).convert("RGBA")
W, H = base.size
LEG = 150  # 하단 범례 높이
canvas = Image.new("RGBA", (W, H + LEG), (255,255,255,255))
canvas.paste(base, (0,0))

overlay = Image.new("RGBA", canvas.size, (0,0,0,0))
d = ImageDraw.Draw(overlay)

def font(sz, bold=True):
    return ImageFont.truetype(FONTB if bold else FONT, sz)

# (x0,y0,x1,y1, 색(RGB), 번호, 이름)
regions = [
    (88, 318, 628, 452,  (0,112,74),   "1", "문제 (수행 지시 + 제한시간)"),
    (88, 458, 628, 1058, (27,100,218), "2", "지문 / 자료 (배경·Flavor Wheel)"),
    (724, 308, 1244, 806, (194,112,10), "3", "작성가이드 (단계·절차)"),
    (724, 822, 1244, 1040,(214,51,108), "4", "루브릭 (채점 기준) ★"),
]

for (x0,y0,x1,y1,c,num,name) in regions:
    fill = c + (38,)         # 반투명 채움 (밑 글자 보이게)
    d.rectangle([x0,y0,x1,y1], fill=fill, outline=c+(255,), width=5)
    # 번호 태그 (좌상단)
    tag = f" {num} "
    fn = font(30)
    tb = d.textbbox((0,0), tag, font=fn)
    tw, th = tb[2]-tb[0], tb[3]-tb[1]
    d.rectangle([x0, y0, x0+tw+14, y0+th+16], fill=c+(255,))
    d.text((x0+7, y0+4), tag, font=fn, fill=(255,255,255,255))

# 범례
ly = H + 22
d.text((40, ly-8), "액티비티 1개 = 문항 1개", font=font(24), fill=(25,31,40,255))
cx = 430
for (x0,y0,x1,y1,c,num,name) in regions:
    d.rectangle([cx, ly, cx+26, ly+26], fill=c+(255,), outline=c+(255,))
    d.text((cx+5, ly-1), num, font=font(20), fill=(255,255,255,255))
    d.text((cx+36, ly+1), name, font=font(18, bold=False), fill=(25,31,40,255))
    cx += 0  # 줄바꿈식 배치
    ly += 32

out = Image.alpha_composite(canvas, overlay).convert("RGB")
out.save(OUT)
print("saved:", OUT, out.size)
