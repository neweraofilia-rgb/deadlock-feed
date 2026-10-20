from pathlib import Path
from io import BytesIO
import textwrap
import requests
from PIL import Image, ImageDraw, ImageFont, ImageEnhance, ImageFilter, ImageOps
from pptx import Presentation
from pptx.util import Inches

W, H = 1920, 1080
OUT_DIR = Path("artifacts/_museum_build")
OUT_DIR.mkdir(parents=True, exist_ok=True)
PPTX_PATH = Path("artifacts/muzei_promyshlennoi_kultury_v2.pptx")

COL = {
    "ink": "#17191C",
    "ink2": "#24282C",
    "paper": "#F4F0E8",
    "paper2": "#E9E2D7",
    "rust": "#C85C3A",
    "amber": "#E4A442",
    "teal": "#4E8581",
    "muted": "#8C949A",
    "white": "#FFFFFF",
    "line": "#CFC7BB",
}

FONT_REG = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
FONT_BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"

def font(size, bold=False):
    return ImageFont.truetype(FONT_BOLD if bold else FONT_REG, size)

def get_image(url, name):
    p = OUT_DIR / name
    headers = {"User-Agent": "Mozilla/5.0 (school presentation; contact: no-reply@example.com)"}
    r = requests.get(url, headers=headers, timeout=45)
    r.raise_for_status()
    img = Image.open(BytesIO(r.content))
    img = ImageOps.exif_transpose(img).convert("RGB")
    img.save(p, quality=94)
    return img

URLS = {
    "demidov": "https://idemvmuzei.ru/storage/museum/32/01/4f/11/fc/41/ae/d0/kolesov.jpg",
    "severskaya": "https://nashural.ru/assets/uploads/DSC_0158-1-1-scaled.jpg",
    "motovilikha": "https://img-fotki.yandex.ru/get/197700/355089976.49/0_11dd40_a5a95787_XXL.jpg",
}

images = {}
fallbacks = {
    "demidov": ("#655044", "СТАРЫЙ ДЕМИДОВСКИЙ ЗАВОД"),
    "severskaya": ("#5D463C", "СЕВЕРСКАЯ ДОМНА"),
    "motovilikha": ("#59625D", "МОТОВИЛИХИНСКИЕ ЗАВОДЫ"),
}
for key, url in URLS.items():
    try:
        images[key] = get_image(url, key + ".jpg")
    except Exception as e:
        bg, label = fallbacks[key]
        img = Image.new("RGB", (1600, 1000), bg)
        d = ImageDraw.Draw(img)
        d.text((100, 450), label, font=font(58, True), fill=COL["white"])
        images[key] = img
        print(f"Image fallback for {key}: {e}")

def rgb(hex_color):
    h = hex_color.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def canvas(color="paper"):
    return Image.new("RGB", (W, H), rgb(COL[color]))

def crop_fill(img, box, focus_y=0.5):
    x, y, w, h = box
    iw, ih = img.size
    scale = max(w / iw, h / ih)
    nw, nh = int(iw * scale), int(ih * scale)
    resized = img.resize((nw, nh), Image.Resampling.LANCZOS)
    left = max(0, (nw - w) // 2)
    top = max(0, min(nh - h, int((nh - h) * focus_y)))
    return resized.crop((left, top, left + w, top + h))

def paste_photo(slide, img, box, focus_y=0.5, dark=0, radius=0):
    x, y, w, h = box
    piece = crop_fill(img, box, focus_y)
    if dark:
        overlay = Image.new("RGBA", piece.size, (10, 12, 14, dark))
        piece = Image.alpha_composite(piece.convert("RGBA"), overlay).convert("RGB")
    if radius:
        mask = Image.new("L", (w, h), 0)
        ImageDraw.Draw(mask).rounded_rectangle((0, 0, w, h), radius=radius, fill=255)
        slide.paste(piece, (x, y), mask)
    else:
        slide.paste(piece, (x, y))

def fit_text(draw, text, max_width, start_size, min_size=18, bold=False):
    size = start_size
    while size > min_size and draw.textbbox((0,0), text, font=font(size,bold))[2] > max_width:
        size -= 1
    return font(size, bold)

def wrap_lines(draw, text, fnt, max_width):
    words = text.split()
    lines, cur = [], ""
    for word in words:
        test = word if not cur else cur + " " + word
        if draw.textbbox((0,0), test, font=fnt)[2] <= max_width:
            cur = test
        else:
            if cur:
                lines.append(cur)
            cur = word
    if cur:
        lines.append(cur)
    return lines

def paragraph(draw, xy, text, max_width, size, color, bold=False, spacing=1.28, max_lines=None):
    x, y = xy
    fnt = font(size, bold)
    lines = []
    for raw in text.split("\n"):
        lines.extend(wrap_lines(draw, raw, fnt, max_width) if raw else [""])
    if max_lines:
        lines = lines[:max_lines]
    step = int(size * spacing)
    for line in lines:
        draw.text((x, y), line, font=fnt, fill=rgb(COL[color]))
        y += step
    return y

def label(draw, x, y, text, color="rust"):
    draw.rounded_rectangle((x, y, x+270, y+42), radius=21, fill=rgb(COL[color]))
    draw.text((x+20, y+8), text.upper(), font=font(20, True), fill=rgb(COL["white"]))

def footer(draw, num, dark=False):
    color = COL["white"] if dark else COL["ink"]
    draw.text((78, 1025), "Зинов Илья  •  10 «Б» класс", font=font(16, True), fill=rgb(color))
    draw.text((1810, 1025), f"{num:02d}", font=font(18, True), fill=rgb(COL["amber"]))

def title(draw, text, dark=False, y=90, size=50):
    color = "white" if dark else "ink"
    fnt = fit_text(draw, text, 1700, size, 34, True)
    draw.text((78, y), text, font=fnt, fill=rgb(COL[color]))
    draw.rectangle((78, y+78, 1842, y+82), fill=rgb(COL["rust"]))

def card(draw, box, fill="white", outline="line", radius=24):
    draw.rounded_rectangle(box, radius=radius, fill=rgb(COL[fill]), outline=rgb(COL[outline]), width=2)

slides = []

# 1
s = canvas("ink")
paste_photo(s, images["demidov"], (0, 0, W, H), focus_y=0.42, dark=150)
d = ImageDraw.Draw(s)
d.rectangle((0, 0, 24, H), fill=rgb(COL["rust"]))
label(d, 92, 88, "Музеи России")
paragraph(d, (92, 225), "ЗАВОД\nКАК МУЗЕЙ", 1080, 92, "white", True, spacing=1.05)
paragraph(d, (98, 470), "Музеи промышленной культуры России", 980, 34, "paper2", False)
d.rounded_rectangle((92, 770, 690, 920), radius=24, fill=(17,19,22,210), outline=rgb(COL["rust"]), width=3)
d.text((128, 803), "ЗИНОВ ИЛЬЯ", font=font(30, True), fill=rgb(COL["white"]))
d.text((128, 858), "ученик 10 «Б» класса", font=font(22), fill=rgb(COL["paper2"]))
d.text((1430, 1025), "Фото: Культура.РФ", font=font(15), fill=rgb(COL["paper2"]))
slides.append(s)

# 2
s = canvas()
d = ImageDraw.Draw(s)
label(d, 78, 58, "Понятие", "teal")
title(d, "Что такое промышленная культура?", y=125)
paragraph(d, (78, 245), "Это наследие производства: здания, машины, технологии и истории людей, которые работали на предприятиях.", 900, 31, "ink")
paste_photo(s, images["demidov"], (1100, 205, 742, 710), focus_y=0.45, dark=20, radius=34)
d.rounded_rectangle((1120, 845, 1818, 915), radius=18, fill=rgb(COL["ink"]))
d.text((1146, 866), "Промышленный объект сам становится экспонатом", font=font(20, True), fill=rgb(COL["white"]))
items = [
    ("01", "ПРОСТРАНСТВО", "Цеха, домны, шахты и заводские посёлки"),
    ("02", "ТЕХНОЛОГИИ", "Станки, чертежи, механизмы и продукция"),
    ("03", "ЛЮДИ", "Профессии, условия труда и личные истории"),
]
for i, (n, h, txt) in enumerate(items):
    y = 520 + i*142
    card(d, (78, y, 970, y+116), fill="white")
    d.rectangle((78, y, 94, y+116), fill=rgb(COL["rust" if i==0 else ("amber" if i==1 else "teal")]))
    d.text((120, y+18), n, font=font(24, True), fill=rgb(COL["rust"]))
    d.text((210, y+15), h, font=font(22, True), fill=rgb(COL["ink"]))
    d.text((210, y+54), txt, font=font(19), fill=rgb(COL["muted"]))
footer(d, 2)
slides.append(s)

# 3
s = canvas("ink")
d = ImageDraw.Draw(s)
label(d, 78, 58, "Обзор")
title(d, "Три музея — три разных подхода", dark=True, y=125)
cards = [
    (78, images["demidov"], "Нижний Тагил", "Старый Демидовский завод", "Завод-музей под открытым небом", "Культура.РФ"),
    (662, images["severskaya"], "Полевской", "Северская домна", "Подлинный доменный цех XIX века", "Наш Урал"),
    (1246, images["motovilikha"], "Пермь", "Мотовилихинские заводы", "История предприятия и артиллерии", "saidpvo / Яндекс.Фотки"),
]
for x, img, city, name, desc, cred in cards:
    paste_photo(s, img, (x, 260, 520, 420), focus_y=0.48, dark=35, radius=28)
    d.rounded_rectangle((x, 650, x+520, 902), radius=26, fill=rgb(COL["ink2"]), outline=rgb(COL["muted"]), width=2)
    d.text((x+30, 684), city.upper(), font=font(18, True), fill=rgb(COL["amber"]))
    paragraph(d, (x+30, 730), name, 460, 28, "white", True, spacing=1.15)
    paragraph(d, (x+30, 810), desc, 460, 19, "paper2")
    d.text((x+30, 872), "Фото: " + cred, font=font(13), fill=rgb(COL["muted"]))
footer(d, 3, True)
slides.append(s)

# 4
s = canvas()
paste_photo(s, images["demidov"], (0, 0, 980, H), focus_y=0.45, dark=45)
d = ImageDraw.Draw(s)
d.rectangle((0, 0, 980, 160), fill=(18,20,22,175))
d.text((72, 50), "ПРИМЕР 1", font=font(20, True), fill=rgb(COL["amber"]))
paragraph(d, (72, 90), "Старый Демидовский завод", 830, 42, "white", True)
d.rectangle((980, 0, W, H), fill=rgb(COL["paper"]))
d.text((1060, 82), "НИЖНИЙ ТАГИЛ", font=font(20, True), fill=rgb(COL["rust"]))
d.text((1060, 142), "1725", font=font(92, True), fill=rgb(COL["ink"]))
d.text((1065, 250), "год основания завода", font=font(24), fill=rgb(COL["muted"]))
paragraph(d, (1060, 340), "Предприятие основал Акинфий Демидов. Здесь выплавляли чугун и производили железо.", 720, 29, "ink", False)
card(d, (1060, 555, 1818, 760), fill="white")
d.text((1100, 590), "ПОЧЕМУ ЭТО МУЗЕЙ?", font=font(24, True), fill=rgb(COL["rust"]))
paragraph(d, (1100, 640), "Сохранились настоящие цеха, доменные печи и оборудование разных эпох. Посетитель проходит по территории бывшего предприятия.", 650, 23, "ink")
d.text((1060, 905), "Фото: Культура.РФ", font=font(16), fill=rgb(COL["muted"]))
footer(d, 4)
slides.append(s)

# 5
s = canvas("ink")
d = ImageDraw.Draw(s)
d.rectangle((0, 0, 870, H), fill=rgb(COL["ink"]))
d.text((78, 66), "ПРИМЕР 2", font=font(20, True), fill=rgb(COL["amber"]))
paragraph(d, (78, 112), "Музейный комплекс\n«Северская домна»", 730, 46, "white", True, spacing=1.12)
d.text((78, 285), "ПОЛЕВСКОЙ", font=font(20, True), fill=rgb(COL["rust"]))
d.text((78, 355), "1860", font=font(76, True), fill=rgb(COL["amber"]))
d.text((80, 445), "год постройки доменного цеха", font=font(21), fill=rgb(COL["paper2"]))
d.text((78, 535), "2009", font=font(76, True), fill=rgb(COL["teal"]))
d.text((80, 625), "год открытия музейного комплекса", font=font(21), fill=rgb(COL["paper2"]))
paragraph(d, (78, 730), "Главный экспонат — подлинная доменная печь и сохранившееся промышленное оборудование XIX века.", 700, 25, "white")
paste_photo(s, images["severskaya"], (870, 0, 1050, H), focus_y=0.48, dark=15)
d.rectangle((900, 870, 1870, 995), fill=(18,20,22,205))
paragraph(d, (935, 900), "Исторический цех помогает увидеть процесс выплавки чугуна в реальном масштабе.", 900, 24, "white", True)
d.text((1590, 1028), "Фото: Наш Урал", font=font(15), fill=rgb(COL["white"]))
footer(d, 5, True)
slides.append(s)

# 6
s = canvas()
d = ImageDraw.Draw(s)
label(d, 78, 58, "Пример 3")
title(d, "Музей истории Мотовилихинских заводов", y=125, size=45)
d.text((78, 235), "ПЕРМЬ", font=font(20, True), fill=rgb(COL["rust"]))
d.text((78, 290), "1976", font=font(72, True), fill=rgb(COL["ink"]))
d.text((80, 380), "год открытия музея", font=font(22), fill=rgb(COL["muted"]))
paragraph(d, (78, 455), "Экспозиция рассказывает почти о трёх веках развития предприятия и промышленности города.", 760, 29, "ink")
card(d, (78, 680, 825, 880), fill="white")
d.text((116, 715), "ДВЕ ЧАСТИ ЭКСПОЗИЦИИ", font=font(22, True), fill=rgb(COL["rust"]))
paragraph(d, (116, 765), "• история завода в здании XIX века\n• артиллерийская техника под открытым небом", 650, 22, "ink")
paste_photo(s, images["motovilikha"], (900, 225, 942, 690), focus_y=0.52, dark=0, radius=32)
d.text((1380, 932), "Фото: saidpvo / Яндекс.Фотки", font=font(14), fill=rgb(COL["muted"]))
footer(d, 6)
slides.append(s)

# 7
s = canvas("ink")
paste_photo(s, images["demidov"], (0, 0, W, H), focus_y=0.45, dark=195)
d = ImageDraw.Draw(s)
label(d, 78, 58, "Значение", "rust")
title(d, "Зачем сохранять старые заводы?", dark=True, y=125)
vals = [
    ("01", "ПАМЯТЬ", "Истории инженеров, мастеров и рабочих"),
    ("02", "ЗНАНИЯ", "Развитие техники и производственных процессов"),
    ("03", "АРХИТЕКТУРА", "Редкие цеха, домны и промышленные сооружения"),
    ("04", "НОВАЯ ЖИЗНЬ", "Экскурсии, выставки и культурные пространства"),
]
positions = [(78,270),(1000,270),(78,555),(1000,555)]
for (n,h,txt),(x,y) in zip(vals,positions):
    d.rounded_rectangle((x,y,x+840,y+220), radius=28, fill=(244,240,232,235))
    d.rectangle((x,y,x+18,y+220), fill=rgb(COL["rust"]))
    d.text((x+52,y+34), n, font=font(26, True), fill=rgb(COL["rust"]))
    d.text((x+140,y+32), h, font=font(27, True), fill=rgb(COL["ink"]))
    paragraph(d, (x+140,y+88), txt, 640, 22, "ink")
d.rounded_rectangle((78, 855, 1840, 970), radius=24, fill=rgb(COL["rust"]))
paragraph(d, (125, 888), "Сохраняя завод, музей сохраняет не только металл и кирпич, но и опыт многих поколений.", 1650, 26, "white", True)
footer(d, 7, True)
slides.append(s)

# 8
s = canvas()
d = ImageDraw.Draw(s)
label(d, 78, 58, "Вывод", "teal")
title(d, "Промышленность — часть культуры России", y=125, size=48)
paragraph(d, (78, 245), "Музеи-заводы показывают историю страны через технологии, пространство и человеческий труд.", 1760, 31, "ink")
for x, key in [(78,"demidov"),(678,"severskaya"),(1278,"motovilikha")]:
    paste_photo(s, images[key], (x, 390, 564, 310), focus_y=0.48, dark=10, radius=24)
names=[("Старый Демидовский завод",78),("Северская домна",678),("Мотовилихинские заводы",1278)]
for name,x in names:
    d.rounded_rectangle((x, 665, x+564, 738), radius=18, fill=rgb(COL["ink"]))
    f=fit_text(d,name,520,21,16,True)
    d.text((x+22,687),name,font=f,fill=rgb(COL["white"]))
d.text((78, 790), "ИСТОЧНИКИ", font=font(20, True), fill=rgb(COL["rust"]))
sources = [
    "ntagil.org/kultura/mesta/muzei/142567 — Старый Демидовский завод",
    "stz.tmk-group.ru/severskaya_domna — музейный комплекс «Северская домна»",
    "culture.ru/institutes/11739 — музей истории Мотовилихинских заводов",
]
paragraph(d, (78, 835), "\n".join(sources), 1710, 17, "muted", spacing=1.5)
footer(d, 8)
slides.append(s)

# Save slide images and assemble a robust single-image-per-slide PPTX.
prs = Presentation()
prs.slide_width = Inches(13.333333)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]
# Remove the default first slide if any (python-pptx starts with none).
for idx, img in enumerate(slides, start=1):
    png_path = OUT_DIR / f"slide_{idx:02d}.png"
    img.save(png_path, "PNG", optimize=True)
    slide = prs.slides.add_slide(blank)
    slide.shapes.add_picture(str(png_path), 0, 0, width=prs.slide_width, height=prs.slide_height)

prs.core_properties.title = "Музеи промышленной культуры России"
prs.core_properties.subject = "Завод как музей"
prs.core_properties.author = "Зинов Илья"
prs.core_properties.comments = "10 «Б» класс"
prs.save(PPTX_PATH)
print(f"Created {PPTX_PATH} ({PPTX_PATH.stat().st_size} bytes)")
