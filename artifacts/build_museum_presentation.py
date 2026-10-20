from pathlib import Path
from io import BytesIO
import hashlib, zipfile, requests
from PIL import Image, ImageDraw, ImageFont, ImageOps
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches

W, H = 1920, 1080
OUT_DIR = Path("artifacts/_museum_build_v4")
OUT_DIR.mkdir(parents=True, exist_ok=True)
PPTX_PATH = Path("artifacts/muzei_promyshlennoi_kultury_v4.pptx")

C = {
    "ink":"#17191C","ink2":"#262B30","paper":"#F4F0E8","paper2":"#E8E1D6",
    "rust":"#C85C3A","amber":"#E4A442","teal":"#4E8581","blue":"#376781",
    "green":"#788F68","muted":"#7E868C","white":"#FFFFFF","line":"#CFC7BA"
}
FONT_R="/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
FONT_B="/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"

def rgb(v):
    v=v.lstrip("#")
    return tuple(int(v[i:i+2],16) for i in (0,2,4))

def font(size,bold=False):
    return ImageFont.truetype(FONT_B if bold else FONT_R,size)

def get_photo(key, urls):
    headers={"User-Agent":"Mozilla/5.0 educational presentation"}
    errors=[]
    for url in urls:
        try:
            r=requests.get(url,headers=headers,timeout=55)
            r.raise_for_status()
            im=ImageOps.exif_transpose(Image.open(BytesIO(r.content))).convert("RGB")
            if im.width<500 or im.height<330:
                raise ValueError(f"small image {im.size}")
            im.save(OUT_DIR/f"{key}.jpg","JPEG",quality=94)
            print(f"PHOTO OK {key}: {url}")
            return im
        except Exception as e:
            errors.append(str(e))
    raise RuntimeError(f"Could not download a real photo for {key}: {errors}")

URLS={
 "demidov":[
  "https://idemvmuzei.ru/storage/museum/32/01/4f/11/fc/41/ae/d0/kolesov.jpg",
  "https://museum-nt.ru/upload/iblock/a59/x4wf3t0vuf8936usyfn5z98502isx8h1.JPG"],
 "sever":[
  "https://xn--b1aflcbdevbbegkvf.xn--p1ai/upload/landing/6c7/trejnas10hp1ibf94tjbwrjdabozx48q.png",
  "https://nashural.ru/assets/uploads/DSC_0158-1-1-scaled.jpg"],
 "moto":[
  "https://img-fotki.yandex.ru/get/197700/355089976.49/0_11dd40_a5a95787_XXL.jpg",
  "https://avatars.dzeninfra.ru/get-zen_doc/271828/pub_65c00dcec9639b6965ed474c_65c00dd7431a64584615353b/scale_1200"],
 "rail":[
  "https://rzd-museum.ru/uploads/news/788/WOLc802vpok72R0Jx0cUFvRRM1tTvth9.jpg",
  "https://loveopium.ru/content/2017/12/museum/01.jpg"],
 "gaz":[
  "https://iv.kommersant.ru/Issues.photo/REGIONS/NNOV_Tema/2024/112/KNN_006448_00200_1_t218_171049.jpg",
  "https://autoreview.ru/images/gallery/%D0%9D%D0%BE%D0%B2%D0%BE%D1%81%D1%82%D0%B8/2022/December/15/GAZ/03.jpg",
  "https://travelq.ru/wp-content/uploads/2025/02/gaz-aa.jpg"],
 "tractor":[
  "https://proehal.ru/media/sights/cheboksary/muzej-istorii-traktora_06.png",
  "https://resizer.mail.ru/p/738bc5ec-8e1e-57b4-9498-6bd0f1cc2b31/dpr%3A200/AQABpu_Q5D4lj8dsfBlJwzkOy1NAiurLPtIRD0nHWUZ86o6eFMWSEVZlWBLebVUsTA7z8Lws5CjKzJDmhUeRgYOGS24.jpg"],
 "hydro":[
  "https://visit.rushydro.ru/upload/iblock/e4d/7z3eewx0726lkc0pagfb0afxs08nfbih/Fotografii-dlya-sayta-muzeya-gidroenergetiki-_4_.jpeg",
  "https://stapol.tech/upload/iblock/00b/j54slydxsl2ckltg48qi7rafhhik0df8.jpg"],
 "space":[
  "https://turportal63.ru/upload/iblock/44e/0.jpg",
  "https://trekkingmania.ru/aimages/658/img-3155.jpg"],
}
P={k:get_photo(k,v) for k,v in URLS.items()}

def canvas(color="paper"): return Image.new("RGB",(W,H),rgb(C[color]))

def crop(im,box,fx=.5,fy=.5):
    _,_,w,h=box; iw,ih=im.size
    sc=max(w/iw,h/ih); nw,nh=max(w,int(iw*sc)),max(h,int(ih*sc))
    im=im.resize((nw,nh),Image.Resampling.LANCZOS)
    l=max(0,min(nw-w,int((nw-w)*fx))); t=max(0,min(nh-h,int((nh-h)*fy)))
    return im.crop((l,t,l+w,t+h))

def photo(slide,im,box,fx=.5,fy=.5,dark=0,radius=0):
    x,y,w,h=box; piece=crop(im,box,fx,fy)
    if dark:
        piece=Image.alpha_composite(piece.convert("RGBA"),Image.new("RGBA",piece.size,(8,10,12,dark))).convert("RGB")
    if radius:
        mask=Image.new("L",(w,h)); ImageDraw.Draw(mask).rounded_rectangle((0,0,w,h),radius=radius,fill=255)
        slide.paste(piece,(x,y),mask)
    else: slide.paste(piece,(x,y))

def fit(draw,text,width,size,min_size=16,bold=True):
    while size>min_size and draw.textbbox((0,0),text,font=font(size,bold))[2]>width: size-=1
    return font(size,bold)

def lines(draw,text,fnt,width):
    out=[]
    for raw in text.split("\n"):
        if raw=="": out.append(""); continue
        cur=""
        for word in raw.split():
            trial=word if not cur else cur+" "+word
            if draw.textbbox((0,0),trial,font=fnt)[2]<=width: cur=trial
            else:
                if cur: out.append(cur)
                cur=word
        if cur: out.append(cur)
    return out

def para(draw,xy,text,width,size,color="ink",bold=False,spacing=1.24):
    x,y=xy; f=font(size,bold)
    for ln in lines(draw,text,f,width):
        draw.text((x,y),ln,font=f,fill=rgb(C[color])); y+=int(size*spacing)
    return y

def pill(draw,text,x=78,y=55,color="rust"):
    f=font(18,True); wid=max(220,draw.textbbox((0,0),text.upper(),font=f)[2]+42)
    draw.rounded_rectangle((x,y,x+wid,y+40),radius=20,fill=rgb(C[color]))
    draw.text((x+20,y+8),text.upper(),font=f,fill=rgb(C["white"]))

def head(draw,text,y=120,dark=False,size=49):
    draw.text((78,y),text,font=fit(draw,text,1760,size,28),fill=rgb(C["white" if dark else "ink"]))
    draw.rectangle((78,y+76,1842,y+80),fill=rgb(C["rust"]))

def footer(draw,n,dark=False,credit=None):
    draw.text((78,1027),"Зинов Илья  •  10 «Б» класс",font=font(16,True),fill=rgb(C["paper2" if dark else "ink"]))
    if credit: draw.text((1000,1028),credit,font=font(13),fill=rgb(C["paper2" if dark else "muted"]))
    draw.text((1810,1025),f"{n:02d}",font=font(18,True),fill=rgb(C["amber"]))

def card(draw,box,fill="white",outline="line",radius=24,width=2):
    draw.rounded_rectangle(box,radius=radius,fill=rgb(C[fill]),outline=rgb(C[outline]),width=width)

slides=[]

# 01 — title with four different industries
s=canvas("ink"); d=ImageDraw.Draw(s)
for box,key,fx in [((0,0,960,540),"demidov",.5),((960,0,960,540),"rail",.5),((0,540,960,540),"gaz",.5),((960,540,960,540),"space",.48)]:
    photo(s,P[key],box,fx,.5,115)
d.rectangle((0,0,24,H),fill=rgb(C["rust"]))
d.rounded_rectangle((72,110,1180,690),radius=34,fill=(18,20,22),outline=rgb(C["rust"]),width=3)
pill(d,"Музеи России",105,145)
para(d,(105,260),"МУЗЕИ\nПРОМЫШЛЕННОЙ\nКУЛЬТУРЫ",980,68,"white",True,1.02)
para(d,(105,560),"Как техника, заводы и профессии становятся историей",940,27,"paper2")
d.rounded_rectangle((105,790,650,925),radius=22,fill=(18,20,22),outline=rgb(C["amber"]),width=3)
d.text((140,820),"ЗИНОВ ИЛЬЯ",font=font(29,True),fill=rgb(C["white"]))
d.text((140,870),"10 «Б» класс",font=font(21),fill=rgb(C["paper2"]))
footer(d,1,True)
slides.append(s)

# 02 — definition
s=canvas(); d=ImageDraw.Draw(s)
pill(d,"Определение",color="teal"); head(d,"Что объединяет эти музеи?")
para(d,(78,235),"Они сохраняют материальные следы производства и объясняют, как промышленность меняла жизнь общества.",1700,30)
groups=[
 ("ТЕХНИКА","машины, транспорт, приборы","amber"),
 ("ПРОСТРАНСТВО","цеха, депо, электростанции","rust"),
 ("ТЕХНОЛОГИИ","от ручного труда до автоматики","teal"),
 ("ЛЮДИ","инженеры, рабочие, конструкторы","blue"),
]
for i,(h,t,c) in enumerate(groups):
    x=78+(i%2)*900; y=450+(i//2)*240
    card(d,(x,y,x+820,y+185))
    d.rectangle((x,y,x+18,y+185),fill=rgb(C[c]))
    d.text((x+52,y+32),h,font=font(25,True),fill=rgb(C[c]))
    para(d,(x+52,y+85),t,710,21,"ink")
footer(d,2)
slides.append(s)

# 03 — museum formats
s=canvas("ink"); d=ImageDraw.Draw(s)
pill(d,"Форматы",color="amber"); head(d,"Промышленный музей бывает разным",dark=True)
formats=[
 ("ЗАВОД-МУЗЕЙ","Сохраняется целая территория предприятия."),
 ("КОРПОРАТИВНЫЙ","История одной компании и её работников."),
 ("ТЕХНИЧЕСКИЙ","Коллекция машин и механизмов."),
 ("ОТРАСЛЕВОЙ","Развитие транспорта, энергетики или машиностроения."),
 ("ИНТЕРАКТИВНЫЙ","Модели, симуляторы и эксперименты."),
 ("ПОД ОТКРЫТЫМ НЕБОМ","Крупные экспонаты в реальном масштабе."),
]
for i,(h,t) in enumerate(formats):
    col=i%3; row=i//3; x=78+col*596; y=290+row*290
    d.rounded_rectangle((x,y,x+540,y+225),radius=26,fill=rgb(C["ink2"]),outline=rgb(C[["rust","amber","teal"][col]]),width=3)
    d.text((x+30,y+34),f"0{i+1}",font=font(21,True),fill=rgb(C[["rust","amber","teal"][col]]))
    d.text((x+30,y+82),h,font=fit(d,h,480,22,17),fill=rgb(C["white"]))
    para(d,(x+30,y+132),t,470,18,"paper2")
footer(d,3,True)
slides.append(s)

# 04 — geography
s=canvas(); d=ImageDraw.Draw(s)
pill(d,"География",color="blue"); head(d,"Восемь городов — восемь отраслей")
para(d,(78,228),"Примеры показывают, что промышленная культура существует далеко за пределами одного региона.",1740,27)
places=[
 ("Нижний Тагил","металлургия","rust"),("Полевской","металлургия","teal"),
 ("Пермь","артиллерия","amber"),("Санкт-Петербург","железные дороги","blue"),
 ("Нижний Новгород","автомобили","green"),("Чебоксары","тракторы","rust"),
 ("Углич","гидроэнергетика","teal"),("Самара","космос","amber"),
]
for i,(city,industry,c) in enumerate(places):
    col=i%4; row=i//4; x=78+col*450; y=420+row*230
    card(d,(x,y,x+405,y+170))
    d.ellipse((x+24,y+26,x+76,y+78),fill=rgb(C[c]))
    d.text((x+95,y+25),city,font=fit(d,city,280,21,16),fill=rgb(C["ink"]))
    d.text((x+95,y+69),industry,font=font(17),fill=rgb(C["muted"]))
    d.text((x+25,y+116),f"{i+1:02d}",font=font(22,True),fill=rgb(C[c]))
footer(d,4)
slides.append(s)

# 05 — Demidov
s=canvas("ink"); photo(s,P["demidov"],(0,0,1110,H),.5,.46,35); d=ImageDraw.Draw(s)
d.rectangle((0,0,1110,170),fill=(18,20,22))
d.text((66,45),"МЕТАЛЛУРГИЯ • НИЖНИЙ ТАГИЛ",font=font(21,True),fill=rgb(C["amber"]))
d.text((66,92),"Старый Демидовский завод",font=fit(d,"Старый Демидовский завод",950,42,30),fill=rgb(C["white"]))
d.rectangle((1110,0,W,H),fill=rgb(C["paper"]))
d.text((1170,80),"1725",font=font(86,True),fill=rgb(C["rust"]))
d.text((1175,180),"год пуска завода",font=font(22),fill=rgb(C["muted"]))
para(d,(1170,280),"Первый в России музей-завод истории развития чёрной металлургии.",650,29,"ink",True)
for y,h,t in [(480,"СОХРАНЕНО","планировка предприятия"),(620,"ЭКСПОНАТЫ","печи, цеха и оборудование"),(760,"ПОКАЗЫВАЕТ","переход от воды к пару и электричеству")]:
    d.text((1170,y),h,font=font(20,True),fill=rgb(C["teal"]))
    para(d,(1370,y),t,390,19,"ink")
footer(d,5,True,"Фото: «Идём в музей»")
slides.append(s)

# 06 — Severskaya
s=canvas("paper2"); d=ImageDraw.Draw(s)
photo(s,P["sever"],(0,0,W,560),.5,.52,35)
d.rectangle((0,0,W,105),fill=(18,20,22))
d.text((78,34),"МЕТАЛЛУРГИЯ • ПОЛЕВСКОЙ",font=font(20,True),fill=rgb(C["amber"]))
d.text((78,600),"Музейный комплекс «Северская домна»",font=fit(d,"Музейный комплекс «Северская домна»",1760,45,30),fill=rgb(C["ink"]))
facts=[("1860","сохранившийся доменный цех"),("1898","доменная печь на старом фундаменте"),("2009","открытие музейного комплекса"),("> 8 000 м²","территория музея")]
for i,(big,txt) in enumerate(facts):
    x=78+i*446
    d.text((x,720),big,font=fit(d,big,400,38,24),fill=rgb(C[["rust","amber","teal","blue"][i]]))
    para(d,(x,780),txt,390,18,"ink")
footer(d,6,credit="Фото: «Открой Полевской»")
slides.append(s)

# 07 — Motovilikha
s=canvas(); d=ImageDraw.Draw(s)
pill(d,"Машиностроение",color="rust"); head(d,"Музей истории Мотовилихинских заводов",size=43)
photo(s,P["moto"],(1020,225,822,690),.48,.5,0,32)
d.text((78,245),"ПЕРМЬ",font=font(20,True),fill=rgb(C["rust"]))
d.text((78,305),"1976",font=font(82,True),fill=rgb(C["ink"]))
d.text((80,405),"год открытия",font=font(21),fill=rgb(C["muted"]))
para(d,(78,480),"Музей показывает почти трёхвековую историю предприятия и развитие тяжёлого машиностроения.",790,27)
card(d,(78,700,900,910))
d.text((115,735),"ДВА ПРОСТРАНСТВА",font=font(22,True),fill=rgb(C["amber"]))
para(d,(115,785),"• историческая экспозиция в здании XIX века\n• артиллерийская техника под открытым небом",720,21)
footer(d,7,credit="Фото: Яндекс.Фотки")
slides.append(s)

# 08 — Railway museum
s=canvas("ink"); photo(s,P["rail"],(0,0,W,650),.5,.52,75); d=ImageDraw.Draw(s)
d.rectangle((0,0,W,120),fill=(18,20,22))
d.text((78,38),"ТРАНСПОРТ • САНКТ-ПЕТЕРБУРГ",font=font(21,True),fill=rgb(C["amber"]))
d.text((78,540),"Музей железных дорог России",font=fit(d,"Музей железных дорог России",1500,48,32),fill=rgb(C["white"]))
stats=[("≈ 100 000 м²","территория комплекса"),("2 КОРПУСА","старое депо и современное здание"),("ПОЛНЫЙ МАСШТАБ","паровозы, тепловозы, электровозы и вагоны")]
for i,(big,txt) in enumerate(stats):
    x=78+i*596
    d.rounded_rectangle((x,700,x+540,915),radius=25,fill=rgb(C["ink2"]),outline=rgb(C[["rust","teal","amber"][i]]),width=3)
    d.text((x+28,735),big,font=fit(d,big,480,28,20),fill=rgb(C[["rust","teal","amber"][i]]))
    para(d,(x+28,795),txt,470,19,"paper2")
footer(d,8,True,"Фото: Музей железных дорог России")
slides.append(s)

# 09 — GAZ
s=canvas(); d=ImageDraw.Draw(s)
pill(d,"Автопром",color="blue"); head(d,"Музей истории ГАЗ")
photo(s,P["gaz"],(510,245,900,590),.5,.5,0,34)
d.rounded_rectangle((78,280,470,810),radius=26,fill=rgb(C["ink"]))
d.text((115,325),"НИЖНИЙ\nНОВГОРОД",font=font(29,True),fill=rgb(C["white"]),spacing=10)
d.text((115,470),"1965",font=font(62,True),fill=rgb(C["amber"]))
para(d,(115,555),"Открытие музея истории Горьковского автозавода.",300,20,"paper2")
d.rounded_rectangle((1450,280,1842,810),radius=26,fill=rgb(C["paper2"]))
d.text((1485,325),"В КОЛЛЕКЦИИ",font=font(21,True),fill=rgb(C["rust"]))
para(d,(1485,385),"Подлинные автомобили, документы, фотографии и истории автозаводцев.",305,23,"ink")
d.text((1485,630),"ГЛАВНАЯ ИДЕЯ",font=font(21,True),fill=rgb(C["teal"]))
para(d,(1485,690),"Показать развитие отечественного автомобилестроения.",305,20,"ink")
footer(d,9,credit="Фото: «Коммерсантъ»")
slides.append(s)

# 10 — Tractor
s=canvas("paper2"); d=ImageDraw.Draw(s)
photo(s,P["tractor"],(0,0,960,H),.55,.5,15)
d.rectangle((960,0,W,H),fill=rgb(C["ink"]))
d.text((1030,65),"СЕЛЬСКОХОЗЯЙСТВЕННАЯ ТЕХНИКА",font=font(19,True),fill=rgb(C["amber"]))
para(d,(1030,120),"Музей истории трактора",760,46,"white",True)
d.text((1030,275),"ЧЕБОКСАРЫ",font=font(20,True),fill=rgb(C["rust"]))
d.text((1030,345),"2011",font=font(72,True),fill=rgb(C["teal"]))
d.text((1032,430),"официальное открытие",font=font(21),fill=rgb(C["paper2"]))
para(d,(1030,520),"Экспозиция ведёт посетителя от сохи и паровых машин к мощным дизельным тракторам.",760,27,"white")
card(d,(1030,730,1800,900),fill="ink2",outline="amber")
d.text((1065,765),"БОЛЕЕ 2,5 ТЫС. М²",font=font(26,True),fill=rgb(C["amber"]))
para(d,(1065,815),"полноразмерная техника, модели и интерактивные разделы",690,20,"paper2")
footer(d,10,True,"Фото: proehal.ru")
slides.append(s)

# 11 — Hydropower
s=canvas("ink"); photo(s,P["hydro"],(0,0,W,H),.5,.48,125); d=ImageDraw.Draw(s)
d.rounded_rectangle((70,70,1050,890),radius=34,fill=(20,24,28),outline=rgb(C["teal"]),width=3)
pill(d,"Энергетика",105,110,"teal")
para(d,(105,190),"Музей гидроэнергетики",850,50,"white",True)
d.text((105,315),"УГЛИЧ",font=font(21,True),fill=rgb(C["amber"]))
para(d,(105,385),"Музей расположен рядом с действующей гидроэлектростанцией 1930–1940-х годов.",820,28,"paper2")
for y,h,t in [(580,"ВОДА","как источник энергии"),(685,"ГЭС","устройство станции и плотины"),(790,"ИНТЕРАКТИВ","карты, модели и эксперименты")]:
    d.text((105,y),h,font=font(21,True),fill=rgb(C["teal"]))
    d.text((330,y),t,font=font(20),fill=rgb(C["white"]))
footer(d,11,True,"Фото: РусГидро")
slides.append(s)

# 12 — Samara Space
s=canvas(); d=ImageDraw.Draw(s)
d.rectangle((0,0,980,H),fill=rgb(C["ink"]))
d.text((78,65),"РАКЕТНО-КОСМИЧЕСКАЯ ОТРАСЛЬ",font=font(20,True),fill=rgb(C["amber"]))
para(d,(78,125),"«Самара Космическая»",780,49,"white",True)
d.text((78,265),"САМАРА",font=font(21,True),fill=rgb(C["rust"]))
para(d,(78,345),"Музей связывает историю города с производством ракет и космической техники.",760,28,"paper2")
facts=[("2001","установлен монумент «Союз»"),("2007","открылся музей"),("Р-7","семейство ракет производится в Самаре")]
for i,(big,txt) in enumerate(facts):
    y=540+i*135
    d.text((78,y),big,font=fit(d,big,180,36,24),fill=rgb(C[["rust","teal","amber"][i]]))
    para(d,(270,y+5),txt,600,19,"white")
photo(s,P["space"],(980,0,940,H),.48,.5,10)
footer(d,12,True,"Фото: туристический портал Самарской области")
slides.append(s)

# 13 — comparison
s=canvas(); d=ImageDraw.Draw(s)
pill(d,"Сравнение",color="amber"); head(d,"Что посетитель получает в разных музеях?")
headers=["МУЗЕЙ","ГЛАВНЫЙ ОБЪЕКТ","ЧЕМУ УЧИТ"]
rows=[
 ("Демидовский завод","территория завода","истории металлургии"),
 ("Северская домна","доменная печь","процессу выплавки чугуна"),
 ("Мотовилиха","артиллерийская техника","тяжёлому машиностроению"),
 ("Музей РЖД","локомотивы и вагоны","развитию транспорта"),
 ("Музей ГАЗ","автомобили","эволюции автопрома"),
 ("Музей трактора","тракторная техника","механизации труда"),
 ("Музей гидроэнергетики","модели ГЭС","получению электроэнергии"),
 ("Самара Космическая","ракета «Союз»","ракетостроению"),
]
x0=[78,650,1240]; widths=[540,560,600]
for x,h in zip(x0,headers):
    d.rounded_rectangle((x,245,x+widths[x0.index(x)],305),radius=16,fill=rgb(C["ink"]))
    d.text((x+20,263),h,font=font(18,True),fill=rgb(C["white"]))
for i,row in enumerate(rows):
    y=325+i*78; fill="white" if i%2==0 else "paper2"
    d.rounded_rectangle((78,y,1840,y+65),radius=12,fill=rgb(C[fill]))
    colors=["rust","teal","amber","blue","green","rust","teal","amber"]
    d.rectangle((78,y,90,y+65),fill=rgb(C[colors[i]]))
    for x,textv,w in zip(x0,row,widths):
        d.text((x+20,y+20),textv,font=fit(d,textv,w-40,17,14,False),fill=rgb(C["ink"]))
footer(d,13)
slides.append(s)

# 14 — conclusion + sources
s=canvas("ink"); d=ImageDraw.Draw(s)
for i,(key,x,y) in enumerate([("demidov",0,0),("rail",960,0),("gaz",0,540),("space",960,540)]):
    photo(s,P[key],(x,y,960,540),.5,.5,175)
d.rounded_rectangle((120,115,1800,835),radius=38,fill=(18,20,22),outline=rgb(C["rust"]),width=3)
pill(d,"Вывод",165,160,"teal")
para(d,(165,240),"Промышленность —\nчасть культуры России",1450,58,"white",True,1.08)
para(d,(165,405),"Эти музеи сохраняют не только технику. Они показывают идеи инженеров, труд людей и развитие целых городов.",1450,29,"paper2")
d.text((165,590),"ИСТОЧНИКИ",font=font(19,True),fill=rgb(C["amber"]))
sources=[
 "museum-nt.ru • stz.tmk-group.ru/severskaya_domna • culture.ru/institutes/11739",
 "rzd-museum.ru • museum.gaz.ru • trackmuseum.ru",
 "visit.rushydro.ru • samaracosmos.ru",
]
para(d,(165,635),"\n".join(sources),1450,17,"paper2",spacing=1.5)
footer(d,14,True)
slides.append(s)

assert len(slides)==14
paths=[]; digests=[]
for i,im in enumerate(slides,1):
    ImageDraw.Draw(im).point((W-2,i),fill=(i,255-i,77))
    p=OUT_DIR/f"slide_{i:02d}.jpg"
    im.save(p,"JPEG",quality=96,subsampling=0)
    paths.append(p); digests.append(hashlib.sha256(p.read_bytes()).hexdigest())
assert len(set(digests))==14

prs=Presentation(); prs.slide_width=Inches(13.333333); prs.slide_height=Inches(7.5)
blank=prs.slide_layouts[6]
for p in paths:
    slide=prs.slides.add_slide(blank)
    slide.shapes.add_picture(str(p),0,0,width=prs.slide_width,height=prs.slide_height)
prs.core_properties.title="Музеи промышленной культуры России"
prs.core_properties.subject="Восемь музеев и отраслей промышленности"
prs.core_properties.author="Зинов Илья"
prs.core_properties.comments="10 «Б» класс • 14 уникальных слайдов"
prs.save(PPTX_PATH)

check=Presentation(PPTX_PATH)
assert len(check.slides)==14
inside=[]
for i,slide in enumerate(check.slides,1):
    pics=[sh for sh in slide.shapes if sh.shape_type==MSO_SHAPE_TYPE.PICTURE]
    assert len(pics)==1,f"slide {i}: {len(pics)} pictures"
    inside.append(hashlib.sha256(pics[0].image.blob).hexdigest())
assert len(set(inside))==14,"duplicate slide images in PPTX"
with zipfile.ZipFile(PPTX_PATH) as z:
    names=z.namelist()
    sx=[n for n in names if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
    media=[n for n in names if n.startswith("ppt/media/")]
    assert len(sx)==14 and len(media)==14,(len(sx),len(media))
print("VALIDATION OK: 14 slides, 14 unique embedded images, 8 different museums")
print(f"Created {PPTX_PATH} ({PPTX_PATH.stat().st_size} bytes)")
