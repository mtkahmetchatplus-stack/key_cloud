import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy

wb = openpyxl.Workbook()

# ── Stil sabitleri ────────────────────────────────────────────────────────────
H1  = PatternFill("solid", fgColor="1F4E79")
H2  = PatternFill("solid", fgColor="2E75B6")
H3  = PatternFill("solid", fgColor="BDD7EE")
ALT = PatternFill("solid", fgColor="DEEAF1")
WH  = PatternFill("solid", fgColor="FFFFFF")
YEL = PatternFill("solid", fgColor="FFFF99")
GRN = PatternFill("solid", fgColor="E2EFDA")
ORG = PatternFill("solid", fgColor="FCE4D6")
PNK = PatternFill("solid", fgColor="FFD9D9")

fw  = Font(bold=True, color="FFFFFF", size=10)
fb  = Font(bold=True, color="1F4E79", size=10)
fn  = Font(size=10)
fn9 = Font(size=9)
fwl = Font(bold=True, color="FFFFFF", size=12)

CA = Alignment(horizontal="center", vertical="center", wrap_text=True)
LA = Alignment(horizontal="left",   vertical="center", wrap_text=True)
RA = Alignment(horizontal="right",  vertical="center")

th = Side(style="thin",   color="4472C4")
mk = Side(style="medium", color="1F4E79")
TB = Border(left=th, right=th, top=th, bottom=th)
MB = Border(left=mk, right=mk, top=mk, bottom=mk)

def sw(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def C(ws, r, c, v, fill=WH, font=fn, align=CA, border=TB):
    cl = ws.cell(row=r, column=c, value=v)
    cl.fill=copy(fill); cl.font=copy(font); cl.alignment=copy(align); cl.border=copy(border)
    return cl

def HDR(ws, r, cols, texts, fill=H2, font=fw, h=30):
    for c, t in zip(cols, texts):
        C(ws, r, c, t, fill, font, CA, TB)
    ws.row_dimensions[r].height = h

def TITLE(ws, r, c1, c2, text, fill=H1, font=fwl, h=24):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    cl = ws.cell(row=r, column=c1, value=text)
    cl.fill=copy(fill); cl.font=copy(font); cl.alignment=copy(CA); cl.border=copy(MB)
    ws.row_dimensions[r].height = h

def SUB(ws, r, c1, c2, text, fill=H3):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    cl = ws.cell(row=r, column=c1, value=text)
    cl.fill=copy(fill); cl.font=copy(fb); cl.alignment=copy(CA); cl.border=copy(TB)
    ws.row_dimensions[r].height = 18

def DATA(ws, r, cols, vals, alt=False):
    fill = ALT if alt else WH
    for c, v in zip(cols, vals):
        C(ws, r, c, v, fill, fn, CA, TB)

def TOT(ws, r, c1, cv, label, val, fill=H1):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=cv-1)
    cl = ws.cell(row=r, column=c1, value=label)
    cl.fill=copy(fill); cl.font=copy(fw); cl.alignment=copy(RA); cl.border=copy(TB)
    C(ws, r, cv, val, fill, fw, CA, TB)

def NOTE(ws, r, c1, c2, text):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    cl = ws.cell(row=r, column=c1, value=text)
    cl.fill=copy(WH); cl.font=copy(fn9); cl.alignment=copy(LA); cl.border=copy(TB)

# ═══════════════════════════════════════════════════════════════════
# SAYFA 1 – TASARIM VERİLERİ
# ═══════════════════════════════════════════════════════════════════
ws = wb.active; ws.title = "Tasarım Verileri"; ws.sheet_view.showGridLines = False
TITLE(ws, 1, 1, 4, "7000 Bbl ÜRETİM TANKI T-202  —  TASARIM VERİLERİ")
HDR(ws, 2, [1,2,3,4], ["PARAMETRE", "DEĞER", "PARAMETRE", "DEĞER"])
tasarim = [
    ("KAPASİTE","7000 Bbl","YÖNETMELİK","API 650"),
    ("SIVI","HAM PETROL","KOROZYON PAYI","3 mm"),
    ("TANK ÇAPI","12.730 m","YOĞUNLUK","920 kg/m³"),
    ("TANK YÜKSEK.","8.920 m","OPER. BASINCI","ATMOSFERİK"),
    ("ÇATI TİPİ","DESTEKLİ KONİK ÇATI","TAS. BASINCI","1.5 / -0.2 kPa"),
    ("TABAN TİPİ","MERKEZDEN DIŞA EĞİMLİ","OPER. SICAKLIK","AMBİANT"),
    ("RÜZGAR HIZI","28 m/s","TAS. SICAKLIK","40 °C / -7 °C"),
    ("SAC MALZEMESİ","S275J2","KAYN. VERİMİ","1"),
    ("ANKRAJ MALZEMESİ","AISI C1030","YALITIM","YOK"),
]
for i, (p1,v1,p2,v2) in enumerate(tasarim, start=3):
    fill = ALT if i%2==0 else WH
    for col, val, bold in [(1,p1,True),(2,v1,False),(3,p2,True),(4,v2,False)]:
        cl = ws.cell(row=i, column=col, value=val)
        cl.fill=fill; cl.font=Font(bold=bold, size=10, color="1F4E79") if bold else fn
        cl.alignment=LA if col in [1,3] else CA; cl.border=TB
sw(ws, [32,22,32,22])

# ═══════════════════════════════════════════════════════════════════
# SAYFA 2 – NOZUL LİSTESİ (Genel + Yerleşim)
# ═══════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("Nozul Listesi"); ws2.sheet_view.showGridLines = False
TITLE(ws2, 1, 1, 11, "7000 Bbl ÜRETİM TANKI — NOZUL LİSTESİ VE YERLEŞİM PLANI")

# Sol – Nozul Listesi
SUB(ws2, 2, 1, 7, "NOZUL LİSTESİ (Tanım ve Flanş Bilgileri)")
HDR(ws2, 3, list(range(1,8)), ["POZ","ADET","TANIMI","ÇAP","BASINÇ","FLANŞ","NOTLAR"])
nozul_liste = [
    ("N1",  1, "Petrol Giriş Nozulu",     '12"', "LB 150", "SO-RF", ""),
    ("N4",  1, "Petrol Çıkış Nozulu",     '12"', "LB 150", "SO-RF", "BOTTOM"),
    ("N7A/B",2,"Drain (Tahliye)",         ' 6"', "LB 150", "SO-RF", "2 adet"),
    ("N8",  1, "Köpük Yapıcı",            ' 8"', "LB 150", "SO-RF", ""),
    ("N9",  1, "Sıcaklık Ölçümü (TI)",   '1.5"',"LB 150", "SO-RF", ""),
    ("N14", 1, "Mekanik Samandıra",       "–",   "–",      "–",     ""),
    ("N18", 1, "Yüksek Seviye Sensörü",   ' 1"', "LB 150", "SO-RF", ""),
    ("N19", 1, "Radarla Ölçüm",           ' 6"', "LB 150", "SO-RF", "ROOF"),
    ("N21", 4, "Hava Nozulu",             '12"', "–",      "–",     "4 adet"),
    ("M1",  1, "Gövde Manholu",           '24"', "API 650","RF",    ""),
    ("M2",  1, "Tavan Manholu",           '24"', "API 650","RF",    ""),
    ("M3",  1, "Temizlik Manholu",        '24"', "API 650","RF",    ""),
]
for i, row in enumerate(nozul_liste, start=4):
    DATA(ws2, i, list(range(1,8)), row, alt=(i%2==0))

# Sağ – Yerleşim Tablosu
SUB(ws2, 2, 9, 15, "NOZZLE ELEVATION AND PROJECTION")
HDR(ws2, 3, list(range(9,16)), ["MARK","SIZE","A (mm)","B (mm)","COURSE","REINF.PAD","NOT"])
nozul_elev = [
    ("N1",  '12"', 225, 449,  "SHELL", "t=8 ø685",  ""),
    ("N4",  '12"', 225, 449,  "BOTTOM","t=6 ø685",  ""),
    ("N7A/B",' 6"',200, 306,  "SHELL", "t=8 ø400",  ""),
    ("N8",  ' 8"', 225, 8500, "SHELL", "t=6 ø485",  ""),
    ("N9",  '1.5"',150, 1200, "SHELL", "–",          ""),
    ("N14", "–",   "–",6068,  "ROOF",  "–",          ""),
    ("N18", ' 1"', 150, 8600, "SHELL", "–",          ""),
    ("N19", ' 6"', 200, 0,    "ROOF",  "t=8 ø400",  ""),
    ("N21", '12"', 225, 0,    "ROOF",  "t=8 ø685",  ""),
    ("M1",  '24"', 150, 750,  "SHELL", "t=8 ø1255", ""),
    ("M2",  '24"', 250, 5750, "ROOF",  "t=8 ø1150", ""),
    ("M3",  '24"', 200, 310,  "SHELL", "t=8 ø1830", ""),
]
for i, row in enumerate(nozul_elev, start=4):
    DATA(ws2, i, list(range(9,16)), row, alt=(i%2==0))
sw(ws2, [8,8,28,8,10,12,14, 2, 8,8,10,12,10,14,14])

# ═══════════════════════════════════════════════════════════════════
# SAYFA 3 – GÖVDE PLAKALARI
# ═══════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("Gövde Plakaları"); ws3.sheet_view.showGridLines = False
TITLE(ws3, 1, 1, 8, "GÖVDE (SHELL) PLAKALARI MALZEME LİSTESİ  (Çizim: KEY199MECEQP2001)")
HDR(ws3, 2, list(range(1,9)),
    ["POZ","ADET","AÇIKLAMA  (En×Kalınlık ......... Boy)  mm","MALZEME",
     "BİRİM\nAĞIRLIK(Kg)","TOPLAM\nAĞIRLIK(Kg)","KALINLIK\n(mm)","İŞLEM TÜRÜ"])
ws3.row_dimensions[2].height = 40
govde = [
    ("1",   1, "⊏ 1480×8 ......... 6000",  "S275J2", 557.6,   557.6,  8, "Kesim + Silindirik Bükme"),
    ("1A",  1, "⊏ 1480×8 ......... 6000",  "S275J2", 557.0,   557.0,  8, "Kesim + Silindirik Bükme"),
    ("1B",  1, "⊏ 1480×8 ......... 6000",  "S275J2", 552.3,   552.3,  8, "Kesim + Silindirik Bükme"),
    ("1C",  1, "⊏ 1480×8 ......... 6000",  "S275J2", 539.1,   539.1,  8, "Kesim + Silindirik Bükme"),
    ("1D",  1, "⊏ 1480×8 ......... 6000",  "S275J2", 414.0,   414.0,  8, "Kesim + Silindirik Bükme"),
    ("1E",  1, "⊏ 1480×8 ......... 6000",  "S275J2", 557.6,   557.6,  8, "Kesim + Silindirik Bükme"),
    ("2",   1, "⊏ 1480×8 ......... 3997",  "S275J2", 371.5,   371.5,  8, "Kesim + Silindirik Bükme"),
    ("3",   30,"⊏ 1480×6 ......... 6000",  "S275J2", 418.2, 12546.0,  6, "Kesim + Silindirik Bükme"),
    ("3A",  1, "⊏ 1480×6 ......... 6000",  "S275J2", 418.0,   418.0,  6, "Kesim + Silindirik Bükme"),
]
for i, row in enumerate(govde, start=3):
    DATA(ws3, i, list(range(1,9)), row, alt=(i%2==0))
TOT(ws3, 12, 1, 6, "TOPLAM AĞIRLIK:", 16513.0)
for c in [7,8]: C(ws3, 12, c, "", H1, fw, CA, TB)
ws3.row_dimensions[12].height = 18

# Kurs özeti
r=14
SUB(ws3, r, 1, 8, "KURS (COURSE) ÖZETI — Gövde Plakası")
r+=1
HDR(ws3, r, list(range(1,8)), ["KURS","KALINLIK","YÜKSEK.(mm)","ÇEVRENİN PLAKA ADET","AÇIKLAMA","MALZEME","AĞIRLIK (kg)"])
ws3.row_dimensions[r].height=30; r+=1
kurs_data=[
    ("1. Kurs (Alt)","8 mm",1480,"7 adet (POZ 1,1A,1B,1C,1D,1E,2)","Nozul kesimleri mevcut","S275J2",
     557.6+557.0+552.3+539.1+414.0+557.6+371.5),
    ("2-6. Kurslar","6 mm",1480,"31 adet (POZ 3×30 + POZ 3A×1)","Standart plakalar","S275J2",
     12546.0+418.0),
]
for i, kd in enumerate(kurs_data, start=r):
    for c, v in enumerate(kd, 1):
        fill = ALT if i%2==0 else WH
        cl = ws3.cell(row=i, column=c, value=v)
        cl.fill=fill; cl.font=fn; cl.alignment=CA; cl.border=TB
r += len(kurs_data)

NOTE(ws3,r,1,8,"NOT: Kurs yüksekliği 1480mm, toplam gövde yüksekliği ≈ 6 × 1480mm = 8880mm ≈ 8.920m. POZ 3 = 30 adet standart silindirik plaka (8mm'lik kursun nozulsuz plakaları 8mm'dir).")
sw(ws3,[8,8,40,14,14,14,12,24])

# ═══════════════════════════════════════════════════════════════════
# SAYFA 4 – TABAN PLAKALARI
# ═══════════════════════════════════════════════════════════════════
ws4 = wb.create_sheet("Taban Plakaları"); ws4.sheet_view.showGridLines = False
TITLE(ws4, 1, 1, 8, "TABAN (BOTTOM) PLAKALARI MALZEME LİSTESİ  (Çizim: KEY199MECEQP2003)")
HDR(ws4, 2, list(range(1,9)),
    ["POZ","ADET","AÇIKLAMA  (En×Kalınlık ......... Boy)  mm","MALZEME",
     "BİRİM\nAĞIRLIK(Kg)","TOPLAM\nAĞIRLIK(Kg)","KALINLIK\n(mm)","İŞLEM TÜRÜ"])
ws4.row_dimensions[2].height = 40
taban = [
    ("1",   7, "⊏ 1500×10 ......... 5990",  "S275J2",  705.3, 4937.1, 10, "Kesim + Düz Hadde"),
    ("2",   2, "⊏ 1500×10 ......... 2684",  "S275J2",  314.0,  628.0, 10, "Kesim + Düz Hadde"),
    ("3",   4, "⊏ 1500×10 ......... 635",   "S275J2",   55.5,  222.0, 10, "Kesim"),
    ("4",   3, "⊏ 1500×10 ......... 2262",  "S275J2",  219.2,  657.6, 10, "Kesim + Düz Hadde"),
    ("5",   1, "⊏ 1500×10 ......... 2675.5","S275J2",  222.0,  222.0, 10, "Kesim + Düz Hadde"),
    ("6",   1, "⊏ 1500×10 ......... 4394.5","S275J2",  557.0,  557.0, 10, "Kesim + Düz Hadde"),
    ("7",   2, "⊏ 1500×10 ......... 2994",  "S275J2",  347.4,  694.8, 10, "Kesim + Düz Hadde"),
    ("8",   2, "⊏ 632×10 .......... 5207.5","S275J2",  174.1,  348.2, 10, "Kesim + Düz Hadde"),
    ("9",   6, "⊏ 800×12 .......... 5937",  "S275J2",  436.1, 2616.6, 12, "Kesim + Düz Hadde (Anüler)"),
    ("9A",  1, "⊏ 800×12 .......... 3327",  "S275J2",  237.8,  237.8, 12, "Kesim + Düz Hadde (Anüler)"),
    ("10",  8, "⊏ 50×5 ............. 720",  "S275J2",    1.42,  11.4,  5, "Kesim (Drenaj strip)"),
    ("11",  1, "⊏ 50×5 ............. 1301.5","S275J2",   2.55,   2.55, 5, "Kesim (Drenaj strip)"),
    ("12", 13, "⊏ 50×5 ............. 1400", "S275J2",    2.8,   36.4,  5, "Kesim (Drenaj strip)"),
    ("13",  1, "⊏ 50×5 ............. 2641.5","S275J2",   5.18,   5.18, 5, "Kesim (Drenaj strip)"),
]
for i, row in enumerate(taban, start=3):
    DATA(ws4, i, list(range(1,9)), row, alt=(i%2==0))
TOT(ws4, 17, 1, 6, "TOPLAM AĞIRLIK:", 11176.7)
for c in [7,8]: C(ws4, 17, c, "", H1, fw, CA, TB)

# Drenaj çukuru
r=19
SUB(ws4, r, 1, 8, "MALZEME LİSTESİ (DRENAJ ÇUKURU İÇİN)  —  Toplam: 76.1 Kg"); r+=1
NOTE(ws4, r, 1, 8, "Drenaj çukuru detay tablosu: Toplam ağırlık 76.1 Kg (detay kaynak çizimde)"); r+=1

# Ankrajlar
r+=1
SUB(ws4, r, 1, 8, "MALZEME LİSTESİ (ANKRAJLAR İÇİN)  —  NOT: 1 komple için, 16 adet imal edilecektir."); r+=1
HDR(ws4, r, list(range(1,8)),
    ["POZ","ADET","AÇIKLAMA","MALZEME","BİRİM\nAĞIRLIK(Kg)","TOPLAM\nAĞIRLIK(Kg)","İŞLEM"])
ws4.row_dimensions[r].height=30; r+=1
ankraj=[
    (1,1,"⊏ 200×8 ........... 205","S275J2",2.6,2.6,"Kesim"),
    (2,2,"⊏ 175×15 .......... 180","S275J2",2.7,5.5,"Kesim"),
    (3,1,"⊏ 150×20 .......... 175","S275J2",4.0,4.0,"Kesim"),
    (4,1,"⊏ 80×20 ........... 80", "S275J2",0.95,0.95,"Kesim"),
]
for i, row in enumerate(ankraj, start=r):
    DATA(ws4, i, list(range(1,8)), row, alt=(i%2==0))
r += len(ankraj)
TOT(ws4, r, 1, 6, "TOPLAM AĞIRLIK (1 Komple): 13.5 Kg  |  16 Adet → 216 Kg", 216.0)
C(ws4, r, 7, "", H1, fw, CA, TB)
sw(ws4,[8,8,42,14,14,14,12,22])

# ═══════════════════════════════════════════════════════════════════
# SAYFA 5 – TAVAN PLAKALARI
# ═══════════════════════════════════════════════════════════════════
ws5 = wb.create_sheet("Tavan Plakaları"); ws5.sheet_view.showGridLines = False
TITLE(ws5, 1, 1, 8, "KONİK TAVAN PLAKALARI MALZEME LİSTESİ  (Çizim: KEY199MECEQP2002)")
HDR(ws5, 2, list(range(1,9)),
    ["POZ","ADET","AÇIKLAMA  (En×Kalınlık ......... Boy)  mm","MALZEME",
     "BİRİM\nAĞIRLIK(Kg)","TOPLAM\nAĞIRLIK(Kg)","KALINLIK\n(mm)","İŞLEM TÜRÜ"])
ws5.row_dimensions[2].height = 40
tavan = [
    ("1",   4, "⊏ 1500×8 ......... 6000",  "S275J2", 565.2, 1695.6, 8, "Kesim + Konik Bükme"),
    ("1A",  1, "⊏ 1500×8 ......... 6000",  "S275J2", 560.0,  560.0, 8, "Kesim + Konik Bükme"),
    ("1B",  1, "⊏ 1500×8 ......... 6000",  "S275J2", 560.0,  560.0, 8, "Kesim + Konik Bükme"),
    ("1C",  1, "⊏ 1500×8 ......... 6000",  "S275J2", 560.0,  560.0, 8, "Kesim + Konik Bükme"),
    ("2",   2, "⊏ 1500×8 ......... 3560.5","S275J2", 334.0,  668.0, 8, "Kesim + Konik Bükme"),
    ("3",   4, "⊏ 1500×8 ......... 548",   "S275J2",  38.3,  153.2, 8, "Kesim + Konik Bükme"),
    ("4",   4, "⊏ 1500×8 ......... 3195.5","S275J2", 269.3, 1077.2, 8, "Kesim + Konik Bükme"),
    ("5",   2, "⊏ 1500×8 ......... 2873",  "S275J2", 257.3,  514.6, 8, "Kesim + Konik Bükme"),
    ("5A",  1, "⊏ 1500×8 ......... 2873",  "S275J2", 254.5,  254.5, 8, "Kesim + Konik Bükme"),
    ("5B",  1, "⊏ 1500×8 ......... 2873",  "S275J2", 254.5,  254.5, 8, "Kesim + Konik Bükme"),
    ("6",   3, "⊏ 1500×8 ......... 2545.5","S275J2", 208.0,  624.0, 8, "Kesim + Konik Bükme"),
    ("6A",  1, "⊏ 1500×8 ......... 2545.5","S275J2", 190.0,  190.0, 8, "Kesim + Konik Bükme"),
    ("7",   4, "⊏ 1000×8 ......... 1824.5","S275J2",  90.7,  362.8, 8, "Kesim + Konik Bükme"),
    ("8",   4, "⊏ 1056.5×8 ...... 874.5", "S275J2",  31.0,  124.0, 8, "Kesim + Konik Bükme"),
    ("9",   2, "⊏ 1500×8 ......... 2910.5","S275J2", 272.8,  545.6, 8, "Kesim + Konik Bükme"),
]
for i, row in enumerate(tavan, start=3):
    DATA(ws5, i, list(range(1,9)), row, alt=(i%2==0))
TOT(ws5, 18, 1, 6, "TOPLAM AĞIRLIK:", 8147.0)
for c in [7,8]: C(ws5, 18, c, "", H1, fw, CA, TB)
sw(ws5,[8,8,42,14,14,14,12,22])

# ═══════════════════════════════════════════════════════════════════
# SAYFA 6 – NOZUL MALZEME LİSTELERİ (Detay)
# ═══════════════════════════════════════════════════════════════════
ws6 = wb.create_sheet("Nozul Malzeme Detay"); ws6.sheet_view.showGridLines = False
TITLE(ws6, 1, 1, 8, "NOZUL DETAY MALZEME LİSTELERİ")

all_nozzles = [
    ("(12\") N1 = PETROL GİRİŞ NOZULU", 1, "1 adet", [
        (1,1,'12" SCH XS BORU ..... 223',"ASTM A 106 Gr.B",21.5,21.5),
        (2,1,'12" 150# SO. RF. FLANS',"ASTM A 105",29.0,29.0),
        (3,1,"⊏ 8×327×685  PLAKA","S275J2",18.0,18.0),
    ], 68.5),
    ("(12\") N4 = PETROL ÇIKIŞ NOZULU", 1, "1 adet — BOTTOM (Taban)", [
        (1,1,'12" SCH XS BORU ..... 223',"ASTM A 106 Gr.B",21.5,21.5),
        (2,1,'12" 150# SO. RF. FLANS',"ASTM A 105",29.0,29.0),
        (3,1,"⊏ 6×327×685  PLAKA","S275J2",14.0,14.0),
    ], 64.5),
    ("(8\") N8 = KÖPÜK YAPICI NOZULU", 1, "1 adet", [
        (1,1,'8" SCH XS BORU ..... 221',"ASTM A 106 Gr.B",11.4,11.4),
        (2,1,'8" 150# SO. RF. FLANS',"ASTM A 105",13.5,13.5),
        (3,1,"⊏ 6×485×222  PLAKA","S275J2",7.0,7.0),
    ], 31.9),
    ("(6\") N7A/B = DRAIN NOZULU", 2, "2 adet imal edilecektir", [
        (1,1,'6" 150# SO. RF. FLANS',"ASTM A 105",5.9,5.9),
        (2,1,'6" SCH XS BORU ..... 360',"ASTM A 106 Gr.B",9.6,9.6),
        (3,1,'6" SCH XS BORU ..... 870',"ASTM A 106 Gr.B",23.2,23.2),
        (4,1,'6" SCH XS DİRSEK 90° LR BW.',"ASTM A 234WPB",5.4,5.4),
        (5,1,'6" SCH XS BORU ..... 540',"ASTM A 106 Gr.B",14.4,14.4),
        (6,1,"⊏ 50×10 ............ 150","S275J2",0.6,0.6),
        (7,4,"M12 SOMUN","A 194 Gr.2H","–","–"),
        (8,1,"M12×594  U CİVATA","A 193 Gr.B7","–","–"),
        (9,1,"L 60×60×6 ..... 140","S235JR",0.76,0.76),
        (10,1,"L 60×60×6 ..... 280","S235JR",1.52,1.52),
        (11,1,"⊏ 90×10 ............ 90","S275J2",0.65,0.65),
    ], 62.0),
    ("(1.5\") N9 = SICAKLIK ÖLÇÜMÜ (TI)", 1, "1 adet", [
        (1,1,'1½" SCH XS BORU ..... 165',"ASTM A 106 Gr.B",0.70,0.70),
        (2,1,'1½" 150# SO. RF. FLANS',"ASTM A 105",1.95,1.95),
    ], 2.65),
    ("(1\") N18 = YÜKSEK SEVİYE SENSÖRÜ", 1, "1 adet", [
        (1,1,'1" SCH XS BORU ..... 165',"ASTM A 106 Gr.B",0.70,0.70),
        (2,1,'1" 150# SO. RF. FLANS',"ASTM A 105",1.0,1.0),
    ], 1.7),
    ("(12\") N21 = HAVA NOZULU", 4, "4 adet imal edilecektir — ROOF", [
        (1,1,'12" SCH XS BORU ..... 290',"ASTM A 106 Gr.B",28.0,28.0),
        (2,1,'12"-90° Elbow LR BW ASME B16.5 SCH40',"ASTM A 234 Gr.WPB",15.1,15.1),
        (3,1,"⊏ 8×327×685  PLAKA","S275J2",18.0,18.0),
    ], 61.1),
]

row = 2
for noz, imalat, not_t, items, toplam in all_nozzles:
    SUB(ws6, row, 1, 8, f"{noz}   |   İmalat: {imalat}   |   {not_t}")
    row += 1
    HDR(ws6, row, list(range(1,9)),
        ["POZ","ADET","AÇIKLAMA","MALZEME","BİRİM\nAĞIRLIK(kg)","TOPLAM\nAĞIRLIK(kg)",
         "OP. TÜRÜ","NOT"], h=30)
    row += 1
    for i,(poz,adet,acik,malz,birim,top) in enumerate(items):
        op=""
        if "BORU" in acik and "YUVARLAK" not in acik: op="Kesim+Kaynak"
        elif "FLANS" in acik: op="Kaynak"
        elif "PLAKA" in acik or "⊏" in acik: op="Kesim+Kaynak"
        elif "DİRSEK" in acik or "Elbow" in acik: op="Kaynak"
        elif any(x in acik for x in ["SOMUN","CİVATA","NPT","L 6"]): op="Montaj"
        DATA(ws6, row, list(range(1,9)), [poz,adet,acik,malz,birim,top,op,""], alt=(i%2==1))
        row += 1
    TOT(ws6, row, 1, 6,
        f"TOPLAM (1 adet): {toplam} Kg  |  {imalat} → {round(toplam*imalat,1)} Kg",
        round(toplam*imalat,1))
    for c in [7,8]: C(ws6, row, c, "", H1, fw, CA, TB)
    row += 2
sw(ws6,[8,8,42,20,14,14,20,14])

# ═══════════════════════════════════════════════════════════════════
# SAYFA 7 – MANHOL + TEMİZLEME NOZULU
# ═══════════════════════════════════════════════════════════════════
ws7 = wb.create_sheet("Manhol+Temizleme"); ws7.sheet_view.showGridLines = False
TITLE(ws7, 1, 1, 7, "MANHOL VE TEMİZLEME NOZULU MALZEME LİSTELERİ")

manhol_all = [
    ('24" GÖVDE MANHOLU (M1)  —  1 Adet', [
        (1,1,'⊂ ø176×8 .........1885',"S275J2",20.8,20.8,"Silindirik Bükme+Kaynak"),
        (2,1,'⊏ ø820\\ø616×10  HALKA',"S275J2",18.0,18.0,"Kesim+Kaynak"),
        (3,1,"ø820×12  DAİRE PLAKA","S275J2",49.8,49.8,"Kesim"),
        (4,1,'⊏ 1255×8 .........1525',"S275J2",69.5,69.5,"Silindirik Bükme+Kaynak"),
        (5,1,"ø735×ø610×3  GASKET","KLİNGERİT","–","–","Montaj"),
        (6,2,"ø16×318  YUVARLAK BORU","S235JR",1.0,2.0,"Kesim"),
        (7,1,"ø40×1083  BORU","S235JR",11.0,11.0,"Kesim+Kaynak"),
        (8,1,'⊏ 100×20 .........200',"S275J2",1.6,1.6,"Kesim"),
        (9,1,'⊏ 100×20 .........200',"S275J2",1.6,1.6,"Kesim"),
        (10,1,"⊏ ø70/ø42×30","S275J2",1.0,1.0,"Kesim"),
        (11,1,'⊏ 60×25 .........70',"S275J2",1.0,1.0,"Kesim"),
        (12,1,'⊏ 67×14 .........80',"S275J2",0.5,0.5,"Kesim"),
        (13,2,'⊏ 50×14 .........56',"S275J2",0.5,1.0,"Kesim"),
        (14,1,'⊏ 46×14 .........50',"S275J2",0.5,0.5,"Kesim"),
        (15,1,"ø24×200  YUVARLAK BORU","S235JR",1.0,1.0,"Kesim"),
        (16,28,"CİVATA M20×75","DIN 601","–","–","Montaj"),
        (17,1,"CİVATA M20×110","DIN 601","–","–","Montaj"),
        (18,2,"SOMUN M24","DIN 555","–","–","Montaj"),
        (19,58,"SOMUN M20","DIN 555","–","–","Montaj"),
        (20,58,"RONDELA D 22/37  S=3","DIN 126 Gr.C","–","–","Montaj"),
        (21,1,'1/4" NPT TAPA',"–","–","–","Montaj"),
    ], 180.0),
    ('24" TAVAN MANHOLU (M2)  —  1 Adet', [
        (1,1,'⊏ 400×8 .........1910',"S275J2",48.0,48.0,"Silindirik Bükme+Kaynak"),
        (2,1,'⊏ ø750/ø616×8  HALKA',"S275J2",10.0,10.0,"Kesim+Kaynak"),
        (3,1,"ø750×8  DAİRE PLAKA","S275J2",28.0,28.0,"Kesim"),
        (4,1,'⊏ ø1150/ø616×8  HALKA',"S275J2",46.5,46.5,"Kesim+Kaynak"),
        (5,1,"ø750×ø616×1.5  GASKET","KLİNGERİT","–","–","Montaj"),
        (6,2,"ø16×318  YUVARLAK BORU","S235JR",1.0,2.0,"Kesim"),
        (7,20,"CİVATA M16×50","DIN 601","–","–","Montaj"),
        (8,20,"SOMUN M16","DIN 555","–","–","Montaj"),
        (9,20,"RONDELA D 17/30  S=3","DIN 126 Gr.C","–","–","Montaj"),
        (10,1,'1/4" NPT TAPA',"–","–","–","Montaj"),
    ], 132.5),
    ('24" TEMİZLEME NOZULU (M3)  —  1 Adet', [
        (1,1,"⊏ 433×8 ......... 1585","S275J2",43.0,43.0,"Silindirik Bükme+Kaynak"),
        (2,1,"⊏ 808×13 ......... 2124","S275J2",145.6,145.6,"Silindirik Bükme+Kaynak"),
        (3,1,"⊏ 915×8 ......... 1830","S275J2",68.5,68.5,"Silindirik Bükme+Kaynak"),
        (4,1,"⊏ 807×13 ......... 814","S275J2",24.3,24.3,"Kesim"),
        (5,1,"⊏ 807×814×3  NON ASBESTOS GASKET","KLİNGERİT","–","–","Montaj"),
        (6,1,"⊏ 807×13 ......... 814","S275J2",60.0,60.0,"Kesim"),
        (7,2,"ø16×316  ÇUBUK DEMİR","S275J2",1.0,2.0,"Kesim"),
        (8,36,"M20×70  CİVATA","DIN 601","–","–","Montaj"),
        (9,74,"M20  SOMUN","DIN 555","–","–","Montaj"),
        (10,37,"D 22/37  S=3  RONDELA","DIN 126","–","–","Montaj"),
        (11,1,"ø40×1179  ÇUBUK DEMİR","S275J2",11.6,11.6,"Kesim"),
        (12,1,"⊏ ø42×ø80×25","S275J2",0.6,0.6,"Kesim"),
        (13,1,"⊏ 60×60×25","S275J2",0.7,0.7,"Kesim"),
        (14,1,"ø20×451  ÇUBUK DEMİR","S275J2",1.3,1.3,"Kesim"),
        (15,1,"⊏ 200×112×10","S275J2",1.8,1.8,"Kesim"),
        (16,1,"⊏ 655×100×15","S275J2",7.4,7.4,"Kesim"),
        (17,1,"⊏ ø80×ø42×10","S275J2",0.2,0.2,"Kesim"),
        (18,1,"ø8×75  PİM","S275J2","–","–","Montaj"),
        (19,1,'1/4" NPT TAPA',"–","–","–","Montaj"),
    ], 367.0),
]

row = 2
for mname, items, toplam in manhol_all:
    SUB(ws7, row, 1, 7, mname); row += 1
    HDR(ws7, row, list(range(1,8)),
        ["POZ","ADET","AÇIKLAMA","MALZEME","BİRİM\nAĞIRLIK(kg)","TOPLAM\nAĞIRLIK(kg)","İŞLEM TÜRÜ"], h=30)
    row += 1
    for i, it in enumerate(items):
        DATA(ws7, row, list(range(1,8)), it, alt=(i%2==1)); row += 1
    TOT(ws7, row, 1, 6, f"TOPLAM AĞIRLIK: {toplam} Kg", toplam)
    C(ws7, row, 7, "", H1, fw, CA, TB); row += 2
# ── Tank İsim Plakası & Merdiven bölümlerini ws7'ye ekle ────────────

# Tank İsim Plakası
ti_items = [
    (1,1,"⊏ 215×6 ......... 230","S275J2",2.33,2.33,"Kesim"),
    (2,1,"⊏ 230×6 ......... 230","S275J2",2.33,2.33,"Kesim"),
    (3,2,"⊏ 25×6 ......... 230","S275J2",0.26,0.52,"Kesim"),
    (4,4,'BOLT/NVT NO.5-40 UNC',"304 S.S","–","–","Montaj"),
]
SUB(ws7, row, 1, 7, "TANK İSİM PLAKASI  (1 Adet)"); row += 1
HDR(ws7, row, list(range(1,8)),
    ["POZ","ADET","AÇIKLAMA","MALZEME","BİRİM\nAĞIRLIK(kg)","TOPLAM\nAĞIRLIK(kg)","İŞLEM TÜRÜ"], h=30)
row += 1
for i, it in enumerate(ti_items):
    DATA(ws7, row, list(range(1,8)), it, alt=(i%2==1)); row += 1
TOT(ws7, row, 1, 6, "TOPLAM AĞIRLIK: 5.18 Kg", 5.18)
C(ws7, row, 7, "", H1, fw, CA, TB); row += 2

# Merdiven (Staircase) Malzeme Listesi
merd_items = [
    (1, 1,"⊏ 8×200 ......... 7550.5","S235JR", 94.8,  94.8,"Kesim"),
    (2, 1,"⊏ 8×200 ......... 972.5", "S235JR", 12.2,  12.2,"Kesim"),
    (3, 1,"⊏ 8×200 ......... 6190",  "S235JR", 77.7,  77.7,"Kesim"),
    (4, 1,"⊏ 8×200 ......... 899.5", "S235JR", 11.3,  11.3,"Kesim"),
    (5, 1,"⊏ 8×200 ......... 8057.5","S235JR",101.4, 101.4,"Kesim"),
    (6, 1,"⊏ 8×200 ......... 1069.5","S235JR", 13.4,  13.4,"Kesim"),
    (7, 1,"⊏ 8×200 ......... 6603.5","S235JR", 82.9,  82.9,"Kesim"),
    (8, 1,"⊏ 8×200 ......... 1002.5","S235JR", 12.6,  12.6,"Kesim"),
    (9,37,"L60×6 ......... 1079",    "S235JR",  5.4, 248.4,"Kesim"),
    (10,1,"⊏ 8×200 ......... 750",   "S235JR",  9.4,   9.4,"Kesim"),
    (11,1,"NPU100 ......... 1080",   "S235JR", 11.5, 138.0,"Kesim"),
    (12,12,"NPU100 ......... 1178.5","S235JR", 12.5, 150.0,"Kesim"),
    (13,24,"⊏ 8×100 ......... 260",  "S355J2",  1.6,  38.4,"Kesim"),
    (14,22,"L60×6 ......... 325.5",  "S235JR",  1.7,  37.4,"Kesim"),
    ("14A",2,"L60×6 ......... 325.5","S235JR",  1.7,   3.4,"Kesim"),
    (15,12,"⊏ 8×70 ......... 180",   "S235JR",  0.8,   9.6,"Kesim"),
    (16, 2,"L60×6 ......... 899.5",  "S235JR",  4.9,   9.8,"Kesim"),
    (17, 4,"L60×6 ......... 630",    "S235JR",  5.4,  21.6,"Kesim"),
    (18, 2,"L60×6 ......... 1002.5", "S235JR",  5.5,  11.0,"Kesim"),
    (19, 5,"L60×6 ......... 1238",   "S235JR",  5.4,  27.0,"Kesim"),
    (20,26,"L60×6 ......... 1020",   "S235JR",  5.4, 140.4,"Kesim"),
    (21, 1,"HANDRAİL (1 1/4\" STD.SCH PIPE) 70000","S235JR",178.5,178.5,"Kesim+Bükme"),
    (22, 1,"⊏ 5×50 ......... 70000", "S235JR",137.5, 137.5,"Kesim"),
    (23, 1,"⊏ 5×150 ......... 34000","S235JR",200.0, 200.0,"Kesim"),
    # Bağlantı elemanları
    (None,48,"M10×30  BOLT", "A193 GR B7","–","–","Montaj"),
    (None,48,"M10  NUT",     "A194 GR 2H","–","–","Montaj"),
    (None,48,"M10  WASHER",  "A194 GR 2H","–","–","Montaj"),
    (None,136,"M12×40  BOLT","A193 GR B7","–","–","Montaj"),
    (None,136,"M12  NUT",    "A194 GR 2H","–","–","Montaj"),
    (None,136,"M12  WASHER", "A194 GR 2H","–","–","Montaj"),
]
SUB(ws7, row, 1, 7, "MERDİVEN + KORKULUK  (1 Komple)"); row += 1
HDR(ws7, row, list(range(1,8)),
    ["POZ","ADET","AÇIKLAMA","MALZEME","BİRİM\nAĞIRLIK(kg)","TOPLAM\nAĞIRLIK(kg)","İŞLEM TÜRÜ"], h=30)
row += 1
for i, it in enumerate(merd_items):
    DATA(ws7, row, list(range(1,8)), it, alt=(i%2==1)); row += 1
TOT(ws7, row, 1, 6, "TOPLAM AĞIRLIK: 1777.0 Kg", 1777.0)
C(ws7, row, 7, "", H1, fw, CA, TB); row += 2

# Mekanik Samandıra (N14) Malzeme Listesi
saman_items = [
    (1,  1, "ø21 BORU ......... 26",             "A 283 Gr.C","–","–",""),
    (2,  2, "ø27 BORU ......... 130",             "S.S. A316L","–","–",""),
    (3,  2, "ø400×6 PLAKA",                       "S.S. A316L","–","–",""),
    (4,  1, "⊏ 300×6×1257 PLAKA",                "S.S. A316L","–","–",""),
    (5,  2, "⊏ 100×6×100 PLAKA",                 "S.S. A316L","–","–",""),
    (6, "–","⊏ 30×5×80 PLAKA",                   "A 283 Gr.C","–","–",""),
    (7, "–","⊏ 50×10×298 PLAKA",                 "A 283 Gr.C","–","–",""),
    (8, "–","⊏ 50×10×292 PLAKA",                 "A 283 Gr.C","–","–",""),
    (9, "–","⊏ 50×5×268 PLAKA",                  "A 283 Gr.C","–","–",""),
    (10, 1, "⊏ 200×2×9000 PLAKA  (Seviye Göstergesi)","A 283 Gr.C","–","–",""),
    (11,"–","⊏ 50×2×200 PLAKA",                  "A 283 Gr.C","–","–",""),
    (12, 3, "L 60×60×6×9514 PROFİL",             "A 283 Gr.C","–","–",""),
    (13, 1, "[ 140×7×625 PROFİL",                "A 283 Gr.C","–","–",""),
    (14, 2, "ø50 MAKARA",                         "A 283 Gr.C","–","–",""),
    (15, 4, "⊏ 30×5×60 PLAKA",                   "A 283 Gr.C","–","–",""),
    (16, 1, "⊏ 150×15×300 PLAKA",               "A 283 Gr.C","–","–",""),
    (17, 8, "⊏ 30×3×60 PLAKA",                   "A 283 Gr.C","–","–",""),
    (18,"–","ø5×20 mm KONİK VİDA / PUL",         "A 283 Gr.C","–","–",""),
    (19, 2, "M8 SAPLAMA-SOMUN",                   "A 283 Gr.C","–","–",""),
    (20, 2, "M18 SAPLAMA-SOMUN",                  "A 283 Gr.C","–","–",""),
    (21, 1, "ø8 ÇELİK HALAT",                    "–","–","–",""),
    (22, 2, "ø12 ÇELİK HALAT",                   "–","–","–",""),
    (23, 4, "ø12 DEMİR",                          "S.S. A316L","–","–",""),
    (24, 8, "HALAT KLEMENSİ",                     "S.S. A316L","–","–",""),
    (25,"–","L 60×60×6×195 PROFİL",              "A 283 Gr.C","–","–",""),
]
SUB(ws7, row, 1, 7, "MEKANİK SAMANDIRA (N14)  —  1 Adet  (A283 Gr.C + S.S. A316L)"); row += 1
HDR(ws7, row, list(range(1,8)),
    ["POZ","ADET","AÇIKLAMA","MALZEME","BİRİM\nAĞIRLIK(kg)","TOPLAM\nAĞIRLIK(kg)","NOT"], h=30)
row += 1
for i, it in enumerate(saman_items):
    DATA(ws7, row, list(range(1,8)), it, alt=(i%2==1)); row += 1
NOTE(ws7, row, 1, 7, "NOT: Malzeme listesi 1 adet samandıra için verilmiştir. Ağırlık değerleri çizimde belirtilmemiştir.")
row += 2

sw(ws7,[8,8,44,20,14,14,22])

# ═══════════════════════════════════════════════════════════════════
# SAYFA 8 – KONİK ÇATI İSKELETİ (Yapısal)  — Çizim image 5 verileri
# ═══════════════════════════════════════════════════════════════════
ws8 = wb.create_sheet("Konik Çatı İskeleti"); ws8.sheet_view.showGridLines = False
TITLE(ws8, 1, 1, 9, "KONİK ÇATI YAPISAL İSKELET MALZEME LİSTESİ  (TOPLAM: 5454.0 kg)")

# Her grup: (grup adı, toplam adet, [(poz, adet/asm, profil, malz, boy_mm, birim_kg, 1asm_toplam)])
catim_grp = [
    ("CK1  —  UNP240 Rafter  (12 çatım adedi)", 12, [
        ("P1",  1, "UNP240",         "S235JR", 5685, 155.5, 155.5),
        ("P20", 2, "PL10×113  Plaka","S235JR",  221,   1.8,   3.6),
        ("P21", 2, "PL10×95   Plaka","S235JR",  152,   1.3,   2.5),
        ("P22", 1, "PL10×221  Plaka","S235JR",  193,   2.7,   2.7),
        ("P23", 1, "PL10×162  Plaka","S235JR",  216,   1.9,   1.9),
    ]),
    ("CK2  —  UNP240 Rafter  (12 çatım adedi)", 12, [
        ("P1",  1, "UNP240",         "S235JR", 5685, 155.5, 155.5),
        ("P20", 2, "PL10×113  Plaka","S235JR",  221,   1.8,   3.6),
        ("P21", 2, "PL10×95   Plaka","S235JR",  152,   1.3,   2.5),
        ("P24", 1, "PL10×189  Plaka","S235JR",  232,   2.5,   2.5),
        ("P25", 1, "PL10×139  Plaka","S235JR",  244,   2.0,   2.0),
    ]),
    ("CK3  —  UNP200 Purlin  (24 adet)", 24, [
        ("P5",  1, "UNP200",         "S235JR", 1227,  31.0,  31.0),
    ]),
    ("CK4  —  UNP200 Purlin  (24 adet)", 24, [
        ("P6",  1, "UNP200",         "S235JR",  677,  17.1,  17.1),
    ]),
    ("KL1  —  Merkezi Kolon PIP 323.9×8  (1 adet)", 1, [
        ("P2",  1, "PIP 323.9×8  Boru",  "S235JR", 9369, 580.2, 580.2),
        ("P9",  1, "PL25×1800  Plaka",   "S235JR", 1800, 497.3, 497.3),
        ("P10", 1, "PL15×250   Plaka",   "S235JR", 4618, 137.1, 137.1),
        ("P11", 1, "PL10×200   Plaka",   "S235JR", 5341,  84.2,  84.2),
        ("P12", 1, "PL25×750   Plaka",   "S235JR",  750,  56.0,  56.0),
        ("P16",24, "PL10×150   Plaka",   "S235JR",  250,   2.3,  55.3),
        ("P17", 4, "PL10×513   Plaka",   "S235JR",  675,  13.3,  53.2),
        ("P18", 1, "PL20×600   Plaka",   "S235JR",  600,  43.9,  43.9),
    ]),
    ("W1  —  Ayak Gusset Grubu  (24 adet)", 24, [
        ("P3",  1, "PL10×300  Plaka",    "S235JR",  300,   7.1,   7.1),
        ("P13", 1, "PL10×200  Plaka",    "S235JR",  250,   3.9,   3.9),
        ("P14", 1, "PL10×200  Plaka",    "S235JR",  100,   1.6,   1.6),
        ("P15", 1, "PL10×248  Plaka",    "S235JR",  299,   4.2,   4.2),
    ]),
    ("W2  —  Üst Halka Plakası  (1 adet)", 1, [
        ("P4",  1, "PL15×750  Plaka",    "S235JR",  750,  51.6,  51.6),
        ("P19", 1, "PL15×100  Plaka",    "S235JR", 1916,  23.0,  23.0),
    ]),
    ("YC1  —  Yatay Çelik 1  (12 adet)", 12, [
        ("P7",  1, "L60×5  Köşebent",    "S235JR", 1931,  15.6,  15.6),
    ]),
    ("YC2  —  Yatay Çelik 2  (12 adet)", 12, [
        ("P8",  1, "L60×5  Köşebent",    "S235JR", 1954,  19.2,  19.2),
    ]),
]

row8 = 2
for grp_name, mult, items in catim_grp:
    SUB(ws8, row8, 1, 9, grp_name); row8 += 1
    HDR(ws8, row8, list(range(1,10)),
        ["POZ","ADET\n(1 asm)","PROFİL / AÇIKLAMA","MALZEME",
         "BOY\n(mm)","BİRİM\nKg","1 ASM\nTOPLAM kg",f"×{mult}\nTOPLAM kg","İŞLEM"], h=35)
    row8 += 1
    asm_tot = 0
    for i,(p,a,ac,m,b,birim,asm1) in enumerate(items):
        all_tot = round(asm1 * mult, 1)
        DATA(ws8, row8, list(range(1,10)),
             [p, a, ac, m, b, birim, round(asm1,1), all_tot, "Kesim+Kaynak"], alt=(i%2==1))
        asm_tot += asm1; row8 += 1
    tot_all = round(asm_tot * mult, 1)
    TOT(ws8, row8, 1, 8, f"ALT TOPLAM  (×{mult}):", tot_all)
    C(ws8, row8, 9, "", H1, fw, CA, TB); row8 += 2

NOTE(ws8, row8, 1, 9,
     "NOT: Çizimden alınan TOPLAM AĞIRLIK = 5454.0 Kg  (1 komple çatı iskeleti için verilmiştir).")
row8 += 2

# Çatım Civata-Somun-Rondela Listesi
SUB(ws8, row8, 1, 9, "ÇATIM CİVATA-SOMUN-RONDELA LİSTESİ  (Standart: DIN 7990 / 7989 / 4032)"); row8 += 1
HDR(ws8, row8, [1,2,3,4,5,6],
    ["CİVATA TANIMI","ÇATIM-1","ÇATIM-2","MALZEME\nSINIFI","ADET","STANDART"], h=25)
row8 += 1
bolt_list = [
    ("BOLT 16×45", "CK3", "CK1", "8.8", 48, "7990/7989/4032"),
    ("BOLT 16×45", "CK3", "CK2", "8.8", 48, "7990/7989/4032"),
    ("BOLT 16×45", "CK4", "CK1", "8.8", 48, "7990/7989/4032"),
    ("BOLT 16×45", "CK4", "CK2", "8.8", 48, "7990/7989/4032"),
    ("BOLT 20×50", "YC1", "CK1", "8.8", 24, "7990/7989/4032"),
    ("BOLT 20×50", "YC1", "CK2", "8.8", 24, "7990/7989/4032"),
    ("BOLT 20×50", "YC2", "CK1", "8.8", 24, "7990/7989/4032"),
    ("BOLT 20×50", "YC2", "CK2", "8.8", 24, "7990/7989/4032"),
]
for i, rd in enumerate(bolt_list):
    fill = ALT if i%2==1 else WH
    for c, v in enumerate(rd, 1):
        C(ws8, row8, c, v, fill, fn, CA, TB)
    row8 += 1
# Totals
b16_tot = 4*48; b20_tot = 4*24
SUB(ws8, row8, 1, 9, f"TOPLAM: BOLT 16×45 → {b16_tot} adet  |  BOLT 20×50 → {b20_tot} adet  (her civata için 1 somun + 2 rondela)"); row8 += 1

sw(ws8,[8,10,30,12,10,12,14,16,18])

# ═══════════════════════════════════════════════════════════════════
# SAYFA 9 – TOPLAM AĞIRLIK ÖZETİ
# ═══════════════════════════════════════════════════════════════════
ws9 = wb.create_sheet("Toplam Ağırlık Özeti"); ws9.sheet_view.showGridLines = False
TITLE(ws9, 1, 1, 6, "7000 Bbl ÜRETİM TANKI  —  TOPLAM AĞIRLIK ÖZETİ  (1 TANK)")
HDR(ws9, 2, list(range(1,7)),
    ["BÖLÜM","AÇIKLAMA","AĞIRLIK\n(kg)","ADET\n(Tank)","TOPLAM\nAĞIRLIK (kg)","KAYNAK ÇİZİM"])

ozet = [
    ("GÖVDE","Silindirik Gövde Sacları  (8mm + 6mm)",16513.0,1,16513.0,"KEY199MECEQP2001"),
    ("TABAN","Taban Plakaları  (10mm + 12mm annüler)",11176.7,1,11176.7,"KEY199MECEQP2003"),
    ("TABAN","Drenaj Çukuru",76.1,1,76.1,"KEY199MECEQP2003"),
    ("TAVAN","Konik Tavan Sacları  (8mm)",8147.0,1,8147.0,"KEY199MECEQP2002"),
    ("TAVAN İSK.","Konik Çatı Yapısal (CK1-CK4+KL1+W1+W2+YC1+YC2)",5454.0,1,5454.0,"KEY199MECEQP2004"),
    ("NOZUL N1","12\" Petrol Giriş Nozulu",68.5,1,68.5,"MKD-205"),
    ("NOZUL N4","12\" Petrol Çıkış Nozulu (BOTTOM)",64.5,1,64.5,"MKD-205"),
    ("NOZUL N7A/B","6\" Drain Nozulu",62.0,2,124.0,"MKD-205"),
    ("NOZUL N8","8\" Köpük Yapıcı Nozulu",31.9,1,31.9,"MKD-205"),
    ("NOZUL N9","1.5\" Sıcaklık Ölçüm Nozulu",2.65,1,2.65,"MKD-205"),
    ("NOZUL N18","1\" Yüksek Seviye Sensörü",1.7,1,1.7,"MKD-205"),
    ("NOZUL N19","6\" Radar Ölçüm Nozulu","~62",1,"~62","MKD-205"),
    ("NOZUL N21","12\" Hava Nozulu",61.1,4,244.4,"MKD-205"),
    ("MANHOL M1","24\" Gövde Manholu",180.0,1,180.0,"MKD-204"),
    ("MANHOL M2","24\" Tavan Manholu",132.5,1,132.5,"MKD-204"),
    ("MANHOL M3","24\" Temizleme Manholu",367.0,1,367.0,"MKD-211"),
    ("ANKRAJ","Ankraj (16 komple × 13.5 Kg)",13.5,16,216.0,"KEY199MECEQP2001"),
    ("MERDİVEN","Merdiven + Korkuluk  (1 Komple)",1777.0,1,1777.0,"MKD-210"),
    ("İSİM PLAK.","Tank İsim Plakası (TI-1 ~ TI-4)",5.18,1,5.18,"–"),
    ("SAMANDIRA","Mekanik Samandıra (N14)","Bekleniyor","1","Bekleniyor","MKD-206"),
]

known = sum(v[2] for v in ozet if isinstance(v[2], (int,float)))
known_tot = sum(v[4] for v in ozet if isinstance(v[4], (int,float)))

for i, row_d in enumerate(ozet, start=3):
    bolum,acik,ag,adet,top,ciz = row_d
    fill = ALT if i%2==0 else WH
    s_fill = GRN if isinstance(ag,(int,float)) else YEL
    for c,v in enumerate([bolum,acik,ag,adet,top,ciz],1):
        cl = ws9.cell(row=i, column=c, value=v)
        cl.fill = s_fill if c==3 else fill
        cl.font = fn; cl.alignment = CA if c!=2 else LA; cl.border = TB

r = 3+len(ozet)+1
ws9.merge_cells(start_row=r,start_column=1,end_row=r,end_column=2)
cl=ws9.cell(row=r,column=1,value="BİLİNEN TOPLAM (yaklaşık, eksikler hariç):")
cl.fill=H1;cl.font=fw;cl.alignment=RA;cl.border=TB
C(ws9,r,3,f"~{known:.0f}",H1,fw,CA,TB)
C(ws9,r,4,"–",H1,fw,CA,TB)
C(ws9,r,5,f"~{known_tot:.0f}",H1,fw,CA,TB)
C(ws9,r,6,"",H1,fw,CA,TB)
sw(ws9,[18,44,14,10,16,24])

# ═══════════════════════════════════════════════════════════════════
# SAYFA 10 – MALİYET TAHMİN ÇİZELGESİ
# ═══════════════════════════════════════════════════════════════════
ws10 = wb.create_sheet("Maliyet Tahmini"); ws10.sheet_view.showGridLines = False
TITLE(ws10, 1, 1, 10, "7000 Bbl ÜRETİM TANKI  —  MALİYET TAHMİN ÇİZELGESİ  (1 TANK)")

# Malzeme maliyeti bölümü
SUB(ws10, 2, 1, 10, "A. MALZEME MALİYETİ")
HDR(ws10, 3, list(range(1,11)),
    ["BÖLÜM","AÇIKLAMA","MALZEME\nSINIFI","KALINLIK\n(mm)","AĞIRLIK\n(kg)",
     "BİRİM FİYAT\n($/kg) ★","MALZEME\nMALİYETİ ($)","İŞÇİLİK\nFİYAT($/kg) ★",
     "İŞÇİLİK\nMALİYETİ ($)","TOPLAM ($)"], h=40)

cost_items = [
    ("GÖVDE","Gövde Sacları (Shell)","S275J2","6-8",16513.0),
    ("TABAN","Taban Plakaları (Orta 10mm + Annüler 12mm + Strip 5mm)","S275J2","5–12",11176.7),
    ("TABAN","Drenaj Çukuru (Sump) — Ayrı İmalat Kalemi","S275J2","–",76.1),
    ("TAVAN","Konik Tavan Sacları","S275J2","8",8147.0),
    ("TAVAN İSK.","Çatı İskeleti (UNP+PL+PIP)","S235JR","–",5454.0),
    ("NOZULLAR","Tüm Nozullar + RF Takviye Sacları (N1+N4+N7A/B+N8+N9+N18+N19+N21)","A106/A105/S275J2","–",
     round(68.5+64.5+62.0*2+31.9+2.65+1.7+62+61.1*4,1)),
    ("MANHOLLER","M1+M2+M3 Manhol","S275J2","–",180+132.5+367),
    ("ANKRAJ","Ankraj (16 Komple)","AISI C1030","–",216.0),
    ("MERDİVEN","Merdiven + Korkuluk","S235JR","–",1777.0),
]

for i, (bolum,acik,malz,thk,ag) in enumerate(cost_items, start=4):
    fill = ALT if i%2==0 else WH
    for c,v in enumerate([bolum,acik,malz,thk,ag,"","","","",""],1):
        cl = ws10.cell(row=i, column=c, value=v)
        cl.fill = YEL if c in [6,8] else fill
        cl.font = fn; cl.alignment = CA if c!=2 else LA; cl.border = TB
        if c==7 and isinstance(ag,(int,float)):
            cl.value = f"=E{i}*F{i}"
        elif c==9 and isinstance(ag,(int,float)):
            cl.value = f"=E{i}*H{i}"
        elif c==10 and isinstance(ag,(int,float)):
            cl.value = f"=G{i}+I{i}"

r = 4 + len(cost_items) + 1

# Bükme / Haddeleme / Kesim bölümü
SUB(ws10, r, 1, 10, "B. İMALAT İŞÇİLİĞİ (Kesim / Bükme / Haddeleme)"); r+=1
HDR(ws10, r, list(range(1,8)),
    ["İŞLEM TÜRÜ","AÇIKLAMA","MİKTAR","BİRİM","BİRİM FİYAT ★","TUTAR ($)","NOT"], h=30)
r+=1
imalat_iscilik=[
    ("Silindirik Bükme","Gövde Sacları 6mm","31","adet","","","POZ 3×30 + POZ 3A×1  |  1480×6000mm"),
    ("Silindirik Bükme","Gövde Sacları 8mm","7","adet","","","POZ 1-1E×6 (6000mm) + POZ 2×1 (3997mm)"),
    ("Konik Bükme","Tavan Sacları 8mm","35","adet","","","POZ 1-9 toplam 35 parça, 1500×6000mm stokundan"),
    ("Kesim (Plazma/Lazer)","Tüm Ana Sac Plakalar (Gövde+Taban+Tavan+Sump)",35913,"kg","","","16513+11253+8147 = 35913 kg S275J2"),
    ("Profil Kesim","Yapısal Çelik Çatı İskeleti (UNP+PL+PIP+L)",5454,"kg","","","Çizim TOPLAM = 5454 kg S235JR"),
    ("Boru Kesim","PIP 323.9×8 + Nozul Boruları (A106 Gr.B)",700,"kg","","","Yaklaşık; kesin ağırlık nozul detay listesinden"),
]
for i,rd in enumerate(imalat_iscilik):
    fill=ALT if i%2==1 else WH
    for c,v in enumerate(rd,1):
        cl=ws10.cell(row=r,column=c,value=v)
        cl.fill=YEL if c==5 else fill; cl.font=fn; cl.alignment=CA if c!=2 else LA; cl.border=TB
    r+=1
r+=1

# Kaynak bölümü
SUB(ws10, r, 1, 10, "C. KAYNAK MALİYETİ"); r+=1
HDR(ws10, r, list(range(1,8)),
    ["KAYNAK TÜRÜ","AÇIKLAMA","MİKTAR (m)","KAYNAK KALİTESİ","BİRİM FİYAT ★","TUTAR ($)","NOT"], h=30)
r+=1
kaynak=[
    ("Dikey Dikiş (SAW/SMAW)","Gövde Dikey Birleşimler","~220 m","RT %100","","",""),
    ("Yatay Dikiş (SAW)","Gövde Yatay/Çevre Birleşimler","~164 m","RT %100","","",""),
    ("Taban Plaka Kaynağı","Taban Birleşimleri (Lap)","~320 m","VT","","",""),
    ("Tavan Plaka Kaynağı","Konik Tavan Birleşimleri","~130 m","VT","","",""),
    ("Nozul Kaynağı","Tüm Nozullar","~12 adet","PT+RT","","",""),
    ("Manhol Kaynağı","M1+M2+M3","3 adet","PT","","",""),
    ("Yapısal Kaynak","Çatı İskeleti","~500 m","VT","","",""),
]
for i,rd in enumerate(kaynak):
    fill=ALT if i%2==1 else WH
    for c,v in enumerate(rd,1):
        cl=ws10.cell(row=r,column=c,value=v)
        cl.fill=YEL if c==5 else fill; cl.font=fn; cl.alignment=CA if c!=2 else LA; cl.border=TB
    r+=1
r+=1

# Boya bölümü
SUB(ws10, r, 1, 10, "D. BOYA VE YÜZEY HAZIRLAMA"); r+=1
HDR(ws10, r, list(range(1,8)),
    ["YÜZEYİ","AÇIKLAMA","ALAN (m²)","KAT SAYISI","BİRİM FİYAT ★","TUTAR ($)","BOYA SİSTEMİ"], h=30)
r+=1
boya_data=[
    ("İÇ YÜZEY (ISLAK)","Gövde İç + Taban İç","~520 m²",3,"","","Epoksi Astar + Ara Kat + Finish"),
    ("DIŞ YÜZEY","Gövde Dış + Tavan Dış","~525 m²",3,"","","Korozyon Önleyici Sistem"),
    ("İÇ TAVAN","Tavan İç Yüzey","~130 m²",2,"","","Epoksi Astar"),
    ("YAPISAL","Çatı İskeleti Dış Yüzey","~150 m²",3,"","","Alkid/Epoksi"),
]
for i,rd in enumerate(boya_data):
    fill=ALT if i%2==1 else WH
    for c,v in enumerate(rd,1):
        cl=ws10.cell(row=r,column=c,value=v)
        cl.fill=YEL if c==5 else fill; cl.font=fn; cl.alignment=CA if c!=2 else LA; cl.border=TB
    r+=1

# Not
r+=1
ws10.merge_cells(start_row=r,start_column=1,end_row=r,end_column=10)
cl=ws10.cell(row=r,column=1,
    value="★ BİRİM FİYAT GİRİŞİ: Sarı hücrelere piyasa fiyatı giriniz. Formüller otomatik hesaplayacaktır.")
cl.fill=YEL; cl.font=Font(bold=True,size=11,color="FF0000")
cl.alignment=CA; cl.border=TB

sw(ws10,[16,36,16,16,14,14,28,14,14,14])

# ═══════════════════════════════════════════════════════════════════
# SAYFA 11 – SATINALMA METRAJ ÖZETİ
# ═══════════════════════════════════════════════════════════════════
ws11 = wb.create_sheet("Satınalma Metrajı"); ws11.sheet_view.showGridLines = False
TITLE(ws11, 1, 1, 8, "7000 Bbl ÜRETİM TANKI  —  NET SATINALMA METRAJ ÖZETİ  (1 TANK)")

# ── A. SAC / PLAKA ─────────────────────────────────────────────────
SUB(ws11, 2, 1, 8, "A. SAC PLAKA — Kalınlık ve Malzeme Sınıfına Göre")
HDR(ws11, 3, list(range(1,9)),
    ["#","MALZEME","KALINLIK\n(mm)","KULLANIM YERİ","PARÇA\nADETİ","NET\nAĞIRLIK (kg)",
     "STANDART STOK ÖNERİSİ","NOTLAR"], h=35)

# Gövde
govde_8mm_wt  = 557.6+557.0+552.3+539.1+414.0+557.6+371.5   # 3549.1
govde_6mm_wt  = 30*418.2 + 418.0                             # 12964.0
# Tavan
tavan_8mm_wt  = 8147.0
tavan_8mm_cnt = 4+1+1+1+2+4+4+2+1+1+3+1+4+4+2              # 35
# Taban orta (10mm)
taban_10mm_wt = 4937.1+628.0+222.0+657.6+222.0+557.0+694.8+348.2  # 8266.7
taban_10mm_cnt= 7+2+4+3+1+1+2+2                                    # 22
# Taban annüler (12mm)
taban_12mm_wt = 2616.6+237.8                                 # 2854.4
taban_12mm_cnt= 6+1                                          # 7
# Taban drenaj stripleri (5mm)
taban_5mm_wt  = 11.4+2.55+36.4+5.18                         # 55.5
taban_5mm_cnt = 8+1+13+1                                     # 23

sac_rows = [
    (1,"S275J2", 6,"Gövde Silindirik Kurslar (POZ 3, 3A)",31,
     govde_6mm_wt,"32× 1500×6000×6mm (%2 fire)",
     "1480mm enine kesilecek; 30 adet tam + 1 adet kısmi"),
    (2,"S275J2", 8,"Gövde Alt Kurs (POZ 1,1A,1B,1C,1D,1E,2)",7,
     govde_8mm_wt,"7× 1500×6000×8mm",
     "6 adet 6000mm tam, 1 adet 3997mm kısmi"),
    (3,"S275J2", 8,"Tavan Konik Sektör Plakaları (POZ 1-9)",tavan_8mm_cnt,
     tavan_8mm_wt,f"40× 1500×6000×8mm (%15 fire)",
     "Sektör şekil; 8mm+6mm gövde ile birleşik sipariş önerilir"),
    ("→","S275J2","8mm TOPLAM","Gövde (7 pk) + Tavan (35 pk) birleşik",42,
     round(govde_8mm_wt+tavan_8mm_wt,1),"~47 plaka 1500×6000×8mm","★ Sipariş miktarı"),
    (4,"S275J2",10,"Taban Orta Plakalar (POZ 1-8)",taban_10mm_cnt,
     round(taban_10mm_wt,1),"12× 1500×6000×10mm",
     "Büyük plakalardan nested kesim — fire düşük"),
    (5,"S275J2",12,"Taban Annüler (Çevre) Plakalar (POZ 9, 9A)",taban_12mm_cnt,
     round(taban_12mm_wt,1),"4× 1500×6000×12mm",
     "800mm en; 1500mm'lik plakadan boyuna bölünecek"),
    (6,"S275J2", 5,"Taban Drenaj Strip Parçaları (POZ 10-13)",taban_5mm_cnt,
     round(taban_5mm_wt,1),"50mm × çeşitli boy şerit",
     "Atık sacdan kesilebilir"),
    (None,"S275J2","—","S275J2 TOPLAM SAC AĞIRLIĞI","—",
     round(govde_6mm_wt+govde_8mm_wt+tavan_8mm_wt+taban_10mm_wt+taban_12mm_wt+taban_5mm_wt,1),
     "—","★ Ana sac siparişi toplam"),
]

rr = 4
for rd in sac_rows:
    is_tot = rd[0] in ["→", None]
    fill_ = H2 if is_tot else (ALT if rr%2==0 else WH)
    fnt_  = fw if is_tot else fn
    for c,v in enumerate(rd,1):
        C(ws11, rr, c, v, fill_, fnt_, CA if c!=4 else LA, TB)
    rr += 1

# ── B. YAPISAL PROFİL / BORU ───────────────────────────────────────
rr += 1
SUB(ws11, rr, 1, 8, "B. YAPISAL PROFİL VE BORU"); rr += 1
HDR(ws11, rr, list(range(1,8)),
    ["#","PROFİL / BORU","MALZEME","KULLANIM YERİ","TOPLAM ADET","TOPLAM BOY (mm)","TOPLAM AĞIRLIK (kg)"], h=30)
rr += 1
profil_rows = [
    (1,"UNP240","S235JR","Çatı CK1+CK2 Rafterlar",24,24*5685,"~3732"),
    (2,"UNP200","S235JR","Çatı CK3 Purlin",24,24*1227,"~744"),
    (3,"UNP200","S235JR","Çatı CK4 Purlin",24,24*677,"~410"),
    (4,"PIP 323.9×8 Boru","S235JR","Çatı KL1 Merkezi Kolon",1,9369,580.2),
    (5,"L60×5 Köşebent","S235JR","Çatı YC1+YC2",24,"~46740","~417"),
    (6,"⊏ 8×200 (C200×8)","S235JR","Merdiven Borda",8,"çeşitli","~388"),
    (7,"NPU100","S235JR","Merdiven Basamak",13,"çeşitli","~288"),
    (8,"L60×6 Köşebent","S235JR","Merdiven Korkuluk/Detay","~100+","çeşitli","~560"),
    (9,'1 1/4" SCH Boru (Handrail)',"S235JR","Merdiven Korkuluk",1,70000,178.5),
]
for i,rd in enumerate(profil_rows):
    DATA(ws11, rr, list(range(1,8)), rd, alt=(i%2==1)); rr += 1

# ── C. NOZUL BORULARI VE FLANŞLARI ─────────────────────────────────
rr += 1
SUB(ws11, rr, 1, 8, "C. NOZUL BORULARI VE FLANŞLARI  (ASTM)"); rr += 1
HDR(ws11, rr, list(range(1,8)),
    ["#","MALZEME / TİP","STANDART","KULLANIM YERİ","ADET","AĞIRLIK (kg)","NOTLAR"], h=30)
rr += 1
noz_mat = [
    (1,'12" SCH XS Boru',"ASTM A106 Gr.B","N1+N4+N21","6 adet","~129","2×N1,2×N4 + N21×4"),
    (2,' 8" SCH XS Boru',"ASTM A106 Gr.B","N8","1 adet","11.4",""),
    (3,' 6" SCH XS Boru+Dirsek',"ASTM A106+A234WPB","N7A/B","2 komple set","~107.2","Dirsek dahil"),
    (4,' 1" / 1.5" Boru',"ASTM A106 Gr.B","N9+N18","2 adet","~8",""),
    (5,'12" 150# SO-RF Flanş',"ASTM A105","N1+N4+N21","6 adet","~174",""),
    (6,' 8" 150# SO-RF Flanş',"ASTM A105","N8","1 adet",13.5,""),
    (7,' 6" 150# SO-RF Flanş',"ASTM A105","N7A/B","2 adet",11.8,""),
    (8,' 1" / 1.5" Flanş',"ASTM A105","N9+N18","2 adet","~5",""),
]
for i,rd in enumerate(noz_mat):
    DATA(ws11, rr, list(range(1,8)), rd, alt=(i%2==1)); rr += 1

# ── D. SAC/PLAKA TOPLAM ÖZET (Satınalma için) ──────────────────────
rr += 2
SUB(ws11, rr, 1, 8, "D. KONSOLİDE AĞIRLIK — Çelik Sınıfına Göre Satınalma Özeti"); rr += 1
HDR(ws11, rr, [1,2,3,4,5],
    ["ÇELİK SINIFI","KULLANIM","NET AĞIRLIK (kg)","FİRE PAYI (%3)","SATINALMA (kg)"], h=30)
rr += 1
tot_s275j2 = round(govde_6mm_wt+govde_8mm_wt+tavan_8mm_wt+taban_10mm_wt+taban_12mm_wt+taban_5mm_wt,1)
konsolidat = [
    ("S275J2","Ana Yapı Sacları (Gövde+Taban+Tavan)",tot_s275j2,
     round(tot_s275j2*0.03,1),round(tot_s275j2*1.03,1)),
    ("S235JR","Yapısal (Çatı+Merdiven profil+plakalar)","~7230",
     "~220","~7450"),
    ("ASTM A106 Gr.B","Nozul Boruları","~255","–","~260"),
    ("ASTM A105","Nozul Flanşları","~205","–","~210"),
    ("A193 Gr.B7 / A194 Gr.2H","Cıvata-Somun-Rondela","–","–","Çizimden sipariş"),
    ("A 283 Gr.C + S.S. A316L","Mekanik Samandıra","–","–","1 komple set"),
    ("AISI C1030","Ankraj (16 adet)","216","-","220"),
]
for i,rd in enumerate(konsolidat):
    fill_ = GRN if i==0 else (ALT if i%2==0 else WH)
    for c,v in enumerate(rd,1):
        C(ws11,rr,c,v,fill_,fn,CA if c!=2 else LA,TB)
    rr += 1
NOTE(ws11, rr, 1, 8, "★ S275J2 sac siparişi: 6mm (~12964 kg) + 8mm (~11696 kg) + 10mm (~8267 kg) + 12mm (~2854 kg). Net toplam ~35781 kg. %3 fire ile ~36855 kg sipariş önerilir.")
sw(ws11,[6,18,14,40,12,16,22,24])

out = "/home/user/key_cloud/7000Bbl_Tank_Malzeme_ve_Maliyet.xlsx"
wb.save(out)
print(f"Kaydedildi: {out}")
