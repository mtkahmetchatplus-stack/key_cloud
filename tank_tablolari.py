import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# ─── Yardımcı fonksiyonlar ────────────────────────────────────────────────────
HEADER_FILL  = PatternFill("solid", fgColor="1F4E79")
HEADER2_FILL = PatternFill("solid", fgColor="2E75B6")
SUB_FILL     = PatternFill("solid", fgColor="BDD7EE")
ALT_FILL     = PatternFill("solid", fgColor="DEEAF1")
WHITE_FILL   = PatternFill("solid", fgColor="FFFFFF")
BOLD_WHITE   = Font(bold=True, color="FFFFFF", size=10)
BOLD_DARK    = Font(bold=True, color="1F4E79", size=10)
NORMAL       = Font(size=10)
CENTER       = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT         = Alignment(horizontal="left",   vertical="center", wrap_text=True)

thin = Side(style="thin", color="4472C4")
thick= Side(style="medium", color="1F4E79")
THIN_BORDER  = Border(left=thin, right=thin, top=thin, bottom=thin)
THICK_BORDER = Border(left=thick, right=thick, top=thick, bottom=thick)

def header_row(ws, row, cols, texts, fill=HEADER_FILL, font=BOLD_WHITE):
    for col, text in zip(cols, texts):
        c = ws.cell(row=row, column=col, value=text)
        c.fill = fill; c.font = font; c.alignment = CENTER; c.border = THIN_BORDER

def data_row(ws, row, cols, values, alt=False):
    fill = ALT_FILL if alt else WHITE_FILL
    for col, val in zip(cols, values):
        c = ws.cell(row=row, column=col, value=val)
        c.fill = fill; c.font = NORMAL; c.alignment = CENTER; c.border = THIN_BORDER

def title_row(ws, row, col_start, col_end, text):
    ws.merge_cells(start_row=row, start_column=col_start,
                   end_row=row,   end_column=col_end)
    c = ws.cell(row=row, column=col_start, value=text)
    c.fill = HEADER_FILL; c.font = Font(bold=True, color="FFFFFF", size=12)
    c.alignment = CENTER; c.border = THICK_BORDER

def set_col_widths(ws, widths):
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w

# ══════════════════════════════════════════════════════════════════════════════
# SAYFA 1 – TASARIM VERİLERİ  (T202ABC)
# ══════════════════════════════════════════════════════════════════════════════
ws1 = wb.active
ws1.title = "Tasarım Verileri"
ws1.sheet_view.showGridLines = False

title_row(ws1, 1, 1, 4, "7000 Bbl ÜRETİM TANKI T-202A/B/C – TASARIM VERİLERİ")
ws1.row_dimensions[1].height = 28

header_row(ws1, 2, [1,2,3,4], ["PARAMETRE", "DEĞER", "PARAMETRE", "DEĞER"], HEADER2_FILL, BOLD_WHITE)

tasarim = [
    ("KAPASİTE",               "7000 bbl",            "YÖNETMELİK",          "API 650"),
    ("TANK ÇAPI",               "12.730 m",            "YOĞUNLUK",            "800 kg/m³"),
    ("TANK YÜKSEK.",            "9.900 m",             "OPER. BASINCI",       "ATMOSFERİK"),
    ("ÇATI TİPİ",               "KESMİ KONİ GR",       "TAŞ. BASINCI",        "1.6 / ~5.2 bar"),
    ("TABAN TİPİ",              "MİNİMUM (RİSK ÜRÜ)", "OPER. SICAKLIK",      "AMBİENT"),
    ("RÜZGAR HIZI",             "45 m/s",              "TAŞ. SICAKLIK",       "50 °C / -7 °C"),
    ("SAC MALZEMESİ",          "S275J0",               "KAYN. VERİMİ",        "1"),
    ("ANKRAJ MALZEMESİ",       "AÇS C1030",            "TALITIM (CATHODIC)",  "YOK"),
    ("SIVININ CİNSİ",          "HAM PETROL",           "KOROZYON PAYI",       "3 mm"),
]
for i, (p1, v1, p2, v2) in enumerate(tasarim, start=3):
    alt = (i % 2 == 0)
    fill = ALT_FILL if alt else WHITE_FILL
    for col, val, is_param in [(1,p1,True),(2,v1,False),(3,p2,True),(4,v2,False)]:
        c = ws1.cell(row=i, column=col, value=val)
        c.fill = fill
        c.font = Font(bold=True, size=10, color="1F4E79") if is_param else NORMAL
        c.alignment = LEFT if is_param else CENTER
        c.border = THIN_BORDER

set_col_widths(ws1, [32, 22, 32, 22])

# ══════════════════════════════════════════════════════════════════════════════
# SAYFA 2 – NOZUL LİSTESİ  (T202ABC)
# ══════════════════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("Nozul Listesi")
ws2.sheet_view.showGridLines = False

title_row(ws2, 1, 1, 11, "7000 Bbl ÜRETİM TANKI T-202A/B/C – NOZUL LİSTESİ")
ws2.row_dimensions[1].height = 28

# ── Sol tablo: geometri ──────────────────────────────────────────────
header_row(ws2, 2, [1,2,3,4,5,6], ["POZ","ÇAP","A (mm)","B (mm)","COURSE","REINF.PAD"], HEADER2_FILL)
nozul_geo = [
    ("N1",       '12"', 225,   440,  "SHELL", "t=6 #MS"),
    ("N2",       '12"', 225,   440,  "SHELL", "t=6 #MS"),
    ("N3A/B",    ' 4"', 200,   350,  "SHELL", "t=4 #MS"),
    ("N4",       ' 8"', 225,  8500,  "SHELL", "t=6 #MS"),
    ("N5",       '1½"', 150,  1200,  "SHELL", "–"),
    ("N6",        "–",  "–",   "–", "SHELL", "–"),
    ("N7/1~N7/5",'¾"',  "–",   "–", "SHELL", "–"),
    ("N8",       '12"', 225,     0,  "ROOF",  "t=6 #MS"),
    ("M1",       '24"', 150,   750,  "SHELL", "t=6 #I150"),
    ("M2",       '24"', 250,  5750,  "SHELL", "t=6 #I150"),
    ("M3",       '24"', 200,   310,  "SHELL", "t=6 #I130"),
]
for i, row in enumerate(nozul_geo, start=3):
    data_row(ws2, i, [1,2,3,4,5,6], row, alt=(i%2==0))

# ── Sağ tablo: tanım & flanş ──────────────────────────────────────────
header_row(ws2, 2, [8,9,10,11,12,13], ["POZ","ADET","TANIM","ÇAP","BASINÇ / FLANŞ","NOTLAR"], HEADER2_FILL)
nozul_tanim = [
    ("N1",       1, "Petrol Giriş Nozulu",          '12"', "LB 150 / SO-RF",     ""),
    ("N2",       1, "Petrol Çıkış Nozulu",           '12"', "LB 150 / SO-RF",     ""),
    ("N3A/B",    2, "Drain (Tahliye)",               ' 4"', "LB 150 / SO-RF",     ""),
    ("N4",       1, "Köpük Yapıcı",                  ' 8"', "LB 150 / SO-RF",     ""),
    ("N5",       1, "Sıcaklık Ölçümü (TG)",          '1½"', "LB 150 / SO-RF",     ""),
    ("N6",       1, "Mekanik Samandıra (LG)",         "–",  "–",                  ""),
    ("N7/1~N7/5",5, "Numune Alma Nozulu",            '¾"',  "LB 3000 / BSPT-Plug",""),
    ("N8",       4, "Hava Nozulu",                   '12"', "–",                  ""),
    ("M1",       1, "Çökük Manho (Sump Cleanout)",   '24"', "API 650 / RF",        ""),
    ("M2",       1, "Takim Manholu",                 '24"', "API 650 / RF",        ""),
    ("M3",       1, "Temizlik Manholu",              '24"', "API 650 / RF",        ""),
]
for i, row in enumerate(nozul_tanim, start=3):
    data_row(ws2, i, [8,9,10,11,12,13], row, alt=(i%2==0))

set_col_widths(ws2, [12,8,10,12,12,14, 2, 12,8,32,8,20,16])

# ══════════════════════════════════════════════════════════════════════════════
# SAYFA 3 – DEPREM PARAMETRELERİ (BAK-001 & BAK-002)
# ══════════════════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("Deprem Parametreleri")
ws3.sheet_view.showGridLines = False

title_row(ws3, 1, 1, 3, "DEPREM TASARIM PARAMETRELERİ")
ws3.row_dimensions[1].height = 28
header_row(ws3, 2, [1,2,3], ["PARAMETRE","SEMBOL","DEĞER"], HEADER2_FILL)

deprem = [
    ("Kısa periyot harita spektral ivme katsayısı",     "SS",   0.364),
    ("1.0 s periyot harita spektral ivme katsayısı",    "S1",   0.123),
    ("Yerel Zemin Sınıfı",                              "–",    "ZC"),
    ("Kısa periyot yerel zemin etki katsayısı",         "FS",   1.300),
    ("1.0 s periyot yerel zemin etki katsayısı",        "F1",   1.500),
    ("Kısa periyot tasarım spektral ivme katsayısı",    "SDS",  0.473),
    ("Bina Kullanım Sınıfı",                            "BKS",  1),
    ("Deprem Tasarım Sınıfı",                           "DTS",  "3a"),
]
for i, (p, s, v) in enumerate(deprem, start=3):
    alt = (i % 2 == 0)
    fill = ALT_FILL if alt else WHITE_FILL
    for col, val, bold in [(1,p,True),(2,s,True),(3,v,False)]:
        c = ws3.cell(row=i, column=col, value=val)
        c.fill = fill
        c.font = Font(bold=bold, size=10, color="1F4E79") if bold else NORMAL
        c.alignment = LEFT if col==1 else CENTER
        c.border = THIN_BORDER

set_col_widths(ws3, [52, 10, 12])

# ══════════════════════════════════════════════════════════════════════════════
# SAYFA 4 – MALZEME LİSTESİ – GC1 GÖMÜLÜ ÇELİK  (BAK-002)
# ══════════════════════════════════════════════════════════════════════════════
ws4 = wb.create_sheet("GC1 Malzeme Listesi")
ws4.sheet_view.showGridLines = False

title_row(ws4, 1, 1, 7, "GC1 GÖMÜLÜ ÇELİK – MALZEME LİSTESİ  (456-1100-15-BAK-002)")
ws4.row_dimensions[1].height = 28
header_row(ws4, 2, [1,2,3,4,5,6,7],
           ["POZ","ADET","PARÇA SINIRLARI","BOY (m)","NET AĞIRLIK (kg)","MALZEME","NOT"])

gc1 = [
    (1,  24, "PL8×…",       0.11,  26.5,  "S275J0", "Gömülü"),
    (2,  24, "PL9×…",       0.13,  25.1,  "S275J0", "Gömülü"),
    (3,  40, "L50×5",       1.18,  46.6,  "S275J0", "Gömülü"),
    (4,   5, "–",           11.40, 113.0,  "S275J0", "Gömülü"),
    (5,  12, "Ø12 (Bar)",   5.44,   7.1,  "S275J0", "Gömülü"),
    (6,  12, "L50×3",       3.54,  26.3,  "S275J0", "Gömülü"),
    (7,  12, "PL50×8",      1.56,  11.3,  "S275J0", "Gömülü"),
    (8,  20, "Ø8 (Bar)",    1.14,  14.0,  "S275J0", "Gömülü/Beton"),
    (9,   2, "L64×…",       1.25,   1.8,  "S275J0", ""),
    (10, 20, "ML×… 100",    0.04,   1.3,  "Duplex", "ABD"),
]
for i, row in enumerate(gc1, start=3):
    data_row(ws4, i, [1,2,3,4,5,6,7], row, alt=(i%2==0))

# Toplam satırı
tr = len(gc1) + 3
ws4.merge_cells(start_row=tr, start_column=1, end_row=tr, end_column=4)
c = ws4.cell(row=tr, column=1, value="TOPLAM AĞIRLIK")
c.fill = HEADER2_FILL; c.font = BOLD_WHITE; c.alignment = CENTER; c.border = THIN_BORDER
total_kg = sum(r[4] for r in gc1)
for col, val in [(5, total_kg),(6,""),(7,"")]:
    c = ws4.cell(row=tr, column=col, value=val)
    c.fill = HEADER2_FILL; c.font = BOLD_WHITE; c.alignment = CENTER; c.border = THIN_BORDER

set_col_widths(ws4, [8, 8, 22, 12, 20, 14, 18])

# ══════════════════════════════════════════════════════════════════════════════
# SAYFA 5 – MALZEME LİSTESİ – ANKRAJ BULONU  (BAK-001)
# ══════════════════════════════════════════════════════════════════════════════
ws5 = wb.create_sheet("Ankraj Malzeme Listesi")
ws5.sheet_view.showGridLines = False

title_row(ws5, 1, 1, 7, "ANK1 ANKRAJ BULONU – MALZEME LİSTESİ  (456-1200-15-BAK-001)")
ws5.row_dimensions[1].height = 28
header_row(ws5, 2, [1,2,3,4,5,6,7],
           ["POZ","ADET","PARÇA / EŞDEĞER","ADET AĞIRLIK (kg)","TOPLAM AĞIRLIK (kg)","MALZEME","NOT"])

ank = [
    (1, 16, "Ø30 × 1140 mm (Ankraj Çubuğu)", 6.32, 101.1, "AÇS C1030", "Dayanıklı"),
    (2, 16, "PL 40×20 × 40 mm (Pul / Plaka)", 0.57,   9.1, "S275J2",    ""),
    ("–","48 adet","M30 Somun","–","–","EN ISO 4032",""),
]
for i, row in enumerate(ank, start=3):
    data_row(ws5, i, [1,2,3,4,5,6,7], row, alt=(i%2==0))

tr5 = len(ank) + 3
ws5.merge_cells(start_row=tr5, start_column=1, end_row=tr5, end_column=4)
c = ws5.cell(row=tr5, column=1, value="TOPLAM AĞIRLIK (1 TANK İÇİN)")
c.fill = HEADER2_FILL; c.font = BOLD_WHITE; c.alignment = CENTER; c.border = THIN_BORDER
for col, val in [(5,"110.2 kg"),(6,""),(7,"")]:
    c = ws5.cell(row=tr5, column=col, value=val)
    c.fill = HEADER2_FILL; c.font = BOLD_WHITE; c.alignment = CENTER; c.border = THIN_BORDER

set_col_widths(ws5, [8, 10, 34, 22, 24, 16, 14])

# ══════════════════════════════════════════════════════════════════════════════
# SAYFA 6 – DONATI METRAJ LİSTESİ  (BAK-001  BAR LIST/XL LIST)
# ══════════════════════════════════════════════════════════════════════════════
ws6 = wb.create_sheet("Donatı Metraj Listesi")
ws6.sheet_view.showGridLines = False

title_row(ws6, 1, 1, 9, "PIT 1 DONATI METRAJ LİSTESİ  (456-1200-15-BAK-001)")
ws6.row_dimensions[1].height = 28
header_row(ws6, 2, list(range(1,10)),
           ["POS","#","ADET","FORM (mm)","L (mm)","Ø8","Ø10","Ø12","Ø16"])

donati = [
    (1, 12,  1, "–", 7000, "7.0", "–",   "–",   "–"),
    (2, 14,  8, "–",  300, "–",   "–",   "–",  "12.6"),
    (3, 12, 20, "–", 1045, "–",   "–",  "20.8", "–"),
]
for i, row in enumerate(donati, start=3):
    data_row(ws6, i, list(range(1,10)), row, alt=(i%2==0))

# Toplam / birim ağırlık / ağırlık satırları
info_rows = [
    ("TOPLAM BOY (m)",       "", "", "", "", "7.0", "–", "20.8", "12.6"),
    ("BİRİM AĞIRLIK (kg/m)", "", "", "", "", "0.395","0.616","0.888","1.578"),
    ("AĞIRLIK (kg)",         "", "", "", "", "2.8",  "–",   "18.5","19.9"),
    ("TOPLAM AĞIRLIK (kg)",  "35.9", "", "", "", "", "", "", ""),
]
row_idx = len(donati) + 3
for r in info_rows:
    for col, val in enumerate(r, 1):
        c = ws6.cell(row=row_idx, column=col, value=val)
        c.fill = SUB_FILL; c.font = BOLD_DARK; c.alignment = CENTER; c.border = THIN_BORDER
    row_idx += 1

set_col_widths(ws6, [8,6,8,14,10,10,10,10,10])

# ══════════════════════════════════════════════════════════════════════════════
# SAYFA 7 – REFERANS BELGELER
# ══════════════════════════════════════════════════════════════════════════════
ws7 = wb.create_sheet("Referans Belgeler")
ws7.sheet_view.showGridLines = False

title_row(ws7, 1, 1, 3, "REFERANS & İLGİLİ DOKÜMANLAR")
ws7.row_dimensions[1].height = 28
header_row(ws7, 2, [1,2,3], ["DOKÜMAN NO","AÇIKLAMA","TÜR"])

refs = [
    ("456-1100-15-BAK-001","Taşırma ve Üretim Tankları Dayk Alanı Kalıp Planı ve Detayları","İlgili"),
    ("456-1100-15-BAK-002","Taşırma ve Üretim Tankları Dayk Alanı Kesitler ve Gömülü Çelik Detayları","İlgili"),
    ("456-1100-15-BAK-003","Taşırma Tankı Temel Kalıp Planı ve Detayları","İlgili"),
    ("456-1200-15-BAK-001","Üretim Tankı, Üretim Taşırma Tankı Temel Kalıp Planı ve Detayları","İlgili"),
    ("456-0000-10-SVP-001","Saha Genel Yerleşim ve Vaziyet Planı","Referans"),
    ("456-0000-19-SPP-001","Yağmur Suyu ve Yağlı Su Drenaj Planı","Referans"),
    ("API 650","Welded Tanks for Oil Storage","Standart"),
    ("TS500","Betonarme Yapıların Tasarım ve Yapım Kuralları","Standart"),
    ("TBDY 2018","Türkiye Bina Deprem Yönetmeliği","Standart"),
    ("TS EN 13670","Beton Yapıların Yapımı","Standart"),
    ("TS EN 10080","Beton İçin Çelik – Kaynak Edilebilir Nervürlü Çelik","Standart"),
    ("TS708","Çelik Hasır ve Çubuklar","Standart"),
    ("TS EN 206","Beton – Özellik, Performans, Üretim ve Uygunluk","Standart"),
    ("TS 6165 / 6173 / 6991","İmalat Toleransları","Standart"),
]
for i, row in enumerate(refs, start=3):
    data_row(ws7, i, [1,2,3], row, alt=(i%2==0))
    ws7.cell(row=i, column=2).alignment = LEFT

set_col_widths(ws7, [30, 68, 12])

# ── Kaydet ────────────────────────────────────────────────────────────────────
out = "/home/user/key_cloud/7000Bbl_Uretim_Tanki_T202ABC_Tablolar.xlsx"
wb.save(out)
print(f"Kaydedildi: {out}")
