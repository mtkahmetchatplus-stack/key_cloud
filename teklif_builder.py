import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy
import shutil

# Şablonu kopyala
src = "/root/.claude/uploads/e4179682-a3a6-4538-95c4-68a43315bb48/da13fdb8-TANK_TEKLIF_DOKUMANTASYONU.xlsx"
dst = "/home/user/key_cloud/TANK_TEKLIF_DOLDURULMUS.xlsx"
shutil.copy2(src, dst)

wb = openpyxl.load_workbook(dst)

# ── Stil sabitleri ────────────────────────────────────────────────
H1   = PatternFill("solid", fgColor="1F4E79")   # koyu mavi - ana başlık
H2   = PatternFill("solid", fgColor="2E75B6")   # orta mavi - bölüm başlığı
H3   = PatternFill("solid", fgColor="BDD7EE")   # açık mavi - kolon başlığı
ALT  = PatternFill("solid", fgColor="F2F2F2")   # alternatif satır (çok hafif gri)
WH   = PatternFill("solid", fgColor="FFFFFF")
YEL  = PatternFill("solid", fgColor="FFFF99")   # birim fiyat girişi
GRN  = PatternFill("solid", fgColor="E2EFDA")   # toplam
ORG  = PatternFill("solid", fgColor="FCE4D6")   # uyarı

# Tank renkleri - sadece bölüm başlıkları
T201_HDR = PatternFill("solid", fgColor="92D050")  # yeşil - T-201
T202_HDR = PatternFill("solid", fgColor="FFFF00")  # sarı  - T-202
T301_HDR = PatternFill("solid", fgColor="FFC000")  # turuncu- T-301

fw   = Font(bold=True, color="FFFFFF", size=10)
fw_b = Font(bold=True, color="000000", size=10)  # siyah bold (sarı/yeşil başlık üstü)
fb   = Font(bold=True, color="1F4E79", size=10)
fb2  = Font(bold=True, color="FFFFFF", size=11)
fn   = Font(size=10)
fn9  = Font(size=9)

CA = Alignment(horizontal="center", vertical="center", wrap_text=True)
LA = Alignment(horizontal="left",   vertical="center", wrap_text=True)
RA = Alignment(horizontal="right",  vertical="center")

th = Side(style="thin",   color="4472C4")
mk = Side(style="medium", color="1F4E79")
TB = Border(left=th, right=th, top=th, bottom=th)
MB = Border(left=mk, right=mk, top=mk, bottom=mk)

def sw(ws, widths, start=1):
    for i, w in enumerate(widths, start):
        ws.column_dimensions[get_column_letter(i)].width = w

def cell(ws, r, c, v=None, fill=WH, font=fn, align=CA, border=TB, h=None):
    cl = ws.cell(row=r, column=c, value=v)
    cl.fill=copy(fill); cl.font=copy(font)
    cl.alignment=copy(align); cl.border=copy(border)
    if h: ws.row_dimensions[r].height = h
    return cl

def title_row(ws, r, c1, c2, text, fill=H1, font=fb2, h=22):
    ws.merge_cells(start_row=r,start_column=c1,end_row=r,end_column=c2)
    cl = ws.cell(row=r,column=c1,value=text)
    cl.fill=copy(fill); cl.font=copy(font); cl.alignment=copy(CA); cl.border=copy(MB)
    ws.row_dimensions[r].height = h

def sec_header(ws, r, c1, c2, text, fill=H2, font_color="FFFFFF"):
    ws.merge_cells(start_row=r,start_column=c1,end_row=r,end_column=c2)
    cl = ws.cell(row=r,column=c1,value=text)
    fnt = Font(bold=True, color=font_color, size=10)
    cl.fill=copy(fill); cl.font=fnt; cl.alignment=copy(CA); cl.border=copy(TB)
    ws.row_dimensions[r].height = 20

def tot_row(ws, r, c1, cv, label, kg_val, fill=H1):
    ws.merge_cells(start_row=r,start_column=c1,end_row=r,end_column=cv-1)
    cl = ws.cell(row=r,column=c1,value=label)
    cl.fill=copy(fill); cl.font=copy(fw); cl.alignment=copy(RA); cl.border=copy(TB)
    cell(ws,r,cv,kg_val,fill,fw,CA,TB)

# ═══════════════════════════════════════════════════════════════════════
# TEKLIF SAYFASI — Ana sipariş tablosu (3 tank)
# Sütunlar: A(boşluk) B(No) C(Tanım) D(Adet) E(Boyut) F(BirimKg) G(ToplamKg)
#            H(MatFiyat★) I(İşçilik★) J(MatTutar) K(İşçTutar) L(TOPLAM)
# ═══════════════════════════════════════════════════════════════════════
ws = wb["TEKLIF"]
for mc in list(ws.merged_cells.ranges):
    ws.unmerge_cells(str(mc))
ws.delete_rows(1, ws.max_row)

sw(ws, [3, 6, 52, 9, 22, 12, 13, 12, 12, 14, 14, 16], start=1)

# Başlık
title_row(ws,1,2,12,"TANK SATINALMA SİPARİŞ TABLOSU  —  SAC / PLAKA METRAJİ  (3 TANK)",h=26)
ws.merge_cells(start_row=2,start_column=2,end_row=2,end_column=12)
cl=ws.cell(2,2,"Sac Malzemecisi → Kolon H (Malzeme Birim Fiyatı $/kg) doldurun.   "
                "Kesim/Büküm Fabrikatörü → Kolon I (İşçilik $/kg) doldurun.")
cl.fill=copy(YEL); cl.font=Font(bold=True,size=9,color="CC0000")
cl.alignment=copy(CA); cl.border=copy(TB); ws.row_dimensions[2].height=18

# Kolon başlıkları
r=3
for c,txt in enumerate(["S.NO","TANIM  /  İŞLEM TÜRÜ","ADET","BOYUT (mm)",
                         "BİRİM\nKg","TOPLAM\nKg","MAT. FİYAT\n★($/kg)",
                         "İŞÇİLİK\n★($/kg)","MAT. TUTARI\n($)","İŞÇİLİK\n($)","TOPLAM\n($)"],1):
    cell(ws,r,c+2,txt,H3,fb,CA,TB,h=38)

# ── Veri yazma yardımcısı ──────────────────────────────────────────
def data_row(ws, r, no, tanim, adet, boyut, birim_kg, alt=False, toplam_override=None):
    fill_ = ALT if alt else WH
    cell(ws,r,2,no,fill_,fn,CA,TB)
    cl=ws.cell(row=r,column=3,value=tanim)
    cl.fill=copy(fill_); cl.font=copy(fn9); cl.alignment=copy(LA); cl.border=copy(TB)
    cell(ws,r,4,adet,fill_,fn,CA,TB)
    cell(ws,r,5,boyut,fill_,fn9,CA,TB)
    cell(ws,r,6,birim_kg,fill_,fn,CA,TB)
    if toplam_override is not None:
        tot_kg = toplam_override
    elif birim_kg and adet:
        tot_kg = round(float(adet) * float(birim_kg), 1)
    else:
        tot_kg = ""
    cell(ws,r,7,tot_kg,fill_,fn,CA,TB)
    cell(ws,r,8,"",YEL,fn,CA,TB)
    cell(ws,r,9,"",YEL,fn,CA,TB)
    cl=ws.cell(row=r,column=10,value=f"=G{r}*H{r}" if tot_kg!="" else "")
    cl.fill=copy(fill_); cl.font=copy(fn); cl.alignment=copy(CA); cl.border=copy(TB)
    cl=ws.cell(row=r,column=11,value=f"=G{r}*I{r}" if tot_kg!="" else "")
    cl.fill=copy(fill_); cl.font=copy(fn); cl.alignment=copy(CA); cl.border=copy(TB)
    cl=ws.cell(row=r,column=12,value=f"=J{r}+K{r}" if tot_kg!="" else "")
    cl.fill=copy(GRN); cl.font=copy(fn); cl.alignment=copy(CA); cl.border=copy(TB)
    ws.row_dimensions[r].height=16

no_counter=1

# ════════════════════════════════════════════════════════════════════
# T-201 TAŞIRMA TANKI  —  GÖVDE  (Çizim: Malzeme Listesi 63,030 kg)
# ════════════════════════════════════════════════════════════════════
r=4
sec_header(ws,r,2,12,
    "T-201  TAŞIRMA TANKI (25,000 BBL)  —  GÖVDE SAC  |  S275J2  |  TOPLAM GÖVDE: 63,030 kg",
    fill=T201_HDR, font_color="000000"); r+=1

t201_govde = [
    #  adet    boyut                  birim_kg  islem
    (1,   "1500×12×12000",   1698.0, "GÖVDE 12mm — POZ 1 — Kesim + Silindirik Bükme"),
    (1,   "1500×12×12000",   1698.0, "GÖVDE 12mm — POZ 2 — Kesim + Silindirik Bükme"),
    (1,   "1500×12×12000",   1698.0, "GÖVDE 12mm — POZ 3 — Kesim + Silindirik Bükme"),
    (1,   "1500×12×12000",   1698.0, "GÖVDE 12mm — POZ 4 — Kesim + Silindirik Bükme"),
    (1,   "1500×12×11678",   1652.0, "GÖVDE 12mm — POZ 5 — Kesim + Silindirik Bükme (kapama)"),
    (3,   "1500×12×12000",   1698.0, "GÖVDE 12mm — POZ 6 ×3 — Kesim + Silindirik Bükme"),
    (1,   "1500×12×12000",   1698.0, "GÖVDE 12mm — POZ 7 — Kesim + Silindirik Bükme"),
    (1,   "1500×12×11678",   1652.0, "GÖVDE 12mm — POZ 8 — Kesim + Silindirik Bükme (kapama)"),
    (10,  "1500×10×12000",   1415.0, "GÖVDE 10mm — POZ 9 ×10 — Kesim + Silindirik Bükme"),
    (3,   "1500×10×11678",   1377.0, "GÖVDE 10mm — POZ 10 ×3 — Kesim + Silindirik Bükme (kapama)"),
    (1,   "1500×10×12000",   1415.0, "GÖVDE 10mm — POZ 11 — Kesim + Silindirik Bükme"),
    (1,   "1500×10×12000",   1415.0, "GÖVDE 10mm — POZ 12 — Kesim + Silindirik Bükme"),
    (9,   "1500×8×12000",    1132.0, "GÖVDE 8mm  — POZ 13 ×9 — Kesim + Silindirik Bükme"),
    (1,   "1500×8×12000",    1132.0, "GÖVDE 8mm  — POZ 14 — Kesim + Silindirik Bükme"),
    (3,   "1500×8×11678",    1102.0, "GÖVDE 8mm  — POZ 15 ×3 — Kesim + Silindirik Bükme (kapama)"),
    (1,   "1500×8×12000",    1132.0, "GÖVDE 8mm  — POZ 16 — Kesim + Silindirik Bükme"),
    (1,   "1500×8×12000",    1132.0, "GÖVDE 8mm  — POZ 17 — Kesim + Silindirik Bükme"),
    (1,   "1500×10×12000",   1415.0, "GÖVDE 10mm — POZ 18 — Kesim + Silindirik Bükme"),
    (40,  "225×10×245",         3.0, "GÖVDE Detay — POZ 19 ×40 — Kesim (Düz, takviye parçası)"),
    (1,   "1500×10×10250",   1208.0, "GÖVDE 10mm — POZ 20 — Kesim + Silindirik Bükme (kapama)"),
    (1,   "1500×10×5050",     595.0, "GÖVDE 10mm — POZ 20A — Kesim + Silindirik Bükme (kapama)"),
    (5,   "425×10×12000",     401.0, "GÖVDE 10mm — POZ 21 ×5 — Kesim (RF/Yapısal)"),
    (1,   "425×10×11644",     389.0, "GÖVDE 10mm — POZ 22 — Kesim (RF/Yapısal kapama)"),
    (5,   "L160×160×15×12000",480.0, "GÖVDE Profil — POZ 23 ×5 — Kesim (L profil, 12m çubuk)"),
    (1,   "L160×160×15×231",    9.0, "GÖVDE Profil — POZ 24 — Kesim (L profil, kapama)"),
]
for i,(adet,boyut,kg,islem) in enumerate(t201_govde):
    data_row(ws,r,no_counter,f"[T-201]  {islem}",adet,boyut,kg,alt=(i%2==1))
    no_counter+=1; r+=1

# POZ 25 — NPT Plug (ağırlık yok, tedarik kalemi)
i=len(t201_govde)
fill_=ALT if i%2==1 else WH
cell(ws,r,2,no_counter,fill_,fn,CA,TB)
cl=ws.cell(row=r,column=3,value="[T-201]  GÖVDE Bağlantı — POZ 25 ×5 — 1/4\" NPT PLUG — Tedarik (ASTM A 105)")
cl.fill=copy(fill_); cl.font=copy(fn9); cl.alignment=copy(LA); cl.border=copy(TB)
cell(ws,r,4,5,fill_,fn,CA,TB)
cell(ws,r,5,"1/4\" NPT PLUG",fill_,fn9,CA,TB)
cell(ws,r,6,"—",fill_,fn,CA,TB)
cell(ws,r,7,"—",fill_,fn,CA,TB)
cell(ws,r,8,"",YEL,fn,CA,TB); cell(ws,r,9,"",YEL,fn,CA,TB)
cell(ws,r,10,"",fill_,fn,CA,TB); cell(ws,r,11,"",fill_,fn,CA,TB)
cl=ws.cell(row=r,column=12,value=""); cl.fill=copy(GRN); cl.font=copy(fn); cl.alignment=copy(CA); cl.border=copy(TB)
ws.row_dimensions[r].height=16; no_counter+=1; r+=1

# T-201 Gövde toplam
tot_row(ws,r,2,7,"T-201 GÖVDE TOPLAM AĞIRLIK:",63030.0,H1); r+=1

# T-201 TABAN
sec_header(ws,r,2,12,
    "T-201  TAŞIRMA TANKI  —  TABAN SAC  |  S275J2  |  TOPLAM TABAN: 15,959 kg",
    fill=T201_HDR, font_color="000000"); r+=1

t201_taban = [
    (8,  "1475×14×6606",   670.0, "TABAN 14mm Annüler/Çevre — POZ 1 ×8 — Kesim"),
    (1,  "1475×14×6606",   537.0, "TABAN 14mm Annüler/Çevre — POZ 2 — Kesim (kapama)"),
    (9,  "1500×12×6000",   849.0, "TABAN 12mm Orta Plaka — POZ 3 ×9 — Kesim"),
    (2,  "1500×12×6000",   849.0, "TABAN 12mm Orta Plaka — POZ 4 ×2 — Kesim"),
    (2,  "441×12×5472",    114.0, "TABAN 12mm Orta Plaka — POZ 5 ×2 — Kesim (köşe)"),
    (2,  "1500×12×5706",   651.0, "TABAN 12mm Orta Plaka — POZ 6 ×2 — Kesim"),
    (4,  "1500×12×5673",   788.0, "TABAN 12mm Orta Plaka — POZ 7 ×4 — Kesim"),
    (4,  "1500×12×5409",   758.0, "TABAN 12mm Orta Plaka — POZ 8 ×4 — Kesim"),
    (4,  "1500×12×4853",   675.0, "TABAN 12mm Orta Plaka — POZ 9 ×4 — Kesim"),
    (4,  "1500×12×3937",   462.0, "TABAN 12mm Orta Plaka — POZ 10 ×4 — Kesim"),
    (4,  "1410×12×2482",   186.0, "TABAN 12mm Orta Plaka — POZ 11 ×4 — Kesim"),
    (1,  "50×6×261000",    626.0, "TABAN 6mm Drenaj Şerit — POZ 12 — Kesim (261m strip, atık sacdan)"),
]
for i,(adet,boyut,kg,islem) in enumerate(t201_taban):
    data_row(ws,r,no_counter,f"[T-201]  {islem}",adet,boyut,kg,alt=(i%2==1))
    no_counter+=1; r+=1

tot_row(ws,r,2,7,"T-201 TABAN TOPLAM AĞIRLIK:",15959.0,H1); r+=1

# T-201 TAVAN
sec_header(ws,r,2,12,
    "T-201  TAŞIRMA TANKI  —  TAVAN SAC  |  S275J2  |  TOPLAM TAVAN: 18,797 kg",
    fill=T201_HDR, font_color="000000"); r+=1

t201_tavan = [
    (45, "1300×8×4148",   264.0, "TAVAN 8mm Sektör — POZ 1 ×45 — Kesim + Bükme"),
    (1,  "1300×8×4148",   234.0, "TAVAN 8mm Sektör — POZ 2 — Kesim + Bükme (kapama)"),
    (1,  "1300×8×4148",   264.0, "TAVAN 8mm Sektör — POZ 3 — Kesim + Bükme"),
    (1,  "1300×8×4148",   264.0, "TAVAN 8mm Sektör — POZ 4 — Kesim + Bükme"),
    (24, "1466×8×4774",   252.0, "TAVAN 8mm Sektör — POZ 5 ×24 — Kesim + Bükme"),
    (1,  "ø1450×8",       107.0, "TAVAN 8mm Merkez Disk — POZ 6 — Kesim (Daire plaka)"),
]
for i,(adet,boyut,kg,islem) in enumerate(t201_tavan):
    data_row(ws,r,no_counter,f"[T-201]  {islem}",adet,boyut,kg,alt=(i%2==1))
    no_counter+=1; r+=1

tot_row(ws,r,2,7,"T-201 TAVAN TOPLAM AĞIRLIK:",18797.0,H1); r+=1

# ════════════════════════════════════════════════════════════════════
# T-202 ÜRETİM TANKI  —  GÖVDE
# ════════════════════════════════════════════════════════════════════
r+=1
sec_header(ws,r,2,12,
    "T-202  ÜRETİM TANKI (7,000 BBL)  —  GÖVDE SAC  |  D=12.730m  H=8.920m  |  S275J2  |  TOPLAM GÖVDE: 16,513 kg",
    fill=T202_HDR, font_color="000000"); r+=1

t202_govde = [
    (6,  "1480×8×6000",   557.6, "GÖVDE 8mm — 1. Kurs (POZ 1,1A,1B,1C,1D,1E ×6) — Kesim + Silindirik Bükme R=6365mm  [Her plakada farklı nozul/manhole kesimi]"),
    (1,  "1480×8×3997",   371.5, "GÖVDE 8mm — 1. Kurs Kapama (POZ 2) — Kesim + Silindirik Bükme R=6365mm"),
    (29, "1480×6×6000",   418.2, "GÖVDE 6mm — 2–6. Kurs (POZ 3 ×29) — Kesim + Silindirik Bükme R=6365mm"),
    (1,  "1480×6×6000",   418.0, "GÖVDE 6mm — 2–6. Kurs Kapama (POZ 3A) — Kesim + Silindirik Bükme R=6365mm"),
    (1,  "1480×6×6000",   418.0, "GÖVDE 6mm — 2–6. Kurs Kapama (POZ 3B) — Kesim + Silindirik Bükme R=6365mm"),
]
for i,(adet,boyut,kg,islem) in enumerate(t202_govde):
    data_row(ws,r,no_counter,f"[T-202]  {islem}",adet,boyut,kg,alt=(i%2==1))
    no_counter+=1; r+=1

tot_row(ws,r,2,7,"T-202 GÖVDE TOPLAM AĞIRLIK:",16513.0,H1); r+=1

# T-202 TABAN
sec_header(ws,r,2,12,
    "T-202  ÜRETİM TANKI  —  TABAN SAC  |  S275J2  (Çizim: KEY199MECEQP2003)  |  TOPLAM TABAN: 11,252.8 kg",
    fill=T202_HDR, font_color="000000"); r+=1

t202_taban = [
    (7,  "1500×10×5990",   705.3, "TABAN 10mm Orta — POZ 1 ×7 — Kesim"),
    (2,  "1500×10×2684",   314.0, "TABAN 10mm Orta — POZ 2 ×2 — Kesim"),
    (4,  "1500×10×635",     55.5, "TABAN 10mm Orta — POZ 3 ×4 — Kesim (kısa strip)"),
    (3,  "1500×10×2262",   219.2, "TABAN 10mm Orta — POZ 4 ×3 — Kesim"),
    (1,  "1500×10×2676",   222.0, "TABAN 10mm Orta — POZ 5 — Kesim"),
    (1,  "1500×10×4395",   557.0, "TABAN 10mm Orta — POZ 6 — Kesim"),
    (2,  "1500×10×2994",   347.4, "TABAN 10mm Orta — POZ 7 ×2 — Kesim"),
    (2,  "632×10×5208",    174.1, "TABAN 10mm Orta — POZ 8 ×2 — Kesim (dar plaka)"),
    (6,  "800×12×5937",    436.1, "TABAN 12mm Annüler/Çevre — POZ 9 ×6 — Kesim"),
    (1,  "800×12×3327",    237.8, "TABAN 12mm Annüler/Çevre — POZ 9A — Kesim (kapama)"),
    (8,  "50×5×720",         1.42,"TABAN 5mm Drenaj Strip — POZ 10 ×8 — Kesim"),    # toplam 11.4 kg override
    (1,  "50×5×1301.5",      2.55,"TABAN 5mm Drenaj Strip — POZ 11 — Kesim"),
    (13, "50×5×1400",        2.80,"TABAN 5mm Drenaj Strip — POZ 12 ×13 — Kesim"),  # toplam 36.4 kg override
    (1,  "50×5×2641.5",      5.18,"TABAN 5mm Drenaj Strip — POZ 13 — Kesim"),
    (1,  "Sump Çukuru",     76.1, "TABAN Sump Drenaj Çukuru — Ayrı İmalat Kalemi (çeşitli parçalar)"),
]
toplam_overrides_202 = {10: 11.4, 12: 36.4}
for i,(adet,boyut,kg,islem) in enumerate(t202_taban):
    ov = toplam_overrides_202.get(i)
    data_row(ws,r,no_counter,f"[T-202]  {islem}",adet,boyut,kg,alt=(i%2==1),toplam_override=ov)
    no_counter+=1; r+=1

tot_row(ws,r,2,7,"T-202 TABAN TOPLAM AĞIRLIK:",11252.8,H1); r+=1

# T-202 TAVAN
sec_header(ws,r,2,12,
    "T-202  ÜRETİM TANKI  —  TAVAN SAC (Konik)  |  S275J2  (Çizim: KEY199MECEQP2002)  |  TOPLAM TAVAN: 8,147 kg",
    fill=T202_HDR, font_color="000000"); r+=1

t202_tavan = [
    (4,  "1500×8×6000",   565.2, "TAVAN 8mm Konik Sektör — POZ 1 ×4 — Kesim + Konik Bükme"),
    (1,  "1500×8×6000",   560.0, "TAVAN 8mm Konik Sektör — POZ 1A — Kesim + Konik Bükme"),
    (1,  "1500×8×6000",   560.0, "TAVAN 8mm Konik Sektör — POZ 1B — Kesim + Konik Bükme"),
    (1,  "1500×8×6000",   560.0, "TAVAN 8mm Konik Sektör — POZ 1C — Kesim + Konik Bükme"),
    (2,  "1500×8×3561",   334.0, "TAVAN 8mm Konik Sektör — POZ 2 ×2 — Kesim + Konik Bükme"),
    (4,  "1500×8×548",     38.3, "TAVAN 8mm Konik Sektör — POZ 3 ×4 — Kesim + Konik Bükme (küçük)"),
    (4,  "1500×8×3196",   269.3, "TAVAN 8mm Konik Sektör — POZ 4 ×4 — Kesim + Konik Bükme"),
    (2,  "1500×8×2873",   257.3, "TAVAN 8mm Konik Sektör — POZ 5 ×2 — Kesim + Konik Bükme"),
    (1,  "1500×8×2873",   254.5, "TAVAN 8mm Konik Sektör — POZ 5A — Kesim + Konik Bükme"),
    (1,  "1500×8×2873",   254.5, "TAVAN 8mm Konik Sektör — POZ 5B — Kesim + Konik Bükme"),
    (3,  "1500×8×2546",   208.0, "TAVAN 8mm Konik Sektör — POZ 6 ×3 — Kesim + Konik Bükme"),
    (1,  "1500×8×2546",   190.0, "TAVAN 8mm Konik Sektör — POZ 6A — Kesim + Konik Bükme"),
    (4,  "1000×8×1825",    90.7, "TAVAN 8mm Konik Sektör — POZ 7 ×4 — Kesim + Konik Bükme (dar)"),
    (4,  "1057×8×875",     31.0, "TAVAN 8mm Konik Sektör — POZ 8 ×4 — Kesim + Konik Bükme (köşe)"),
    (2,  "1500×8×2911",   272.8, "TAVAN 8mm Konik Sektör — POZ 9 ×2 — Kesim + Konik Bükme"),
]
for i,(adet,boyut,kg,islem) in enumerate(t202_tavan):
    data_row(ws,r,no_counter,f"[T-202]  {islem}",adet,boyut,kg,alt=(i%2==1))
    no_counter+=1; r+=1

tot_row(ws,r,2,7,"T-202 TAVAN TOPLAM AĞIRLIK:",8147.0,H1); r+=1

# ════════════════════════════════════════════════════════════════════
# T-301 ATIK SU TANKI
# ════════════════════════════════════════════════════════════════════
r+=1
sec_header(ws,r,2,12,
    "T-301  ATIK SU TANKI (~2,000 BBL)  —  GÖVDE SAC  |  St52-2",
    fill=T301_HDR, font_color="000000"); r+=1

t301_govde = [
    (3,  "1500×10×6000",  706.5, "GÖVDE 10mm — Kesim + Silindirik Bükme"),
    (1,  "1500×10×4000",  471.0, "GÖVDE 10mm — Kesim + Silindirik Bükme (kapama)"),
    (3,  "1500×8×6000",   565.2, "GÖVDE 8mm  — Kesim + Silindirik Bükme"),
    (1,  "1500×8×4000",   376.8, "GÖVDE 8mm  — Kesim + Silindirik Bükme (kapama)"),
    (12, "1500×6×6000",   423.9, "GÖVDE 6mm  — Kesim + Silindirik Bükme"),
    (4,  "1500×6×4000",   282.6, "GÖVDE 6mm  — Kesim + Silindirik Bükme (kapama)"),
]
for i,(adet,boyut,kg,islem) in enumerate(t301_govde):
    data_row(ws,r,no_counter,f"[T-301]  {islem}",adet,boyut,kg,alt=(i%2==1))
    no_counter+=1; r+=1

tot_row(ws,r,2,7,"T-301 GÖVDE TOPLAM AĞIRLIK:",7453.2,H1); r+=1

# T-301 TABAN
sec_header(ws,r,2,12,
    "T-301  ATIK SU TANKI  —  TABAN SAC  |  St52-2",
    fill=T301_HDR, font_color="000000"); r+=1

t301_taban = [
    (1,  "⊏ 1500×8×6000", 565.2, "TABAN 8mm Orta Plaka — POZ 1 — Kesim"),
    (1,  "⊏ 1500×8×6000", 565.2, "TABAN 8mm Orta Plaka — POZ 2 — Kesim"),
    (1,  "⊏ 1500×8×6000", 565.2, "TABAN 8mm Orta Plaka — POZ 3 — Kesim"),
    (1,  "⊏ 1500×8×6000", 565.2, "TABAN 8mm Orta Plaka — POZ 4 — Kesim"),
    (1,  "⊏ 1500×8×4550", 430.0, "TABAN 8mm Orta Plaka — POZ 5 (kapama) — Kesim"),
    (4,  "⊏ 700×8×4470",  158.0, "TABAN 8mm Annüler/Çevre — POZ 6 ×4 — Kesim"),
    (1,  "⊏ 700×8×1830",   65.0, "TABAN 8mm Annüler/Çevre — POZ 7 (kapama) — Kesim"),
]
for i,(adet,boyut,kg,islem) in enumerate(t301_taban):
    data_row(ws,r,no_counter,f"[T-301]  {islem}",adet,boyut,kg,alt=(i%2==1))
    no_counter+=1; r+=1

tot_row(ws,r,2,7,"T-301 TABAN TOPLAM AĞIRLIK:",2913.8,H1); r+=1

# T-301 TAVAN
sec_header(ws,r,2,12,
    "T-301  ATIK SU TANKI  —  TAVAN SAC  |  St52-2",
    fill=T301_HDR, font_color="000000"); r+=1

t301_tavan = [
    (1,  "⊏ 1500×6×5700", 403.7, "TAVAN 6mm Sektör — POZ 1 — Kesim + Bükme"),
    (1,  "⊏ 1500×6×5700", 403.7, "TAVAN 6mm Sektör — POZ 2 — Kesim + Bükme"),
    (1,  "⊏ 1500×6×5700", 403.7, "TAVAN 6mm Sektör — POZ 3 — Kesim + Bükme"),
    (1,  "⊏ 1500×6×5700", 403.7, "TAVAN 6mm Sektör — POZ 4 — Kesim + Bükme"),
]
for i,(adet,boyut,kg,islem) in enumerate(t301_tavan):
    data_row(ws,r,no_counter,f"[T-301]  {islem}",adet,boyut,kg,alt=(i%2==1))
    no_counter+=1; r+=1

tot_row(ws,r,2,7,"T-301 TAVAN TOPLAM AĞIRLIK:",1614.8,H1); r+=1

# ── Genel Toplam ─────────────────────────────────────────────────────
r+=1
ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=6)
cl=ws.cell(r,2,"3 TANK  GENEL TOPLAM SAC AĞIRLIĞI  (T-201 + T-202 + T-301)")
cl.fill=copy(H1); cl.font=copy(fw); cl.alignment=copy(RA); cl.border=copy(MB)
cell(ws,r,7,f"=SUM(G4:G{r-1})",H1,fw,CA,MB)
cell(ws,r,8,"",H1,fw,CA,MB); cell(ws,r,9,"",H1,fw,CA,MB)
cell(ws,r,10,f"=SUM(J4:J{r-1})",H1,fw,CA,MB)
cell(ws,r,11,f"=SUM(K4:K{r-1})",H1,fw,CA,MB)
cell(ws,r,12,f"=SUM(L4:L{r-1})",H1,fw,CA,MB)
ws.row_dimensions[r].height=22

ws.sheet_view.showGridLines = False
ws.freeze_panes = "C4"

# ═══════════════════════════════════════════════════════════════════════
# TAŞIRMA sayfası — T-201 Gövde POZ detayı (fabrikatör için)
# ═══════════════════════════════════════════════════════════════════════
wst = wb["TAŞIRMA"]
for mc in list(wst.merged_cells.ranges):
    wst.unmerge_cells(str(mc))
wst.delete_rows(1, wst.max_row)

sw(wst, [3, 8, 8, 46, 14, 14, 14, 12, 26], start=1)
title_row(wst,1,2,9,"T-201 TAŞIRMA TANKI (25,000 BBL)  —  GÖVDE + TABAN + TAVAN PLAKA DETAYI  (Kesim/Bükme Fabrikatörü)", h=22)

HDR = ["POZ","ADET","AÇIKLAMA  (En × Kalınlık × Boy)  mm","MALZEME","BİRİM Kg","TOPLAM Kg","KALINLIK (mm)","İŞLEM TÜRÜ"]
for c,t in enumerate(HDR,2):
    cell(wst,2,c,t,H3,fb,CA,TB,h=36)

# Gövde
wst.merge_cells(start_row=3,start_column=2,end_row=3,end_column=9)
cl=wst.cell(3,2,"— GÖVDE SAC —  TOPLAM: 63,030 kg")
cl.fill=copy(T201_HDR); cl.font=Font(bold=True,color="000000",size=10)
cl.alignment=copy(CA); cl.border=copy(TB); wst.row_dimensions[3].height=18

t201_govde_detail = [
    ("1",   1,"⊏ 1500×12 × 12000","S275J2",1698.0, 1698.0,12,"Kesim + Silindirik Bükme"),
    ("2",   1,"⊏ 1500×12 × 12000","S275J2",1698.0, 1698.0,12,"Kesim + Silindirik Bükme"),
    ("3",   1,"⊏ 1500×12 × 12000","S275J2",1698.0, 1698.0,12,"Kesim + Silindirik Bükme"),
    ("4",   1,"⊏ 1500×12 × 12000","S275J2",1698.0, 1698.0,12,"Kesim + Silindirik Bükme"),
    ("5",   1,"⊏ 1500×12 × 11678","S275J2",1652.0, 1652.0,12,"Kesim + Silindirik Bükme (kapama)"),
    ("6",   3,"⊏ 1500×12 × 12000","S275J2",1698.0, 5094.0,12,"Kesim + Silindirik Bükme"),
    ("7",   1,"⊏ 1500×12 × 12000","S275J2",1698.0, 1698.0,12,"Kesim + Silindirik Bükme"),
    ("8",   1,"⊏ 1500×12 × 11678","S275J2",1652.0, 1652.0,12,"Kesim + Silindirik Bükme (kapama)"),
    ("9",  10,"⊏ 1500×10 × 12000","S275J2",1415.0,14150.0,10,"Kesim + Silindirik Bükme"),
    ("10",  3,"⊏ 1500×10 × 11678","S275J2",1377.0, 4131.0,10,"Kesim + Silindirik Bükme (kapama)"),
    ("11",  1,"⊏ 1500×10 × 12000","S275J2",1415.0, 1415.0,10,"Kesim + Silindirik Bükme"),
    ("12",  1,"⊏ 1500×10 × 12000","S275J2",1415.0, 1415.0,10,"Kesim + Silindirik Bükme"),
    ("13",  9,"⊏ 1500×8  × 12000","S275J2",1132.0,10188.0, 8,"Kesim + Silindirik Bükme"),
    ("14",  1,"⊏ 1500×8  × 12000","S275J2",1132.0, 1132.0, 8,"Kesim + Silindirik Bükme"),
    ("15",  3,"⊏ 1500×8  × 11678","S275J2",1102.0, 3306.0, 8,"Kesim + Silindirik Bükme (kapama)"),
    ("16",  1,"⊏ 1500×8  × 12000","S275J2",1132.0, 1132.0, 8,"Kesim + Silindirik Bükme"),
    ("17",  1,"⊏ 1500×8  × 12000","S275J2",1132.0, 1132.0, 8,"Kesim + Silindirik Bükme"),
    ("18",  1,"⊏ 1500×10 × 12000","S275J2",1415.0, 1415.0,10,"Kesim + Silindirik Bükme"),
    ("19", 40,"⊏ 225×10 × 245",   "S275J2",   3.0,  120.0,10,"Kesim (Düz — takviye parçası)"),
    ("20",  1,"⊏ 1500×10 × 10250","S275J2",1208.0, 1208.0,10,"Kesim + Silindirik Bükme (kapama)"),
    ("20A", 1,"⊏ 1500×10 × 5050", "S275J2", 595.0,  595.0,10,"Kesim + Silindirik Bükme (kapama)"),
    ("21",  5,"⊏ 425×10 × 12000", "S275J2", 401.0, 2005.0,10,"Kesim (RF/Yapısal)"),
    ("22",  1,"⊏ 425×10 × 11644", "S275J2", 389.0,  389.0,10,"Kesim (RF/Yapısal kapama)"),
    ("23",  5,"L 160×160×15 × 12000","S275J2",480.0,2400.0,15,"Kesim (L profil 12m çubuk)"),
    ("24",  1,"L 160×160×15 × 231","S275J2",   9.0,    9.0,15,"Kesim (L profil kapama)"),
    ("25",  5,'1/4" NPT PLUG',     "ASTM A105","—","—","—","Tedarik"),
]
dr=4
for i,(poz,adet,acik,malz,birim,toplam,thk,islem) in enumerate(t201_govde_detail):
    fill_=ALT if i%2==0 else WH
    for c,v in enumerate([poz,adet,acik,malz,birim,toplam,thk,islem],2):
        cell(wst,dr,c,v,fill_,fn9 if c==4 else fn,CA,TB,h=18)
    dr+=1

# Gövde toplam
wst.merge_cells(start_row=dr,start_column=2,end_row=dr,end_column=6)
cl=wst.cell(dr,2,"GÖVDE TOPLAM AĞIRLIK:")
cl.fill=copy(H1);cl.font=copy(fw);cl.alignment=copy(RA);cl.border=copy(TB)
cell(wst,dr,7,63030.0,H1,fw,CA,TB); cell(wst,dr,8,"",H1,fw,CA,TB); cell(wst,dr,9,"",H1,fw,CA,TB)
wst.row_dimensions[dr].height=18; dr+=2

# Taban özet
wst.merge_cells(start_row=dr,start_column=2,end_row=dr,end_column=9)
cl=wst.cell(dr,2,"— TABAN SAC —  TOPLAM: 15,959 kg")
cl.fill=copy(T201_HDR); cl.font=Font(bold=True,color="000000",size=10)
cl.alignment=copy(CA); cl.border=copy(TB); wst.row_dimensions[dr].height=18; dr+=1

t201_taban_detail = [
    ("1", 8, "⊏ 1475×14 × 6606","S275J2",670.0, 5361.0,14,"Kesim (Annüler/Çevre)"),
    ("2", 1, "⊏ 1475×14 × 6606","S275J2",537.0,  537.0,14,"Kesim (Annüler kapama)"),
    ("3", 9, "⊏ 1500×12 × 6000","S275J2",849.0, 7641.0,12,"Kesim (Orta plaka)"),
    ("4", 2, "⊏ 1500×12 × 6000","S275J2",849.0, 1698.0,12,"Kesim (Orta plaka)"),
    ("5", 2, "⊏ 441×12 × 5472", "S275J2",114.0,  228.0,12,"Kesim (Köşe)"),
    ("6", 2, "⊏ 1500×12 × 5706","S275J2",651.0, 1302.0,12,"Kesim"),
    ("7", 4, "⊏ 1500×12 × 5673","S275J2",788.0, 3152.0,12,"Kesim"),
    ("8", 4, "⊏ 1500×12 × 5409","S275J2",758.0, 3032.0,12,"Kesim"),
    ("9", 4, "⊏ 1500×12 × 4853","S275J2",675.0, 2700.0,12,"Kesim"),
    ("10",4, "⊏ 1500×12 × 3937","S275J2",462.0, 1848.0,12,"Kesim"),
    ("11",4, "⊏ 1410×12 × 2482","S275J2",186.0,  744.0,12,"Kesim"),
    ("12",1, "⊏ 50×6 × 261000", "S275J2",626.0,  626.0, 6,"Kesim (261m drenaj şerit)"),
]
for i,(poz,adet,acik,malz,birim,toplam,thk,islem) in enumerate(t201_taban_detail):
    fill_=ALT if i%2==0 else WH
    for c,v in enumerate([poz,adet,acik,malz,birim,toplam,thk,islem],2):
        cell(wst,dr,c,v,fill_,fn9 if c==4 else fn,CA,TB,h=18)
    dr+=1

wst.merge_cells(start_row=dr,start_column=2,end_row=dr,end_column=6)
cl=wst.cell(dr,2,"TABAN TOPLAM AĞIRLIK:")
cl.fill=copy(H1);cl.font=copy(fw);cl.alignment=copy(RA);cl.border=copy(TB)
cell(wst,dr,7,15959.0,H1,fw,CA,TB); cell(wst,dr,8,"",H1,fw,CA,TB); cell(wst,dr,9,"",H1,fw,CA,TB)
wst.row_dimensions[dr].height=18; dr+=2

# Tavan özet
wst.merge_cells(start_row=dr,start_column=2,end_row=dr,end_column=9)
cl=wst.cell(dr,2,"— TAVAN SAC —  TOPLAM: 18,797 kg")
cl.fill=copy(T201_HDR); cl.font=Font(bold=True,color="000000",size=10)
cl.alignment=copy(CA); cl.border=copy(TB); wst.row_dimensions[dr].height=18; dr+=1

t201_tavan_detail = [
    ("1",45,"⊏ 1300×8 × 4148","S275J2",264.0,11880.0,8,"Kesim + Bükme (Sektör)"),
    ("2", 1,"⊏ 1300×8 × 4148","S275J2",234.0,  234.0,8,"Kesim + Bükme (kapama)"),
    ("3", 1,"⊏ 1300×8 × 4148","S275J2",264.0,  264.0,8,"Kesim + Bükme"),
    ("4", 1,"⊏ 1300×8 × 4148","S275J2",264.0,  264.0,8,"Kesim + Bükme"),
    ("5",24,"⊏ 1466×8 × 4774","S275J2",252.0, 6048.0,8,"Kesim + Bükme (Sektör)"),
    ("6", 1,"ø1450×8 (Daire)", "S275J2",107.0,  107.0,8,"Kesim (Merkez disk)"),
]
for i,(poz,adet,acik,malz,birim,toplam,thk,islem) in enumerate(t201_tavan_detail):
    fill_=ALT if i%2==0 else WH
    for c,v in enumerate([poz,adet,acik,malz,birim,toplam,thk,islem],2):
        cell(wst,dr,c,v,fill_,fn9 if c==4 else fn,CA,TB,h=18)
    dr+=1

wst.merge_cells(start_row=dr,start_column=2,end_row=dr,end_column=6)
cl=wst.cell(dr,2,"TAVAN TOPLAM AĞIRLIK:")
cl.fill=copy(H1);cl.font=copy(fw);cl.alignment=copy(RA);cl.border=copy(TB)
cell(wst,dr,7,18797.0,H1,fw,CA,TB); cell(wst,dr,8,"",H1,fw,CA,TB); cell(wst,dr,9,"",H1,fw,CA,TB)
wst.row_dimensions[dr].height=18; dr+=1

# Grand total TAŞIRMA
wst.merge_cells(start_row=dr,start_column=2,end_row=dr,end_column=6)
cl=wst.cell(dr,2,"T-201 GENEL TOPLAM  (Gövde + Taban + Tavan):")
cl.fill=copy(H1);cl.font=copy(fw);cl.alignment=copy(RA);cl.border=copy(MB)
cell(wst,dr,7,63030+15959+18797,H1,fw,CA,MB); cell(wst,dr,8,"",H1,fw,CA,MB); cell(wst,dr,9,"",H1,fw,CA,MB)
wst.row_dimensions[dr].height=20

wst.sheet_view.showGridLines = False

# ═══════════════════════════════════════════════════════════════════════
# ÜRETİM sayfası — T-202 Shell POZ detayı
# ═══════════════════════════════════════════════════════════════════════
wsu = wb["ÜRETİM"]
for mc in list(wsu.merged_cells.ranges):
    wsu.unmerge_cells(str(mc))
wsu.delete_rows(1, wsu.max_row)

sw(wsu, [3, 8, 8, 44, 14, 14, 14, 12, 26], start=1)
title_row(wsu,1,2,9,"T-202 ÜRETİM TANKI  —  GÖVDE PLAKA DETAYI  (Kesim + Bükme Fabrikatörü İçin)",h=22)
for c,t in enumerate(["POZ","ADET","AÇIKLAMA  (En × Kalınlık × Boy)  mm",
                       "MALZEME","BİRİM\nKg","TOPLAM\nKg","KALINLIK\n(mm)","İŞLEM TÜRÜ"],2):
    cell(wsu,2,c,t,H3,fb,CA,TB,h=36)

govde_detail = [
    ("1",   1,"⊏ 1480×8 × 6000","S275J2", 557.6,  557.6, 8,"Kesim + Silindirik Bükme R=6365mm  |  Nozul N1 (Ø685) + N8 kesim"),
    ("1A",  1,"⊏ 1480×8 × 6000","S275J2", 557.0,  557.0, 8,"Kesim + Silindirik Bükme R=6365mm"),
    ("1B",  1,"⊏ 1480×8 × 6000","S275J2", 552.3,  552.3, 8,"Kesim + Silindirik Bükme R=6365mm  |  Küçük nozul kesimi"),
    ("1C",  1,"⊏ 1480×8 × 6000","S275J2", 539.1,  539.1, 8,"Kesim + Silindirik Bükme R=6365mm  |  Drain nozul kesimi"),
    ("1D",  1,"⊏ 1480×8 × 6000","S275J2", 414.0,  414.0, 8,"Kesim + Silindirik Bükme R=6365mm  |  Manhole M1 (Ø1255) kesim"),
    ("1E",  1,"⊏ 1480×8 × 6000","S275J2", 557.6,  557.6, 8,"Kesim + Silindirik Bükme R=6365mm"),
    ("2",   1,"⊏ 1480×8 × 3997","S275J2", 371.5,  371.5, 8,"Kesim + Silindirik Bükme R=6365mm  (Kapama plakası)"),
    ("3",  29,"⊏ 1480×6 × 6000","S275J2", 418.2,12127.8, 6,"Kesim + Silindirik Bükme R=6365mm"),
    ("3A",  1,"⊏ 1480×6 × 6000","S275J2", 418.0,  418.0, 6,"Kesim + Silindirik Bükme R=6365mm  (Kapama plakası)"),
    ("3B",  1,"⊏ 1480×6 × 6000","S275J2", 418.0,  418.0, 6,"Kesim + Silindirik Bükme R=6365mm  (Kapama plakası)"),
]
for i,(poz,adet,acik,malz,birim,toplam,thk,islem) in enumerate(govde_detail,start=3):
    fill_=ALT if i%2==0 else WH
    for c,v in enumerate([poz,adet,acik,malz,birim,toplam,thk,islem],2):
        cell(wsu,i,c,v,fill_,fn9 if c==4 else fn,CA,TB,h=22)

tr=3+len(govde_detail)
wsu.merge_cells(start_row=tr,start_column=2,end_row=tr,end_column=6)
cl=wsu.cell(tr,2,"GÖVDE TOPLAM AĞIRLIK:")
cl.fill=copy(H1);cl.font=copy(fw);cl.alignment=copy(RA);cl.border=copy(TB)
cell(wsu,tr,7,16513.0,H1,fw,CA,TB)
cell(wsu,tr,8,"",H1,fw,CA,TB); cell(wsu,tr,9,"",H1,fw,CA,TB)
wsu.row_dimensions[tr].height=18

tr+=2
wsu.merge_cells(start_row=tr,start_column=2,end_row=tr,end_column=9)
cl=wsu.cell(tr,2,"KURS (COURSE) ÖZETİ")
cl.fill=copy(H2);cl.font=copy(fw);cl.alignment=copy(CA);cl.border=copy(TB)
wsu.row_dimensions[tr].height=18; tr+=1
for c,t in enumerate(["KURS","KALINLIK","YÜKSEKLİK","PLAKA ADET","AÇIKLAMA","MALZEME","AĞIRLIK (kg)"],2):
    cell(wsu,tr,c,t,H3,fb,CA,TB,h=30)
tr+=1
for i,kd in enumerate([
    ("1. Kurs (Alt)","8 mm",1480,"7 adet (POZ 1,1A,1B,1C,1D,1E,2)","Nozul/Manhole kesimli","S275J2",3549.1),
    ("2–6. Kurslar", "6 mm",1480,"31 adet (POZ 3×29 + POZ 3A×1 + POZ 3B×1)","Standart silindirik","S275J2",12963.8),
]):
    fill_=ALT if i%2==0 else WH
    for c,v in enumerate(kd,2): cell(wsu,tr,c,v,fill_,fn,CA,TB,h=18)
    tr+=1

wsu.sheet_view.showGridLines = False

# ═══════════════════════════════════════════════════════════════════════
# ATIKSU sayfası — T-301
# ═══════════════════════════════════════════════════════════════════════
wsx = wb["ATIKSU"]
for mc in list(wsx.merged_cells.ranges):
    wsx.unmerge_cells(str(mc))
wsx.delete_rows(1, wsx.max_row)
sw(wsx,[3,8,8,44,14,14,14,12,26],start=1)
title_row(wsx,1,2,9,"T-301 ATIK SU TANKI (~2,000 BBL)  —  GÖVDE + TABAN + TAVAN PLAKA DETAYI",h=22)
for c,t in enumerate(["POZ","ADET","AÇIKLAMA","MALZEME","BİRİM Kg","TOPLAM Kg","KALINLIK (mm)","İŞLEM TÜRÜ"],2):
    cell(wsx,2,c,t,H3,fb,CA,TB,h=36)

wsx.merge_cells(start_row=3,start_column=2,end_row=3,end_column=9)
cl=wsx.cell(3,2,"— GÖVDE SAC —  St52-2")
cl.fill=copy(T301_HDR); cl.font=Font(bold=True,color="000000",size=10)
cl.alignment=copy(CA); cl.border=copy(TB); wsx.row_dimensions[3].height=18

t301_govde_det = [
    ("1",3,"⊏ 1500×10 × 6000","St52-2",706.5,2119.5,10,"Kesim + Silindirik Bükme"),
    ("2",1,"⊏ 1500×10 × 4000","St52-2",471.0, 471.0,10,"Kesim + Silindirik Bükme (kapama)"),
    ("3",3,"⊏ 1500×8  × 6000","St52-2",565.2,1695.6, 8,"Kesim + Silindirik Bükme"),
    ("4",1,"⊏ 1500×8  × 4000","St52-2",376.8, 376.8, 8,"Kesim + Silindirik Bükme (kapama)"),
    ("5",12,"⊏ 1500×6 × 6000","St52-2",423.9,5086.8, 6,"Kesim + Silindirik Bükme"),
    ("6",4,"⊏ 1500×6  × 4000","St52-2",282.6,1130.4, 6,"Kesim + Silindirik Bükme (kapama)"),
]
dr=4
for i,(poz,adet,acik,malz,birim,toplam,thk,islem) in enumerate(t301_govde_det):
    fill_=ALT if i%2==0 else WH
    for c,v in enumerate([poz,adet,acik,malz,birim,toplam,thk,islem],2):
        cell(wsx,dr,c,v,fill_,fn,CA,TB,h=18)
    dr+=1

wsx.merge_cells(start_row=dr,start_column=2,end_row=dr,end_column=6)
cl=wsx.cell(dr,2,"GÖVDE TOPLAM AĞIRLIK:")
cl.fill=copy(H1);cl.font=copy(fw);cl.alignment=copy(RA);cl.border=copy(TB)
cell(wsx,dr,7,10880.1,H1,fw,CA,TB); cell(wsx,dr,8,"",H1,fw,CA,TB); cell(wsx,dr,9,"",H1,fw,CA,TB)
wsx.row_dimensions[dr].height=18; dr+=2

wsx.merge_cells(start_row=dr,start_column=2,end_row=dr,end_column=9)
cl=wsx.cell(dr,2,"★ T-301 TABAN ve TAVAN: Çizim temin edildiğinde güncellenecektir (yaklaşık veriler TEKLIF sayfasındadır).")
cl.fill=copy(ORG); cl.font=Font(bold=True,size=10,color="CC0000")
cl.alignment=copy(CA); cl.border=copy(TB); wsx.row_dimensions[dr].height=24

wsx.sheet_view.showGridLines = False

wb.save(dst)
print(f"Kaydedildi: {dst}")
