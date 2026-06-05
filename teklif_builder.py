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
H1   = PatternFill("solid", fgColor="1F4E79")  # koyu mavi - başlık
H2   = PatternFill("solid", fgColor="2E75B6")  # orta mavi - bölüm başlığı
H3   = PatternFill("solid", fgColor="BDD7EE")  # açık mavi - sütun başlığı
ALT  = PatternFill("solid", fgColor="DEEAF1")  # alternatif satır
WH   = PatternFill("solid", fgColor="FFFFFF")
YEL  = PatternFill("solid", fgColor="FFFF99")  # birim fiyat girişi
GRN  = PatternFill("solid", fgColor="E2EFDA")  # toplam/özet
ORG  = PatternFill("solid", fgColor="FCE4D6")  # uyarı/not

fw  = Font(bold=True, color="FFFFFF", size=10)
fb  = Font(bold=True, color="1F4E79", size=10)
fb2 = Font(bold=True, color="FFFFFF", size=11)
fn  = Font(size=10)
fn9 = Font(size=9)

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

def sec_header(ws, r, c1, c2, text, fill=H2):
    ws.merge_cells(start_row=r,start_column=c1,end_row=r,end_column=c2)
    cl = ws.cell(row=r,column=c1,value=text)
    cl.fill=copy(fill); cl.font=copy(fw); cl.alignment=copy(CA); cl.border=copy(TB)
    ws.row_dimensions[r].height = 18

def note_row(ws, r, c1, c2, text, fill=ORG):
    ws.merge_cells(start_row=r,start_column=c1,end_row=r,end_column=c2)
    cl = ws.cell(row=r,column=c1,value=text)
    cl.fill=copy(fill); cl.font=copy(fn9); cl.alignment=copy(LA); cl.border=copy(TB)
    ws.row_dimensions[r].height = 14

def tot_row(ws, r, c1, cv, label, kg_val, fill=H1):
    ws.merge_cells(start_row=r,start_column=c1,end_row=r,end_column=cv-1)
    cl = ws.cell(row=r,column=c1,value=label)
    cl.fill=copy(fill); cl.font=copy(fw); cl.alignment=copy(RA); cl.border=copy(TB)
    cell(ws,r,cv,kg_val,fill,fw,CA,TB)

# ═══════════════════════════════════════════════════════════════════════
# TEKLIF SAYFASI — Tüm tanklar (sac malzemecisi + kesim/büküm fabrikatörü)
# ═══════════════════════════════════════════════════════════════════════
ws = wb["TEKLIF"]

# Mevcut içeriği temizle ve sıfırdan yaz
ws.delete_rows(1, ws.max_row)
for mc in list(ws.merged_cells.ranges):
    ws.unmerge_cells(str(mc))

# Sütun genişlikleri:
# A(boşluk) | B(No) | C(Tanım+İşlem) | D(Adet) | E(Boyut mm) | F(Net kg/adet) | G(Toplam kg) | H(Mat.Fiyat★) | I(İşçilik★) | J(Mat.Tutar) | K(İşç.Tutar) | L(TOPLAM $)
sw(ws, [3, 6, 54, 9, 22, 12, 13, 12, 12, 14, 14, 16], start=1)

# ── Başlık ──────────────────────────────────────────────────────────
title_row(ws, 1, 2, 12, "TANK SATINALMA SİPARİŞ TABLOSU  —  SAC/PLAKA METRAJİ  (3 TANK)", h=26)
ws.merge_cells(start_row=2,start_column=2,end_row=2,end_column=12)
cl=ws.cell(2,2,"Sac Malzemecisi: Kolon H (Malzeme Birim Fiyatı $/kg) girin.   "
                "Büküm/Kesim Fabrikatörü: Kolon I (İşçilik Birim Fiyatı $/kg) girin.")
cl.fill=copy(YEL); cl.font=Font(bold=True,size=9,color="CC0000")
cl.alignment=copy(CA); cl.border=copy(TB); ws.row_dimensions[2].height=18

# Kolon başlıkları
r=3
for c,txt in enumerate(["S.NO","TANIM  /  İŞLEM TÜRÜ","ADET","BOYUT (mm)",
                         "BİRİM\nKg","TOPLAM\nKg","MAT. FİYAT\n★($/kg)",
                         "İŞÇİLİK\n★($/kg)","MAT. TUTARI\n($)","İŞÇİLİK\n($)","TOPLAM\n($)"],1):
    cell(ws,r,c+2,txt,H3,fb,CA,TB,h=38)

# ── Veri yazma yardımcıları ──────────────────────────────────────────
def data_row(ws, r, no, tanim, adet, boyut, birim_kg, alt=False, first_of_tank=False, tank_label=""):
    fill_ = ALT if alt else WH
    # No
    cell(ws,r,2,no,fill_,fn,CA,TB)
    # Tanım
    cl=ws.cell(row=r,column=3,value=tanim)
    cl.fill=copy(fill_); cl.font=copy(fn9); cl.alignment=copy(LA); cl.border=copy(TB)
    # Adet
    cell(ws,r,4,adet,fill_,fn,CA,TB)
    # Boyut
    cell(ws,r,5,boyut,fill_,fn9,CA,TB)
    # Birim kg
    cell(ws,r,6,birim_kg,fill_,fn,CA,TB)
    # Toplam kg = adet × birim_kg
    tot_kg = round(adet * birim_kg, 1) if isinstance(adet,int) and isinstance(birim_kg,float) else ""
    cell(ws,r,7,tot_kg,fill_,fn,CA,TB)
    # Malzeme fiyat (sarı)
    cell(ws,r,8,"",YEL,fn,CA,TB)
    # İşçilik fiyat (sarı)
    cell(ws,r,9,"",YEL,fn,CA,TB)
    # Mat tutarı = TotKg × MatFiyat
    cl=ws.cell(row=r,column=10,value=f"=G{r}*H{r}")
    cl.fill=copy(fill_); cl.font=copy(fn); cl.alignment=copy(CA); cl.border=copy(TB)
    # İşçilik tutarı = TotKg × İşçilik
    cl=ws.cell(row=r,column=11,value=f"=G{r}*I{r}")
    cl.fill=copy(fill_); cl.font=copy(fn); cl.alignment=copy(CA); cl.border=copy(TB)
    # Toplam
    cl=ws.cell(row=r,column=12,value=f"=J{r}+K{r}")
    cl.fill=copy(GRN); cl.font=copy(fn); cl.alignment=copy(CA); cl.border=copy(TB)
    ws.row_dimensions[r].height=16

# ════════════════════════════════════════════════════════════════════
# T-201 TAŞIRMA TANKI  (Mevcut veri korundu - çizim eklendiğinde güncellenecek)
# ════════════════════════════════════════════════════════════════════
r=4
sec_header(ws,r,2,12,"T-201  TAŞIRMA TANKI  (Wastewater/Overflow)  —  GÖVDE SAC  |  Malzeme: S275J2"); r+=1

t201_govde = [
    # Adet, Boyut, BirimKg - mevcut template verisinden
    (8,  "1500×12×12000", 1698.0, "GÖVDE — Kesim + Silindirik Bükme"),
    (2,  "1500×12×11678", 1652.0, "GÖVDE — Kesim + Silindirik Bükme (kapama)"),
    (12, "1500×10×12000", 1415.0, "GÖVDE — Kesim + Silindirik Bükme"),
    (3,  "1500×10×11678", 1377.0, "GÖVDE — Kesim + Silindirik Bükme (kapama)"),
    (12, "1500×8×12000",  1132.0, "GÖVDE — Kesim + Silindirik Bükme"),
    (3,  "1500×8×11678",  1102.0, "GÖVDE — Kesim + Silindirik Bükme (kapama)"),
]
no_counter=1
for i,(adet,boyut,kg,islem) in enumerate(t201_govde):
    tanim = f"[T-201]  {islem}"
    data_row(ws, r, no_counter, tanim, adet, boyut, kg, alt=(i%2==1))
    no_counter+=1; r+=1

note_row(ws,r,2,12,"★ T-201 TABAN ve TAVAN sacları: Çizimler temin edildiğinde eklenecektir."); r+=1

# ════════════════════════════════════════════════════════════════════
# T-202 ÜRETİM TANKI  —  GÖVDE
# ════════════════════════════════════════════════════════════════════
r+=1
sec_header(ws,r,2,12,"T-202  ÜRETİM TANKI (T-202)  —  GÖVDE SAC  |  D=12.730m  H=8.920m  |  S275J2"); r+=1

t202_govde = [
    # 8mm kurs 1 (alt):
    (6,  "1480×8×6000",  557.6, "GÖVDE 8mm — 1. Kurs (POZ 1,1A,1B,1C,1D,1E) — Kesim + Silindirik Bükme R=6365mm\n"
                                 "NOT: Her plakada farklı nozul/manhole kesimi var → Çizim KEY199MECEQP2001"),
    (1,  "1480×8×3997",  371.5, "GÖVDE 8mm — 1. Kurs Kapama (POZ 2) — Kesim + Silindirik Bükme R=6365mm"),
    # 6mm kurslar 2-6:
    (30, "1480×6×6000",  418.2, "GÖVDE 6mm — 2–6. Kurs (POZ 3 ×30) — Kesim + Silindirik Bükme R=6365mm"),
    (1,  "1480×6×6000",  418.0, "GÖVDE 6mm — 2–6. Kurs Kapama (POZ 3A) — Kesim + Silindirik Bükme R=6365mm"),
]
for i,(adet,boyut,kg,islem) in enumerate(t202_govde):
    data_row(ws,r,no_counter,f"[T-202]  {islem}",adet,boyut,kg,alt=(i%2==1))
    no_counter+=1; r+=1

# ── T-202 TABAN ──────────────────────────────────────────────────
sec_header(ws,r,2,12,"T-202  ÜRETİM TANKI  —  TABAN SAC  |  S275J2  (Çizim: KEY199MECEQP2003)"); r+=1

t202_taban = [
    # 10mm orta plakalar
    (7,  "1500×10×5990",   705.3, "TABAN 10mm Orta (POZ 1 ×7) — Kesim  (Düz, kaynak payı dahil)"),
    (2,  "1500×10×2684",   314.0, "TABAN 10mm Orta (POZ 2 ×2) — Kesim"),
    (4,  "1500×10×635",     55.5, "TABAN 10mm Orta (POZ 3 ×4) — Kesim  (Kısa strip)"),
    (3,  "1500×10×2262",   219.2, "TABAN 10mm Orta (POZ 4 ×3) — Kesim"),
    (1,  "1500×10×2676",   222.0, "TABAN 10mm Orta (POZ 5 ×1) — Kesim"),
    (1,  "1500×10×4395",   557.0, "TABAN 10mm Orta (POZ 6 ×1) — Kesim"),
    (2,  "1500×10×2994",   347.4, "TABAN 10mm Orta (POZ 7 ×2) — Kesim"),
    (2,  "632×10×5208",    174.1, "TABAN 10mm Orta (POZ 8 ×2) — Kesim  (Dar plaka)"),
    # 12mm annüler (çevre)
    (6,  "800×12×5937",    436.1, "TABAN 12mm Annüler/Çevre (POZ 9 ×6) — Kesim  (1500mm plakadan boyuna bölünecek)"),
    (1,  "800×12×3327",    237.8, "TABAN 12mm Annüler/Çevre (POZ 9A ×1) — Kesim (Kapama)"),
    # Drenaj şeritler (5mm) - küçük parçalar
    (23, "50×5×çeşitli",     2.4, "TABAN 5mm Drenaj Strip (POZ 10-13) — Kesim  (Atık sacdan kesilebilir)"),
]
for i,(adet,boyut,kg,islem) in enumerate(t202_taban):
    data_row(ws,r,no_counter,f"[T-202]  {islem}",adet,boyut,kg,alt=(i%2==1))
    no_counter+=1; r+=1

# ── T-202 TAVAN ──────────────────────────────────────────────────
sec_header(ws,r,2,12,"T-202  ÜRETİM TANKI  —  TAVAN SAC (Konik Sektörler)  |  S275J2  (Çizim: KEY199MECEQP2002)"); r+=1
note_row(ws,r,2,12,
    "İŞLEM: Tüm tavan sacları KESİM + KONİK BÜKME gerektirir.  "
    "Koniklik açısı çizimden kontrol edilecek.  Sektör şablonu fabrikatörde yapılacak."); r+=1

t202_tavan = [
    (4,  "1500×8×6000",   565.2, "TAVAN 8mm Konik Sektör (POZ 1 ×4) — Kesim + Konik Bükme"),
    (1,  "1500×8×6000",   560.0, "TAVAN 8mm Konik Sektör (POZ 1A) — Kesim + Konik Bükme"),
    (1,  "1500×8×6000",   560.0, "TAVAN 8mm Konik Sektör (POZ 1B) — Kesim + Konik Bükme"),
    (1,  "1500×8×6000",   560.0, "TAVAN 8mm Konik Sektör (POZ 1C) — Kesim + Konik Bükme"),
    (2,  "1500×8×3561",   334.0, "TAVAN 8mm Konik Sektör (POZ 2 ×2) — Kesim + Konik Bükme"),
    (4,  "1500×8×548",     38.3, "TAVAN 8mm Konik Sektör (POZ 3 ×4) — Kesim + Konik Bükme  (Küçük sektör)"),
    (4,  "1500×8×3196",   269.3, "TAVAN 8mm Konik Sektör (POZ 4 ×4) — Kesim + Konik Bükme"),
    (2,  "1500×8×2873",   257.3, "TAVAN 8mm Konik Sektör (POZ 5 ×2) — Kesim + Konik Bükme"),
    (1,  "1500×8×2873",   254.5, "TAVAN 8mm Konik Sektör (POZ 5A) — Kesim + Konik Bükme"),
    (1,  "1500×8×2873",   254.5, "TAVAN 8mm Konik Sektör (POZ 5B) — Kesim + Konik Bükme"),
    (3,  "1500×8×2546",   208.0, "TAVAN 8mm Konik Sektör (POZ 6 ×3) — Kesim + Konik Bükme"),
    (1,  "1500×8×2546",   190.0, "TAVAN 8mm Konik Sektör (POZ 6A) — Kesim + Konik Bükme"),
    (4,  "1000×8×1825",    90.7, "TAVAN 8mm Konik Sektör (POZ 7 ×4) — Kesim + Konik Bükme  (Dar sektör)"),
    (4,  "1057×8×875",     31.0, "TAVAN 8mm Konik Sektör (POZ 8 ×4) — Kesim + Konik Bükme  (Küçük köşe)"),
    (2,  "1500×8×2911",   272.8, "TAVAN 8mm Konik Sektör (POZ 9 ×2) — Kesim + Konik Bükme"),
]
for i,(adet,boyut,kg,islem) in enumerate(t202_tavan):
    data_row(ws,r,no_counter,f"[T-202]  {islem}",adet,boyut,kg,alt=(i%2==1))
    no_counter+=1; r+=1

# ════════════════════════════════════════════════════════════════════
# T-301 ATIK SU TANKI  (Mevcut veri korundu)
# ════════════════════════════════════════════════════════════════════
r+=1
sec_header(ws,r,2,12,"T-301  ATIK SU TANKI  —  GÖVDE SAC  |  Malzeme: St52-2"); r+=1

t301_govde = [
    (3,  "1500×10×6000",  706.5, "GÖVDE — Kesim + Silindirik Bükme"),
    (1,  "1500×10×4000",  471.0, "GÖVDE — Kesim + Silindirik Bükme (kapama)"),
    (3,  "1500×8×6000",   565.2, "GÖVDE — Kesim + Silindirik Bükme"),
    (1,  "1500×8×4000",   376.8, "GÖVDE — Kesim + Silindirik Bükme (kapama)"),
    (12, "1500×6×6000",   423.9, "GÖVDE — Kesim + Silindirik Bükme"),
    (4,  "1500×6×4000",   282.6, "GÖVDE — Kesim + Silindirik Bükme (kapama)"),
]
for i,(adet,boyut,kg,islem) in enumerate(t301_govde):
    data_row(ws,r,no_counter,f"[T-301]  {islem}",adet,boyut,kg,alt=(i%2==1))
    no_counter+=1; r+=1

note_row(ws,r,2,12,"★ T-301 TABAN ve TAVAN sacları: Çizimler temin edildiğinde eklenecektir."); r+=1

# ── Genel Ağırlık Özet Satırı ────────────────────────────────────
r+=1
ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=6)
cl=ws.cell(r,2,"TOPLAM NET SAC AĞIRLIĞI (Bilinen Kalemler)  —  Fiyat Formülleri Otomatik Çalışır")
cl.fill=copy(H1); cl.font=copy(fw); cl.alignment=copy(RA); cl.border=copy(MB)
# Toplam kg formülü (sütun G)
cell(ws,r,7,f"=SUM(G4:G{r-1})",H1,fw,CA,MB)
cell(ws,r,8,"",H1,fw,CA,MB)
cell(ws,r,9,"",H1,fw,CA,MB)
cell(ws,r,10,f"=SUM(J4:J{r-1})",H1,fw,CA,MB)
cell(ws,r,11,f"=SUM(K4:K{r-1})",H1,fw,CA,MB)
cell(ws,r,12,f"=SUM(L4:L{r-1})",H1,fw,CA,MB)
ws.row_dimensions[r].height=20

ws.sheet_view.showGridLines = False
ws.freeze_panes = "C4"

# ═══════════════════════════════════════════════════════════════════════
# ÜRETİM SAYFASI — T-202 Shell detay (fabrikatör için)
# ═══════════════════════════════════════════════════════════════════════
wsu = wb["ÜRETİM"]
wsu.delete_rows(1, wsu.max_row)
for mc in list(wsu.merged_cells.ranges):
    wsu.unmerge_cells(str(mc))

sw(wsu, [3, 8, 8, 44, 14, 14, 14, 12, 26], start=1)

title_row(wsu, 1, 2, 9, "T-202 ÜRETİM TANKI  —  GÖVDE PLAKA DETAYI  (Kesim + Bükme Fabrikatörü İçin)", h=22)
HDR_TEXTS = ["POZ","ADET","AÇIKLAMA  (En × Kalınlık × Boy)  mm",
             "MALZEME","BİRİM\nKg","TOPLAM\nKg","KALINLIK\n(mm)","İŞLEM TÜRÜ"]
for c,t in enumerate(HDR_TEXTS,2):
    cell(wsu,2,c,t,H3,fb,CA,TB,h=36)

govde_detail = [
    ("1",   1, "⊏ 1480×8 × 6000",  "S275J2", 557.6,   557.6,  8, "Kesim + Silindirik Bükme R=6365mm  |  Nozul N1 (Ø685) + N8 kesim var"),
    ("1A",  1, "⊏ 1480×8 × 6000",  "S275J2", 557.0,   557.0,  8, "Kesim + Silindirik Bükme R=6365mm"),
    ("1B",  1, "⊏ 1480×8 × 6000",  "S275J2", 552.3,   552.3,  8, "Kesim + Silindirik Bükme R=6365mm  |  Küçük nozul kesimi"),
    ("1C",  1, "⊏ 1480×8 × 6000",  "S275J2", 539.1,   539.1,  8, "Kesim + Silindirik Bükme R=6365mm  |  Drain nozul kesimi"),
    ("1D",  1, "⊏ 1480×8 × 6000",  "S275J2", 414.0,   414.0,  8, "Kesim + Silindirik Bükme R=6365mm  |  Manhole M1 (Ø1255) kesim var"),
    ("1E",  1, "⊏ 1480×8 × 6000",  "S275J2", 557.6,   557.6,  8, "Kesim + Silindirik Bükme R=6365mm"),
    ("2",   1, "⊏ 1480×8 × 3997",  "S275J2", 371.5,   371.5,  8, "Kesim + Silindirik Bükme R=6365mm  (Kapama plakası)"),
    ("3",  30, "⊏ 1480×6 × 6000",  "S275J2", 418.2, 12546.0,  6, "Kesim + Silindirik Bükme R=6365mm"),
    ("3A",  1, "⊏ 1480×6 × 6000",  "S275J2", 418.0,   418.0,  6, "Kesim + Silindirik Bükme R=6365mm  (Kapama plakası)"),
]
for i,(poz,adet,acik,malz,birim,toplam,thk,islem) in enumerate(govde_detail,start=3):
    fill_=ALT if i%2==0 else WH
    for c,v in enumerate([poz,adet,acik,malz,birim,toplam,thk,islem],2):
        cell(wsu,i,c,v,fill_,fn9 if c==4 else fn,CA if c!=4 else LA,TB,h=22)

# Toplam
tr=3+len(govde_detail)
wsu.merge_cells(start_row=tr,start_column=2,end_row=tr,end_column=6)
cl=wsu.cell(tr,2,"GÖVDE TOPLAM AĞIRLIK:")
cl.fill=copy(H1);cl.font=copy(fw);cl.alignment=copy(RA);cl.border=copy(TB)
cell(wsu,tr,7,16513.0,H1,fw,CA,TB)
cell(wsu,tr,8,"",H1,fw,CA,TB); cell(wsu,tr,9,"",H1,fw,CA,TB)
wsu.row_dimensions[tr].height=18

# Kurs özeti
tr+=2
wsu.merge_cells(start_row=tr,start_column=2,end_row=tr,end_column=9)
cl=wsu.cell(tr,2,"KURS (COURSE) ÖZETİ")
cl.fill=copy(H2);cl.font=copy(fw);cl.alignment=copy(CA);cl.border=copy(TB); wsu.row_dimensions[tr].height=18; tr+=1
for c,t in enumerate(["KURS","KALINLIK","YÜKSEKLİK","PLAKA ADET","AÇIKLAMA","MALZEME","AĞIRLIK (kg)"],2):
    cell(wsu,tr,c,t,H3,fb,CA,TB,h=30)
tr+=1
kurs=[
    ("1. Kurs (Alt)","8 mm",1480,"7 adet (POZ 1,1A,1B,1C,1D,1E,2)","Nozul/Manhole kesimli","S275J2",3549.1),
    ("2–6. Kurslar","6 mm",1480,"31 adet (POZ 3×30 + POZ 3A×1)","Standart silindirik","S275J2",12964.0),
]
for i,kd in enumerate(kurs):
    fill_=ALT if i%2==0 else WH
    for c,v in enumerate(kd,2): cell(wsu,tr,c,v,fill_,fn,CA,TB,h=18)
    tr+=1

wsu.sheet_view.showGridLines = False

# ═══════════════════════════════════════════════════════════════════════
# TAŞIRMA ve ATIKSU — Başlık güncelle, boş bırak (çizim gelince doldurulacak)
# ═══════════════════════════════════════════════════════════════════════
for sh_name, tank_name in [("TAŞIRMA","T-201 TAŞIRMA TANKI"), ("ATIKSU","T-301 ATIK SU TANKI")]:
    wsx = wb[sh_name]
    wsx.delete_rows(1, wsx.max_row)
    for mc in list(wsx.merged_cells.ranges):
        wsx.unmerge_cells(str(mc))
    sw(wsx,[3,8,8,44,14,14,14,12,26],start=1)
    title_row(wsx,1,2,9,f"{tank_name}  —  GÖVDE PLAKA DETAYI  (Çizimler Bekleniyor)",h=22)
    for c,t in enumerate(["POZ","ADET","AÇIKLAMA","MALZEME","BİRİM Kg","TOPLAM Kg","KALINLIK (mm)","İŞLEM TÜRÜ"],2):
        cell(wsx,2,c,t,H3,fb,CA,TB,h=36)
    wsx.merge_cells(start_row=3,start_column=2,end_row=3,end_column=9)
    cl=wsx.cell(3,2,f"★ {tank_name} çizimleri temin edildiğinde bu sayfa doldurulacaktır.")
    cl.fill=copy(ORG); cl.font=Font(bold=True,size=11,color="CC0000")
    cl.alignment=copy(CA); cl.border=copy(TB); wsx.row_dimensions[3].height=30
    wsx.sheet_view.showGridLines = False

wb.save(dst)
print(f"Kaydedildi: {dst}")
