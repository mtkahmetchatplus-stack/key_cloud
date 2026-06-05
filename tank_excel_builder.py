import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# ── Stil sabitleri ────────────────────────────────────────────────────────────
H1   = PatternFill("solid", fgColor="1F4E79")
H2   = PatternFill("solid", fgColor="2E75B6")
H3   = PatternFill("solid", fgColor="BDD7EE")
ALT  = PatternFill("solid", fgColor="DEEAF1")
WH   = PatternFill("solid", fgColor="FFFFFF")
YEL  = PatternFill("solid", fgColor="FFFF99")
GRN  = PatternFill("solid", fgColor="E2EFDA")
ORG  = PatternFill("solid", fgColor="FCE4D6")
RED  = PatternFill("solid", fgColor="FF0000")

fw  = Font(bold=True, color="FFFFFF", size=10)
fb  = Font(bold=True, color="1F4E79", size=10)
fn  = Font(size=10)
fn9 = Font(size=9)
fwb = Font(bold=True, color="FFFFFF", size=11)

CA = Alignment(horizontal="center", vertical="center", wrap_text=True)
LA = Alignment(horizontal="left",   vertical="center", wrap_text=True)
RA = Alignment(horizontal="right",  vertical="center")

t = Side(style="thin",   color="4472C4")
m = Side(style="medium", color="1F4E79")
TB = Border(left=t,right=t,top=t,bottom=t)
MB = Border(left=m,right=m,top=m,bottom=m)

def sw(ws, widths):
    for i,w in enumerate(widths,1):
        ws.column_dimensions[get_column_letter(i)].width = w

def cell(ws, r, c, v, fill=WH, font=fn, align=CA, border=TB):
    cl = ws.cell(row=r, column=c, value=v)
    cl.fill=fill; cl.font=font; cl.alignment=align; cl.border=border
    return cl

def hdr(ws, r, cols, texts, fill=H2, font=fw):
    for c,t in zip(cols,texts):
        cell(ws,r,c,t,fill,font,CA,TB)

def title(ws, r, c1, c2, text, fill=H1, font=fwb, h=22):
    ws.merge_cells(start_row=r,start_column=c1,end_row=r,end_column=c2)
    cl=ws.cell(row=r,column=c1,value=text)
    cl.fill=fill; cl.font=font; cl.alignment=CA; cl.border=MB
    ws.row_dimensions[r].height=h

def sub_title(ws, r, c1, c2, text):
    ws.merge_cells(start_row=r,start_column=c1,end_row=r,end_column=c2)
    cl=ws.cell(row=r,column=c1,value=text)
    cl.fill=H3; cl.font=fb; cl.alignment=CA; cl.border=TB

def data(ws, r, cols, vals, alt=False):
    fill = ALT if alt else WH
    for c,v in zip(cols,vals):
        cell(ws,r,c,v,fill,fn,CA,TB)

def total_row(ws, r, c1, c2, label, val_col, val, fill=H2):
    ws.merge_cells(start_row=r,start_column=c1,end_row=r,end_column=val_col-1)
    cl=ws.cell(row=r,column=c1,value=label)
    cl.fill=fill;cl.font=fw;cl.alignment=RA;cl.border=TB
    cell(ws,r,val_col,val,fill,fw,CA,TB)

# ═════════════════════════════════════════════════════════════════════════════
# SAYFA 1: NOZUL LİSTESİ
# ═════════════════════════════════════════════════════════════════════════════
ws1 = wb.active; ws1.title="Nozul Detayları"; ws1.sheet_view.showGridLines=False

nozzles = [
    # (Başlık, imalat adedi, notlar, [(poz,adet,açıklama,malzeme,birim,toplam),...], toplam_kg)
    ("(12\") N1 = NOZULU", 1, "Her tank için 1 adet", [
        (1,1,'12" SCH XS BORU ..... 215',"ASTM A 106 Gr.B",21.5,21.5),
        (2,1,'12" 150# SO. RF. FLANS',"ASTM A 105",29.0,29.0),
        (3,1,"⊏ 8×327×685  PLAKA","S275J2",18.0,18.0),
    ], 68.5),
    ("(8\") N8 = NOZULU", 1, "Her tank için 1 adet", [
        (1,1,'8" SCH XS BORU ..... 221',"ASTM A 106 Gr.B",11.4,11.4),
        (2,1,'8" 150# SO. RF. FLANS',"ASTM A 105",13.5,13.5),
        (3,1,"⊏ 6×485×222  PLAKA","S275J2",7.0,7.0),
    ], 31.9),
    ("(4\") N9 = NOZULU", 1, "Her tank için 1 adet", [
        (1,1,'4" SCH XS BORU ..... 215',"ASTM A 106 Gr.B",4.8,4.8),
        (2,1,'4" 150# SO. RF. FLANS',"ASTM A 105",5.9,5.9),
        (3,1,"⊏ 8×305×117  PLAKA","A 283 Gr.C",4.0,4.0),
    ], 14.7),
    ("(1½\") N9 = NOZULU", 1, "Her tank için 1 adet", [
        (1,1,'1½" SCH XS BORU ..... 165',"ASTM A 106 Gr.B",0.70,0.70),
        (2,1,'1½" 150# SO. RF. FLANS',"ASTM A 105",1.95,1.95),
    ], 2.65),
    ("(1\") N18 = NOZULU", 1, "Her tank için 1 adet", [
        (1,1,'1" SCH XS BORU ..... 165',"ASTM A 106 Gr.B",0.70,0.70),
        (2,1,'1" 150# SO. RF. FLANS',"ASTM A 105",1.0,1.0),
    ], 1.7),
    ("(12\") N21 = NOZULU", 4, "4 adet imal edilecektir", [
        (1,1,'12" SCH XS BORU ..... 290',"ASTM A 106 Gr.B",28.0,28.0),
        (2,1,'12"-90° Elbow LR BW ASME B16.5 SCH40',"ASTM A 234 Gr.WPB",29.0,29.0),
        (3,1,"⊏ 8×327×685  PLAKA","S275J2",18.0,18.0),
    ], 61.1),
    ("(6\") N3 = NOZULU", 2, "2 adet imal edilecektir", [
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
]

row = 1
title(ws1, row, 1, 8, "7000 Bbl ÜRETİM TANKI T-202A/B/C — NOZUL DETAY MALZEME LİSTELERİ")
row += 1

for noz_name, imalat_adet, not_text, items, toplam in nozzles:
    sub_title(ws1, row, 1, 8, f"{noz_name}   |   İmalat Adedi: {imalat_adet}   |   {not_text}")
    row += 1
    hdr(ws1, row, [1,2,3,4,5,6,7,8],
        ["POZ","ADET","AÇIKLAMA","MALZEME","BİRİM\nAĞIRLIK (kg)","TOPLAM\nAĞIRLIK (kg)",
         "OP. TÜRÜ","NOT"], H2, fw)
    ws1.row_dimensions[row].height=30; row+=1
    for i,(poz,adet,acik,malz,birim,top) in enumerate(items):
        op = ""
        if "BORU" in acik: op="Kesim + Kaynak"
        elif "FLANS" in acik: op="Kaynak"
        elif "PLAKA" in acik or "⊏" in acik: op="Kesim + Kaynak"
        elif "DİRSEK" in acik or "Elbow" in acik: op="Kaynak"
        elif "SOMUN" in acik or "CİVATA" in acik or "NPT" in acik: op="Montaj"
        data(ws1, row, [1,2,3,4,5,6,7,8],
             [poz,adet,acik,malz,birim,top,op,""], alt=(i%2==1))
        row+=1
    # Toplam
    ws1.merge_cells(start_row=row,start_column=1,end_row=row,end_column=5)
    cl=ws1.cell(row=row,column=1,value=f"TOPLAM AĞIRLIK ({imalat_adet} ADET): {toplam*imalat_adet:.1f} Kg  (1 adet: {toplam} Kg)")
    cl.fill=H1;cl.font=fw;cl.alignment=RA;cl.border=TB
    cell(ws1,row,6,round(toplam*imalat_adet,1),H1,fw,CA,TB)
    for c in [7,8]: cell(ws1,row,c,"",H1,fw,CA,TB)
    row+=2

sw(ws1,[8,8,42,20,14,14,20,20])

# ═════════════════════════════════════════════════════════════════════════════
# SAYFA 2: MANHOL LİSTESİ
# ═════════════════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("Manhol Detayları"); ws2.sheet_view.showGridLines=False

manhols = [
    ('24" GÖVDE MANHOLU (M1)', [
        (1,1,'⊂ ø176×8 .........1885',"S275J2",20.8,20.8),
        (2,1,'⊏ ø820\\ø616×10  HALKA',"S275J2",18.0,18.0),
        (3,1,"ø820×12  DAİRE PLAKA","S275J2",49.8,49.8),
        (4,1,'⊏ 1255×8 .........1525',"S275J2",69.5,69.5),
        (5,1,"ø735×ø610×3  GASKET","KLİNGERİT","–","–"),
        (6,2,"ø16×318  YUVARLAK BORU","S235JR",1.0,2.0),
        (7,1,"ø40×1083  BORU","S235JR",11.0,11.0),
        (8,1,'⊏ 100×20 .........200',"S275J2",1.6,1.6),
        (9,1,'⊏ 100×20 .........200',"S275J2",1.6,1.6),
        (10,1,"⊏ ø70/ø42×30  HALKA","S275J2",1.0,1.0),
        (11,1,'⊏ 60×25 .........70',"S275J2",1.0,1.0),
        (12,1,'⊏ 67×14 .........80',"S275J2",0.5,0.5),
        (13,2,'⊏ 50×14 .........56',"S275J2",0.5,1.0),
        (14,1,'⊏ 46×14 .........50',"S275J2",0.5,0.5),
        (15,1,"ø24×200  YUVARLAK BORU","S235JR",1.0,1.0),
        (16,28,"CİVATA M20×75","DIN 601","–","–"),
        (17,1,"CİVATA M20×110","DIN 601","–","–"),
        (18,2,"SOMUN M24","DIN 555","–","–"),
        (19,58,"SOMUN M20","DIN 555","–","–"),
        (20,58,"RONDELA D 22/37  S=3","DIN 126 Gr.C","–","–"),
        (21,1,'1/4" NPT TAPA',"","–","–"),
    ], 180.0),
    ('24" TAVAN MANHOLU (M2)', [
        (1,1,'⊏ 400×8 .........1910',"S275J2",48.0,48.0),
        (2,1,'⊏ ø750/ø616×8  HALKA',"S275J2",10.0,10.0),
        (3,1,"ø750×8  DAİRE PLAKA","S275J2",28.0,28.0),
        (4,1,'⊏ ø1150/ø616×8  HALKA',"S275J2",46.5,46.5),
        (5,1,"ø750×ø616×1.5  GASKET","KLİNGERİT","–","–"),
        (6,2,"ø16×318  YUVARLAK BORU","S235JR",1.0,2.0),
        (7,20,"CİVATA M16×50","DIN 601","–","–"),
        (8,20,"SOMUN M16","DIN 555","–","–"),
        (9,20,"RONDELA D 17/30  S=3","DIN 126 Gr.C","–","–"),
        (10,1,'1/4" NPT TAPA',"","–","–"),
    ], 132.5),
]

row=1
title(ws2,row,1,7,"7000 Bbl ÜRETİM TANKI T-202A/B/C — MANHOL MALZEME LİSTELERİ")
row+=1
for mnh_name, items, toplam in manhols:
    sub_title(ws2,row,1,7,mnh_name)
    row+=1
    hdr(ws2,row,[1,2,3,4,5,6,7],
        ["POZ","ADET","AÇIKLAMA","MALZEME","BİRİM\nAĞIRLIK (kg)","TOPLAM\nAĞIRLIK (kg)","OP. TÜRÜ"],H2,fw)
    ws2.row_dimensions[row].height=30; row+=1
    for i,(poz,adet,acik,malz,birim,top) in enumerate(items):
        op=""
        if "BORU" in acik and "YUVARLAK" not in acik: op="Kesim + Kaynak"
        elif "PLAKA" in acik or "⊏" in acik or "⊂" in acik: op="Kesim + Kaynak"
        elif "GASKET" in acik: op="Montaj"
        elif any(x in acik for x in ["CİVATA","SOMUN","RONDELA","TAPA"]): op="Montaj"
        else: op="Kaynak"
        data(ws2,row,[1,2,3,4,5,6,7],[poz,adet,acik,malz,birim,top,op],alt=(i%2==1))
        row+=1
    ws2.merge_cells(start_row=row,start_column=1,end_row=row,end_column=5)
    cl=ws2.cell(row=row,column=1,value=f"TOPLAM AĞIRLIK: {toplam} Kg")
    cl.fill=H1;cl.font=fw;cl.alignment=RA;cl.border=TB
    cell(ws2,row,6,toplam,H1,fw,CA,TB)
    cell(ws2,row,7,"",H1,fw,CA,TB)
    row+=2

sw(ws2,[8,8,44,20,14,14,20])

# ═════════════════════════════════════════════════════════════════════════════
# SAYFA 3: TAVAN PLAKALARI
# ═════════════════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("Tavan Plakaları"); ws3.sheet_view.showGridLines=False

tavan_rows = [
    ("1",   4, "⊏ 1500×8 ......... 6000",  "S275J2", 565.2, 1695.6),
    ("1A",  1, "⊏ 1500×8 ......... 6000",  "S275J2", 560.0,  560.0),
    ("1B",  1, "⊏ 1500×8 ......... 6000",  "S275J2", 560.0,  560.0),
    ("1C",  1, "⊏ 1500×8 ......... 6000",  "S275J2", 560.0,  560.0),
    ("2",   2, "⊏ 1500×8 ......... 3560.5","S275J2", 334.0,  668.0),
    ("3",   4, "⊏ 1500×8 ......... 548",   "S275J2",  38.3,  153.2),
    ("4",   4, "⊏ 1500×8 ......... 3195.5","S275J2", 269.3, 1077.2),
    ("5",   2, "⊏ 1500×8 ......... 2873",  "S275J2", 257.3,  514.6),
    ("5A",  1, "⊏ 1500×8 ......... 2873",  "S275J2", 254.5,  254.5),
    ("5B",  1, "⊏ 1500×8 ......... 2873",  "S275J2", 254.5,  254.5),
    ("6",   3, "⊏ 1500×8 ......... 2545.5","S275J2", 208.0,  624.0),
    ("6A",  1, "⊏ 1500×8 ......... 2545.5","S275J2", 190.0,  190.0),
    ("7",   4, "⊏ 1000×8 ......... 1824.5","S275J2",  90.7,  362.8),
    ("8",   4, "⊏ 1056.5×8 ...... 874.5", "S275J2",  31.0,  124.0),
    ("9",   2, "⊏ 1500×8 ......... 2910.5","S275J2", 272.8,  545.6),
]

row=1
title(ws3,row,1,9,"7000 Bbl ÜRETİM TANKI — TAVAN PLAKALARI MALZEME LİSTESİ  (Çizim: 456-1200-12-MKD-208/209)")
row+=1
hdr(ws3,row,[1,2,3,4,5,6,7,8,9],
    ["POZ","ADET","AÇIKLAMA  (En×Kalınlık ......... Boy) mm","MALZEME",
     "BİRİM\nAĞIRLIK\n(Kg)","TOPLAM\nAĞIRLIK\n(Kg)",
     "KALINLIKmm","OP. TÜRÜ","NOT"],H2,fw)
ws3.row_dimensions[row].height=40; row+=1
for i,(poz,adet,acik,malz,birim,top) in enumerate(tavan_rows):
    thk=8 if "1056.5" not in acik else 8
    data(ws3,row,[1,2,3,4,5,6,7,8,9],
         [poz,adet,acik,malz,birim,top,thk,"Kesim + Silindirik Bükme","Konik tavan sacı"],
         alt=(i%2==1))
    row+=1
ws3.merge_cells(start_row=row,start_column=1,end_row=row,end_column=5)
cl=ws3.cell(row=row,column=1,value="TOPLAM AĞIRLIK:")
cl.fill=H1;cl.font=fw;cl.alignment=RA;cl.border=TB
cell(ws3,row,6,8147.0,H1,fw,CA,TB)
for c in [7,8,9]: cell(ws3,row,c,"",H1,fw,CA,TB)
row+=2

# Notlar
sub_title(ws3,row,1,9,"GENEL NOTLAR (Tavan Plakaları)")
row+=1
notlar=[
    "1- Tüm ölçüler 'mm', kotlar 'm' birimindedir.",
    "2- İmalat toleransları API 650, Bölüm 7.5'e uygun olacaktır.",
    "3- Resimde verilen ölçüler büküm öncesi (açılmış plaka) ölçüleridir.",
    "4- Kaynak ağzı açıları için tolerans ±5° dir.",
    "5- Kesim plaka uzunlukları ilgili olduğu konum sac kalınlığının ortasına göre hesaplanmıştır.",
    "6- Tavan plakası birleşimleri kaynak detayına göre yapılacaktır.",
    "7- Tank iç ve dış yüzeyinin boyanmasından önce tüm kaynak çapakalrı temizlenecektir.",
]
for n in notlar:
    ws3.merge_cells(start_row=row,start_column=1,end_row=row,end_column=9)
    cl=ws3.cell(row=row,column=1,value=n)
    cl.fill=WH;cl.font=fn9;cl.alignment=LA;cl.border=TB
    row+=1

sw(ws3,[8,8,44,14,12,12,12,24,20])

# ═════════════════════════════════════════════════════════════════════════════
# SAYFA 4: TABAN + GÖVDE (placeholder - kullanıcıdan beklenecek)
# ═════════════════════════════════════════════════════════════════════════════
ws4 = wb.create_sheet("Taban+Gövde Plakaları"); ws4.sheet_view.showGridLines=False
title(ws4,1,1,7,"TABAN ve GÖVDE PLAKALARI — Detay veriler bekleniyor")
hdr(ws4,2,[1,2,3,4,5,6,7],["POZ","ADET","AÇIKLAMA","MALZEME","BİRİM AĞIRLIK (Kg)","TOPLAM AĞIRLIK (Kg)","OP. TÜRÜ"])
# Toplam satırları (bilinen)
sub_title(ws4,3,1,7,"TABAN PLAKASI  (Çizim: KEY199MECEQP2003)")
ws4.merge_cells(start_row=4,start_column=1,end_row=4,end_column=5)
cl=ws4.cell(row=4,column=1,value="TOPLAM AĞIRLIK (detay tablo bekleniyor):")
cl.fill=YEL;cl.font=fb;cl.alignment=RA;cl.border=TB
cell(ws4,4,6,11176.7,YEL,fb,CA,TB)
cell(ws4,4,7,"Detay eklenecek",YEL,fb,CA,TB)
sub_title(ws4,5,1,7,"DRENAJ ÇUKURU  (Taban dahili)")
ws4.merge_cells(start_row=6,start_column=1,end_row=6,end_column=5)
cl=ws4.cell(row=6,column=1,value="TOPLAM AĞIRLIK (detay tablo bekleniyor):")
cl.fill=YEL;cl.font=fb;cl.alignment=RA;cl.border=TB
cell(ws4,6,6,76.1,YEL,fb,CA,TB)
cell(ws4,6,7,"Detay eklenecek",YEL,fb,CA,TB)
sub_title(ws4,7,1,7,"GÖVDE PLAKALARI  (Çizim: KEY199MECEQP2001)")
ws4.merge_cells(start_row=8,start_column=1,end_row=8,end_column=5)
cl=ws4.cell(row=8,column=1,value="TOPLAM AĞIRLIK (detay tablo bekleniyor):")
cl.fill=YEL;cl.font=fb;cl.alignment=RA;cl.border=TB
cell(ws4,8,6,16513.0,YEL,fb,CA,TB)
cell(ws4,8,7,"Detay eklenecek",YEL,fb,CA,TB)
sub_title(ws4,9,1,7,"ANKRAJ DETAYI")
ws4.merge_cells(start_row=10,start_column=1,end_row=10,end_column=5)
cl=ws4.cell(row=10,column=1,value="TOPLAM AĞIRLIK (detay tablo bekleniyor):")
cl.fill=YEL;cl.font=fb;cl.alignment=RA;cl.border=TB
cell(ws4,10,6,13.5,YEL,fb,CA,TB)
cell(ws4,10,7,"Detay eklenecek",YEL,fb,CA,TB)
sub_title(ws4,11,1,7,"İSİM PLAKASI")
ws4.merge_cells(start_row=12,start_column=1,end_row=12,end_column=5)
cl=ws4.cell(row=12,column=1,value="TOPLAM AĞIRLIK (detay tablo bekleniyor):")
cl.fill=YEL;cl.font=fb;cl.alignment=RA;cl.border=TB
cell(ws4,12,6,5.18,YEL,fb,CA,TB)
cell(ws4,12,7,"Detay eklenecek",YEL,fb,CA,TB)
sw(ws4,[8,8,44,16,22,22,20])

# ═════════════════════════════════════════════════════════════════════════════
# SAYFA 5: KONİK ÇATI YAPI İSKELETİ (Sayfa 10-11 verileri)
# ═════════════════════════════════════════════════════════════════════════════
ws5 = wb.create_sheet("Konik Çatı İskeleti"); ws5.sheet_view.showGridLines=False
title(ws5,1,1,8,"KONİK ÇATI YAPISAL İSKELET MALZEME LİSTESİ  (Çizim: 456-1200-12-MKG-207)")

sub_title(ws5,2,1,8,"CK1 MONTAJ GRUBU  (12 Adet)  — Rafter UNP240")
hdr(ws5,3,[1,2,3,4,5,6,7,8],
    ["POZ","ADET\n(1 grp)","AÇIKLAMA","MALZEME","BOY\n(mm)","BİRİM\nAĞIRLIK(kg)","TOPLAM\n(×12 grp) kg","OP. TÜRÜ"])
ws5.row_dimensions[3].height=35
ck1_items=[
    ("P1",1,"UNP240","S235JR",5685,188.8,12*188.8),
    ("P20",2,"PL10×113  PLAKA","S235JR",221,1.8,12*3.6),
    ("P21",2,"PL10×98   PLAKA","S235JR",182,1.4,12*2.8),
    ("P22",1,"PL10×221  PLAKA","S235JR",193,2.7,12*2.7),
    ("P23",1,"PL10×162  PLAKA","S235JR",216,1.9,12*1.9),
]
r=4
for i,(p,a,ac,m,b,birim,top) in enumerate(ck1_items):
    data(ws5,r,[1,2,3,4,5,6,7,8],[p,a,ac,m,b,birim,round(top,1),"Kesim+Kaynak"],alt=(i%2==1)); r+=1
total_row(ws5,r,1,7,"CK1 TOPLAM (12 grup × 1 adet/grup):",7,
          round(sum(x[6] for x in ck1_items),1)); r+=2

sub_title(ws5,r,1,8,"CK2 MONTAJ GRUBU  (12 Adet)  — Rafter UNP240"); r+=1
hdr(ws5,r,[1,2,3,4,5,6,7,8],
    ["POZ","ADET\n(1 grp)","AÇIKLAMA","MALZEME","BOY\n(mm)","BİRİM\nAĞIRLIK(kg)","TOPLAM\n(×12 grp) kg","OP. TÜRÜ"])
ws5.row_dimensions[r].height=35; r+=1
ck2_items=[
    ("P1",1,"UNP240","S235JR",5685,188.8,12*188.8),
    ("P20",2,"PL10×113  PLAKA","S235JR",221,1.8,12*3.6),
    ("P21",2,"PL10×98   PLAKA","S235JR",182,1.4,12*2.8),
    ("P24",1,"PL10×189  PLAKA","S235JR",232,2.8,12*2.8),
    ("P25",1,"PL10×139  PLAKA","S235JR",244,2.0,12*2.0),
]
for i,(p,a,ac,m,b,birim,top) in enumerate(ck2_items):
    data(ws5,r,[1,2,3,4,5,6,7,8],[p,a,ac,m,b,birim,round(top,1),"Kesim+Kaynak"],alt=(i%2==1)); r+=1
total_row(ws5,r,1,7,"CK2 TOPLAM (12 grup × 1 adet/grup):",7,
          round(sum(x[6] for x in ck2_items),1)); r+=2

sub_title(ws5,r,1,8,"KL1 MONTAJ GRUBU  (1 Adet)  — Merkezi Kolon PIP 323.9×8"); r+=1
hdr(ws5,r,[1,2,3,4,5,6,7,8],
    ["POZ","ADET","AÇIKLAMA","MALZEME","BOY\n(mm)","BİRİM\nAĞIRLIK(kg)","TOPLAM\nAĞIRLIK(kg)","OP. TÜRÜ"])
ws5.row_dimensions[r].height=35; r+=1
kl1_items=[
    ("P2",1,"PIP 323.9×8  BORU","S235JR",9369,580.2,580.2),
    ("P9",1,"PL25×1800  PLAKA","S235JR",1800,497.3,497.3),
    ("P10",1,"PL15×250  PLAKA","S235JR",4618,137.1,137.1),
]
for i,(p,a,ac,m,b,birim,top) in enumerate(kl1_items):
    data(ws5,r,[1,2,3,4,5,6,7,8],[p,a,ac,m,b,birim,top,"Kesim+Kaynak"],alt=(i%2==1)); r+=1
total_row(ws5,r,1,7,"KL1 TOPLAM (1 adet):",7,round(sum(x[6] for x in kl1_items),1)); r+=2

sub_title(ws5,r,1,8,"GENEL PLANTAN DİĞER YAPISAL ELEMANLAR  (Çizim sayfa 10)"); r+=1
hdr(ws5,r,[1,2,3,4,5,6,7,8],
    ["POZ","ADET\n(TOPLAM)","AÇIKLAMA","MALZEME","BOY\n(mm)","BİRİM\nAĞIRLIK(kg)","TOPLAM\nAĞIRLIK(kg)","OP. TÜRÜ"])
ws5.row_dimensions[r].height=35; r+=1
extra=[
    ("P3",24,"PL10×300  PLAKA","S235JR",300,0.3,7.1,"Kesim+Kaynak"),
    ("P4",1,"PL15×750  PLAKA","S235JR",750,51.6,51.6,"Kesim+Kaynak"),
    ("P5",24,"UNP200","S235JR",1227,31.0,744.0,"Kesim+Kaynak"),
    ("P6",24,"UNP200","S235JR",677,17.1,410.4,"Kesim+Kaynak"),
    ("P7",12,"L80×8  KÖŞEBENTi","S235JR",1931,18.6,223.2,"Kesim+Kaynak"),
    ("P8",12,"L80×8  KÖŞEBENTi","S235JR",1984,19.2,230.4,"Kesim+Kaynak"),
    ("P11",1,"PL10×200  PLAKA","S235JR",5341,84.2,84.2,"Kesim+Kaynak"),
    ("P12",1,"PL25×750  PLAKA","S235JR",750,86.0,86.0,"Kesim+Kaynak"),
    ("P13",24,"PL10×200  PLAKA","S235JR",250,0.4,9.6,"Kesim"),
    ("P14",24,"PL10×200  PLAKA","S235JR",100,0.2,4.8,"Kesim"),
    ("P15",24,"PL10×248  PLAKA","S235JR",299,0.6,14.4,"Kesim"),
    ("P16",24,"PL10×150  PLAKA","S235JR",250,0.4,9.6,"Kesim"),
    ("P17",4,"PL10×513  PLAKA","S235JR",675,5.2,20.8,"Kesim+Kaynak"),
    ("P18",1,"PL20×600  PLAKA","S235JR",600,43.9,43.9,"Kesim+Kaynak"),
    ("P19",1,"PL15×100  PLAKA","S235JR",1916,23.0,23.0,"Kesim"),
    ("P20",48,"PL10×113  PLAKA","S235JR",221,0.2,9.6,"Kesim"),
    ("P21",48,"PL10×98   PLAKA","S235JR",182,0.1,4.8,"Kesim"),
    ("P22",12,"PL10×221  PLAKA","S235JR",193,0.3,3.6,"Kesim"),
    ("P23",12,"PL10×162  PLAKA","S235JR",216,0.2,2.4,"Kesim"),
    ("P24",12,"PL10×189  PLAKA","S235JR",232,0.3,3.6,"Kesim"),
    ("P25",12,"PL10×139  PLAKA","S235JR",244,0.2,2.4,"Kesim"),
]
for i,row_data in enumerate(extra):
    data(ws5,r,list(range(1,9)),list(row_data),alt=(i%2==1)); r+=1

sw(ws5,[8,10,28,14,10,14,14,20])

# ═════════════════════════════════════════════════════════════════════════════
# SAYFA 6: MALİYET ÖZET
# ═════════════════════════════════════════════════════════════════════════════
ws6 = wb.create_sheet("MALİYET ÖZET"); ws6.sheet_view.showGridLines=False
title(ws6,1,1,9,"7000 Bbl ÜRETİM TANKI — MALİYET TAHMİN ÇİZELGESİ  (1 TANK İÇİN)")

# Başlık
hdr(ws6,2,[1,2,3,4,5,6,7,8,9],
    ["BÖLÜM","AÇIKLAMA","MALZEME\nSINIFI","AĞIRLIK\n(kg)","MALZEMEBİRİM\nFİYAT ($/kg)","MALZEME\nMALİYETİ ($)","İŞÇİLİK\nFİYAT ($/kg)","İŞÇİLİK\nMALİYETİ ($)","TOPLAM ($)"],
    H1,fw)
ws6.row_dimensions[2].height=40

sections=[
    # (bölüm, açıklama, malzeme sınıfı, ağırlık kg)
    ("GÖVDE","Gövde Silindirik Sacları (Shell Plates)","S275J2",16513.0),
    ("TABAN","Taban Plakaları (Bottom Plates)","S275J2",11176.7),
    ("TABAN","Drenaj Çukuru","S275J2",76.1),
    ("TAVAN","Konik Tavan Sacları (Roof Plates)","S275J2",8147.0),
    ("TAVAN","Konik Çatı Yapısal İskelet (UNP+PL+PIP)","S235JR",None),  # calculated
    ("NOZUL","Nozul N1 (12\") × 1 adet","ASTM A106B/A105",68.5),
    ("NOZUL","Nozul N3 (6\") × 2 adet","ASTM A106B/A105",round(62.0*2,1)),
    ("NOZUL","Nozul N8 (8\") × 1 adet","ASTM A106B/A105",31.9),
    ("NOZUL","Nozul N9 (4\") × 1 adet","ASTM A106B/A105",14.7),
    ("NOZUL","Nozul N9 (1½\") × 1 adet","ASTM A106B/A105",2.65),
    ("NOZUL","Nozul N18 (1\") × 1 adet","ASTM A106B/A105",1.7),
    ("NOZUL","Nozul N21 (12\") × 4 adet","ASTM A106B/A105",round(61.1*4,1)),
    ("MANHOL","24\" Gövde Manholu M1 × 1 adet","S275J2/A105",180.0),
    ("MANHOL","24\" Tavan Manholu M2 × 1 adet","S275J2",132.5),
    ("ANKRAJ","Ankraj Detayı","S275J2",13.5),
    ("İSİM","İsim Plakası","–",5.18),
    ("DİĞER","Merdiven + Korkuluk (detay bekleniyor)","S235JR",None),
    ("DİĞER","Samandıra + Diğer (detay bekleniyor)","–",None),
]

r=3
for i,(bolum,acik,malz,agirlik) in enumerate(sections):
    fill = ALT if i%2==1 else WH
    ag_str = agirlik if agirlik else "—"
    for c,v in zip([1,2,3,4],[bolum,acik,malz,ag_str]):
        cl=ws6.cell(row=r,column=c,value=v)
        cl.fill=fill;cl.font=fn;cl.alignment=CA if c!=2 else LA;cl.border=TB
    for c in [5,6,7,8,9]:
        cl=ws6.cell(row=r,column=c,value="")
        cl.fill=YEL if c in [5,7] else WH
        cl.font=fn;cl.alignment=CA;cl.border=TB
        if c==6 and agirlik:
            cl.value=f"=D{r}*E{r}"
        elif c==8 and agirlik:
            cl.value=f"=D{r}*G{r}"
        elif c==9 and agirlik:
            cl.value=f"=F{r}+H{r}"
    r+=1

# BOYA & YÜZEYİ İŞLEM
ws6.merge_cells(start_row=r,start_column=1,end_row=r,end_column=9)
cl=ws6.cell(row=r,column=1,value="══════ BOYA VE YÜZEY İŞLEM MALİYETİ (m² × birim fiyat) ══════")
cl.fill=H3;cl.font=fb;cl.alignment=CA;cl.border=TB; r+=1

boya_rows=[
    ("İç Yüzey (Gövde+Taban)","m²","~520 m²",None,None),
    ("Dış Yüzey (Gövde+Tavan)","m²","~525 m²",None,None),
    ("İç Tavan Yüzeyi","m²","~130 m²",None,None),
]
hdr(ws6,r,[1,2,3,4,5,6,7,8,9],
    ["ALAN","BİRİM","TAHMİNİ MİKTAR","KAT SAYISI","İÇ BOYA $/m²","İÇ BOYA $","DIŞ BOYA $/m²","DIŞ BOYA $","TOPLAM $"],
    H2,fw); ws6.row_dimensions[r].height=30; r+=1
for i,(alan,birim,mik,_,__) in enumerate(boya_rows):
    for c,v in zip([1,2,3,4,5,6,7,8,9],[alan,birim,mik,2,"","","","",""]):
        cl=ws6.cell(row=r,column=c,value=v)
        cl.fill=ALT if i%2==1 else WH
        cl.font=fn;cl.alignment=CA;cl.border=TB
        if c in [5,7]: cl.fill=YEL
    r+=1

# KAYNAK MALİYETİ
ws6.merge_cells(start_row=r,start_column=1,end_row=r,end_column=9)
cl=ws6.cell(row=r,column=1,value="══════ KAYNAK MALİYETİ ══════")
cl.fill=H3;cl.font=fb;cl.alignment=CA;cl.border=TB; r+=1
kaynak=[
    ("Dikey İç Dikişler (Gövde Shell)","ml","~220 ml"),
    ("Yatay Çevre Dikişleri (Gövde)","ml","~160 ml"),
    ("Taban Plaka Dikişleri","ml","~280 ml"),
    ("Tavan Plaka Dikişleri","ml","~130 ml"),
    ("Nozul Kaynakları","adet","~12 nozul"),
    ("Manhol Kaynakları","adet","~2 manhol"),
]
hdr(ws6,r,[1,2,3,4,5,6,7,8,9],
    ["KAYNAK TÜRÜ","BİRİM","MİKTAR","KAYNAK KALİTESİ","BİRİM FİYAT ($/ml)","TUTAR ($)","","",""],
    H2,fw); ws6.row_dimensions[r].height=30; r+=1
for i,(tur,birim,mik) in enumerate(kaynak):
    for c,v in zip([1,2,3,4,5,6],[tur,birim,mik,"SAW/SMAW","",""]):
        cl=ws6.cell(row=r,column=c,value=v)
        cl.fill=ALT if i%2==1 else WH;cl.font=fn;cl.alignment=CA;cl.border=TB
        if c==5: cl.fill=YEL
    for c in [7,8,9]:
        cl=ws6.cell(row=r,column=c,value="");cl.fill=WH;cl.font=fn;cl.border=TB
    r+=1

sw(ws6,[14,42,16,12,14,14,14,14,14])
for r2 in range(2,50):
    ws6.row_dimensions[r2].height=20

# ═════════════════════════════════════════════════════════════════════════════
# SAYFA 7: TOPLAM ÖZET
# ═════════════════════════════════════════════════════════════════════════════
ws7 = wb.create_sheet("Toplam Ağırlık Özeti"); ws7.sheet_view.showGridLines=False
title(ws7,1,1,5,"7000 Bbl ÜRETİM TANKI — TOPLAM AĞIRLIK ÖZETİ  (1 TANK İÇİN)")
hdr(ws7,2,[1,2,3,4,5],["BÖLÜM","AÇIKLAMA","AĞIRLIK (kg)","DURUM","KAYNAK ÇİZİM"])
ozet=[
    ("GÖVDE","Silindirik Gövde Sacları",16513.0,"Onaylı","KEY199MECEQP2001"),
    ("TABAN","Taban Plakaları",11176.7,"Onaylı","KEY199MECEQP2003"),
    ("TABAN","Drenaj Çukuru (Pit)",76.1,"Onaylı","KEY199MECEQP2003"),
    ("TAVAN","Konik Tavan Sacları",8147.0,"Onaylı","KEY199MECEQP2002 (MKD-208/209)"),
    ("TAVAN İSKELET","Konik Çatı Yapısal (UNP+Boru+PL)","Hesaplanıyor","Kontrol gerekli","456-1200-12-MKG-207"),
    ("NOZUL N1","12\" Giriş/Çıkış Nozulu × 1",68.5,"Onaylı","MKD-205"),
    ("NOZUL N3","6\" Nozulu × 2 adet",62.0*2,"Onaylı","MKD-205"),
    ("NOZUL N8","8\" Nozulu × 1",31.9,"Onaylı","MKD-205"),
    ("NOZUL N9","4\" Nozulu × 1",14.7,"Onaylı","MKD-205"),
    ("NOZUL N9","1.5\" Nozulu × 1",2.65,"Onaylı","MKD-205"),
    ("NOZUL N18","1\" Nozulu × 1",1.7,"Onaylı","MKD-205"),
    ("NOZUL N21","12\" Nozulu × 4 adet",61.1*4,"Onaylı","MKD-205"),
    ("MANHOL M1","24\" Gövde Manholu × 1",180.0,"Onaylı","MKD-204"),
    ("MANHOL M2","24\" Tavan Manholu × 1",132.5,"Onaylı","MKD-204"),
    ("ANKRAJ","Ankraj Detayı",13.5,"Onaylı","KEY199MECEQP2001"),
    ("İSİM PLAKASI","İsim Plakası",5.18,"Onaylı","KEY199MECEQP2001"),
    ("MERDİVEN","Merdiven + Korkuluk (bekleniyor)","—","Bekleniyor","MKD-210"),
    ("SAMANDIRA","Samandıra + Diğer (bekleniyor)","—","Bekleniyor","MKD-206"),
    ("TEMİZLEME","24\" Temizleme Nozulu",367.0,"Onaylı","MKD-211"),
]
known_total = sum(x[2] for x in ozet if isinstance(x[2],(int,float)))

for i,row_d in enumerate(ozet):
    fill = ALT if i%2==1 else WH
    st_fill = GRN if row_d[3]=="Onaylı" else YEL
    for c,v in enumerate(row_d,1):
        cl=ws7.cell(row=3+i,column=c,value=v)
        cl.fill=st_fill if c==4 else fill
        cl.font=fn;cl.alignment=CA if c!=2 else LA;cl.border=TB

tr=3+len(ozet)
ws7.merge_cells(start_row=tr,start_column=1,end_row=tr,end_column=2)
cl=ws7.cell(row=tr,column=1,value="BİLİNEN TOPLAM AĞIRLIK (1 TANK):")
cl.fill=H1;cl.font=fw;cl.alignment=RA;cl.border=TB
cell(ws7,tr,3,round(known_total,1),H1,fw,CA,TB)
cell(ws7,tr,4,"Kısmi (Eksik kalemler dahil değil)",H1,fw,CA,TB)
cell(ws7,tr,5,"",H1,fw,CA,TB)

sw(ws7,[18,42,18,18,28])

out="/home/user/key_cloud/7000Bbl_Tank_Malzeme_ve_Maliyet.xlsx"
wb.save(out)
print(f"Kaydedildi: {out}")
