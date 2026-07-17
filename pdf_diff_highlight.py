#!/usr/bin/env python3
"""Iki PDF'i ust uste karsilastirir, farkli bolgeleri kirmizi ile isaretler,
ayni kalan bolgeleri oldugu gibi birakir ve sonucu kucuk boyutlu bir PDF olarak kaydeder.

Kullanim:
    python3 pdf_diff_highlight.py dosya1.pdf dosya2.pdf cikti.pdf
    python3 pdf_diff_highlight.py dosya1.pdf dosya2.pdf cikti.pdf --dpi 150 --threshold 30 --quality 60
"""
import argparse
import io

import cv2
import fitz
import numpy as np
from PIL import Image


def render_pdf_pages(path, dpi):
    doc = fitz.open(path)
    zoom = dpi / 72
    matrix = fitz.Matrix(zoom, zoom)
    pages = []
    for page in doc:
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        if pix.n == 4:
            img = cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)
        pages.append(img)
    doc.close()
    return pages


def highlight_diff(img1, img2, threshold, min_area, dilate_px, color, alpha, box_mode):
    if img2.shape[:2] != img1.shape[:2]:
        img2 = cv2.resize(img2, (img1.shape[1], img1.shape[0]))

    gray1 = cv2.cvtColor(img1, cv2.COLOR_RGB2GRAY)
    gray2 = cv2.cvtColor(img2, cv2.COLOR_RGB2GRAY)
    diff = cv2.absdiff(gray1, gray2)
    _, mask = cv2.threshold(diff, threshold, 255, cv2.THRESH_BINARY)

    if dilate_px > 0:
        kernel = np.ones((dilate_px, dilate_px), np.uint8)
        mask = cv2.dilate(mask, kernel, iterations=1)

    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    result = img1.copy()
    color_bgr = np.array(color, dtype=np.uint8)

    for cnt in contours:
        area = cv2.contourArea(cnt)
        if area < min_area:
            continue
        x, y, w, h = cv2.boundingRect(cnt)
        if box_mode:
            cv2.rectangle(result, (x, y), (x + w, y + h), tuple(int(c) for c in color), 3)
        else:
            roi = result[y:y + h, x:x + w]
            blended = (roi.astype(np.float32) * (1 - alpha) + color_bgr.astype(np.float32) * alpha)
            result[y:y + h, x:x + w] = blended.astype(np.uint8)

    return result


def save_compressed_pdf(images, out_path, quality):
    pil_images = [Image.fromarray(img).convert("RGB") for img in images]
    if not pil_images:
        raise ValueError("Kaydedilecek sayfa yok")
    buffers = []
    for img in pil_images:
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=quality, optimize=True)
        buf.seek(0)
        buffers.append(Image.open(buf))
    buffers[0].save(out_path, format="PDF", save_all=True, append_images=buffers[1:])


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("pdf1")
    ap.add_argument("pdf2")
    ap.add_argument("output")
    ap.add_argument("--dpi", type=int, default=150, help="Render cozunurlugu (varsayilan 150)")
    ap.add_argument("--threshold", type=int, default=30, help="Fark hassasiyeti, 0-255 (varsayilan 30)")
    ap.add_argument("--min-area", type=int, default=40, help="Gurultuyu elemek icin en kucuk fark alani (piksel^2)")
    ap.add_argument("--dilate", type=int, default=5, help="Fark bolgelerini birlestirmek icin genisletme (piksel)")
    ap.add_argument("--color", default="255,0,0", help="Vurgu rengi R,G,B (varsayilan kirmizi)")
    ap.add_argument("--alpha", type=float, default=0.45, help="Dolgu modunda kirmizinin opakligi 0-1")
    ap.add_argument("--box", action="store_true", help="Dolgu yerine sadece kirmizi cerceve ciz")
    ap.add_argument("--quality", type=int, default=60, help="Cikti JPEG kalitesi, dosya boyutunu belirler (varsayilan 60)")
    args = ap.parse_args()

    color = tuple(int(c) for c in args.color.split(","))

    pages1 = render_pdf_pages(args.pdf1, args.dpi)
    pages2 = render_pdf_pages(args.pdf2, args.dpi)

    if len(pages1) != len(pages2):
        print(f"Uyari: sayfa sayilari farkli ({len(pages1)} vs {len(pages2)}), "
              f"kisa olana kadar karsilastirilacak.")

    n = min(len(pages1), len(pages2))
    results = []
    for i in range(n):
        print(f"Sayfa {i + 1}/{n} karsilastiriliyor...")
        results.append(highlight_diff(
            pages1[i], pages2[i],
            threshold=args.threshold,
            min_area=args.min_area,
            dilate_px=args.dilate,
            color=color,
            alpha=args.alpha,
            box_mode=args.box,
        ))

    for i in range(n, len(pages1)):
        results.append(pages1[i])

    save_compressed_pdf(results, args.output, args.quality)
    print(f"Kaydedildi: {args.output}")


if __name__ == "__main__":
    main()
