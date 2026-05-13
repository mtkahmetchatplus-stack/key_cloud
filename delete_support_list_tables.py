"""
DWG dosyalarındaki "SUPPORT LIST" başlıklı tabloları siler.
Kullanım: python delete_support_list_tables.py
"""

import os
import sys
import shutil
from pathlib import Path

try:
    import ezdxf
    from ezdxf.entities import Table
except ImportError:
    print("Hata: ezdxf kurulu degil. Lutfen: pip install ezdxf")
    sys.exit(1)

TARGET_DIR = r"P:\1-KEY-199_TPAO_SIRNAK_MIG\01_KEY\8-MTO\WORKING\KEY199-PIP-MTO-0001 YERÜSTÜ BORULAMA MALZEME LİSTESİ\NATIVE_EXISTING"
SEARCH_TEXT = "SUPPORT LIST"
BACKUP_SUFFIX = ".bak"


def get_cell_text(table_entity, row: int, col: int) -> str:
    """TABLE entity'nin belirtilen hücresindeki metni döndürür."""
    try:
        cell = table_entity.get_cell(row, col)
        if cell is None:
            return ""
        # Hücre içeriği metin bloğu referansı olabilir
        if hasattr(cell, "content") and cell.content:
            return str(cell.content).strip()
        # attribute_data üzerinden dene
        content = cell.dxf.get("value", "")
        return str(content).strip()
    except Exception:
        return ""


def get_first_row_text(table_entity) -> str:
    """Tablonun ilk satırındaki tüm hücre metinlerini birleştirir."""
    parts = []
    try:
        num_cols = table_entity.dxf.num_cols
    except Exception:
        num_cols = 10  # varsayılan

    for col in range(num_cols):
        text = get_cell_text(table_entity, 0, col)
        if text:
            parts.append(text)
    return " ".join(parts)


def has_support_list_in_first_row(table_entity) -> bool:
    """İlk satırda SUPPORT LIST geçiyor mu kontrol eder."""
    try:
        # ezdxf TABLE entity: cells metodunu dene
        cells = list(table_entity.cells())
        row0_texts = []
        for cell in cells:
            if cell.dxf.get("row_index", -1) == 0:
                val = cell.dxf.get("value", "")
                if val:
                    row0_texts.append(str(val).strip())

        combined = " ".join(row0_texts).upper()
        if SEARCH_TEXT in combined:
            return True

        # Alternatif: get_first_row_text ile dene
        first_row = get_first_row_text(table_entity).upper()
        return SEARCH_TEXT in first_row

    except Exception as e:
        # Son çare: entity'nin tüm dxf attribute'larına bak
        try:
            raw = str(table_entity.dxf.all_existing_dxf_attribs()).upper()
            return SEARCH_TEXT in raw
        except Exception:
            return False


def process_dwg(filepath: Path, dry_run: bool = False) -> int:
    """
    Tek bir DWG dosyasını işler.
    Returns: silinen tablo sayısı
    """
    try:
        doc = ezdxf.readfile(str(filepath))
    except Exception as e:
        print(f"  [HATA] Dosya okunamadi: {e}")
        return 0

    deleted = 0
    msp = doc.modelspace()

    tables_to_delete = []
    for entity in msp.query("ACAD_TABLE"):
        if has_support_list_in_first_row(entity):
            tables_to_delete.append(entity)

    if not tables_to_delete:
        print(f"  Silinecek tablo bulunamadi.")
        return 0

    for tbl in tables_to_delete:
        print(f"  -> Siliniyor: tablo handle={tbl.dxf.handle}")
        if not dry_run:
            msp.delete_entity(tbl)
        deleted += 1

    if not dry_run and deleted > 0:
        # Orijinal dosyayı yedekle
        backup = filepath.with_suffix(BACKUP_SUFFIX + filepath.suffix)
        shutil.copy2(filepath, backup)
        print(f"  Yedek olusturuldu: {backup.name}")

        # Kaydet (DXF olarak kaydet, sonra DWG'ye çevir)
        save_path = filepath.with_suffix(".dxf")
        doc.saveas(str(save_path))
        print(f"  Kaydedildi (DXF): {save_path.name}")
        print(f"  NOT: DXF formatinda kaydedildi. DWG gerekiyorsa ODA Converter kullanin.")

    return deleted


def main():
    target = Path(TARGET_DIR)

    if not target.exists():
        print(f"Dizin bulunamadi: {target}")
        sys.exit(1)

    dwg_files = list(target.glob("*.dwg")) + list(target.glob("*.DWG"))

    if not dwg_files:
        print("Dizinde DWG dosyasi bulunamadi.")
        sys.exit(0)

    print(f"Dizin: {target}")
    print(f"Bulunan DWG sayisi: {len(dwg_files)}")
    print(f"Aranacak metin (ilk satir): '{SEARCH_TEXT}'")
    print("-" * 60)

    # Once dry-run: ne silineceğini göster
    print("\n[ONIZLEME - hicbir sey silinmiyor]\n")
    total_preview = 0
    for dwg in sorted(dwg_files):
        print(f"Dosya: {dwg.name}")
        count = process_dwg(dwg, dry_run=True)
        total_preview += count

    if total_preview == 0:
        print("\nHiçbir tabloda 'SUPPORT LIST' bulunamadi. Islem yapilmadi.")
        return

    print(f"\nToplam {total_preview} tablo silinecek.")
    confirm = input("\nDevam etmek istiyor musunuz? (evet/hayir): ").strip().lower()
    if confirm not in ("evet", "e", "yes", "y"):
        print("Iptal edildi.")
        return

    print("\n[SILME ISLEMI BASLIYOR]\n")
    total_deleted = 0
    for dwg in sorted(dwg_files):
        print(f"Dosya: {dwg.name}")
        count = process_dwg(dwg, dry_run=False)
        total_deleted += count
        print()

    print(f"\nTamamlandi. Toplam {total_deleted} tablo silindi.")


if __name__ == "__main__":
    main()
