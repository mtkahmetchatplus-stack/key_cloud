;;; delete_support_list_tables.lsp
;;; Modeldeki tabloların 1. satırında "SUPPORT LIST" geçiyorsa o tabloyu siler.
;;; Kullanım: AutoCAD'de (load "delete_support_list_tables.lsp") ardından (delete-support-list-tables)

(defun get-table-first-row-text (ent / tblobj numcols col celltext combined)
  "Tablo entity'sinin ilk satırındaki hücre metinlerini birleştirir."
  (setq tblobj (vlax-ename->vla-object ent))
  (setq combined "")
  (vl-catch-all-apply
    (function
      (lambda ()
        (setq numcols (vla-get-columns tblobj))
        (setq col 0)
        (while (< col numcols)
          (setq celltext
            (vl-catch-all-apply
              (function (lambda () (vla-GetText tblobj 0 col)))
            )
          )
          (if (and (not (vl-catch-all-error-p celltext))
                   (stringp celltext)
                   (/= celltext ""))
            (setq combined (strcat combined " " celltext))
          )
          (setq col (1+ col))
        )
      )
    )
  )
  (strcase (vl-string-trim " " combined))
)

(defun delete-support-list-tables (/ ss i ent firstrow count fname)
  "Açık DWG dosyasındaki SUPPORT LIST tablolarını siler."
  (vl-load-com)
  (setq count 0)
  (setq ss (ssget "_X" '((0 . "ACAD_TABLE"))))
  (if (null ss)
    (progn
      (princ "\nModelde hiç tablo bulunamadı.")
      (princ)
    )
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq ent (ssname ss i))
        (setq firstrow (get-table-first-row-text ent))
        (if (wcmatch firstrow "*SUPPORT LIST*")
          (progn
            (princ (strcat "\n-> Siliniyor: handle=" (cdr (assoc 5 (entget ent)))))
            (princ (strcat "  | İlk satır: " firstrow))
            (entdel ent)
            (setq count (1+ count))
          )
        )
        (setq i (1+ i))
      )
      (princ (strcat "\n\nTamamlandı. " (itoa count) " tablo silindi."))
      (if (> count 0)
        (progn
          (command "_.QSAVE")
          (princ "\nDosya kaydedildi.")
        )
      )
      (princ)
    )
  )
)

;;; Toplu işlem: Bir klasördeki tüm DWG'leri işle
(defun delete-support-list-tables-batch (folder / files f)
  "Belirtilen klasördeki tüm DWG dosyalarını işler."
  (vl-load-com)
  (if (null folder)
    (setq folder
      "P:\\1-KEY-199_TPAO_SIRNAK_MIG\\01_KEY\\8-MTO\\WORKING\\KEY199-PIP-MTO-0001 YERÜSTÜ BORULAMA MALZEME LİSTESİ\\NATIVE_EXISTING"
    )
  )
  (setq files (vl-directory-files folder "*.dwg" 1))
  (if (null files)
    (princ (strcat "\nDizinde DWG bulunamadı: " folder))
    (foreach f files
      (progn
        (princ (strcat "\n\n=== İşleniyor: " f " ==="))
        (open-and-process-dwg (strcat folder "\\" f))
      )
    )
  )
  (princ)
)

(defun open-and-process-dwg (fullpath / )
  "Dosyayı aç, tabloları sil, kaydet, kapat."
  (vl-catch-all-apply
    (function
      (lambda ()
        (command "_.OPEN" fullpath "")
        (command "_.ZOOM" "_E")
        (delete-support-list-tables)
      )
    )
  )
)

(princ "\nYüklendi. Komutlar:")
(princ "\n  (delete-support-list-tables)        -> Aktif DWG'yi işler")
(princ "\n  (delete-support-list-tables-batch nil) -> Tüm klasörü işler")
(princ)
