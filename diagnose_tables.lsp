;;; diagnose_tables.lsp
;;; Çizimdeki entity türlerini ve tablo benzeri nesneleri listeler.
;;; Kullanım: (diagnose-tables)

(defun diagnose-tables (/ ss i ent ed etype counts pair found-types)
  (vl-load-com)
  (princ "\n=== TÜM ENTITY TÜRLERİ (Model Space) ===")

  ;; Tüm entity'leri say
  (setq ss (ssget "_X"))
  (setq counts '())
  (if ss
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq ent (ssname ss i))
        (setq etype (cdr (assoc 0 (entget ent))))
        (setq pair (assoc etype counts))
        (if pair
          (setq counts (subst (cons etype (1+ (cdr pair))) pair counts))
          (setq counts (cons (cons etype 1) counts))
        )
        (setq i (1+ i))
      )
      (foreach p (vl-sort counts (function (lambda (a b) (> (cdr a) (cdr b)))))
        (princ (strcat "\n  " (car p) " : " (itoa (cdr p))))
      )
    )
    (princ "\n  Model space bos!")
  )

  ;; ACAD_TABLE ara
  (princ "\n\n=== ACAD_TABLE SORGUSU ===")
  (setq ss (ssget "_X" '((0 . "ACAD_TABLE"))))
  (if ss
    (princ (strcat "\n  Bulundu: " (itoa (sslength ss)) " adet"))
    (princ "\n  ACAD_TABLE BULUNAMADI!")
  )

  ;; INSERT (blok referanslari) - tablo benzeri olabilir
  (princ "\n\n=== INSERT (BLOK) İSİMLERİ ===")
  (setq ss (ssget "_X" '((0 . "INSERT"))))
  (if ss
    (progn
      (setq i 0)
      (setq found-types '())
      (while (< i (sslength ss))
        (setq ent (ssname ss i))
        (setq etype (cdr (assoc 2 (entget ent))))  ; blok adı
        (if (not (member etype found-types))
          (progn
            (setq found-types (cons etype found-types))
            (princ (strcat "\n  BLOK: " etype))
          )
        )
        (setq i (1+ i))
      )
    )
    (princ "\n  INSERT bulunamadi.")
  )

  ;; MTEXT ve TEXT içinde SUPPORT LIST ara
  (princ "\n\n=== 'SUPPORT LIST' İÇEREN METİNLER ===")
  (setq ss (ssget "_X" '((0 . "MTEXT,TEXT"))))
  (setq found-types '())
  (if ss
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq ent (ssname ss i))
        (setq ed (entget ent))
        (setq etype (cdr (assoc 1 ed)))
        (if (null etype) (setq etype (cdr (assoc 3 ed))))
        (if (and etype (wcmatch (strcase etype) "*SUPPORT LIST*"))
          (progn
            (princ (strcat "\n  Bulundu! Tür: " (cdr (assoc 0 ed))))
            (princ (strcat " | Metin: " etype))
            (princ (strcat " | Handle: " (cdr (assoc 5 ed))))
            (setq found-types (cons ent found-types))
          )
        )
        (setq i (1+ i))
      )
    )
  )
  (if (null found-types)
    (princ "\n  'SUPPORT LIST' metni bulunamadi!")
  )

  (princ "\n\n=== TEŞHİS TAMAMLANDI ===")
  (princ)
)

(princ "\nYüklendi. Calistirmak icin: (diagnose-tables)")
(princ)
