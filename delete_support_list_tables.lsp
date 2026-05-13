;;; delete_support_list_tables.lsp
;;; SUPPORT LIST baslıklı ACAD_TABLE entity'lerini siler.
;;; Kullanım: (delete-support-list-tables)

(defun c:DELETESUPPORTLIST ( / ) (delete-support-list-tables))

(defun table-first-row-text (ent / obj ncols col txt result)
  "Tablo entity'sinin 0. satırındaki tüm hücre metinlerini birleştirir."
  (setq result "")
  (setq obj (vlax-ename->vla-object ent))
  (vl-catch-all-apply
    (function (lambda ()
      (setq ncols (vla-get-Columns obj))
      (setq col 0)
      (while (< col ncols)
        (setq txt
          (vl-catch-all-apply
            (function (lambda () (vla-GetText obj 0 col)))
          )
        )
        (if (and (not (vl-catch-all-error-p txt)) (stringp txt))
          (setq result (strcat result txt " "))
        )
        (setq col (1+ col))
      )
    ))
  )
  (strcase (vl-string-trim " " result))
)

(defun delete-support-list-tables ( / ent ed etype firstrow deleted total)
  (vl-load-com)
  (setq deleted 0)
  (setq total 0)

  ;; entnext ile TÜM entity'leri tara (ssget filtresi atlatılır)
  (setq ent (entnext))
  (while ent
    (setq ed (entget ent))
    (setq etype (cdr (assoc 0 ed)))

    (if (= etype "ACAD_TABLE")
      (progn
        (setq total (1+ total))
        (setq firstrow (table-first-row-text ent))
        (princ (strcat "\n  Tablo bulundu | İlk satır: '" firstrow "' | Handle: " (cdr (assoc 5 ed))))

        (if (wcmatch firstrow "*SUPPORT LIST*")
          (progn
            (princ " --> SİLİNİYOR")
            (entdel ent)
            (setq deleted (1+ deleted))
          )
          (princ " --> atlandı")
        )
      )
    )
    (setq ent (entnext ent))
  )

  (princ (strcat "\n\nTarama tamamlandı. Toplam tablo: " (itoa total) " | Silinen: " (itoa deleted)))

  (if (> deleted 0)
    (progn
      (command "_.QSAVE")
      (princ "\nDosya kaydedildi.")
    )
  )
  (princ)
)

(princ "\nYüklendi.")
(princ "\nKomut: DELETESUPPORTLIST  veya  (delete-support-list-tables)")
(princ)
