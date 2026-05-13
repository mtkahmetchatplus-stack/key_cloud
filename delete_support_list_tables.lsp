;;; delete_support_list_tables.lsp
;;; SUPPORT LIST baslıklı ACAD_TABLE entity'lerini siler.
;;; Kullanım: DELETESUPPORTLIST komutu

(defun c:DELETESUPPORTLIST ( / ) (delete-support-list-tables))

(defun entity-all-strings-contain (ent searchtext / ed found)
  "entget sonucundaki TÜM string değerlerini tarar - birinde searchtext varsa T döner."
  (setq found nil)
  (setq ed (entget ent))
  (foreach pair ed
    (if (and (not found)
             (stringp (cdr pair))
             (wcmatch (strcase (cdr pair)) (strcat "*" searchtext "*")))
      (setq found T)
    )
  )
  found
)

(defun table-contains-support-list (ent / obj nrows ncols row col txt found)
  "ACAD_TABLE entity'sinde SUPPORT LIST metnini çoklu yöntemle arar."
  (setq found nil)

  ; Yöntem 1: entget DXF verilerinde tüm string'leri tara (merged cell sorununu atlatır)
  (if (entity-all-strings-contain ent "SUPPORT LIST")
    (setq found T)
  )

  ; Yöntem 2: vla-GetText ile ilk 5 satır, tüm sütunlar
  (if (not found)
    (vl-catch-all-apply
      (function (lambda ()
        (setq obj (vlax-ename->vla-object ent))
        (setq nrows (min 5 (fix (vla-get-Rows obj))))
        (setq ncols (fix (vla-get-Columns obj)))
        (setq row 0)
        (while (and (< row nrows) (not found))
          (setq col 0)
          (while (< col ncols)
            (setq txt
              (vl-catch-all-apply
                (function (lambda () (vla-GetText obj row col)))
              )
            )
            (if (and (not (vl-catch-all-error-p txt))
                     (stringp txt)
                     (wcmatch (strcase txt) "*SUPPORT LIST*"))
              (setq found T)
            )
            (setq col (1+ col))
          )
          (setq row (1+ row))
        )
      ))
    )
  )

  found
)

(defun delete-support-list-tables ( / ent ed etype deleted total)
  (vl-load-com)
  (setq deleted 0)
  (setq total 0)

  (setq ent (entnext))
  (while ent
    (setq ed (entget ent))
    (setq etype (cdr (assoc 0 ed)))

    (if (= etype "ACAD_TABLE")
      (progn
        (setq total (1+ total))
        (princ (strcat "\n  Tablo: Handle=" (cdr (assoc 5 ed))))

        (if (table-contains-support-list ent)
          (progn
            (princ " --> SUPPORT LIST bulundu, SİLİNİYOR")
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
(princ "\nKomut: DELETESUPPORTLIST")
(princ)
