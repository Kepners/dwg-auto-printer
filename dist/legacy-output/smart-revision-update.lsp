;;; smart-revision-update.lsp
;;; Intelligent paper-space revision updater.
;;; Derived from bellway-revision-update.lsp and hardened with validation checks.
;;; Commands:
;;;   None by default (EXE-driven mode only).
;;; API:
;;;   (dap:exe-run folder batch-size rev-mode exact-rev target-desc close-after)

(vl-load-com)

(setq *dap-config*
  (list
    (cons 'header-row-window 30.0)
    (cons 'max-column-distance 20.0)
    (cons 'rev-headers '("REV" "REVISION"))
    (cons 'date-headers '("DATE"))
    (cons 'desc-headers '("DESCRIPTION" "DESC"))
    (cons 'skip-text '("REV" "REVISION" "DATE" "DESCRIPTION" "DESC" "DRN BY"))
    (cons 'titleblock-rev-tags '("REV" "REVISION" "SHEET_REV" "DRAWING_REV" "SHT_REV" "SHEETREV"))
    (cons 'titleblock-bottom-y-max 80.0)
    (cons 'titleblock-right-x-min 300.0)
    (cons 'default-description "Construction Issue")
  )
)

;;; Defaults used by no-prompt automation entry points.
(setq *dap-auto-folder* "")
(setq *dap-auto-batch-size* 5)
(setq *dap-auto-rev-mode* "NEXT")
(setq *dap-auto-exact-rev* "")
(setq *dap-auto-desc* "")
(setq *dap-auto-close-after* T)
(setq *dap-exe-mode-only* T)

(defun dap:getcfg (key / pair)
  (setq pair (assoc key *dap-config*))
  (if pair (cdr pair) nil)
)

(defun dap:trim (s)
  (if s (vl-string-trim " \t\r\n" s) "")
)

(defun dap:up (s)
  (strcase (dap:trim s))
)

(defun dap:item (k it)
  (cdr (assoc k it))
)

(defun dap:all-digits-p (s / i ch ok)
  (setq s (dap:trim s))
  (if (= s "")
    nil
    (progn
      (setq ok T i 1)
      (while (and ok (<= i (strlen s)))
        (setq ch (substr s i 1))
        (if (not (wcmatch ch "#")) (setq ok nil))
        (setq i (1+ i))
      )
      ok
    )
  )
)

(defun dap:next-rev (rev / u code prefix n)
  (setq u (dap:up rev))
  (cond
    ((= u "") "A")
    ((and (= (strlen u) 1) (wcmatch u "[A-Z]"))
     (setq code (ascii u))
     (if (< code 90) (chr (1+ code)) u)
    )
    ((dap:all-digits-p u)
     (itoa (1+ (atoi u)))
    )
    ((and (> (strlen u) 1)
          (wcmatch (substr u 1 1) "[A-Z]")
          (dap:all-digits-p (substr u 2)))
     (setq prefix (substr u 1 1))
     (setq n (atoi (substr u 2)))
     (strcat prefix (itoa (1+ n)))
    )
    ((= u "-") "A")
    (T u)
  )
)

(defun dap:today-ddmmyy ()
  (menucmd "M=$(edtime,$(getvar,date),DD.MO.YY)")
)

(defun dap:to-xy (pt / p)
  (cond
    ((= (type pt) 'LIST)
     (list (car pt) (cadr pt))
    )
    (T
     (setq p (vlax-safearray->list (vlax-variant-value pt)))
     (list (car p) (cadr p))
    )
  )
)

(defun dap:make-item (obj kind text x y tag)
  (list
    (cons 'obj obj)
    (cons 'kind kind)
    (cons 'text (dap:trim text))
    (cons 'up (dap:up text))
    (cons 'x (float x))
    (cons 'y (float y))
    (cons 'tag (if tag (dap:up tag) ""))
  )
)

(defun dap:collect-layout-items (layout-name / doc layout blk items obj oname pt attrs raw att)
  (setq doc (vla-get-ActiveDocument (vlax-get-acad-object)))
  (setq layout (vla-Item (vla-get-Layouts doc) layout-name))
  (setq blk (vla-get-Block layout))
  (setq items nil)

  (vlax-for obj blk
    (setq oname (vla-get-ObjectName obj))
    (cond
      ((= oname "AcDbText")
       (setq pt (dap:to-xy (vlax-get obj 'InsertionPoint)))
       (setq items
         (cons (dap:make-item obj 'TEXT (vla-get-TextString obj) (car pt) (cadr pt) nil) items))
      )
      ((= oname "AcDbMText")
       (setq pt (dap:to-xy (vlax-get obj 'InsertionPoint)))
       (setq items
         (cons (dap:make-item obj 'MTEXT (vla-get-TextString obj) (car pt) (cadr pt) nil) items))
      )
      ((and (= oname "AcDbBlockReference")
            (= :vlax-true (vla-get-HasAttributes obj)))
       (setq raw (vlax-variant-value (vla-GetAttributes obj)))
       (setq attrs (vlax-safearray->list raw))
       (foreach att attrs
         (setq pt (dap:to-xy (vlax-get att 'InsertionPoint)))
         (setq items
           (cons (dap:make-item att 'ATTRIB (vla-get-TextString att) (car pt) (cadr pt) (vla-get-TagString att)) items))
       )
      )
    )
  )
  items
)

(defun dap:find-header (items accepted / best it)
  (setq best nil)
  (foreach it items
    (if (member (dap:item 'up it) accepted)
      (if (or (null best) (> (dap:item 'y it) (dap:item 'y best)))
        (setq best it)
      )
    )
  )
  best
)

(defun dap:skip-text-p (txt / up skip)
  (setq up (dap:up txt))
  (setq skip (dap:getcfg 'skip-text))
  (or (member up skip) (wcmatch up "*REVISION SCHEDULE*"))
)

(defun dap:find-cell-below-header (header items / hx hy rowwin maxdx best bestdx bestdy it x y dx dy)
  (if (null header)
    nil
    (progn
      (setq hx (dap:item 'x header))
      (setq hy (dap:item 'y header))
      (setq rowwin (dap:getcfg 'header-row-window))
      (setq maxdx (dap:getcfg 'max-column-distance))
      (setq best nil bestdx 1e99 bestdy 1e99)

      (foreach it items
        (setq x (dap:item 'x it))
        (setq y (dap:item 'y it))
        (if (and (< y hy)
                 (> y (- hy rowwin))
                 (/= (dap:item 'up it) "")
                 (not (dap:skip-text-p (dap:item 'text it))))
          (progn
            (setq dx (abs (- x hx)))
            (setq dy (abs (- hy y)))
            (if (and (<= dx maxdx)
                     (or (< dx bestdx)
                         (and (= dx bestdx) (< dy bestdy))))
              (progn
                (setq best it)
                (setq bestdx dx)
                (setq bestdy dy)
              )
            )
          )
        )
      )
      best
    )
  )
)

(defun dap:rev-token-p (txt / u)
  (setq u (dap:up txt))
  (or (wcmatch u "[A-Z]")
      (wcmatch u "[A-Z]#")
      (wcmatch u "[A-Z]##")
      (wcmatch u "#")
      (wcmatch u "##")
      (wcmatch u "###")
      (= u "-"))
)

(defun dap:find-titleblock-rev-items (items / tags yMax xMin out it)
  (setq tags (dap:getcfg 'titleblock-rev-tags))
  (setq yMax (dap:getcfg 'titleblock-bottom-y-max))
  (setq xMin (dap:getcfg 'titleblock-right-x-min))
  (setq out nil)

  (foreach it items
    (if (and (= (dap:item 'kind it) 'ATTRIB)
             (member (dap:item 'tag it) tags))
      (setq out (cons it out))
    )
  )

  (if out
    out
    (progn
      (foreach it items
        (if (and (or (= (dap:item 'kind it) 'TEXT) (= (dap:item 'kind it) 'MTEXT))
                 (< (dap:item 'y it) yMax)
                 (> (dap:item 'y it) 0.0)
                 (> (dap:item 'x it) xMin)
                 (dap:rev-token-p (dap:item 'text it)))
          (setq out (cons it out))
        )
      )
      out
    )
  )
)

(defun dap:set-item-text (it new-text / obj old)
  (setq obj (dap:item 'obj it))
  (setq old (dap:trim (vla-get-TextString obj)))
  (if (/= old new-text)
    (progn
      (vla-put-TextString obj new-text)
      T
    )
    nil
  )
)

(defun dap:get-paper-layout-names (/ doc layouts out lay name)
  (setq doc (vla-get-ActiveDocument (vlax-get-acad-object)))
  (setq layouts (vla-get-Layouts doc))
  (setq out nil)
  (vlax-for lay layouts
    (setq name (vla-get-Name lay))
    (if (/= (strcase name) "MODEL")
      (setq out (append out (list name)))
    )
  )
  out
)

(defun dap:detect-current-rev (layout-name / items revHeader revCell tbItems)
  (setq items (dap:collect-layout-items layout-name))
  (setq revHeader (dap:find-header items (dap:getcfg 'rev-headers)))
  (setq revCell (dap:find-cell-below-header revHeader items))
  (cond
    ((and revCell (/= (dap:item 'text revCell) ""))
     (dap:up (dap:item 'text revCell))
    )
    (T
     (setq tbItems (dap:find-titleblock-rev-items items))
     (if tbItems
       (dap:up (dap:item 'text (car tbItems)))
       ""
     )
    )
  )
)

(defun dap:process-layout (layout-name target-rev target-date target-desc dry-run / items revH dateH descH revCell dateCell descCell tbItems changes warnings changed it)
  (setq changes 0)
  (setq warnings nil)
  (setq changed nil)

  (setvar "CTAB" layout-name)
  (setq items (dap:collect-layout-items layout-name))

  (setq revH (dap:find-header items (dap:getcfg 'rev-headers)))
  (setq dateH (dap:find-header items (dap:getcfg 'date-headers)))
  (setq descH (dap:find-header items (dap:getcfg 'desc-headers)))

  (if (null revH) (setq warnings (cons "Missing REV header" warnings)))
  (if (null dateH) (setq warnings (cons "Missing DATE header" warnings)))
  (if (null descH) (setq warnings (cons "Missing DESCRIPTION header" warnings)))

  (setq revCell (dap:find-cell-below-header revH items))
  (setq dateCell (dap:find-cell-below-header dateH items))
  (setq descCell (dap:find-cell-below-header descH items))

  (if (and revCell (/= (dap:item 'text revCell) target-rev))
    (if dry-run
      (princ (strcat "\n[DRY] " layout-name ": REV " (dap:item 'text revCell) " -> " target-rev))
      (if (dap:set-item-text revCell target-rev)
        (progn
          (setq changes (1+ changes))
          (setq changed T)
        )
      )
    )
  )

  (if dateCell
    (if (or changed (/= (dap:item 'text dateCell) target-date))
      (if dry-run
        (princ (strcat "\n[DRY] " layout-name ": DATE " (dap:item 'text dateCell) " -> " target-date))
        (if (dap:set-item-text dateCell target-date) (setq changes (1+ changes)))
      )
    )
    (setq warnings (cons "No DATE value cell found" warnings))
  )

  (if (and descCell (/= (dap:trim target-desc) ""))
    (if (/= (dap:item 'text descCell) target-desc)
      (if dry-run
        (princ (strcat "\n[DRY] " layout-name ": DESC " (dap:item 'text descCell) " -> " target-desc))
        (if (dap:set-item-text descCell target-desc) (setq changes (1+ changes)))
      )
    )
    (if (null descCell) (setq warnings (cons "No DESCRIPTION value cell found" warnings)))
  )

  (setq tbItems (dap:find-titleblock-rev-items items))
  (if tbItems
    (foreach it tbItems
      (if (/= (dap:item 'text it) target-rev)
        (if dry-run
          (princ (strcat "\n[DRY] " layout-name ": TITLE REV " (dap:item 'text it) " -> " target-rev))
          (if (dap:set-item-text it target-rev)
            (progn
              (setq changes (1+ changes))
              (setq changed T)
            )
          )
        )
      )
    )
    (setq warnings (cons "No title block revision field found" warnings))
  )

  (list
    (cons 'layout layout-name)
    (cons 'changes changes)
    (cons 'warnings (reverse warnings))
  )
)

(defun dap:run (mode forcedScope / scope layouts activeLayout currentRev defaultRev inRev targetRev inDesc targetDesc targetDate doc res totalChanges totalWarnings doSave lay w)
  (vl-load-com)
  (dap:ensure-paperspace)
  (setq activeLayout (getvar "CTAB"))

  (if forcedScope
    (setq scope forcedScope)
    (progn
      (initget "Current All")
      (setq scope (getkword "\nProcess layouts [Current/All] <All>: "))
      (if (null scope) (setq scope "All"))
    )
  )

  (if (= scope "Current")
    (setq layouts (list activeLayout))
    (setq layouts (dap:get-paper-layout-names))
  )

  (if (null layouts)
    (progn
      (princ "\nNo paper space layouts found.")
      (princ)
    )
    (progn
      (setq currentRev (dap:detect-current-rev (car layouts)))
      (setq defaultRev (dap:next-rev currentRev))
      (setq inRev (getstring T (strcat "\nTarget revision <" defaultRev ">: ")))
      (if (= (dap:trim inRev) "")
        (setq targetRev defaultRev)
        (setq targetRev (dap:up inRev))
      )

      (setq inDesc (getstring T (strcat "\nDescription <" (dap:getcfg 'default-description) ">: ")))
      (if (= (dap:trim inDesc) "")
        (setq targetDesc (dap:getcfg 'default-description))
        (setq targetDesc (dap:trim inDesc))
      )

      (setq targetDate (dap:today-ddmmyy))
      (setq totalChanges 0)
      (setq totalWarnings 0)

      (princ (strcat "\n\n=== DAP revision run (" mode ") ==="))
      (princ (strcat "\nLayouts: " (itoa (length layouts))))
      (princ (strcat "\nRevision: " targetRev))
      (princ (strcat "\nDate: " targetDate))
      (princ (strcat "\nDescription: " targetDesc))
      (princ "\n-----------------------------------")

      (foreach lay layouts
        (setq res (dap:process-layout lay targetRev targetDate targetDesc (= mode "DRY")))
        (setq totalChanges (+ totalChanges (dap:item 'changes res)))
        (if (dap:item 'warnings res)
          (progn
            (setq totalWarnings (+ totalWarnings (length (dap:item 'warnings res))))
            (foreach w (dap:item 'warnings res)
              (princ (strcat "\nWARN [" lay "]: " w))
            )
          )
        )
        (princ (strcat "\nOK [" lay "]: " (itoa (dap:item 'changes res)) " changes"))
      )

      (if (/= mode "DRY")
        (progn
          (setq doc (vla-get-ActiveDocument (vlax-get-acad-object)))
          (vla-Regen doc 0)

          (initget "Yes No")
          (setq doSave (getkword "\nSave drawing now? [Yes/No] <Yes>: "))
          (if (or (null doSave) (= doSave "Yes"))
            (progn
              (command "_.QSAVE")
              (princ "\nDrawing saved.")
            )
            (princ "\nChanges kept in memory (not saved).")
          )
        )
      )

      (princ "\n-----------------------------------")
      (princ (strcat "\nDone. Total changes: " (itoa totalChanges)))
      (princ (strcat "\nTotal warnings: " (itoa totalWarnings)))
      (princ "\n")
      (princ)
    )
  )
)

(defun dap:ensure-paperspace (/ layouts)
  (if (/= (getvar "TILEMODE") 0)
    (setvar "TILEMODE" 0)
  )
  (if (= (strcase (getvar "CTAB")) "MODEL")
    (progn
      (setq layouts (dap:get-paper-layout-names))
      (if layouts
        (setvar "CTAB" (car layouts))
      )
    )
  )
)

(defun dap:parse-int-or (s fallback / v)
  (setq s (dap:trim s))
  (if (= s "")
    fallback
    (progn
      (setq v (atoi s))
      (if (> v 0) v fallback)
    )
  )
)

(defun dap:truthy-p (v / u)
  (cond
    ((or (= v T) (= v :vlax-true)) T)
    ((numberp v) (/= v 0))
    ((= (type v) 'STR)
     (progn
       (setq u (dap:up v))
       (or (= u "T")
           (= u "TRUE")
           (= u "Y")
           (= u "YES")
           (= u "1")))
    )
    (T nil)
  )
)

(defun dap:exact-mode-p (rev-mode / u)
  (setq u (dap:up rev-mode))
  (or (= u "EXACT") (= u "FIXED") (= u "SET"))
)

(defun dap:path-join (folder file / f)
  (setq f folder)
  (if (or (null f) (= (dap:trim f) ""))
    file
    (if (or (= (substr f (strlen f) 1) "\\")
            (= (substr f (strlen f) 1) "/"))
      (strcat f file)
      (strcat f "\\" file)
    )
  )
)

(defun dap:list-dwg-files-root (folder / names out n)
  (setq names (vl-directory-files folder "*.dwg" 1))
  (setq out nil)
  (foreach n names
    (setq out (append out (list (dap:path-join folder n))))
  )
  out
)

(defun dap:find-open-doc-by-fullname (docs fullpath / found d)
  (setq found nil)
  (vlax-for d docs
    (if (= (strcase (vla-get-FullName d)) (strcase fullpath))
      (setq found d)
    )
  )
  found
)

(defun dap:run-auto-current-doc (exact-rev target-desc / doc layouts currentRev targetRev targetDate totalChanges totalWarnings res lay w saveErr)
  (setq doc (vla-get-ActiveDocument (vlax-get-acad-object)))
  (dap:ensure-paperspace)
  (setq layouts (dap:get-paper-layout-names))

  (if (null layouts)
    (list
      (cons 'doc (vla-get-Name doc))
      (cons 'rev "")
      (cons 'changes 0)
      (cons 'warnings 1)
      (cons 'status "No paper space layouts found")
    )
    (progn
      (setq currentRev (dap:detect-current-rev (car layouts)))
      (if (and exact-rev (/= (dap:trim exact-rev) ""))
        (setq targetRev (dap:up exact-rev))
        (setq targetRev (dap:next-rev currentRev))
      )
      (setq targetDate (dap:today-ddmmyy))
      (setq totalChanges 0)
      (setq totalWarnings 0)

      (foreach lay layouts
        (setq res (dap:process-layout lay targetRev targetDate target-desc nil))
        (setq totalChanges (+ totalChanges (dap:item 'changes res)))
        (if (dap:item 'warnings res)
          (progn
            (setq totalWarnings (+ totalWarnings (length (dap:item 'warnings res))))
            (foreach w (dap:item 'warnings res)
              (princ (strcat "\nWARN [" (vla-get-Name doc) " | " lay "]: " w))
            )
          )
        )
      )

      (vla-Regen doc 1)
      (setq saveErr (vl-catch-all-apply 'vla-Save (list doc)))
      (if (vl-catch-all-error-p saveErr)
        (progn
          (setq totalWarnings (1+ totalWarnings))
          (princ (strcat "\nWARN [" (vla-get-Name doc) "]: save failed"))
        )
      )

      (list
        (cons 'doc (vla-get-Name doc))
        (cons 'rev targetRev)
        (cons 'changes totalChanges)
        (cons 'warnings totalWarnings)
        (cons 'status "OK")
      )
    )
  )
)

(defun dap:process-dwg-group (group exact-rev target-desc close-after / acad docs opened path existing openedObj pair doc openedByUs res closeErr)
  (setq acad (vlax-get-acad-object))
  (setq docs (vla-get-Documents acad))
  (setq opened nil)

  (foreach path group
    (setq existing (dap:find-open-doc-by-fullname docs path))
    (if existing
      (progn
        (princ (strcat "\nINFO already open: " path))
        (setq opened (append opened (list (list existing nil path))))
      )
      (progn
        (setq openedObj (vl-catch-all-apply 'vla-Open (list docs path)))
        (if (vl-catch-all-error-p openedObj)
          (princ (strcat "\nERROR open failed: " path))
          (setq opened (append opened (list (list openedObj T path))))
        )
      )
    )
  )

  (foreach pair opened
    (setq doc (nth 0 pair))
    (setq openedByUs (nth 1 pair))
    (vl-catch-all-apply 'vla-Activate (list doc))

    (princ (strcat "\n\n--- Processing: " (vla-get-FullName doc) " ---"))
    (setq res (dap:run-auto-current-doc exact-rev target-desc))
    (princ (strcat "\nResult: REV=" (dap:item 'rev res)
                   " changes=" (itoa (dap:item 'changes res))
                   " warnings=" (itoa (dap:item 'warnings res))))

    (if (and close-after openedByUs)
      (progn
        (setq closeErr (vl-catch-all-apply 'vla-Close (list doc)))
        (if (vl-catch-all-error-p closeErr)
          (princ (strcat "\nWARN close failed: " (vla-get-Name doc)))
        )
      )
    )
  )
)

(defun dap:run-batch-core (folder batchSize exactRev targetDesc closeAfter / files idx total j group)
  (if (not (vl-file-directory-p folder))
    (princ (strcat "\nFolder not found: " folder))
    (progn
      (setq files (dap:list-dwg-files-root folder))
      (if (null files)
        (princ (strcat "\nNo DWG files found in: " folder))
        (progn
          (setq batchSize (dap:parse-int-or (vl-princ-to-string batchSize) 5))
          (setq closeAfter (dap:truthy-p closeAfter))
          (if (= (dap:trim targetDesc) "")
            (setq targetDesc (dap:getcfg 'default-description))
          )
          (if (= (dap:trim exactRev) "")
            (setq exactRev nil)
            (setq exactRev (dap:up exactRev))
          )

          (setq idx 0)
          (setq total (length files))
          (princ (strcat "\n\n=== DAP batch revision run ==="))
          (princ (strcat "\nFolder: " folder))
          (princ (strcat "\nDWGs found: " (itoa total)))
          (princ (strcat "\nBatch size: " (itoa batchSize)))
          (if exactRev
            (princ (strcat "\nRevision mode: Exact (" exactRev ")"))
            (princ "\nRevision mode: Next per drawing")
          )
          (princ (strcat "\nDescription: " targetDesc))
          (princ "\n-----------------------------------")

          (while (< idx total)
            (setq group nil)
            (setq j 0)
            (while (and (< (+ idx j) total) (< j batchSize))
              (setq group (append group (list (nth (+ idx j) files))))
              (setq j (1+ j))
            )
            (princ (strcat "\n\nBatch " (itoa (+ 1 (/ idx batchSize)))
                           " (" (itoa (length group)) " drawings)"))
            (dap:process-dwg-group group exactRev targetDesc closeAfter)
            (setq idx (+ idx j))
          )

          (princ "\n-----------------------------------")
          (princ "\nBatch run completed.")
        )
      )
    )
  )
)

(defun dap:run-batch-folder (/ inFolder folder batchIn batchSize revMode exactRev inDesc targetDesc closeAns closeAfter)
  (setq inFolder (getstring T (strcat "\nDWG folder <" (getvar "DWGPREFIX") ">: ")))
  (if (= (dap:trim inFolder) "")
    (setq folder (getvar "DWGPREFIX"))
    (setq folder (dap:trim inFolder))
  )

  (setq batchIn (getstring T "\nBatch size <5>: "))
  (setq batchSize (dap:parse-int-or batchIn 5))

  (initget "Next Exact")
  (setq revMode (getkword "\nRevision mode [Next/Exact] <Next>: "))
  (if (= revMode "Exact")
    (setq exactRev (dap:up (getstring T "\nExact revision value: ")))
    (setq exactRev nil)
  )

  (setq inDesc (getstring T (strcat "\nDescription <" (dap:getcfg 'default-description) ">: ")))
  (if (= (dap:trim inDesc) "")
    (setq targetDesc (dap:getcfg 'default-description))
    (setq targetDesc (dap:trim inDesc))
  )

  (initget "Yes No")
  (setq closeAns (getkword "\nClose each drawing after processing? [Yes/No] <Yes>: "))
  (setq closeAfter (or (null closeAns) (= closeAns "Yes")))

  (dap:run-batch-core folder batchSize exactRev targetDesc closeAfter)
  (princ)
)

(defun dap:set-auto-options (folder batch-size rev-mode exact-rev target-desc close-after)
  (setq *dap-auto-folder* (if folder (dap:trim folder) ""))
  (setq *dap-auto-batch-size* (dap:parse-int-or (vl-princ-to-string batch-size) 5))
  (setq *dap-auto-rev-mode* (if rev-mode (dap:up rev-mode) "NEXT"))
  (setq *dap-auto-exact-rev* (if exact-rev (dap:trim exact-rev) ""))
  (setq *dap-auto-desc* (if target-desc (dap:trim target-desc) ""))
  (setq *dap-auto-close-after* (dap:truthy-p close-after))
  T
)

(defun dap:run-batch-auto (folder batch-size rev-mode exact-rev target-desc close-after / useExact useDesc res)
  ;; Safety mode: do NOT open/close multiple drawings from LISP.
  ;; External controller (EXE) should handle batch document orchestration.
  (if (and folder (/= (dap:trim folder) ""))
    (princ (strcat "\nINFO: folder parameter received (handled by EXE): " folder))
  )
  (if batch-size
    (princ (strcat "\nINFO: batch-size parameter received (handled by EXE): " (vl-princ-to-string batch-size)))
  )
  (if close-after
    (princ "\nINFO: close-after parameter received (handled by EXE).")
  )

  (if (and (dap:exact-mode-p rev-mode) exact-rev (/= (dap:trim exact-rev) ""))
    (setq useExact (dap:up exact-rev))
    (setq useExact nil)
  )
  (if (and target-desc (/= (dap:trim target-desc) ""))
    (setq useDesc (dap:trim target-desc))
    (setq useDesc (dap:getcfg 'default-description))
  )

  (setq res (dap:run-auto-current-doc useExact useDesc))
  (princ (strcat "\nCurrent drawing result: REV=" (dap:item 'rev res)
                 " changes=" (itoa (dap:item 'changes res))
                 " warnings=" (itoa (dap:item 'warnings res))
                 " status=" (dap:item 'status res)))
  (princ)
)

(defun dap:exe-run (folder batch-size rev-mode exact-rev target-desc close-after)
  (if *dap-exe-mode-only*
    (progn
      (dap:set-auto-options folder batch-size rev-mode exact-rev target-desc close-after)
      (dap:run-batch-auto
        *dap-auto-folder*
        *dap-auto-batch-size*
        *dap-auto-rev-mode*
        *dap-auto-exact-rev*
        *dap-auto-desc*
        *dap-auto-close-after*)
    )
    (princ "\nDAP exe mode is disabled.")
  )
)

(princ)
