
(vl-load-com) (progn

(setq empty_file "D:\\AcsModule\\Client\\EmptyDwg.dwg")
(defun c:CapolBindFix() 
  (setq res (getint "是否保存 [1]不保存 [2]保存: <取消>"))
  (cond 
    ((eq res 1) (progn (bind-fix) ))
    ((eq res 2) (progn (bind-fix) (command "qsave")))
  )
);only on auto bind fail 
;(defun superclean() (dictremove(namedobjdict)"ACAD_DGNLINESTYLECOMP")(command "-purge" "a" "*" "n")nil)
(defun in-list(itm lst)(setq res nil)(setq li lst) (while li (if (= (car li) itm)(setq res T li nil) (setq li (cdr li)))) res);判断元素是否在列表中
(defun xref-unload-list ()(setq lst nil)(foreach item (get-blocks) (if (= 4 (cdaddr item)) (setq lst(cons item lst))))lst)
(defun get-blocks (/ lst a);list all blocks
  (SETQ LST (LIST  (tblnext "block" t))) (while (setq a (tblnext "block")) (setq lst (cons a lst))) 
lst)
(defun get-file-name (str / p pa sl xn);从文件路径获取文件名
  (setq p "\\")
  (setq xn (1+ (strlen p)))
  (while (setq pa (vl-string-search p str))
    (setq sl (cons (substr str 1 pa) sl))
    (setq str (substr str (+ pa xn)))
    )
  str
  )
(defun purgefile(filepath) (command "CAPOL_PURGEFILEIN" filepath)(command))
(defun superclean() (command "Capol_PurgeIn")(command "-purge" "a" "*" "n")nil)
(defun xrefsuperclean()
  (setq lst nil blocks (vla-Get-Blocks (vla-Get-ActiveDocument (vlax-Get-Acad-Object))))
  (vlax-For blk blocks
    (if (= (vla-Get-IsXref blk) :vlax-True)
      (setq lst (cons (findfile (get-file-name (vla-get-path blk))) lst))
      )
    );end vlax-for
  (foreach file lst(purgefile file))
);end defun
(defun all-xref-lst()
  (setq lst nil blocks (vla-Get-Blocks (vla-Get-ActiveDocument (vlax-Get-Acad-Object))))
  (vlax-For blk blocks
    (if (= (vla-Get-IsXref blk) :vlax-True);(princ  (list(vla-get-name blk)(vla-get-path blk))))
      (setq lst (cons (vla-get-name blk) lst))
      )
    );end vlax-for
  lst
);end defun
(defun lv1-xref-lst()
  (setq obj (entnext))
  (setq lv1_lst nil)
  (setq a_list (all-xref-lst))
  (while obj
    (setq e_name(cdadr (cddddr (cddddr (entget obj)))))
    (setq e_type(cdadr (entget obj)))
    (if (and (= e_type "INSERT") (in-list e_name a_list))
      (setq lv1_lst (cons e_name lv1_lst))
      )
    (setq obj (entnext obj))
  );end while
lv1_lst
)


(defun bind-fix()
  (superclean)
  (foreach xref (xref-unload-list)
    (setq block_name (cdadr xref))
    (command "-xref" "p" block_name empty_file)
  );end foreach
  (command "-xref" "r" "*")
  (foreach xref (xref-unload-list)
    (setq xref_name (cdadr xref))
    (command "-xref" "d" xref_name)
  );end foreach
  (command "-xref" "b" "*")
  (superclean)
nil
);end defun

(defun bind-fix2();一级深度绑定
  nil
  )


)
  


