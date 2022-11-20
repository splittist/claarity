#|
 This file is a part of claarity
|#

(in-package #:com.splittist.claarity)

(defun w-t->t (wt &optional upcase)
  (docxplora:make-text-element (plump:make-root)
			       (sml::string-xstring (let ((text (docxplora::to-text wt)))
						      (if upcase (string-upcase text) text)))
			       "t"))

(defun make-default-rpr () ; FIXME get/ensure default style
  (let ((rpr (plump:make-element (plump:make-root) "rPr")))
    (docxplora:make-element/attrs rpr "sz" "val" "11")
    (docxplora:make-element/attrs rpr "color" "theme" "1")
    (docxplora:make-element/attrs rpr "rFont" "val" "Calibri")
    (docxplora:make-element/attrs rpr "family" "val" "2")
    (docxplora:make-element/attrs rpr "scheme" "val" "minor")
    rpr))

(defun set-rpr-sz (rpr sz)
  (let ((sz-element (plump-utils:find-child/tag rpr "sz")))
    (setf (plump:attribute sz-element "val") sz)))

(defun set-rpr-color (rpr color)
  (let ((color-element (plump-utils:find-child/tag rpr "color")))
    (plump:remove-attribute color-element "theme")
    (setf (plump:attribute color-element "rgb") color)))

(defun ensure-bold (rpr)
  (unless (docxplora:find-child/tag rpr "b")
    (plump:prepend-child rpr (plump:make-element (plump:make-root) "b"))))

(defun ensure-italic (rpr)
  (unless (docxplora:find-child/tag rpr "i")
    (plump:prepend-child rpr (plump:make-element (plump:make-root) "i"))))

(defun w-rpr->rpr (wrpr)
  (loop with rpr = (make-default-rpr)
	with upcase = nil
	for child across (plump:children wrpr)
	do (serapeum:string-case (plump:tag-name child)
	     (("w:b" "w:bCs") (ensure-bold rpr))
	     ("w:caps" (setf upcase t))
	     ("w:dstrike" (plump:prepend-child rpr (plump:make-element (plump:make-root) "strike")))
	     ("w:em" nil)
	     (("w:i" "w:iCs") (ensure-italic rpr))
	     ("w:position" nil)
	     ("w:smallCaps" (set-rpr-sz rpr "8") ; FIXME ?? - looks ok for 11pt
			    (setf upcase t))
	     ("w:strike" (plump:prepend-child rpr (plump:make-element (plump:make-root) "strike")))
	     ("w:u" (let ((val (plump:attribute child "w:val")))
		      (if (string= "double" val)
			  (plump:prepend-child rpr (docxplora:make-element/attrs (plump:make-root) "u" "val" "double"))
			  (plump:prepend-child rpr (plump:make-element (plump:make-root) "u")))))
	     ("w:rStyle" (serapeum:string-case (plump:attribute child "w:val")
			   ("IPCInsertion" (plump:prepend-child
					    rpr
					    (docxplora:make-element/attrs (plump:make-root) "u" "val" "double"))
					   (set-rpr-color rpr "FF0000FF"))
			   ("IPCDeletion" (plump:prepend-child
					   rpr
					   (plump:make-element (plump:make-root) "strike"))
					  (set-rpr-color rpr "FFFF0000"))
			   ("IPCMoveTo" (plump:prepend-child
					 rpr
					 (docxplora:make-element/attrs (plump:make-root) "u" "val" "double"))
					(set-rpr-color rpr "FF00C000"))
			   ("IPCMoveFrom" (plump:prepend-child
					   rpr
					   (plump:make-element (plump:make-root) "strike"))
					  (set-rpr-color rpr "FF00C000"))
			   (t nil)))
	     ("w:vertAlign" (plump:prepend-child
			     rpr
			     (docxplora:make-element/attrs
			      (plump:make-root) "vertAlign" "val" (plump:attribute child "w:val"))))
	     ("w:highlight" nil)) ; FIXME what to do?
	finally (return (values rpr upcase))))

(defun run->sml (run)
  (loop with rpr = nil
	with upcase = nil
	for child across (plump:children run)
	when (docxplora:tagp child "w:rPr")
	  do (multiple-value-bind (new uc) (w-rpr->rpr child)
	       (setf rpr new
		     upcase uc))
	when (docxplora:tagp child "w:t")
	  collect (w-t->t child upcase) into texts
	finally
	   (return
	     (if (null texts)
		 nil
		 (plump:make-element (plump:make-root)
				     "r"
				     :children (plump:ensure-child-array
						(coerce (if rpr (cons rpr texts) texts) 'vector)))))))

(defun build-array (table-entries)
  (loop with table = '()
	for file in table-entries
	for filename = (getf file :name)
	for entries = (getf file :entries)
	do (loop for entry in entries
		 for reference = (getf entry :reference)
		 for paragraph = (getf entry :paragraph)
		 for annotation = (getf entry :annotation)
		 do (push (list filename reference paragraph annotation) table))
	finally
	   (return (make-array (list (length table) 4) :initial-contents (nreverse table)))))

(defmethod process-paragraph-for-format ((format (eql 'sml)) paragraph)
  (loop for child across (plump:children paragraph)
	when (and (docxplora:tagp child "w:r") (run->sml child))
	  collect it into children
	finally
	   (return (plump:make-root (plump:ensure-child-array (coerce children 'vector))))))

(defun line-break ()
  (let* ((root (plump:make-root))
	 (run (plump:make-element root "r")))
    (plump:append-child run (make-default-rpr))
    (docxplora:make-text-element run (concatenate 'string (list #\Space #\Return)) "t")
    (plump:children root)))

(defun build-annotation (paras)
  (plump:make-root
   (plump:ensure-child-array
    (reduce (lambda (a b) (concatenate 'vector a (line-break) b))
	    (mapcar #'plump:children paras)))))

(defmethod process-annotation-for-format ((format (eql 'sml)) annotation)
  (if annotation
      (loop for paragraph across (plump:children annotation)
	    collecting (process-paragraph-for-format 'sml paragraph) into paras
	    finally (return (case (length paras)
			      (1 (first paras))
			      (t (build-annotation paras)))))
      " "))

(defmethod apply-template ((format (eql 'sml)) &key template-arguments template outfile &allow-other-keys)
  (let* ((we (sml::make-workbook-editor template))
	 (sst (sml::workbook-editor-sst we))
	 (ws (sml::edit-worksheet we "Data"))
	 (table-entries (getf template-arguments :files)))
    (sml::write-array ws 3 2 (build-array table-entries)) ; FIXME Update Table size
    (update-string-table sst template-arguments)
    (sml::write-workbook we outfile)))

;;; vrbl substitution

(defparameter *variable-scanner* (cl-ppcre:create-scanner "{{\\s*([a-zA-Z0-9_]+)\\s*}}"))

(defun replace-variables (string arguments)
  (labels ((replacement (match variable)
	     (declare (ignore match))
	     (if (null variable)
		 ""
		 (let ((result (getf arguments (intern (string-upcase variable) :keyword) "")))
		   (princ-to-string result)))))
    (cl-ppcre:regex-replace-all *variable-scanner* string #'replacement :simple-calls t)))

(defun update-string-table (string-table arguments)
  (let ((entries (alexandria:hash-table-alist (sml::sst-dict string-table))))
    (loop for entry in entries
	  do (setf (car entry)
		   (replace-variables (car entry) arguments))
	  finally (setf (sml::sst-dict string-table)
			(alexandria:alist-hash-table entries :test #'equal)))))


