#|
 This file is part of claarity
|#

(cl:in-package #:com.splittist.claarity)

(defparameter *default-djula-template*
  (asdf:system-relative-pathname "claarity" "templates/MultifileCommentTemplate.html"))

(defun apply-djula-template (&key template-arguments template outfile)
  (djula:add-template-directory
   (uiop:pathname-directory-pathname (or template *default-djula-template*)))
  (with-open-file (out outfile :direction :output :if-exists :supersede)
    (let ((template (djula:compile-template*
		     (file-namestring (or template *default-djula-template*)))))
      (apply #'djula:render-template* template out template-arguments))))

(defmethod apply-template ((format (eql 'html)) &rest args &key &allow-other-keys)
  (apply #'apply-djula-template args))

(lquery:define-lquery-function run->html (run)
  (run->html run))

(defun run->html (run)
  (let* ((rpr (docxplora:get-first-element-by-tag-name run "w:rPr"))
	 (classes
	   (when rpr
	     (loop for prop across (plump:children rpr)
		   collecting (prop->class prop) into names
		   do (plump:remove-child prop)
		   finally (return (serapeum:string-join (remove-duplicates names :test #'equal) #\Space)))))
	 (text (docxplora::to-text run)))
    (loop for child across (plump:children run)
	  do (plump:remove-child child))
    (setf (plump:tag-name run) "span")
    (unless (alexandria:emptyp classes)
      (setf (plump:attribute run "class") classes))
    (plump:make-text-node run text)
    run))

(defun prop->class (prop)
  (serapeum:string-case (plump:tag-name prop)
    (("w:b" "w:bCs") "b")
    ("w:caps" "caps")
    ("w:dstrike" "dstrike")
    ("w:em" (serapeum:string-case (plump:attribute prop "w:val")
	      ("circle" "em-circle")
	      ("comma" "em-comma")
	      ("dot" "em-dot")
	      ("none" "em-none")
	      ("underDot" "em-underdot")))
    (("w:i" "w:iCs") "i")
    ("w:position" (let ((val (plump:attribute prop "w:val")))
		  (format nil "position-~A" val))) ;; FIXME
    ("w:smallCaps" "smallcaps")
    ("w:strike" "strike")
    ("w:u" (let ((val (plump:attribute prop "w:val")))
	     (cond
	       ((serapeum:string^= "dash" val)
		"u-dashed")
	       ((serapeum:string^= "dot" val)
		"u-dotted")
	       ((serapeum:string^= "wav" val)
		"u-wavy")
	       ((string= "double" val)
		"u-double")
	       ((string= "none" val)
		nil)
	       (t
		"u-solid")))) ;; FIXME words, thick etc.
    ("w:rStyle" (serapeum:string-case (plump:attribute prop "w:val")
		  ("IPCInsertion" "ipc-insertion")
		  ("IPCDeletion" "ipc-deletion")
		  ("IPCMoveTo" "ipc-moveto")
		  ("IPCMoveFrom" "ipc-movefrom")
		  (t "")))
    ("w:vertAlign" (serapeum:string-case (plump:attribute prop "w:val")
		     ("subscript" "vertalign-subscript")
		     ("superscript" "vertalign-superscript")))
    ("w:highlight" (let ((val (plump:attribute prop "w:val")))
		     (format nil "highlight-~A" val)))))

(defmethod process-paragraph-for-format ((format (eql 'html)) paragraph)
  (lquery:with-master-document (paragraph)
    (lquery:$
      "w::r" (run->html)))
  (serapeum:string-replace-all
   "w:p>"
   (opc:serialize paragraph nil)
   "p>"))

(defmethod process-annotation-for-format ((format (eql 'html)) annotation)
  (if annotation
      (serapeum:string-replace-all
       "w:tc>"
       (process-paragraph-for-format 'html annotation)
       "td>")
      "<td/>"))
