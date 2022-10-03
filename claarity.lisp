#|
 This file is part of claarity
|#

(cl:in-package #:com.splittist.claarity)

(defun paragraph-has-revisions (paragraph)
  (not
   (alexandria:emptyp
    (clss:select "w::del,w::ins,w::moveTo,w::moveFrom" paragraph))))

(defun paragraph-has-highlighting (paragraph)
  (not
   (alexandria:emptyp
    (clss:select "w::highlight" paragraph))))

(defun paragraph-has-square-brackets (paragraph)
  (let ((text (docxplora::to-text paragraph)))
    (or (find #\[ text :test #'char=)
	(find #\] text :test #'char=))))

(defun paragraph-has-comments (paragraph)
  (not
   (alexandria:emptyp
    (clss:select "w::commentReference" paragraph))))

(defun paragraph-comment-ids (paragraph)
  (let ((comment-refs (plump:get-elements-by-tag-name paragraph "w:commentReference")))
    (mapcar (alexandria:rcurry #'plump:attribute "w:id") comment-refs)))

(defparameter *comment-reference-replacement*
  "<w:r><w:rPr><w:vertAlign w:val=\"superscript\" /></w:rPr><w:t>comment</w:t>")

(defun replace-comment-references (root)
  (lquery:with-master-document (root)
    (lquery:$
      "w::commentReference"
      (parent)
      (replace-with *comment-reference-replacement*))
    root))

(defun paragraph-has-footnotes (paragraph)
  (not
   (alexandria:emptyp
    (clss:select "w::footnoteReference" paragraph))))

(defun paragraph-footnote-ids (paragraph)
  (let ((footnote-refs (plump:get-elements-by-tag-name paragraph "w:footnoteReference")))
    (mapcar (alexandria:rcurry #'plump:attribute "w:id") footnote-refs)))

(defparameter *footnote-reference-replacement*
  "<w:r><w:rPr><w:vertAlign w:val=\"superscript\" /></w:rPr><w:t>footnote</w:t>")

(defun replace-footnote-references (root)
  (lquery:with-master-document (root)
    (lquery:$
      "w::footnoteReference"
      (parent)
      (replace-with *footnote-reference-replacement*))
    root))

(defun replace-footnote-refs (root)
  (lquery:with-master-document (root)
    (lquery:$
      "w::footnoteRef"
      (parent)
      (replace-with *footnote-reference-replacement*))
    root))

(defun paragraph-has-endnotes (paragraph)
  (not
   (alexandria:emptyp
    (clss:select "w::endnoteReference" paragraph))))

(defun paragraph-endnote-ids (paragraph)
  (let ((endnote-refs (plump:get-elements-by-tag-name paragraph "w:endnoteReference")))
    (mapcar (alexandria:rcurry #'plump:attribute "w:id") endnote-refs)))

(defparameter *endnote-reference-replacement*
  "<w:r><w:rPr><w:vertAlign w:val=\"superscript\" /></w:rPr><w:t>endnote</w:t>")

(defun replace-endnote-references (root)
  (lquery:with-master-document (root)
    (lquery:$
      "w::endnoteReference"
      (parent)
      (replace-with *endnote-reference-replacement*))
    root))

(defun replace-endnote-refs (root)
  (lquery:with-master-document (root)
    (lquery:$
      "w::endnoteRef"
      (parent)
      (replace-with *endnote-reference-replacement*))
    root))

(defun paragraph-has-item-of-interest (paragraph &key revisionsp highlightingp square-brackets-p
						   commentsp footnotesp endnotesp)
  (or (and revisionsp (paragraph-has-revisions paragraph))
      (and highlightingp (paragraph-has-highlighting paragraph))
      (and square-brackets-p (paragraph-has-square-brackets paragraph))
      (and commentsp (paragraph-has-comments paragraph))
      (and footnotesp (paragraph-has-footnotes paragraph))
      (and endnotesp (paragraph-has-endnotes paragraph))))

(defun replace-annotation-references (root &key commentsp footnotesp endnotesp)
  (when commentsp (replace-comment-references root))
  (when footnotesp (replace-footnote-references root))
  (when endnotesp (replace-endnote-references root))
  root)

(defun clone-paragraphs (paragraphs)
  (let ((root (plump:make-root)))
    (serapeum:do-each (p paragraphs)
      (plump:append-child
       root
       (plump:clone-node p t)))
    root))

(defparameter *run-properties-whitelist*
  (serapeum:string-join
   '("w::b" "w::bCs"
     "w::caps"
     "w::cs"
     "w::dstrike"
     "w::em"
     "w::i" "w::iCs"
     "w::position"
     "w::smallCaps"
     "w::strike"
     "w::u"
     "w::rStyle"
     "w::vertAlign"
     "w::highlight")
   ","))

(defun trim-run-properties (root)
  (lquery:with-master-document (root)
    (lquery:$ "w::rPr"
              (children)
              (not *run-properties-whitelist*)
              (remove))))

(defun delete-annotations (root)
  (lquery:with-master-document (root)
    (lquery:$ "w::bookmarkEnd,w::bookmarkStart"
              (add "w::commentRangeEnd,w::commentRangeStart")
              (add "w::annotationRef") ; comments
              (add "w::customXmlDelRangeEnd,w::customXmlDelRangeStart")
              (add "w::customXmlInsRangeEnd,w::customXmlInsRangeStart")
              (add "w::customXmlMoveFromRangeEnd,w::customXmlMoveFromRangeStart")
              (add "w::customXmlMoveToRangeEnd,w::customXmlMoveToRangeStart")
              (add "w::permEnd,w::permStart")
              (remove))))

(defun delete-revision-ids (root)
  (lquery:with-master-document (root)
    (lquery:$ "w::p"
              (remove-attr "w:rsidR" "w:rsidRDefault" "w:rsidRpr" "w:rsidP"
                           "w14:textId" "w14:paraId"))
    (lquery:$ "w::r"
              (remove-attr "w:rsidR" "w:rsidRPr"))))

(defun delete-ppr (root)
  (lquery:with-master-document (root)
    (lquery:$ "w::pPr" (remove))))


(defun delete-out-of-band (root)
  (lquery:with-master-document (root)
    (lquery:$ "w::subDoc,w::commentReference,w::footnoteReference,w::endnoteReference,w::contentPart,w::object"
              (remove))))

(defun delete-separators (root)
  (lquery:with-master-document (root)
    (lquery:$ "w::continuationSeparator,w::lastRenderedPageBreak,w::separator"
              (remove))))

(defun unwrap-containers (root)
  (lquery:with-master-document (root)
    (lquery:$ "w::customXml,w::hyperlink,w::smartTag"
              (splice))))

(defun handle-sdt (root)
  (lquery:with-master-document (root)
    (lquery:$ "w::sdtPr" (remove))
    (lquery:$ "w::sdt" (splice))
    (lquery:$ "w::sdtContent" (splice))))

(defun delete-empty-paras (root)
  (lquery:with-master-document (root)
    (lquery:$ "w::p" (filter #'lquery-funcs:empty-p) (remove))))

(defun process-paragraphs (root)
  (delete-revision-ids root)
  (delete-annotations root)
  (delete-out-of-band root)
  (delete-separators root)
  (handle-sdt root)
  (unwrap-containers root)
  (trackpalette:handle-deleted-field-codes root)
  (trackpalette:handle-inserted-math-characters root :remove-highlighting t)
  (trackpalette:handle-deleted-math-characters root :remove-highlighting t)
  (trackpalette:handle-insertions root :remove-highlighting t)
  (trackpalette:handle-deletions root :remove-highlighting t)
  (trackpalette:handle-move-tos root :remove-highlighting t)
  (trackpalette:handle-move-froms root :remove-highlighting t)
  (delete-ppr root)
  (trim-run-properties root)
  (trackpalette:accept-other-changes root)
  (delete-empty-paras root)
  root)

(defun md-pages-and-sections (paragraphs)
  (let ((result (make-array (length paragraphs)))
        (page 1)
        (section 1))
    (dotimes (i (length paragraphs) result)
      (setf (aref result i) (cons page section))
      (let ((paragraph (elt paragraphs i)))
        (incf page (length
                    (plump:get-elements-by-tag-name paragraph "w:lastRenderedPageBreak")))
        (incf section (length
                       (plump:get-elements-by-tag-name paragraph "w:sectPr")))))))

(defun comment-author-para (author)
  (plump:first-child
   (plump:parse (format nil "<w:p><w:r><w:rPr><w:i/><w:iCs/></w:rPr><w:t>- ~A</w:t></w:r></w:p>"
                        (or author "anon")))))

(defun empty-cell ()
  (plump:first-child (plump:parse "<w:tc/>")))

(defun blank-cell ()
  (plump:first-child (plump:parse "<w:tc><w:p/></w:tc>")))

(defun md-annotations (paragraphs document &key commentsp footnotesp endnotesp)
  (let* ((comment-part (docxplora:comments document))
	 (comment-root (and comment-part (opc:xml-root comment-part)))
	 (footnote-part (docxplora:footnotes document))
	 (footnote-root (and footnote-part (opc:xml-root footnote-part)))
	 (endnote-part (docxplora:endnotes document))
	 (endnote-root (and endnote-part (opc:xml-root endnote-part)))
         (result (make-array (length paragraphs))))
    (dotimes (i (length paragraphs) result)
      (let ((paragraph (elt paragraphs i))
	    (tc (empty-cell)))
	(when (and commentsp (paragraph-has-comments paragraph))
	  (let* ((ids (paragraph-comment-ids paragraph))
		 (comments
		   (and comment-root
			(mapcar (alexandria:curry #'docxplora:get-comment-by-id comment-root)
				ids))))
	    (dolist (comment comments)
	      (serapeum:do-each (child (plump:clone-children comment))
		(plump:append-child tc child))
	      (plump:append-child tc (comment-author-para (docxplora:comment-author comment))))))
	(when (and footnotesp (paragraph-has-footnotes paragraph))
	  (let* ((ids (paragraph-footnote-ids paragraph))
		 (footnotes
		   (and footnote-root
			(mapcar (alexandria:curry #'docxplora:get-footnote-by-id footnote-root)
				ids))))
	    (dolist (footnote footnotes)
	      (serapeum:do-each (child (plump:clone-children footnote))
		(replace-footnote-refs child)
		(plump:append-child tc child)))))
	(when (and endnotesp (paragraph-has-endnotes paragraph))
	  (let* ((ids (paragraph-endnote-ids paragraph))
		 (endnotes
		   (and endnote-root
			(mapcar (alexandria:curry #'docxplora:get-endnote-by-id endnote-root)
				ids))))
	    (dolist (endnote endnotes)
	      (serapeum:do-each (child (plump:clone-children endnote))
		(replace-endnote-refs child)
		(plump:append-child tc child)))))
	(when (alexandria:emptyp (plump:children tc))
	  (setf tc (blank-cell)))
	(setf (aref result i) tc)))))

(defun extract-items-of-interest (document &rest items
				           &key commentsp footnotesp endnotesp
					        revisionsp highlightingp square-brackets-p)
  (let* ((md-root (opc:xml-root (docxplora:main-document document)))
         (all-paras (docxplora::paragraphs-in-document-order md-root))
         (annotations (md-annotations all-paras document
				      :commentsp commentsp
				      :footnotesp footnotesp
				      :endnotesp endnotesp))
         (formatted-numbers (docxplora::paragraph-formatted-numbers document all-paras))
         (pages-and-sections (md-pages-and-sections all-paras))
         (clone-root (plump:make-root))
         (extracted-paras (mapcar
                           (lambda (p)
                             (when (apply #'paragraph-has-item-of-interest p items) 
                               (let ((clone
                                       (plump:append-child
                                        clone-root
                                        (plump:clone-node p t))))
                                 (process-paragraphs
				  (replace-annotation-references clone-root :commentsp commentsp
								            :footnotesp footnotesp
									    :endnotesp endnotesp))
                                 clone)))
                           all-paras)))
    (map 'list #'list extracted-paras formatted-numbers pages-and-sections annotations)))

(defun item-entries-to-arguments (item-entries)
  (let ((arguments '())
        (current-number nil))
    (dolist (p item-entries (nreverse arguments))
      (alexandria:when-let ((n (second p))) (setf current-number n))
      (when (first p)
        (let ((reference (or (second p)
                             current-number
                             (format nil "~C~D,~C~D"
                                     #\Pilcrow_sign (car (third p))
                                     #\Section_sign (cdr (third p))))))
          (push (list :reference reference
                      :revision (plump:serialize (first p) nil)
                      :comment (when (fourth p) (plump:serialize (fourth p) nil)))
                arguments))))))

(defun collect-template-arguments-for-file (file &rest items
					         &key revisionsp commentsp footnotesp endnotesp
					              highlightingp square-brackets-p)
  (let* ((indoc (docxplora:open-document file))
	 (item-entries (apply #'extract-items-of-interest indoc items)))
    (item-entries-to-arguments item-entries)))

(defun collect-template-arguments (infiles &rest items
				           &key revisionsp commentsp footnotesp endnotesp
				                highlightingp square-brackets-p)
  (loop for infile in infiles
	collecting
        (list :name (pathname-name infile)
              :entries (apply #'collect-template-arguments-for-file infile items))))

(defun apply-docxdjula-template (&key template-arguments template outfile)
  (let ((djula:*current-compiler* (make-instance 'docxdjula:docx-compiler))
	(djula:*current-store* (make-instance 'docxdjula::docx-file-store)))
    (djula:add-template-directory
     (uiop:pathname-directory-pathname (or template *default-comment-template*)))
    (let ((template (djula:compile-template*
		     (file-namestring (or template *default-comment-template*))))
	  (djula::*template-arguments* template-arguments))
      (funcall template outfile))))

(defparameter *default-comment-template* (asdf:system-relative-pathname "claarity" "templates/MultifileCommentTemplate.docx"))

(defun report (&key infiles outfile arguments template
		 commentsp footnotesp endnotesp highlightingp (revisionsp t) square-brackets-p)
  (let* ((infiles (alexandria:ensure-list infiles))
	 (template-arguments
	   (append
	    (list :files
		  (collect-template-arguments
		   infiles
		   :commentsp commentsp
		   :footnotesp footnotesp
		   :endnotesp endnotesp
		   :highlightingp highlightingp
		   :revisionsp revisionsp
		   :square-brackets-p square-brackets-p))
	    arguments))
	 (outfile (or outfile (merge-pathnames (format nil "~A-com" (pathname-name (first infiles)))
					       (first infiles)))))
    (apply-docxdjula-template :template-arguments template-arguments
			      :template template
			      :outfile outfile)))
