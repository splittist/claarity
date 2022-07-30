#|
 This file is part of claarity
|#

(cl:in-package #:com.splittist.claarity)

(defun paragraph-has-revisions (paragraph)
  (not
   (zerop
    (length
     (clss:select "w::del,w::ins,w::moveTo,w::moveFrom" paragraph)))))

(defun paragraph-has-comments (paragraph)
  (not
   (zerop
    (length
     (clss:select "w::commentReference" paragraph)))))

(defun paragraph-comment-ids (paragraph)
  (let ((comment-refs (plump:get-elements-by-tag-name paragraph "w:commentReference")))
    (mapcar (alexandria:rcurry #'plump:attribute "w:id") comment-refs)))

(defun replace-comment-references (root)
  (lquery:with-master-document (root)
    (lquery:$ "w::commentReference" (replace-with "<w:t>[*]</w:t>"))
    root))

(defun clone-paragraphs (paragraphs)
  (let ((root (plump:make-root)))
    (serapeum:do-each (p paragraphs)
      (plump:append-child
       root
       (plump:clone-node p t)))
    root))

(defvar *run-properties-whitelist*
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
     "w::rStyle")
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
    (lquery:$ "w::subDoc,w::commentReference,w::contentPart,w::object"
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
  (trackpalette:handle-inserted-math-characters root)
  (trackpalette:handle-deleted-math-characters root)
  (trackpalette:handle-insertions root)
  (trackpalette:handle-deletions root)
  (trackpalette:handle-move-tos root)
  (trackpalette:handle-move-froms root)
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

(defun author-para (author)
  (plump:first-child
   (plump:parse (format nil "<w:p><w:r><w:rPr><w:i/><w:iCs/></w:rPr><w:t>- ~A</w:t></w:r></w:p>"
                        author))))

(defun md-comments (paragraphs document)
  (let ((comment-root (opc:xml-root (docxplora:comments document)))
        (result (make-array (length paragraphs))))
    (dotimes (i (length paragraphs) result)
      (setf (aref result i)
            (if (paragraph-has-comments (elt paragraphs i))
                (let* ((paragraph (elt paragraphs i))
                       (ids (paragraph-comment-ids paragraph))
                       (comments
                         (mapcar (alexandria:curry #'docxplora:get-comment-by-id comment-root)
                                 ids))
                       (tc (plump:first-child (plump:parse "<w:tc/>"))))
                  (dolist (comment comments (process-paragraphs tc))
                    (serapeum:do-each (child (plump:clone-children comment))
                      (plump:append-child tc child))
                    (plump:append-child tc (author-para (docxplora:comment-author comment)))
                    (setf tc (process-paragraphs tc))))
                (plump:first-child (plump:parse "<w:tc><w:p/></w:tc>")))))))

(defun md-extract-paragraphs (document)
  (let* ((md (docxplora:main-document document))
         (root (opc:xml-root md))
         (all-paras (docxplora::paragraphs-in-document-order root))
         (list-infos (mapcar
                      (alexandria:curry #'docxplora::make-paragraph-list-info document)
                      all-paras))
         (formatted-numbers (docxplora::paragraph-formatted-numbers list-infos))
         (pages-and-sections (md-pages-and-sections all-paras))
         (clone-root (plump:make-root))
         (extracted-paras (mapcar
                           (lambda (p)
                             (when (paragraph-has-revisions p)
                               (let ((clone
                                       (plump:append-child
                                        clone-root
                                        (plump:clone-node p t))))
                                 (process-paragraphs clone-root)
                                 clone)))
                           all-paras)))
    (map 'list #'list extracted-paras formatted-numbers pages-and-sections)))

(defun extract-revisions-and-comments (document)
  (let* ((md-root (opc:xml-root (docxplora:main-document document)))
         (all-paras (docxplora::paragraphs-in-document-order md-root))
         (comments (md-comments all-paras document))
         (formatted-numbers (docxplora::paragraph-formatted-numbers document all-paras))
         (pages-and-sections (md-pages-and-sections all-paras))
         (clone-root (plump:make-root))
         (extracted-paras (mapcar
                           (lambda (p)
                             (when (or (paragraph-has-revisions p)
                                       (paragraph-has-comments p))
                               (let ((clone
                                       (plump:append-child
                                        clone-root
                                        (plump:clone-node p t))))
                                 (process-paragraphs (replace-comment-references clone-root))
                                 clone)))
                           all-paras)))
    (map 'list #'list extracted-paras formatted-numbers pages-and-sections comments)))

(defun md-entries (md-paras)
  (let ((entries '())
        (current-number nil))
    (dolist (p md-paras (nreverse entries))
      (alexandria:when-let ((n (second p))) (setf current-number n))
      (when (first p)
        (let ((reference (or (second p)
                             current-number
                             (format nil "~C~D,~C~D"
                                     #\Pilcrow_sign (car (third p))
                                     #\Section_sign (cdr (third p))))))
          (push (list :reference reference
                      :revision (plump:serialize (first p) nil))
                entries))))))

(defun revisions-and-comments-entries (revisions-and-comments)
  (let ((entries '())
        (current-number nil))
    (dolist (p revisions-and-comments (nreverse entries))
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
                entries))))))

(defun collect-template-arguments (infiles commentsp)
  (loop for infile in infiles collecting
                              (let* ((indoc (docxplora:open-document infile))
                                     (md-paras (funcall (if commentsp
                                                            #'extract-revisions-and-comments
                                                            #'md-extract-paragraphs)
                                                        indoc)))
                                (list :name (pathname-name infile)
                                      :entries (funcall (if commentsp
                                                            #'revisions-and-comments-entries
                                                            #'md-entries)
                                                        md-paras)))))

(defun issues-list (&key infiles outfile commentsp arguments template)
  (let ((template-arguments (collect-template-arguments infiles commentsp))
        (djula:*current-compiler* (make-instance 'docxdjula:docx-compiler))
        (djula:*current-store* (make-instance 'docxdjula::docx-file-store)))
    (djula:add-template-directory
     (uiop:pathname-directory-pathname (or template *default-comment-template*)))
    (let ((template (djula:compile-template*
                     (file-namestring (or template *default-comment-template*))))
          (djula::*template-arguments* (append (list :files template-arguments) arguments)))
      (funcall template
               (or outfile
                   (merge-pathnames (format nil "~A-com" (pathname-name (first infiles)))
                                    (first infiles)))))))
