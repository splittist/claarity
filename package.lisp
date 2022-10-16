#|
 This file is part of claarity
|#

(in-package #:cl)

(defpackage #:com.splittist.claarity
  (:use #:cl)
  (:local-nicknames (#:trackpalette #:com.splittist.trackpalette))
  (:export
   #:report
   #:*format*
   #:wml
   #:html
   #:*default-docxdjula-template*
   #:*default-djula-template*
   #:*outfile-name-generator*
   #:apply-template
   #:process-paragraph-for-format
   #:process-annotation-for-format))
