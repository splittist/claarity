#|
 This file is a part of claarity
|#

(asdf:defsystem claarity
  :version "0.0.1"
  :license "AGPL3.0"
  :author "John Q. Splittist <splittist@splittist.com>"
  :maintainer "John Q. Splittist <splittist@splittist.com>"
  :description "Automatically create reports from docx files with tracked changes, comments, etc."
  :homepage "https://github.com/splittist/claarity"
  :bug-tracker "https://github.com/splittist/claarity/issues"
  :source-control (:git "https://github.com/splittist/claarity.git")
  :serial T
  :components ((:file "package")
               (:file "claarity"))
  :depends-on (#:alexandria
               #:serapeum
               #:uiop
               #:plump
               #:clss
               #:lquery
               #:docxplora
               #:docxdjula
               #:trackpalette))
