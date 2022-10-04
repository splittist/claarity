# claarity
Automatically create reports from docx files with tracked changes, comments, etc.

claarity uses [Docxplora](https://github.com/splittist/docxplora) to extract paragraphs with items of interest
from `docx` files (technically, files in Open Office XML Wordprocessing markup format)
and populate [docxdjula](https://github.com/splittist/docxdjula) templates with those items.
When the item of interest is a paragraph with [tracked changes](https://support.microsoft.com/en-us/office/track-changes-in-word-197ba630-0f5f-4a8e-9a77-3712475e806a)
the changes are displayed with formatting supplied by [trackpalette](https://github.com/splittist/trackpalette).

## Usage

*function* **REPORT** &key *infiles* *outfile* *arguments* *template*  *(revisionsp t)* *highlightingp* *square-brackets-p* *commentsp* *footnotesp* *endnotesp*

Extracts paragraphs with items of interest
from the files designated by *infiles* (a single path or a list of paths),
passing them, together with the addtional template
arguments in *arguments* (a list) to the docxdjula template designated by *template*,
writing the result to *outfile*.

*template* defaults to `*default-comment-template*`.

*outfile* defaults to the first pathname in *infiles* with "-com" appended to the pathname-name
(i.e. before the extension).

If *revisionsp* is true (the default), paragraphs with tracked changes will be extracted, those tracked changes being styled by **trackpalette**.

If *highlightingp* is true, paragraphs with [highlighting](https://support.microsoft.com/en-us/office/apply-or-remove-highlighting-1747d808-6db7-4d49-86ac-1f0c3cc87e2e) will be extracted.

If *square-brackets-p* is true, paragraphs with `#\[` or `#\]` will be extracted.

If *commentsp* is true, paragraphs with comments will be extracted, as will the text and the author of the associated comment(s).

If *footnotesp* is true, paragraphs with footnotes will be extracted, as will the corresponding text of the footnote.

If *endnotesp* is true, paragraphs with endnotes will be extracted, as will the corresponding text of the endnote.

Note that the style of the extracted paragraphs will be simplified to improve the clarity of the report.

*special* **\*DEFAULT-COMMENT-TEMPLATE\***

Set to the path of a file named "MultifileCommentTemplate.docx" in the "Templates" directory of the "claarity" asdf system directory.

## Templates

The template will be passed a list of template arguments consisting of the arguments passed in as *arguments* to the `REPORT` function and and entry `:FILES`, consiting of one plist for each of the *infiles* consisting of one entry (`:NAME`) for the `PATHNAME-NAME` of the `infile`, and one entry (`:ENTRIES`) containing a list of plists with `:REFERENCE`, `:REVISION` and `:COMMENT` entries, the values being a reference for the paragraph, the text of the paragraph, and the text of any comments, footnotes or endnotes if those items have been selected as being of interest.

The template can, therefore, iterate through the files in `files`, and for each file refer to its `file.name` and the `file.entries`, each entry containing an `entry.reference`, `entry.revision` and (possibly blank) `entry.comment`.

Note that `entry.revision` and `entry.comment` are xml, so need to be passed to the `safe` psuedo-filter.
