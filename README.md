# PdfToWord

Turns a PDF that has a text layer into an editable Word file. Headings are found by font size, paragraphs and tables stay in page order, each PDF page starts a new Word page, and tables come out as real Word tables.

A scanned page has no text to read, so it is not guessed. It is named in the document and in the report.

Every page is checked: the letters and digits on the page are counted against the ones written to the Word file. If a page comes out with fewer or more, the page number is printed and the exit code is 1, so nobody finds out from the client.

## Use

    python3 pdf2word.py input.pdf output.docx --report report.json

Needs PyMuPDF (`pip install pymupdf`). Word does not need to be installed.

## Known limits

Tried on two bank statements (12 pages). Prose pages and the cover pages come out with the same letter count as the PDF. Table pages come out within 7 letters of about 3,000, and every one of those differences is a diagonal SYNTHETIC DATA watermark whose letters fall across table cells. The check still names those pages, which is the point: a watermark is noise to a reader, but the tool does not decide that for the client.

## Test

    bash tests/test.sh

Builds a PDF with a title, a section heading, a ruled table and a scanned page, converts it, and checks the Word file. Two broken copies of the tool are run too, and both must fail.
