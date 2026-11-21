#!/usr/bin/env python3
"""pdf2word: turn a PDF that has a text layer into an editable Word file (.docx).

It keeps headings (found by font size), paragraphs and tables, in page order. A page with no
text layer is a scan, and it is named in the report and in the document instead of being guessed.
The report also counts the letters and digits on each page against the ones written out, so a
dropped line shows up as a number, not as a surprise later.

usage:  python3 pdf2word.py input.pdf output.docx [--report report.json]
exit:   0 done, every page with text fully written · 1 done, but a page lost text or was a scan · 2 bad input
needs:  PyMuPDF (pip install pymupdf). The .docx is written directly, no Word install needed.
"""
import json, re, statistics, sys, zipfile
from xml.sax.saxutils import escape

import pymupdf

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def run(text, bold=False, size=None):
    rpr = ("<w:b/>" if bold else "") + (f'<w:sz w:val="{int(size * 2)}"/>' if size else "")
    return f'<w:r>{"<w:rPr>" + rpr + "</w:rPr>" if rpr else ""}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def para(text, style=None, bold=False, size=None):
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    return f"<w:p>{ppr}{run(text, bold, size)}</w:p>"


def table(rows):
    ncol = max(len(r) for r in rows)
    border = "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="808080"/>'
                     for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
    out = [f'<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/>'
           f'<w:tblBorders>{border}</w:tblBorders></w:tblPr><w:tblGrid>' + '<w:gridCol/>' * ncol + '</w:tblGrid>']
    for i, r in enumerate(rows):
        cells = list(r) + [""] * (ncol - len(r))
        out.append("<w:tr>" + "".join(f"<w:tc><w:p>{run(c or '', bold=(i == 0))}</w:p></w:tc>" for c in cells) + "</w:tr>")
    out.append("</w:tbl><w:p/>")
    return "".join(out)


STYLES = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
          '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>'
          '<w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
          '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>'
          '<w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style>'
          '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style></w:styles>')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DOCRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')


def inside(b, boxes):
    x0, y0, x1, y1 = b
    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
    return any(t[0] <= cx <= t[2] and t[1] <= cy <= t[3] for t in boxes)


def touches(b, boxes):
    x0, y0, x1, y1 = b
    return any(not (x1 <= c[0] or x0 >= c[2] or y1 <= c[1] or y0 >= c[3]) for c in boxes)


def outer_tables(tables):
    """Drop a table that sits inside another one. A statement cover page gave a 3x1 table and a 1x1
    table inside it, and the letters of the inner one were written twice."""
    keep = []
    for t in tables:
        a = t.bbox
        if not any(o is not t and o.bbox[0] <= a[0] + 1 and o.bbox[1] <= a[1] + 1 and o.bbox[2] >= a[2] - 1 and o.bbox[3] >= a[3] - 1
                   and (o.bbox[2] - o.bbox[0]) * (o.bbox[3] - o.bbox[1]) > (a[2] - a[0]) * (a[3] - a[1]) for o in tables):
            keep.append(t)
    return keep


def count(s):
    return len(re.findall(r"\w", s))


def convert(src, dst):
    doc = pymupdf.open(src)
    sizes = []
    for page in doc:
        for b in page.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    sizes += [round(s["size"], 1)] * max(1, len(s["text"].strip()))
    body_size = statistics.median(sizes) if sizes else 11
    parts, report = [], {"source": src, "pages": [], "body_font_size": body_size}
    for n, page in enumerate(doc, 1):
        if n > 1:   # each PDF page starts a new Word page, so page breaks fall where they did
            parts.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
        info = {"page": n, "headings": 0, "paragraphs": 0, "tables": [], "scan": False}
        raw = page.get_text()
        if count(raw) == 0:
            info["scan"] = bool(page.get_images())
            note = f"[Page {n} has no text layer" + (", it is a scanned image. It was not typed or guessed.]" if info["scan"] else ", it is empty.]")
            parts.append(para(note, bold=True))
            info["chars_in"] = info["chars_out"] = 0
            report["pages"].append(info)
            continue
        tables = outer_tables(page.find_tables().tables)
        tboxes = [t.bbox for t in tables]
        items = [(t.bbox[1], "table", t) for t in tables]
        # Masking is per WORD since 2026-11-21. Masking whole blocks lost a page number and a watermark on a
        # real bank statement, because their block reached into the table area. Words outside every table are
        # kept as text, and words inside a table area but in no cell are kept as a line after that table.
        words = page.get_text("words")
        cells = [[c for c in t.cells if c] for t in tables]
        for ti, t in enumerate(tables):
            # a word that touches any cell is left to the cell, which already holds its letters (a watermark
            # letter laid across a cell was counted twice when only the word centre was tested)
            loose = [w[4] for w in words if inside(w[:4], [t.bbox]) and not touches(w[:4], cells[ti])]
            if loose:
                items.append((t.bbox[3] - 0.01, "text", (" ".join(loose), 0)))
        masked = set()
        for b_no, b in enumerate(page.get_text("dict")["blocks"]):
            if b.get("type") == 0 and inside(b["bbox"], tboxes):
                masked.add(b_no)
        for b_no in sorted(masked):
            keep = [w for w in words if w[5] == b_no and not inside(w[:4], tboxes)]
            if keep:
                items.append((keep[0][1], "text", (" ".join(w[4] for w in keep), 0)))
        for b_no, b in enumerate(page.get_text("dict")["blocks"]):
            if b.get("type") != 0 or b_no in masked:
                continue
            text = " ".join(" ".join(s["text"] for s in l["spans"]).strip() for l in b["lines"]).strip()
            text = re.sub(r"\s+", " ", text)
            if text:
                big = max(s["size"] for l in b["lines"] for s in l["spans"])
                items.append((b["bbox"][1], "text", (text, big)))
        out_chars = 0
        for _, kind, obj in sorted(items, key=lambda x: x[0]):
            if kind == "table":
                rows = [[re.sub(r"\s+", " ", c or "").strip() for c in r] for r in obj.extract()]
                # find_tables can return split columns that are empty in every row (seen on the first fixture,
                # 3 real columns came back as 5), so a column or row with no text in any cell is dropped.
                keep = [i for i in range(max(len(r) for r in rows)) if any(i < len(r) and r[i] for r in rows)]
                rows = [[r[i] if i < len(r) else "" for i in keep] for r in rows if any(r)]
                parts.append(table(rows))
                info["tables"].append(f"{len(rows)}x{max(len(r) for r in rows)}")
                out_chars += sum(count(c) for r in rows for c in r)
            else:
                text, big = obj
                if big >= body_size * 1.6:
                    parts.append(para(text, "Heading1")); info["headings"] += 1
                elif big >= body_size * 1.2:
                    parts.append(para(text, "Heading2")); info["headings"] += 1
                else:
                    parts.append(para(text)); info["paragraphs"] += 1
                out_chars += count(text)
        info["chars_in"], info["chars_out"] = count(raw), out_chars
        report["pages"].append(info)
    body = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{"".join(parts)}<w:sectPr/></w:body></w:document>'
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOCRELS); z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", body)
    lost = [p["page"] for p in report["pages"] if p["chars_out"] != p["chars_in"]]  # fewer is a dropped line, more is a doubled one
    scans = [p["page"] for p in report["pages"] if p["scan"]]
    report["pages_short_of_text"], report["scanned_pages"] = lost, scans
    return report


def main(a):
    if len(a) < 2 or not a[0].lower().endswith(".pdf") or not a[1].lower().endswith(".docx"):
        print(__doc__); return 2
    try:
        rep = convert(a[0], a[1])
    except Exception as e:
        print(f"could not read {a[0]}: {e}"); return 2
    if "--report" in a:
        json.dump(rep, open(a[a.index("--report") + 1], "w"), indent=1)
    t = sum(len(p["tables"]) for p in rep["pages"]); h = sum(p["headings"] for p in rep["pages"])
    print(f"{len(rep['pages'])} pages, {h} headings, {t} tables, scans: {rep['scanned_pages'] or 'none'}, "
          f"pages whose letter count changed: {rep['pages_short_of_text'] or 'none'}")
    return 1 if rep["pages_short_of_text"] or rep["scanned_pages"] else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
