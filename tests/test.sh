#!/usr/bin/env bash
# PdfToWord planted case test. P2W overrides the tool so a broken copy can be shown to FAIL.
set -u; H=$(cd "$(dirname "$0")/.." && pwd); P=${P2W:-$H/pdf2word.py}; T=$(mktemp -d); trap 'rm -rf "$T"' EXIT; f=0
python3 "$H/tests/make_fixture.py" "$T/in.pdf" 2>/dev/null
python3 "$P" "$T/in.pdf" "$T/out.docx" --report "$T/r.json" > "$T/o" 2>&1; rc=$?
ok(){ if eval "$2"; then echo "pass $1"; else echo "FAIL $1"; f=1; fi; }
ok "exit 1 because page 2 is a scan" '[ $rc = 1 ]'
ok "docx opens as a zip with a document part" 'python3 -c "import zipfile,sys;open(sys.argv[2],\"wb\").write(zipfile.ZipFile(sys.argv[1]).read(\"word/document.xml\"))" "$T/out.docx" "$T/d.xml" 2>/dev/null'
ok "title is Heading1" 'grep -q "Heading1\"/></w:pPr><w:r><w:t xml:space=\"preserve\">Quarterly Supply Report" "$T/d.xml"'
ok "section is Heading2" 'grep -q "Heading2\"/></w:pPr><w:r><w:t xml:space=\"preserve\">Delivery Summary" "$T/d.xml"'
ok "three tables in all"  '[ "$(grep -o "<w:tbl>" "$T/d.xml" | wc -l)" = 3 ]'
ok "first table has 3 rows and 3 columns" 'python3 -c "import sys;t=open(sys.argv[1]).read().split(\"<w:tbl>\")[1].split(\"</w:tbl>\")[0];sys.exit(0 if t.count(\"<w:tr>\")==3 and t.count(\"<w:gridCol/>\")==3 else 1)" "$T/d.xml"'
ok "table cell South 388 14 kept in order" 'grep -q "South</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t xml:space=\"preserve\">388</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t xml:space=\"preserve\">14<" "$T/d.xml"'
ok "table words not repeated as loose text" '[ "$(grep -o "North" "$T/d.xml" | wc -l)" = 1 ]'
ok "paragraph after the table comes after it" 'python3 -c "import sys;d=open(sys.argv[1]).read();sys.exit(0 if d.index(\"<w:tbl>\")<d.index(\"last working day\") else 1)" "$T/d.xml"'
ok "a page break between each of the four pages" '[ "$(grep -o "w:type=\"page\"" "$T/d.xml" | wc -l)" = 3 ]'
ok "scan page named, not guessed" 'grep -q "Page 2 has no text layer, it is a scanned image" "$T/d.xml"'
ok "report: page 1 wrote every letter once, none lost or doubled" 'python3 -c "import json,sys;r=json.load(open(sys.argv[1]));p=r[\"pages\"][0];sys.exit(0 if p[\"chars_out\"]==p[\"chars_in\"]>0 and r[\"scanned_pages\"]==[2] else 1)" "$T/r.json"'
ok "a table inside another table is dropped" 'python3 -c "
import importlib.util,sys,types
s=importlib.util.spec_from_file_location(\"m\",sys.argv[1]);m=importlib.util.module_from_spec(s);s.loader.exec_module(m)
T=lambda b:types.SimpleNamespace(bbox=b)
a,b,c=T((10,10,300,200)),T((20,20,100,60)),T((310,10,400,80))
k=m.outer_tables([a,b,c]);sys.exit(0 if k==[a,c] else 1)" "$P"'
ok "a word touching a cell is left to the cell" 'python3 -c "
import importlib.util,sys
s=importlib.util.spec_from_file_location(\"m\",sys.argv[1]);m=importlib.util.module_from_spec(s);s.loader.exec_module(m)
sys.exit(0 if m.touches((95,10,105,20),[(0,0,100,100)]) and not m.touches((101,10,110,20),[(0,0,100,100)]) else 1)" "$P"'
ok "cell words placed one by one come out with spaces" 'grep -q ">Explore the effect of force<" "$T/d.xml"'
ok "rotated page keeps its caption and table, every letter once" 'python3 -c "import json,sys;r=json.load(open(sys.argv[1]));p=r[\"pages\"][3];sys.exit(0 if p[\"chars_in\"]==p[\"chars_out\"]>0 and r[\"rotated_pages\"]==[4] and p[\"tables\"] else 1)" "$T/r.json"'
ok "a non pdf input exits 2" 'python3 "$P" "$T/r.json" "$T/x.docx" >/dev/null 2>&1; [ $? = 2 ]'
if [ -z "${P2W:-}" ]; then   # broken copies must fail this test
  sed 's/        tboxes = \[t.bbox for t in tables\]/        tboxes = []/' "$P" > "$T/m1.py"
  P2W="$T/m1.py" bash "$0" >/dev/null 2>&1 && { echo "FAIL a copy that prints table words twice passed"; f=1; } || echo "pass copy with tables not masked fails"
  sed 's/            info\["scan"\] = bool(page.get_images())/            continue/' "$P" > "$T/m2.py"
  P2W="$T/m2.py" bash "$0" >/dev/null 2>&1 && { echo "FAIL a copy that skips scans silently passed"; f=1; } || echo "pass copy that skips scans fails"
  # no planted case found yet that makes extract() drop spaces while words still split (gaps of 0.2 to 4 pt tried
  # 2026-11-21), so the cell_rows fix rests on one real lab guide PDF that cannot be shipped in a public test
  sed 's/^            p.set_rotation(0)/            pass/' "$P" > "$T/m5.py"
  P2W="$T/m5.py" bash "$0" >/dev/null 2>&1 && { echo "FAIL a copy that reads rotated pages as they are passed"; f=1; } || echo "pass copy that reads rotated pages as they are fails"
  sed 's/            keep.append(t)/            keep.append(t)\n    return list(tables)/' "$P" > "$T/m3.py"
  P2W="$T/m3.py" bash "$0" >/dev/null 2>&1 && { echo "FAIL a copy that keeps nested tables passed"; f=1; } || echo "pass copy that keeps nested tables fails"
fi
exit $f
