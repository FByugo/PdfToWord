"""Builds tests/fixture.pdf: page 1 a title, a section heading, two paragraphs and a 3x3 ruled table,
page 2 an image only page (a scan stand in). The expected words are listed so the test can look for them."""
import sys, pymupdf
out = sys.argv[1]
d = pymupdf.open()
p = d.new_page()
p.insert_text((72, 80), "Quarterly Supply Report", fontsize=22)
p.insert_text((72, 120), "Delivery Summary", fontsize=15)
p.insert_text((72, 150), "Orders shipped on time rose in the third quarter.", fontsize=11)
p.insert_text((72, 166), "Two warehouses reported late stock counts.", fontsize=11)
x, y, cw, rh = 72, 200, 120, 22
cells = [["Region", "Orders", "Late"], ["North", "412", "9"], ["South", "388", "14"]]
for r in range(4):
    p.draw_line((x, y + r * rh), (x + 3 * cw, y + r * rh))
for c in range(4):
    p.draw_line((x + c * cw, y), (x + c * cw, y + 3 * rh))
for r, row in enumerate(cells):
    for c, v in enumerate(row):
        p.insert_text((x + c * cw + 6, y + r * rh + 15), v, fontsize=11)
p.insert_text((72, 300), "Counts were taken on the last working day.", fontsize=11)
q = d.new_page()
pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 200, 100), False); pix.clear_with(200)
q.insert_image(pymupdf.Rect(72, 72, 472, 272), pixmap=pix)
d.save(out)
