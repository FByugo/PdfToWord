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
# page 3: a ruled 1x2 table whose cell words are placed one by one with no space characters, the way some
# PDFs store text (a lab guide came out as Exploretheeffect before 2026-11-21)
r = d.new_page()
for i in range(4): r.draw_line((72, 100 + i * 24), (472, 100 + i * 24))
for c in (72, 200, 472): r.draw_line((c, 100), (c, 172))
for i, (a, b) in enumerate([("Item", "Detail"), ("Purpose", "Explore the effect of force"), ("Time", "About fifty minutes")]):
    r.insert_text((80, 117 + i * 24), a, fontsize=11)
    x = 208
    for w in b.split():
        r.insert_text((x, 117 + i * 24), w, fontsize=11); x += pymupdf.get_text_length(w, fontsize=11) + 4
r.insert_text((72, 200), "A plain closing sentence on page three.", fontsize=11)
# page 4: rotated 90 degrees, a landscape page with a caption and a ruled table
s4 = d.new_page(width=792, height=612)
s4.insert_text((72, 80), "Table 2: Orders by month", fontsize=11)
for i in range(3): s4.draw_line((72, 100 + i * 22), (372, 100 + i * 22))
for c in range(3): s4.draw_line((72 + c * 150, 100), (72 + c * 150, 144))
for i, row in enumerate([["Month", "Orders"], ["May", "41"]]):
    for c, v in enumerate(row): s4.insert_text((78 + c * 150, 115 + i * 22), v, fontsize=11)
s4.set_rotation(90)
d.save(out)
