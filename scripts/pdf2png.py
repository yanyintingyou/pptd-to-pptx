# -*- coding: utf-8 -*-
import os, sys
import pypdfium2 as pdfium

pdf_path = sys.argv[1]
outdir = sys.argv[2]
scale = float(sys.argv[3]) if len(sys.argv) > 3 else 1.6
os.makedirs(outdir, exist_ok=True)
doc = pdfium.PdfDocument(pdf_path)
names = []
for i in range(len(doc)):
    page = doc[i]
    bmp = page.render(scale=scale)
    img = bmp.to_pil()
    p = os.path.join(outdir, "p%02d.png" % (i + 1))
    img.save(p)
    names.append("%s %s" % (p, img.size))
with open(os.path.join(outdir, "_list.txt"), "w", encoding="utf-8") as fh:
    fh.write("\n".join(names))
print("pages:", len(doc))
