import pymupdf  # PyMuPDF Library is used to work in the PDF

# Finds the paragraphs after the states and moves them correct location
def moveParasUp(page, manualName):
   if manualName == "Experience Rating Plans Manual":
      paras = [pymupdf.Rect((1.17*72), (7.56*72), (7.53*72), (7.68*72)),
               pymupdf.Rect((1.17*72), (7.76*72), (7.53*72), (8.2*72)),
               pymupdf.Rect((1.17*72), (8.28*72), (7.53*72), (8.89*72)),
               pymupdf.Rect((1.17*72), (8.95*72), (7.53*72), (10.19*72)),
               pymupdf.Rect((1.17*72), (10.22*72), (7.53*72), (10.55*72))]
     
      x0, y0 = (1.17*72), (7.27*72)
   elif manualName == "Scopes® Manuals":
      paras = [pymupdf.Rect((1.11*72), (7.37*72), (6.56*72), (7.52*72)),
               pymupdf.Rect((1.11*72), (7.36*72), (7.4*72), (7.93*72)),
               pymupdf.Rect((1.11*72), (8.02*72), (7.46*72), (8.69*72)),
               pymupdf.Rect((1.11*72), (8.72*72), (7.44*72), (9.96*72)),
               pymupdf.Rect((1.11*72), (10.05*72), (7.51*72), (10.42*72))]
     
      x0, y0 = (1.12*72), (7.3*72)
   else:
      paras = [pymupdf.Rect(79.92, 521.28, 409, 533.52),
               pymupdf.Rect(79.92, 539.28, 529.92, 564.48),
               pymupdf.Rect(79.92, 570.24, 531.36, 615.6),
               pymupdf.Rect(79.92, 621.36, 532.8, 709.92),
               pymupdf.Rect(79.92, 716.4, 522.72, 741.6),]
     
      x0, y0 = 80.64, 510.42
     
   for para in paras:
      paraText = page.get_text("text", clip=para)

      page.add_redact_annot(para)
      page.apply_redactions()

      # Register the fonts on the page        
      normal_font_path = 'C:\\Windows\\Fonts\\Calibri.ttf'
      bold_font_path = 'C:\\Windows\\Fonts\\Calibrib.ttf'
      page.insert_font(fontfile=normal_font_path, fontname="calibri")
      page.insert_font(fontfile=bold_font_path, fontname="calibri-bold")

      if manualName == "Experience Rating Plans Manual":
         if para == paras[0]:
            page.insert_text((x0, y0), paraText, fontsize=9, fontname="calibri", color=(0, 0, 0))
            y0 += 17
         elif para == paras[1]:
            page.insert_text((x0, y0), paraText, fontsize=9, fontname="calibri", color=(0, 0, 0))
            y0 += 38
         elif para == paras[2]:
            page.insert_text((x0, y0), paraText, fontsize=9, fontname="calibri", color=(0, 0, 0))
            y0 += 50
         elif para == paras[3]:
            page.insert_text((x0, y0), paraText, fontsize=9, fontname="calibri-bold", color=(0, 0, 0))
            y0 += 90
         elif para == paras[4]:
            page.insert_text((x0, y0), paraText, fontsize=9, fontname="calibri-bold", color=(0, 0, 0))
            y0 += 22
      elif manualName == "Scopes® Manuals":
         if para == paras[0]:
            page.insert_text((x0, y0), paraText, fontsize=9, fontname="calibri", color=(0, 0, 0))
            y0 += 17
         elif para == paras[1]:
            page.insert_text((x0, y0), paraText, fontsize=9, fontname="calibri", color=(0, 0, 0))
            y0 += 30
         elif para == paras[2]:
            page.insert_text((x0, y0), paraText, fontsize=9, fontname="calibri", color=(0, 0, 0))
            y0 += 50
         elif para == paras[3]:
            page.insert_text((x0, y0), paraText, fontsize=9, fontname="calibri-bold", color=(0, 0, 0))
            y0 += 94
         elif para == paras[4]:
            page.insert_text((x0, y0), paraText, fontsize=9, fontname="calibri-bold", color=(0, 0, 0))       
      else:
         if para == paras[0]:
            page.insert_text((x0, y0), paraText, fontsize=9, fontname="calibri", color=(0, 0, 0))
            y0 += 20
         elif para == paras[1]:
            page.insert_text((x0, y0), paraText, fontsize=9, fontname="calibri", color=(0, 0, 0))
         elif para == paras[2]:
            page.insert_text((x0, 562), paraText, fontsize=9, fontname="calibri", color=(0, 0, 0))
         elif para == paras[3]:
            page.insert_text((x0, 616), paraText, fontsize=9, fontname="calibri-bold", color=(0, 0, 0))
         elif para == paras[4]:
            page.insert_text((x0, 710), paraText, fontsize=9, fontname="calibri-bold", color=(0, 0, 0))