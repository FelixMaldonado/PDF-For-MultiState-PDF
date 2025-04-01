import pymupdf  # PyMuPDF Library is used to work in the PDF
from datetime import datetime

#Looks for Effective Since date and verbiage in coordinates and redacts
def redactEffectiveSince(page, manualName):

    if manualName == 'Scopes® Manuals':
        effectiveIssuedBox = pymupdf.Rect((2.43*72), (4.58*72), (6.17*72), (5.25*72))
        
        current_date = datetime.now()
        formatted_date = current_date.strftime("Published Date: %B %d, %Y")

        page.add_redact_annot(effectiveIssuedBox)
        page.apply_redactions()

        # Register the fonts on the page        
        normal_font_path = 'C:\\Windows\\Fonts\\Calibri.ttf'
        bold_font_path = 'C:\\Windows\\Fonts\\Calibrib.ttf'
        page.insert_font(fontfile=normal_font_path, fontname="calibri")
        page.insert_font(fontfile=bold_font_path, fontname="calibri-bold")
        
        page.insert_text(((2.7*72), (4.97*72)), formatted_date, fontsize=16, fontname="calibri", color=(0, 0, 0))
        
        return
    elif manualName == 'Residual Market Manuals' or  manualName == "Basic Manual":
        effectiveIssuedBox = pymupdf.Rect(79.92, 460.08, 514.8, 482.4)
    elif manualName == "Assigned Risk" or manualName == "Basic Manual User's Guide":
        effectiveIssuedBox = pymupdf.Rect(202.32, 351.76, 416.16, 363.6)
    elif manualName == "Experience Rating Plans Manual":
        effectiveIssuedBox = pymupdf.Rect((3.03*72), (5.61*72), (5.57*72), (5.85*72))
    elif manualName == "Filing Guide":
        effectiveIssuedBox = pymupdf.Rect((2.8*72), (4.55*72), (5.76*72), (4.76*72))

    page.add_redact_annot(effectiveIssuedBox)
    page.apply_redactions()
