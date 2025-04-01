import pymupdf  # PyMuPDF Library is used to work in the PDF

#Looks for Disclaimer verbiage in coordinates and redacts
def redactDisclaimer(page, manualName):
    if manualName == 'Residual Market Manuals':
        disclaimerBox = pymupdf.Rect(163.44, 294.48, 421.92, 309.6)
    elif manualName == "Basic Manual":
        disclaimerBox = pymupdf.Rect((2.69 *72), (4.56*72), (5.87*72), (4.78*72))
    elif manualName == "Assigned Risk" or manualName == "Basic Manual User's Guide" or manualName == "Filing Guide":
        disclaimerBox = pymupdf.Rect(79.92, 459.36, 526.32, 482.4)
    elif manualName == "Experience Rating Plans Manual":
        disclaimerBox = pymupdf.Rect((1.18*72), (6.71*72), (7.24*72), (7*72))
    elif manualName == 'Scopes® Manuals':
        disclaimerBox = pymupdf.Rect((1.11*72), (6.65*72), (6.85*72), (6.81*72))

    page.add_redact_annot(disclaimerBox)
    page.apply_redactions()

