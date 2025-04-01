import pymupdf  # PyMuPDF Library is used to work in the PDF

#Looks for Effective Since date and verbiage in coordinates and redacts
def findAndRedactStates(page, manual):

    if manual == "Experience Rating Plans Manual":
        statesXY1 = pymupdf.Rect((3.31*72), (7.19*72), (7.37*72), (7.37*72))
        statesXY2 = pymupdf.Rect((1.17*72), (7.35*72), (7.43*72), (7.52*72))
    elif manual == "Scopes® Manuals":
        statesXY1 = pymupdf.Rect((3.27*72), (6.98*72), (7.34*72), (7.12*72))  #7.34
        statesXY2 = pymupdf.Rect((1.13*72), (7.13*72), (7.35*72), (7.26*72))  #1.13
    else:
        statesXY1 = pymupdf.Rect(236.16, 493.91, 532.8, 505.44)
        statesXY2 = pymupdf.Rect(79.92, 504.72, 531.36, 514.8)

    statesFirstRow = page.get_text("text", clip=statesXY1)
    statesSecondRow = page.get_text("text", clip=statesXY2)

    page.add_redact_annot(statesXY1)
    page.apply_redactions() 

    page.add_redact_annot(statesXY2)
    page.apply_redactions()

    inputAndFormatStates(statesFirstRow, statesSecondRow, page, manual)


# Redacts states and moves them to new location
def inputAndFormatStates(statesFirstRow, statesSecondRow, page, manual):
    # List of State Abreviations for comparison
    states = [
    "AK", "AL", "AR", "AZ", "CA", "CO", "CT", "DC", "DE", "FL",
    "GA", "HI", "IA", "ID", "IL", "IN", "KS", "KY", "LA", "MA",
    "MD", "ME", "MI", "MN", "MO", "MS", "MT", "NC", "ND", "NE",
    "NH", "NJ", "NM", "NV", "NY", "OH", "OK", "OR", "PA", "PR",
    "RI", "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WI",
    "WV", "WY"
    ]

    # Get the Para with the states in it and redact it
    if manual == "Experience Rating Plans Manual":
        verbiageBox = pymupdf.Rect((1.17*72), (6.90*72), (7.43*72), (7.52*72))
    else:
        verbiageBox = pymupdf.Rect(79.92, 480.24, 532.08, 517.68)
    
    verbiage = page.get_text("text", clip=verbiageBox)
    
    page.add_redact_annot(verbiageBox)
    page.apply_redactions()

    # Register the fonts on the page        
    normal_font_path = 'C:\\Windows\\Fonts\\Calibri.ttf'
    bold_font_path = 'C:\\Windows\\Fonts\\Calibrib.ttf'
    page.insert_font(fontfile=normal_font_path, fontname="calibri")
    page.insert_font(fontfile=bold_font_path, fontname="calibri-bold")

    # Insert verbiage
    if manual == "Experience Rating Plans Manual":
        page.insert_text(((1.17*72), (6.75*72)), verbiage, fontsize=9, fontname="calibri", color=(0, 0, 0))
    elif manual == "Scopes® Manuals":
        page.insert_text(((1.12*72), (6.72*72)), verbiage, fontsize=9, fontname="calibri", color=(0, 0, 0))
    else:
        page.insert_text((80.64, 469.36), verbiage, fontsize=9, fontname="calibri", color=(0, 0, 0))
    
            
    #split text by word, parts is an array of words in textbox
    parts = statesFirstRow.split()

    #Calculate position of state abrev. used to determine spacing
    if manual == "Experience Rating Plans Manual":
        current_x, y0  = 239.02, (6.9*72)
    elif manual == "Scopes® Manuals":
        current_x, y0  = 237, (6.87*72)
    else:
        current_x, y0  = 237, 480.7
        
    counter = 0

    #insert text dynamic color formatting
    for part in parts:
        part_width = pymupdf.get_text_length(part, fontsize=9)
        if part.strip(",") in states:
            if counter == 0:
                page.insert_text((current_x, y0), part, fontsize=9, fontname="calibri", color = (0, 0, 1)) #Blue color for State abbeviations
                current_x += part_width
                counter += 1
            else:
                page.insert_text((current_x, y0), part, fontsize=9, fontname="calibri", color = (0, 0, 1)) #Blue color for State abbeviations
                current_x += part_width
        else:
            page.insert_text((current_x, y0), part + " ", fontsize=9, fontname="calibri", color = (0, 0, 1)) #Blue color for "and"


    ################## 2nd ROW of States #########################################
    # Get the states on 2nd Row
    parts = statesSecondRow.split()
    if manual == "Experience Rating Plans Manual":
        x0, y0  = (1.17*72), (7.05*72)
    elif manual == "Scopes® Manuals":
        x0, y0  = (1.12*72), (7.05*72)
    else:
        x0, y0 = 80.64, 490.7
    counter = 0
    #insert text dynamic color formatting
    for part in parts:
        part_width = pymupdf.get_text_length(part, fontsize=9)
        if part.strip(",") in states:
            if counter == 0:
                page.insert_text((x0, y0), part, fontsize=9, fontname="calibri", color = (0, 0, 1)) #Blue color for State abbeviations
                x0 += part_width
                counter += 1
            else:
                page.insert_text((x0, y0), part, fontsize=9, fontname="calibri", color = (0, 0, 1)) #Blue color for State abbeviations
                x0 += part_width
        else:
            page.insert_text((x0, y0), part + " ", fontsize=9, fontname="calibri", color = (0, 0, 1)) #Blue color for "and"
            x0 += part_width
    y0 += 20
    boldManualTitle(page, manual)


# Finds the Title in the verbiage and bolds it
def boldManualTitle(page, manual):
    if manual == 'Residual Market Manuals':
        manualTitle = "Residual Market Manual."
    elif manual == 'Assigned Risk':
        manualTitle = 'Assigned Risk Supplement.'
    elif manual == "Basic Manual User's Guide":
        manualTitle = "Basic Manual User's Guide."
    elif manual == "Basic Manual" or manual == "Scopes® Manuals":
        manualTitle = "Basic Manual."
    elif manual == "Experience Rating Plans Manual":
        manualTitle = "Experience Rating Plan Manual."
    elif manual == "Filing Guide":
        manualTitle = "Filing Guide."

    # Make variable based on Manual Name
    title = page.search_for(manualTitle)

    for inst in title:
        # Redact the original text
        page.add_redact_annot(inst)
        page.apply_redactions()

        # Register the bold Italic font
        bold_ital_font_path = 'C:\\Windows\\Fonts\\Calibriz.ttf'
        page.insert_font(fontfile=bold_ital_font_path, fontname="calibri-bold-ital")

        # Insert the text with the bold font
        x0, y0, x1, y1 = inst
        if manual == "Experience Rating Plans Manual":
            page.insert_text((x0, 485.5), manualTitle, fontsize=9, fontname="calibri-bold-ital", color=(0, 0, 0))
        elif manual == "Scopes® Manuals":
            page.insert_text((x0, 483.7), manualTitle, fontsize=9, fontname="calibri-bold-ital", color=(0, 0, 0))
        else:
            page.insert_text((x0, 469.36), manualTitle, fontsize=9, fontname="calibri-bold-ital", color=(0, 0, 0))




