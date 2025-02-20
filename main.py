# Program Files need to be here C:\\Users\\CSFIM\\PDF-FOR_MULTISTATE-PDF.py
# To Run in Terminal: python main.py
# Multiply coordinates in PDF by 72 to get coordinates for code
# Scopes, Basic, RM, Filing Guide, (See Procedures)

import fitz  # PyMuPDF Library is used to work in the PDF

# This defines a function named pdf_for_multistate_pdf that takes five parameters: 
# 1. the input PDF file path = input_pdf
# 2. the output PDF file path = output_pdf
# 3. the copyright verbiage to be replaced = copyrightSpan
# 4. the new copyright verbiage = replaceCopyright
# 3. the disclaimer verbiage to delete = disclaimer
# 4. the Effective Since Verbiage to delete = EffectiveSince

def pdf_for_multistate_pdf(input_pdf, output_pdf, copyrightSpan, replaceCopyright, disclaimer, EffectiveSince):
    # Open the input PDF
    doc = fitz.open(input_pdf)

    # Coordinates for RM Manual content
    paras = [fitz.Rect(79.92, 480.24, 532.08, 517.68),
             fitz.Rect(79.92, 521.28, 409, 533.52),
             fitz.Rect(79.92, 539.28, 529.92, 564.48),
             fitz.Rect(79.92, 570.24, 531.36, 615.6),
             fitz.Rect(79.92, 621.36, 532.8, 709.92),
             fitz.Rect(79.92, 716.4, 522.72, 741.6),]

    states = [
    "AK", "AL", "AR", "AZ", "CA", "CO", "CT", "DC", "DE", "FL",
    "GA", "HI", "IA", "ID", "IL", "IN", "KS", "KY", "LA", "MA",
    "MD", "ME", "MI", "MN", "MO", "MS", "MT", "NC", "ND", "NE",
    "NH", "NJ", "NM", "NV", "NY", "OH", "OK", "OR", "PA", "PR",
    "RI", "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WI",
    "WV", "WY"
    ]
    page = doc[0]
    
    ################### DISCLAIMER #################################   
    # Search for the disclaimer text and redact it
    foundDisclaimer = page.search_for(disclaimer)
    for inst in foundDisclaimer:
        page.add_redact_annot(inst)
        page.apply_redactions()

    ################### EFFECTIVE DATE ################################
    # Search for "Effective Since" and redact it
    foundEffectiveSince = page.search_for(EffectiveSince)
    for inst in foundEffectiveSince:
        page.add_redact_annot(inst)
        page.apply_redactions() 

    ################### GET AND REDACT STATES ################################    
    statesXY1 = fitz.Rect(236.16, 493.91, 532.8, 505.44)
    statesXY2 = fitz.Rect(79.92, 504.72, 531.36, 514.8)

    statesFirstRow = page.get_text("text", clip=statesXY1)
    statesSecondRow = page.get_text("text", clip=statesXY2)

    page.add_redact_annot(statesXY1)
    page.apply_redactions() 

    page.add_redact_annot(statesXY2)
    page.apply_redactions() 

    ################### Cover Page Content ############################### 
    # Loop through the paras array
    # extract text from text box - Store it in paraText and redact the content
    # Determine where the content belongs and insert the content to coordinates based on coordinates   
    for para in paras:

        paraText = page.get_text("text", clip=para)

        page.add_redact_annot(para)
        page.apply_redactions()

        # Register the fonts on the page        
        normal_font_path = 'C:\\Windows\\Fonts\\Calibri.ttf'
        bold_font_path = 'C:\\Windows\\Fonts\\Calibrib.ttf'
        bold_ital_font_path = 'C:\\Windows\\Fonts\\Calibriz.ttf'
        page.insert_font(fontfile=normal_font_path, fontname="calibri")
        page.insert_font(fontfile=bold_font_path, fontname="calibri-bold")
        page.insert_font(fontfile=bold_ital_font_path, fontname="calibri-bold-ital")

        #Insert Text
        if para == paras[0]:
            # Insert First Row of States
            page.insert_text((80.64, 469.36), paraText, fontsize=9, fontname="calibri", color=(0, 0, 0))
            
            #split text by word, parts is an array of words in textbox
            parts = statesFirstRow.split()
            # Y coordinate for inserted text for first line of states
            y0 = 480.7

            #Calculate position of state abrev. used to determine spacing
            current_x = 237
            counter = 0
            #insert text dynamic color formatting
            for part in parts:
                part_width = fitz.get_text_length(part, fontsize=9)
                if part.strip(",") in states:
                    if counter == 0:
                        page.insert_text((237, y0), part, fontsize=9, fontname="calibri", color = (0, 0, 1)) #Blue color for State abbeviations
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
            x0, y0 = 80.64, 490.7
            current_x = 80.64
            counter = 0
            #insert text dynamic color formatting
            for part in parts:
                part_width = fitz.get_text_length(part, fontsize=9)
                if part.strip(",") in states:
                    if counter == 0:
                        page.insert_text((x0, y0), part, fontsize=9, fontname="calibri", color = (0, 0, 1)) #Blue color for State abbeviations
                        current_x += part_width
                        counter += 1
                    else:
                        page.insert_text((current_x, y0), part, fontsize=9, fontname="calibri", color = (0, 0, 1)) #Blue color for State abbeviations
                        current_x += part_width
                else:
                    page.insert_text((current_x, y0), part + " ", fontsize=9, fontname="calibri", color = (0, 0, 1)) #Blue color for "and"
                    current_x += part_width
            y0 += 20            
        elif para == paras[1]:
            page.insert_text((x0, y0), paraText, fontsize=9, fontname="calibri", color=(0, 0, 0))
            y0 += 20
        elif para == paras[2]:
            page.insert_text((x0, y0), paraText, fontsize=9, fontname="calibri", color=(0, 0, 0))
        elif para == paras[3]:
            page.insert_text((x0, 562), paraText, fontsize=9, fontname="calibri", color=(0, 0, 0))
        elif para == paras[4]:
            page.insert_text((x0, 616), paraText, fontsize=9, fontname="calibri-bold", color=(0, 0, 0))
        elif para == paras[5]:
            page.insert_text((x0, 710), paraText, fontsize=9, fontname="calibri-bold", color=(0, 0, 0))

############ Bold The Manual Name #######################

    # Make variable based on Manual Name
    foundManualTitle = page.search_for("Residual Market Manual.")

    for inst in foundManualTitle:
        # Redact the original text
        page.add_redact_annot(inst)
        page.apply_redactions()

        # Register the bold Italic font
        page.insert_font(fontfile=bold_ital_font_path, fontname="calibri-bold-ital")

        # Insert the text with the bold font
        x0, y0, x1, y1 = inst
        page.insert_text((x0, 469.36), "Residual Market Manual.", fontsize=9, fontname="calibri-bold-ital", color=(0, 0, 0))

###################### Copyright Span ########################
    # Seach for the copyright text and replace it
    foundCopyright = page.search_for(copyrightSpan) 

    # Loops through the text_instances found 
    for inst in foundCopyright:
        page = doc.load_page()
    
        # Redact the old text
        page.add_redact_annot(inst) #Puts annotation to be redacted
        page.apply_redactions()     #Redacts the content

        # Insert the new text
        # Gets the coordinates of the redacted Text and insert new content here
        x0, y0, x1, y1 = inst
        font_size = 7  # Adjust the font size as needed
        y_offset = 6.7  # Adjust the vertical position as needed
        page.insert_text((x0, y0 + y_offset), replaceCopyright, fontsize=font_size, fontname="calibri", color=(0, 0, 0))

    # Save the output PDF
    doc.save(output_pdf)

# Call the function to run with your parameters
pdf_for_multistate_pdf(
    'C:\\Users\\CSFIM\\OneDrive - NCCI Holdings, Inc\\Desktop\\CO Notes\\Apps in Draft\\PDF-FOR-MULTISTATE-PDF\\RM-MASTR-PUB_GoldCopy.pdf',           #Directory of File to be changed
    'C:\\Users\\CSFIM\\OneDrive - NCCI Holdings, Inc\\Desktop\\CO Notes\\Apps in Draft\\PDF-FOR-MULTISTATE-PDF\\RM-MASTR-PUB_GoldCopy_Updated.pdf',      #Directory of where your file will go and be named
    '© 2021–2025 National Council on Compensation Insurance, Inc. All Rights Reserved.',                                                                   #Content you want to be replaced    
    '© 2025 National Council on Compensation Insurance, Inc. All Rights Reserved.',                                                                        #Content you are replacing the above with
    'This is a compilation of all NCCI filed and approved state Residual Market Manuals effective since November 1, 2021, published February 13, 2025.',
    'Effective Since: November 1, 2021'   
    )

