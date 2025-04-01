# Program Files need to be here C:\\Users\\CSFIM\\PDF-FOR_MULTISTATE-PDF.py
# To Run in Terminal: python main.py
# Multiply coordinates in PDF by 72 to get coordinates for code

import pymupdf
from coverPage.disclaimer.redactDisclaimer import redactDisclaimer
from coverPage.effectiveDate.redactEffectiveSince import redactEffectiveSince
from coverPage.disclaimer.handleStates import findAndRedactStates
from coverPage.disclaimer.moveParasUp import moveParasUp
from copyrightDate.copyrightDate import unspanCopyrightDates

# Handles the process of updating PDF
def pdf_for_multistate_pdf(input_pdf, output_pdf):
    # Open the input PDF
    doc = pymupdf.open(input_pdf)
    page = doc[0]

    # Gets the Manual you are modifying -- The order of these manuals is important. Please do not move.
    manuals= [ "Scopes® Manuals", 'Residual Market Manual', 'Assigned Risk', "Basic Manual User's Guide", "Basic Manual", "Experience Rating Plans Manual", "Electronic Transmission User's Guide", "Filing Guide"]
    manualResult = [page.search_for(manualName)for manualName in manuals]

    for manualXY in manualResult:
        if manualXY:
            manualName = page.get_text("text", clip=manualXY[0]).strip()
            print("manualName: ", manualName)
            if manualName:
                break
            
    if manualName != "Electronic": 
        redactDisclaimer(page, manualName)  

        redactEffectiveSince(page, manualName)
 
        findAndRedactStates(page, manualName)

        moveParasUp(page, manualName)

    
    unspanCopyrightDates(doc)

    # Save the output PDF
    doc.save(output_pdf)

# Call the function to run with your parameters
pdf_for_multistate_pdf(
    'C:\\Users\\CSFIM\\OneDrive - NCCI Holdings, Inc\\Desktop\\CO Notes\\Apps in Draft\\PDF-FOR-MULTISTATE-PDF\\ARS\\AR-MASTER-PUB.pdf',
    'C:\\Users\\CSFIM\\OneDrive - NCCI Holdings, Inc\\Desktop\\CO Notes\\Apps in Draft\\PDF-FOR-MULTISTATE-PDF\\ARS\\Test.pdf',
    )

