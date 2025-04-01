import re
import os

# Finds Spanned Copyright Dates and Unspanns them leaving the most current year
def unspanCopyrightDates(doc):
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        
        # Define the regex pattern  
        if page_num == 0:
            pattern = r"© Copyright (\d{4})–(\d{4}) National Council on Compensation Insurance, Inc\. All Rights Reserved\."
        else:
            pattern = r"© (\d{4})–(\d{4}) National Council on Compensation Insurance, Inc\. All Rights Reserved\."
        
        

        # Extract the entire page text
        pageText = page.get_text("text").strip()

        # Search for the pattern in the copyright
        match = re.search(pattern, pageText)

        if match:
            # Extract second year
            second_year = match.group(2)

            # Replace the spanned date with the second year
            if page_num == 0:
                new_copyright_text = f"© Copyright {second_year} National Council on Compensation Insurance, Inc. All Rights Reserved."
            else:

                new_copyright_text = f"© {second_year} National Council on Compensation Insurance, Inc. All Rights Reserved."



            # Replace the text in the PDF
            rects = page.search_for(match.group(0))
            if rects:
                x, y, x0, y0 = rects[0]   #.6, 10.34

                page.add_redact_annot(rects[0])
                page.apply_redactions()

                # Register the fonts on the page        
                normal_font_path = 'C:\\Windows\\Fonts\\calibri.ttf'
                if not os.path.exists(normal_font_path):
                    raise FileNotFoundError(f"Font file not found: {normal_font_path}")
                page.insert_font(fontfile=normal_font_path, fontname="calibri")

                if page_num == 0:
                    font_size = 9
                else:
                    font_size = 7

                page.insert_text((x, y + 4), new_copyright_text, fontsize=font_size, fontname="calibri", color=(0, 0, 0))
