import os
import pandas as pd
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Mm
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.shared import Inches
from PIL import Image, ImageDraw, ImageFont
import io

def use_external_logo(logo_path='resources/logo.png', src_image=None):
    """Use an external logo image or create one if not provided."""
    # If the source image is provided, save it to the logo path
    if src_image and os.path.exists(src_image):
        # Ensure directory exists
        os.makedirs(os.path.dirname(logo_path), exist_ok=True)
        
        # Copy the image
        try:
            from shutil import copyfile
            copyfile(src_image, logo_path)
            return logo_path
        except Exception as e:
            print(f"Error copying logo: {e}")
    
    # Create a logo if one doesn't exist
    return create_circular_logo(logo_path)

def create_circular_logo(output_path='resources/logo.png'):
    """Create a circular logo image similar to the one in the reference."""
    # Check if logo already exists
    if os.path.exists(output_path):
        return output_path
    
    # Ensure directory exists
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    # Create a circular logo with a silhouette
    size = (300, 300)
    img = Image.new('RGBA', size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    
    # Draw golden circle
    circle_color = (192, 155, 85, 255)  # Golden color
    draw.ellipse((0, 0, 300, 300), fill=circle_color)
    
    # Draw white silhouette (simplified)
    white = (255, 255, 255, 255)
    # Draw a meditation pose silhouette (simplified)
    # Head
    draw.ellipse((125, 60, 175, 110), fill=white)
    # Shoulders
    draw.ellipse((110, 110, 130, 130), fill=white)
    draw.ellipse((170, 110, 190, 130), fill=white)
    # Body
    draw.rectangle((135, 110, 165, 180), fill=white)
    # Arms
    draw.polygon([(110, 130), (70, 160), (85, 180), (130, 140)], fill=white)
    draw.polygon([(190, 130), (230, 160), (215, 180), (170, 140)], fill=white)
    # Legs (folded position)
    draw.ellipse((110, 180, 190, 220), fill=white)
    
    img.save(output_path)
    return output_path

def add_border_to_table(table, color="C09B55", size=12):
    """Add a border to a table."""
    tbl = table._tbl
    for cell in table._cells:
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = OxmlElement('w:tcBorders')
        
        # Make border 3 times thicker than lines between fields
        for border_name in ['top', 'left', 'bottom', 'right']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), str(size))  # Border width in 1/8 points
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), color)
            tcBorders.append(border)
        
        tcPr.append(tcBorders)
        
        # Add background color (very light cream/beige to match image)
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), 'FFF9E7')  # Light cream color
        shd.set(qn('w:val'), 'clear')
        tcPr.append(shd)

def add_horizontal_line(paragraph, color="C09B55"):
    """Add a horizontal line to a paragraph."""
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    
    # Horizontal line thickness (1/3 of the border thickness)
    border = OxmlElement('w:bottom')
    border.set(qn('w:val'), 'single')
    border.set(qn('w:sz'), '4')  # Border width 
    border.set(qn('w:space'), '1')
    border.set(qn('w:color'), color)
    
    pBdr = OxmlElement('w:pBdr')
    pBdr.append(border)
    pPr.append(pBdr)

def create_card(doc, data, card_width_pt, logo_path):
    """Create a single card with the given data."""
    # Create a table for the card layout with border
    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    table.allow_autofit = False
    
    # Set table dimensions and column widths (65% left column, 35% right column)
    left_width = card_width_pt * 0.55
    right_width = card_width_pt * 0.45
    
    table.columns[0].width = Pt(left_width)
    table.columns[1].width = Pt(right_width)
    
    # Add thicker golden border to the table (3x the field lines)
    add_border_to_table(table, color="C09B55", size=12)
    
    # Get the left column for text content
    left_cell = table.cell(0, 0)
    right_cell = table.cell(0, 1)
    
    # Set cell margins (reduced for compactness)
    for cell in [left_cell, right_cell]:
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcMar = OxmlElement('w:tcMar')
        for margin_name, margin_val in [('top', 60), ('left', 80), ('bottom', 60), ('right', 80)]:
            node = OxmlElement(f'w:{margin_name}')
            node.set(qn('w:w'), str(margin_val))
            node.set(qn('w:type'), 'dxa')
            tcMar.append(node)
        tcPr.append(tcMar)
    
    # Remove cell borders between columns
    for row in table.rows:
        for cell in row.cells:
            # Get cell properties
            tc_pr = cell._tc.get_or_add_tcPr()
            
            # Remove vertical borders between cells
            for border_side in ['left', 'right']:
                border_elem = OxmlElement(f'w:{border_side}')
                border_elem.set(qn('w:val'), 'nil')
                
                # Check if tcBorders element already exists
                tc_borders = None
                for child in tc_pr:
                    if child.tag.endswith('tcBorders'):
                        tc_borders = child
                        break
                
                # If tcBorders doesn't exist, create it
                if tc_borders is None:
                    tc_borders = OxmlElement('w:tcBorders')
                    tc_pr.append(tc_borders)
                
                # Replace or add border element
                for i, child in enumerate(tc_borders):
                    if child.tag.endswith(border_side):
                        tc_borders[i] = border_elem
                        break
                else:
                    tc_borders.append(border_elem)
    
    # Define gold color to use consistently
    gold_color = RGBColor(192, 155, 85)
    
    # LEFT COLUMN CONTENT
    # Reduced header size by 30% (from Pt(12) to Pt(8.4))
    # Reduced spacing between fields by 50% (from Pt(3) to Pt(1.5))
    header_font_size = Pt(8.5)
    data_font_size = Pt(8)
    name_font_size = Pt(10)
    field_spacing = Pt(1.5)
    
    # Title (LAABHARTHI NAME)
    title_para = left_cell.paragraphs[0]
    title_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    title_run = title_para.add_run("LAABHARTHI NAME")
    title_run.bold = True
    title_run.font.size = header_font_size
    title_run.font.color.rgb = gold_color
    title_para.paragraph_format.space_after = Pt(0)
    
    # Add name data under title
    name_data_para = left_cell.add_paragraph()
    name_data_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    name_data_run = name_data_para.add_run(data['LAABHARTHI_NAME'])
    name_data_run.font.size = name_font_size
    name_data_run.font.color.rgb = gold_color  # Same gold color as headings
    name_data_para.paragraph_format.space_after = Pt(0)
    
    # Add horizontal line after name
    name_line = left_cell.add_paragraph()
    add_horizontal_line(name_line)
    name_line.paragraph_format.space_after = field_spacing
    
    # Contact Number label
    contact_label = left_cell.add_paragraph()
    contact_label.alignment = WD_ALIGN_PARAGRAPH.LEFT
    contact_run = contact_label.add_run("CONTACT NUMBER")
    contact_run.bold = True
    contact_run.font.size = header_font_size
    contact_run.font.color.rgb = gold_color
    contact_label.paragraph_format.space_after = Pt(0)
    
    # Add contact data
    contact_data_para = left_cell.add_paragraph()
    contact_data_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    contact_data_run = contact_data_para.add_run(data['CONTACT_NUMBER'])
    contact_data_run.font.size = data_font_size
    contact_data_run.font.color.rgb = gold_color  # Same gold color as headings
    contact_data_para.paragraph_format.space_after = Pt(0)
    
    # Add horizontal line after contact
    contact_line = left_cell.add_paragraph()
    add_horizontal_line(contact_line)
    contact_line.paragraph_format.space_after = field_spacing
    
    # Arpit Group label
    group_label = left_cell.add_paragraph()
    group_label.alignment = WD_ALIGN_PARAGRAPH.LEFT
    group_run = group_label.add_run("ARPIT GROUP (if applicable)")
    group_run.bold = True
    group_run.font.size = header_font_size
    group_run.font.color.rgb = gold_color
    group_label.paragraph_format.space_after = Pt(0)
    
    # Add group data
    group_data_para = left_cell.add_paragraph()
    group_data_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    group_data_run = group_data_para.add_run(data['ARPIT_GROUP'])
    group_data_run.font.size = data_font_size
    group_data_run.font.color.rgb = gold_color  # Same gold color as headings
    group_data_para.paragraph_format.space_after = Pt(0)
    
    # Add horizontal line after group
    group_line = left_cell.add_paragraph()
    add_horizontal_line(group_line)
    group_line.paragraph_format.space_after = field_spacing
    
    # Area label
    area_label = left_cell.add_paragraph()
    area_label.alignment = WD_ALIGN_PARAGRAPH.LEFT
    area_run = area_label.add_run("AREA")
    area_run.bold = True
    area_run.font.size = header_font_size
    area_run.font.color.rgb = gold_color
    area_label.paragraph_format.space_after = Pt(0)
    
    # Add area data
    area_data_para = left_cell.add_paragraph()
    area_data_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    area_data_run = area_data_para.add_run(data['AREA'])
    area_data_run.font.size = data_font_size
    area_data_run.font.color.rgb = gold_color  # Same gold color as headings
    area_data_para.paragraph_format.space_after = Pt(0)
    
    # No horizontal line after the last field (area)
    
    # RIGHT COLUMN CONTENT (logo and celebrating text)
    # Add the logo image
    logo_para = right_cell.paragraphs[0]
    logo_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    logo_run = logo_para.add_run()
    
    # Insert logo image
    if os.path.exists(logo_path):
        # Smaller logo for better fit
        logo_run.add_picture(logo_path, width=Inches(1.5))
    
    # Create a container for the celebration text and amount
    celebration_container = right_cell.add_paragraph()
    celebration_container.alignment = WD_ALIGN_PARAGRAPH.LEFT
    celebration_container.paragraph_format.space_before = Pt(20)
    

    # Amount text aligned with Celebrating text (left aligned)
    amount_run = celebration_container.add_run("Amount Rs. 1000/-")
    amount_run.font.size = Pt(8)
    amount_run.bold = True
    amount_run.font.color.rgb = gold_color
    amount_run.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    return table

# def paragraph_format_run(cell):
#     paragraph = cell.paragraphs[0]
#     paragraph = cell.add_paragraph()
#     format = paragraph.paragraph_format
#     run = paragraph.add_run()
    
#     format.space_before = Pt(1)
#     format.space_after = Pt(10)
#     format.line_spacing = 1.0
#     format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
#     return paragraph, format, run

def generate_cards(csv_file, output_file, rows=4, cols=2, logo_path=None):
    """Generate a Word document with cards from CSV data."""
    # Generate or use existing logo
    if logo_path and os.path.exists(logo_path):
        final_logo_path = use_external_logo(src_image=logo_path)
    else:
        final_logo_path = create_circular_logo()
    
    # Read CSV file
    df = pd.read_csv(csv_file)
    
    # Convert all columns to string
    for col in df.columns:
        df[col] = df[col].astype(str)
    
    # Create document
    doc = Document()
    
    # Set page size to A4 and minimized margins
    section = doc.sections[0]
    section.page_height = Cm(29.7)  # A4 height
    section.page_width = Cm(21.0)   # A4 width
    section.left_margin = Cm(0.8)   # Further reduced margins
    section.right_margin = Cm(0.8)
    section.top_margin = Cm(0.8)
    section.bottom_margin = Cm(0.8)
    
    # Calculate available space on the page (in points)
    available_width = Pt(section.page_width.pt - section.left_margin.pt - section.right_margin.pt)
    available_height = Pt(section.page_height.pt - section.top_margin.pt - section.bottom_margin.pt)

    card_width_pt = (available_width.pt / cols) - 4  # 8pt gap between cards
    
    # Calculate card dimensions with minimal gaps between cards
    card_height_pt = (available_height.pt / rows) - 4  # 8pt gap between cards
    
    # Generate cards (8 per page = 4 rows x 2 columns)
    cards_per_page = rows * cols
    
    # Process each record in the DataFrame
    for i, record in enumerate(df.to_dict('records')):
        # Determine position on the page
        page_position = i % cards_per_page
        row = page_position // cols
        col = page_position % cols
        
        # Add page break if needed
        if page_position == 0 and i > 0:
            doc.add_page_break()
        
        # Create card
        card = create_card(doc, record, card_width_pt, final_logo_path)
        
        # Add minimal spacing after each card except for the last one on the page
        if page_position < cards_per_page - 1:
            spacing_para = doc.add_paragraph()
            if (page_position + 1) % cols == 0:  # If at the end of a row
                spacing_para.paragraph_format.space_after = Pt(4)  # Minimal space between rows
            else:
                spacing_para.paragraph_format.space_after = Pt(0)  # No space between columns
    
    # Save the document
    doc.save(output_file)
    print(f"Cards generated successfully and saved to {output_file}")

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description='Generate cards from CSV data')
    parser.add_argument('--csv', type=str, default='data/sample_data.csv', help='Path to CSV file')
    parser.add_argument('--output', type=str, default='output_cards.docx', help='Output Word document path')
    parser.add_argument('--rows', type=int, default=4, help='Number of rows per page')
    parser.add_argument('--cols', type=int, default=2, help='Number of columns per page')
    parser.add_argument('--logo', type=str, help='Path to custom logo image')
    
    args = parser.parse_args()
    
    # Ensure CSV file exists
    if not os.path.exists(args.csv):
        csv_path = os.path.join(os.path.dirname(__file__), args.csv)
        if not os.path.exists(csv_path):
            print(f"Error: CSV file not found at {args.csv} or {csv_path}")
            exit(1)
        args.csv = csv_path
        
    # Generate output path if it's relative
    if not os.path.isabs(args.output):
        args.output = os.path.join(os.path.dirname(__file__), args.output)
        
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.output)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    generate_cards(args.csv, args.output, args.rows, args.cols, args.logo) 