import os
import win32com.client
import PyPDF2
import tkinter as tk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF
#print(dir(fitz))
from pathlib import Path


def explore_pdf_coordinates(pdf_path, page_num=0):
    """Display all text with coordinates on a specific page."""
    doc = fitz.open(pdf_path)
    page = doc.load_page(page_num)

    print(f"\nText on page {page_num + 1} with coordinates:")
    print("(x0, y0, x1, y1) 'text'")
    print("----------------------------------")

    for block in page.get_text("blocks"):
        rect = block[:4]  # The rectangle coordinates
        text = block[4]  # The text content
        print(f"{rect} '{text.strip()}'")

    doc.close()


def extract_text_by_coordinates(pdf_path, coordinates):
    """
    Extract text from a PDF based on specific coordinates.

    Parameters:
    - pdf_path: Path to the PDF file
    - coordinates: A dictionary containing page number and rectangle coordinates
                   Example: {'page': 0, 'x0': 100, 'y0': 100, 'x1': 200, 'y1': 200}

    Returns:
    - Extracted text within the specified rectangle
    """
    doc = fitz.open(pdf_path)
    page = doc.load_page(coordinates['page'])

    # Create a rectangle object
    rect = fitz.Rect(coordinates['x0'], coordinates['y0'],
                     coordinates['x1'], coordinates['y1'])

    # Extract text within the rectangle
    text = page.get_text("text", clip=rect)

    doc.close()
    return text.strip()

def delete_pdf_files(pdf_files):
    """Deletes all specified PDF files"""
    deleted_count = 0
    for pdf_file in pdf_files:
        try:
            os.remove(pdf_file)
            deleted_count += 1
            print(f"Deleted: {os.path.basename(pdf_file)}")
        except Exception as e:
            print(f"Error deleting {pdf_file}: {e}")
    return deleted_count

# Example usage



# Usage:
#explore_pdf_coordinates("example.pdf", page_num=0)


# Set up the file dialog
root = tk.Tk()
root.withdraw()  # Hide the main window

# Ask for PDF file directory
directory_path = filedialog.askdirectory(
    title="Select a directory",
    initialdir="/"  # You can change the starting directory
)




if directory_path:  # If user didn't cancel
    print("Selected directory:", directory_path)
else:
    print("No directory selected")
# Get all PDF files in the directory
pdf_files = [f for f in os.listdir(directory_path) if f.lower().endswith('.pdf')]
renamed_files = []
modified_files=[]


#print("Readable?", os.access(pdf_files[0], os.R_OK))
#print("Writable?", os.access(pdf_files[0], os.W_OK))

#delete_pdf_files(pdf_files)


# Sort files alphabetically for consistent ordering
pdf_files.sort()

prefix="certificate_"
start_number=1


for i, filename in enumerate(pdf_files, start=start_number):
        # Create new filename
        new_name = f"{prefix}{i}.pdf"

# Get full paths
        old_path = os.path.join(directory_path, filename)
        new_path = os.path.join(directory_path, new_name)
# Get full paths
        old_path = os.path.join(directory_path, filename)
        new_path = os.path.join(directory_path, new_name)
# Rename the file
        os.rename(old_path, new_path)
        renamed_files.append(new_path)

# Delete/Clear the certificate
        doc = fitz.open(new_path)
        first_page = doc.load_page(0)
        full_text = first_page.get_text()
        text_to_remove = "НЕДЕЙСТВИТЕЛЬНЫЙ ДОКУМЕНТ!!!!!!!!!"
        text_to_remove_1 = "INVALID DOCUMENT !!!!!!!!!!!!!!!!!"
        text_to_remove_2 = "!!!!! INVALID DOCUMENT !!!!!"
        contract_number = "GG3023"
# Text removal
        for page in doc:
            # Search for the text
            text_instances = page.search_for(text_to_remove)
            text_instances_1 = page.search_for(text_to_remove_1)
            text_instances_2 = page.search_for(text_to_remove_2)
            text_instances_3 = page.search_for(contract_number)
            # Add redaction annotations (white rectangles) over each found text
            for inst in text_instances:
                redaction = page.add_redact_annot(inst, fill=(1, 1, 1))  # White fill
                redaction.update()

            # Apply the redactions
            page.apply_redactions()
            for inst in text_instances_1:
                redaction = page.add_redact_annot(inst, fill=(1, 1, 1))  # White fill
                redaction.update()

            # Apply the redactions
            page.apply_redactions()

            for inst in text_instances_2:
                redaction = page.add_redact_annot(inst, fill=(1, 1, 1))  # White fill
                redaction.update()

            # Apply the redactions
            page.apply_redactions()

        # Create output path (same directory with "_modified" suffix)
        directory, filename = os.path.split(new_path)
        name, ext = os.path.splitext(filename)
        output_path = os.path.join(directory, f"{name}_modified{ext}")
        output_pdf = output_path
        modified_files.append(output_path)


        doc.save(output_path)
        doc.close()
        print(f"Renamed: {filename} -> {new_name}")





#Extract contract number form certificate
target_area_1 = {
    'page': 0,  # Page number (0-based index)
    'x0': 467.760009765625,  # Left coordinate
    'y0': 186.21563720703125,  # Top coordinate
    'x1': 510.96746826171875,  # Right coordinate
    'y1': 199.80014038085938  # Bottom coordinate
}

target_area_2 = {
    'page': 0,  # Page number (0-based index)
    'x0': 79.31999969482422,  # Left coordinate
    'y0': 353.4956359863281,  # Top coordinate
    'x1': 136.9303436279297,  # Right coordinate
    'y1': 367.08013916015625  # Bottom coordinate
}

extracted_text_1 = extract_text_by_coordinates(modified_files[0], target_area_1)
extracted_text_2 = extract_text_by_coordinates(modified_files[0], target_area_2)
text_3= "сертификаты на вагон"

"""Create and display an Outlook email with attachment"""

outlook = win32com.client.Dispatch('Outlook.Application')
mail = outlook.CreateItem(0)  # 0 = olMailItem
mail.Subject = f"{extracted_text_1} {text_3} {extracted_text_2}"
#attachment_path = output_pdf
#pdf_files = [f for f in os.listdir(directory_path) if f.lower().endswith('.pdf')]
#if os.path.exists(attachment_path):
    #mail.Attachments.Add(attachment_path)

for pdf_file in modified_files:
    file_path = os.path.join(directory_path, pdf_file)
    mail.Attachments.Add(file_path)
mail.Display(True)  # Display the email (True keeps it open)



results = []



    # Search for the text


for inst in text_instances_3:
    results.append({
        'page': + 1,  # 1-based page numbering
        'x0': inst.x0,
        'y0': inst.y0,
        'x1': inst.x1,
        'y1': inst.y1,
        'width': inst.width,
        'height': inst.height,
        'text': contract_number
    })

positions = results
#rect = fitz.Rect(results['x0'], results['y0'], results['x1'], results['y1'])
#text = page.get_text("text", clip=rect)
#print("Номер контракта нашелся!!!!!", text)
#mail_subject = page.get_text("text", clip=positions)

#print("НОМЕР КОНТРАКТА", mail_subject)

if positions:

    for i, pos in enumerate(positions, 1):
        print(f"\nLocation {i}:")
        print(f"Page: {pos['page']}")
        print(f"Coordinates (x0,y0,x1,y1): ({pos['x0']:.2f}, {pos['y0']:.2f}, {pos['x1']:.2f}, {pos['y1']:.2f})")
        print(f"Width: {pos['width']:.2f}, Height: {pos['height']:.2f}")
else:
    print(f"Text '{contract_number}' not found in the PDF.")

explore_pdf_coordinates(modified_files[0], page_num=0)
#(467.760009765625, 186.21563720703125, 510.96746826171875, 199.80014038085938) 'GG3023'


    # Define the coordinates of the area you want to extract
    # You can get these coordinates using PDF tools or by trial and error

#print(f"Extracted text: '{extracted_text}'")

#clean the directory
pdf_files = [f for f in os.listdir(directory_path) if f.lower().endswith('.pdf')]
print(pdf_files)
#delete_pdf_files(pdf_files)

#input("Press Enter to exit...")
