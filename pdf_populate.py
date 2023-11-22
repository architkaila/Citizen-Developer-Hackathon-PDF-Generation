import csv
import pdfrw
from datetime import date
from mapping import sheet_mapping
import openpyxl
from pdfrw import PdfDict, PdfObject
import shutil   
from pikepdf import Pdf

# Get today's date
today = date.today()

def read_csv(filename):
    """
    Reads a CSV file and returns a list of dictionaries

    Args:
        filename (str): The name of the CSV file to read

    Returns:    
        list: A list of dictionaries where each item in the list
    """
    with open(filename, 'r') as f:
        reader = csv.DictReader(f)
        for row in reader:
            yield row

def read_excel(filename):
    """
    Reads an Excel file and returns a list of dictionaries

    Args:
        filename (str): The name of the Excel file to read

    Yields:    
        dict: Each row in the Excel sheet as a dictionary where the keys are the column headers
    """
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook.active

    # Get the headers from the first row
    headers = [cell.value for cell in sheet[1]]

    for row in sheet.iter_rows(min_row=2, values_only=True):  # Start from the second row to skip headers
        yield dict(zip(headers, row))

    workbook.close()

def zip_folder(folder_path, output_filename):
    """Zip the contents of an entire folder."""
    shutil.make_archive(output_filename, 'zip', folder_path)


def populate_pdf(input_pdf_path, output_pdf_path, data_dict, approval_image_path=None):
    """
    Populates a PDF form with data from a dictionary

    Args:
        input_pdf_path (str): The input PDF form to populate
        output_pdf_path (str): The output PDF to save the data to
        data_dict (dict): A dictionary of data to map into the PDF form

    Returns:
        None
    """

    ## Read the PDF template
    template = pdfrw.PdfReader(input_pdf_path)

    ## Populate the PDF with the data
    for page in template.pages:

        ## Get the annotations
        annotations = page.get('/Annots')  # Using .get() to avoid KeyError
        if annotations is None:
            continue

        ## Loop over the annotations
        for annotation in annotations:

            ## Get the annotation name/object ID
            field_key = annotation.get('/T')

            ## Get the data from the mapping
            if field_key is not None:
                key = field_key[1:-1]  # Remove leading and trailing parentheses
                
                ## This is where we populate the PDF with the data
                if sheet_mapping[key] in data_dict:
                    annotation.update(pdfrw.PdfDict(AP=str(data_dict[sheet_mapping[key]]), V=str(data_dict[sheet_mapping[key]])) )
                
                ## This is where we populate the PDF with the data (session data)
                if key in ["fall_1", "fall_2", "spring_1", "spring_2"] and data_dict["Session"] == sheet_mapping[key]:
                    annotation.update(pdfrw.PdfDict(AP="Yes", V="Yes"))

                ## This is where we populate the PDF with the date
                if key in ["date", "date_2", "date_sign"]:
                    annotation.update(pdfrw.PdfDict(AP=str(data_dict[sheet_mapping[key]]).split()[0], V=str(data_dict[sheet_mapping[key]]).split()[0]))

                ## This is where we populate the PDF with the credit/audit data
                if key in ["credit", "audit"] and data_dict["Credit/Audit"] == sheet_mapping[key]:
                    annotation.update(pdfrw.PdfDict(AP="Yes", V="Yes"))

    ## Save the PDF with the data
    #template.Root.AcroForm.update(PdfDict(NeedAppearances=PdfObject("true")))

    if "/AcroForm" in template.Root:
        print("true")
        template.Root.AcroForm.update(PdfDict(NeedAppearances=PdfObject("true")))
    else:
        print("false")
        # If there's no AcroForm, add one with the /NeedAppearances flag set
        template.Root.update(PdfDict(AcroForm=PdfDict(NeedAppearances=PdfObject("true"))))

    
    # Create the final PDF writer
    final_pdf_writer = pdfrw.PdfWriter()

    # Add the template pages
    final_pdf_writer.addpages(template.pages)

    # Add the approval image pages if they exist
    if approval_image_path is not None:
        # Read the approval image
        reader = pdfrw.PdfReader(approval_image_path)
        
        # Add the approval image pages
        final_pdf_writer.addpages(reader.pages)

    # Write the PDF to disk
    temp_output_path = './data/temp_output.pdf'
    final_pdf_writer.write(temp_output_path)
    

    reader_final = pdfrw.PdfReader(temp_output_path)
    if "/AcroForm" in reader_final.Root:
        print("true")
        reader_final.Root.AcroForm.update(PdfDict(NeedAppearances=PdfObject("true")))
    else:
        print("false")
        # If there's no AcroForm, add one with the /NeedAppearances flag set
        reader_final.Root.update(PdfDict(AcroForm=PdfDict(NeedAppearances=PdfObject("true"))))

    
    final_pdf_writer = pdfrw.PdfWriter(output_pdf_path)
    final_pdf_writer.addpages(reader_final.pages)
    final_pdf_writer.write()

    with Pdf.open(temp_output_path) as pdf:
        pdf.generate_appearance_streams()
        pdf.save(output_pdf_path)

# Main function
def main():
    # Read the data from the file
    data = read_excel('./data/Fuqua Form Automation Excel.xlsx')
    
    for i, row in enumerate(data):
        input_pdf_path = './data/template_fuqua.pdf'
        output_pdf_path = f'results/{row["Full name"]}.pdf'
        
        populate_pdf(input_pdf_path, output_pdf_path, row)
        print(f"[INFO] Generated PDF: {row['Full name']}")

if __name__ == '__main__':
    main()