import streamlit as st
import pandas as pd
from pdf_populate import populate_pdf, read_csv, read_excel, zip_folder
import os
import glob
import zipfile
import shutil
import json
import re
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter


def convert_image_to_pdf(image_path, pdf_path):
    """
    Converts an image to pdf
    """
    # create a new PDF with Reportlab
    pdf_canvas = canvas.Canvas(pdf_path, pagesize=letter)

    # draw the image at x, y. I positioned the x,y to be where i like here
    pdf_canvas.drawImage(image_path, x=50, y=50, width=500, height=500) # Adjust position and size as needed

    # Save the PDF
    pdf_canvas.save()

def sanitize_filename(filename):
    """
    Sanitize the filename by removing or replacing special characters.
    """
    # Replace or remove special characters as needed
    # This is a simple example, you can customize it as per your requirements
    filename = re.sub(r'[^a-zA-Z0-9.\-_]', '', filename)
    return filename

## Landing page UI
def run_UI():
    """
    The main UI function to display the UI for the webapp
    """

    # Set the page tab title
    st.set_page_config(page_title="Fuqua Course Enrollment", page_icon="🤖", layout="wide")

    # Set the page title
    st.header("Fuqua Course Enrollment Forms")

    # Sidebar menu
    with st.sidebar:
        st.subheader("Student Authorization Forms Menu")

        # Data file uploader
        data_file = st.file_uploader("Upload the Student Info Excel File", type=["xlsx"], key="upload-1", accept_multiple_files=False)
        approval_images = st.file_uploader("Upload the Student Approval Screenshots Zip file", type=["zip"], key="upload-2", accept_multiple_files=False)

        # Process the document
        if st.button("Generate Forms ✨"):

            # Delete the files in the approval images folder
            if os.path.exists("./data/approval_images"):
                shutil.rmtree("./data/approval_images")

            # Make sure approval screenshots are uploaded
            if approval_images is not None:
                print("[INFO] Approval images uploaded")
                os.makedirs("./data/approval_images/", exist_ok=True)

                # Extract the files               
                with zipfile.ZipFile(approval_images, 'r') as z:
                    print(z.infolist())
                    for file_info in z.infolist():
                        # Sanitize each filename
                        sanitized_name = sanitize_filename(file_info.filename)
                        # Extract the file with the new sanitized filename
                        source = z.open(file_info.filename)
                        target = open(os.path.join("./data/approval_images/", sanitized_name), "wb")
                        with source, target:
                            shutil.copyfileobj(source, target)

            # Check if the data file is uploaded
            if data_file is not None:
                # Add a progress spinner
                with st.spinner("Processing"):
                    
                    # Read the data
                    data = read_excel(data_file)

                    # Delete the files in the results folder
                    files = glob.glob('./results/*')
                    for f in files:
                        os.remove(f)

                    # Loop over the data
                    for i, row in enumerate(data):
                        
                        # Path to pdf template
                        input_pdf_path = './data/full_template_fuqua.pdf'
                        # Path to output pdf. Unique Key: student-id, class number, course-schedule
                        output_pdf_path = f'results/{row["Duke Unique ID#"]}-{row[" Class Number #"]}-{row["Course Schedule"]}.pdf'
                        # Path to approval image
                        approval_image_path = None

                        # Check if approval image exists
                        if row["Professor Approval Screenshot"] != None and approval_images is not None:
                            # Get the approval image path
                            link_dict = json.loads(row["Professor Approval Screenshot"])
                            approval_image_path = f'./data/approval_images/{link_dict[0]["id"]}.pdf'

                            # Sanitize the filename
                            file_name = sanitize_filename(link_dict[0]["name"])
                            path_components = file_name.split(".")
                            print("Processing: ", file_name)
                            
                            if path_components[-1] != "pdf":
                                # Convert the image to pdf
                                convert_image_to_pdf(f'./data/approval_images/{approval_images.name[:-4]}{sanitize_filename(link_dict[0]["name"])}', f'./data/approval_images/{link_dict[0]["id"]}.pdf')
                            else:
                                approval_image_path=None
                        # Populate the PDF with the data
                        populate_pdf(input_pdf_path, output_pdf_path, row, approval_image_path)
                        print(f"[INFO] Generated PDF: {row['Full name']}")
                    
                    folder_path = "./results"  # Replace with your folder path

                    ## Zip the pdfs
                    zip_folder(folder_path, "Student_Forms")
                    
                    ## Download the zip file
                    with open("Student_Forms.zip", "rb") as fp:
                        btn = st.download_button(
                            label="Download Forms 📥",
                            data=fp,
                            file_name="Student_Forms.zip",
                            mime="application/zip"
                        )
        
    # Display the data
    if data_file is not None:
        df = pd.read_excel(data_file)
        #df = df.drop_duplicates(subset='Student ID#', keep='last')
        df.rename(columns={'Student ID#': 'Student ID', 'Full name':'Student Name', 'Duke e-mail address':"Email ID", "Approve/Reject":"Status", "Credit/Audit":"Enrollment Type"}, inplace=True)
        st.table(data=df[["Student Name", "Email ID", "Enrollment Type", "Status"]])

if __name__ == "__main__":
    run_UI()