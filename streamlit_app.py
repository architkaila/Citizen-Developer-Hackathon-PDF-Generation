import streamlit as st
import pandas as pd
from pdf_populate import populate_pdf, read_csv, read_excel, zip_folder
import os
import glob

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
        data_file = st.file_uploader("Upload the Student Info Excel File", type="xlsx", key="upload", accept_multiple_files=False)

        # Process the document
        if st.button("Generate Forms ✨"):
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
                        input_pdf_path = './data/full_template_fuqua.pdf'
                        output_pdf_path = f'results/{row["Full name"]}.pdf'
                        
                        # Populate the PDF with the data
                        populate_pdf(input_pdf_path, output_pdf_path, row)
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
        df = df.drop_duplicates(subset='Student ID#', keep='last')
        df.rename(columns={'Student ID#': 'Student ID', 'Full name':'Student Name', 'Duke e-mail address':"Email ID", "Approve/Reject":"Status", "Credit/Audit":"Enrollment Type"}, inplace=True)
        st.table(data=df[["Student Name", "Email ID", "Enrollment Type", "Status"]])

if __name__ == "__main__":
    run_UI()