import streamlit as st
import pandas as pd
from docx import Document
import io
import zipfile

def generate_document(student_name, code_p1, code_p2):
    # Create a new Word document
    doc = Document()
    
    # Add headings for both practical examinations
    doc.add_heading(f'Pediatric Clerkship Practical Examination #1 Instructions for {code_p1}', level=1)
    doc.add_heading(f'Pediatric Clerkship Practical Examination #2 Instructions for {code_p2}', level=1)
    
    # Optionally add a paragraph that shows the student's legal name
    doc.add_paragraph(f"Student Name: {student_name}")
    
    # Add the main instructions text
    instructions = (
        "Welcome to the Pediatric Clerkship Practical Examination!\n\n"
        "This examination provides a unique opportunity to demonstrate your skills in biomedical knowledge, clinical reasoning, and systems-based thinking as you engage with simulated cases involving undifferentiated pediatric patients.\n\n"
        "You will be presented with prompts containing key details and asked to respond to specific questions. Please note that all answers are required. While the system will not prompt you if an answer is left blank, your responses are mandatory in order to proceed through the examination. Ensure that you fill in all required fields before moving forward (unless otherwise specified). The purpose of this summative assessment is to evaluate your ability to synthesize data, interpret findings, and apply critical thinking in clinical decision-making.\n\n"
        "All essential resources, including immunization schedules and pharmacologic references, will be provided to assist you in formulating your responses. While you are welcome to take notes during the assessment, please note that these notes cannot be taken out of the exam room.\n\n"
        "Learning Objectives\n"
        "Upon achieving an acceptable score on this examination, you will have demonstrated the following key learning objectives of the Pediatric Clerkship Course:\n\n"
        "• Obtain, synthesize, and interpret comprehensive medical histories for newborns, children, and adolescents in various clinical settings. (Patient Care 1.1, EPA 1)\n"
        "• Create an assessment, problem list, differential diagnosis, and management plan for common pediatric complaints. (Patient Care 1.2, EPA 2)\n"
        "• Integrate and adapt knowledge of growth and development to develop individualized pediatric care strategies that address physical, emotional, and psychosocial needs. (Patient Care 1.2)\n"
        "• Analyze social determinants of health, evaluate their impact on pediatric health outcomes, and formulate a plan to address these needs effectively. (Health Humanities 7.1, Systems Based Practice 6.4)\n\n"
        "Good luck, and let this be an opportunity to demonstrate the knowledge, skills, and professionalism you have cultivated throughout your clerkship."
    )
    doc.add_paragraph(instructions)
    
    # Add instructions on how to access the exam website
    access_instructions = "Please first access https://redcap.ctsi.psu.edu/surveys and then enter your code."
    doc.add_paragraph(access_instructions)
    
    return doc

st.title("Pediatric Clerkship Practical Examination Document Generator")

# File uploader widget
uploaded_file = st.file_uploader("Upload your CSV file", type="csv")

if uploaded_file is not None:
    # Read the CSV file into a DataFrame
    df = pd.read_csv(uploaded_file)
    st.write("Data Preview:", df.head())
    
    # Create an in-memory ZIP file to store all documents
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        # Process each row in the dataset
        for index, row in df.iterrows():
            student_name = row['legal_name']
            code_p1 = row['code_p1']
            code_p2 = row['code_p2']
            
            # Generate the document for the current record
            doc = generate_document(student_name, code_p1, code_p2)
            
            # Save the document to an in-memory buffer
            doc_io = io.BytesIO()
            doc.save(doc_io)
            doc_io.seek(0)
            
            # Use the student's legal name as the filename with a .docx extension
            filename = f"{student_name}.docx"
            zip_file.writestr(filename, doc_io.read())
    
    # Reset buffer position for download
    zip_buffer.seek(0)
    
    # Provide a download button for the ZIP file
    st.download_button(
        label="Download All Documents as ZIP",
        data=zip_buffer,
        file_name="generated_documents.zip",
        mime="application/zip"
    )

