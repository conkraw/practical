import streamlit as st
import pandas as pd
from docx import Document
import io
import zipfile
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH


def generate_document(student_name, code, exam_number):
    """
    Generate a Word document for a given exam number.
    """
    doc = Document()

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(11)
    
    # Add heading for the specific practical examination
    
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(f'Pediatric Clerkship Practical Examination #{exam_number}')
    run.bold = True
    run.font.all_caps = True  # This will display the text in all uppercase
    run.underline = True  
        
    # Add a paragraph with the student's legal name (with the name in bold)
    paragraph = doc.add_paragraph()
    paragraph.add_run("Student Name: ")
    run = paragraph.add_run(str(student_name))
    run.bold = True

    # Add access instructions heading
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run("Practical Examination Access Instructions")
    run.bold = True
    run.underline = True
    
    # Add website access instruction
    doc.add_paragraph("1. Please first enter the following website: https://redcap.ctsi.psu.edu/surveys")
    
    # Add code entry instruction with the code in bold
    paragraph = doc.add_paragraph()
    paragraph.add_run("2. Please enter the following code ")
    run = paragraph.add_run(str(code))
    run.bold = True
    
    # Add the main instructions text (common for both exams)
    instructions = (
        "Welcome to the Pediatric Clerkship Practical Examination!\n\n"
        "This examination provides a unique opportunity to demonstrate your skills in biomedical knowledge, clinical reasoning, "
        "and systems-based thinking as you engage with simulated cases involving undifferentiated pediatric patients.\n\n"
        "You will be presented with prompts containing key details and asked to respond to specific questions. Please note that "
        "all answers are required. While the system will not prompt you if an answer is left blank, your responses are mandatory "
        "in order to proceed through the examination. Ensure that you fill in all required fields before moving forward (unless otherwise specified). "
        "The purpose of this summative assessment is to evaluate your ability to synthesize data, interpret findings, and apply critical thinking in clinical decision-making.\n\n"
        "All essential resources, including immunization schedules and pharmacologic references, will be provided to assist you in formulating your responses. "
        "While you are welcome to take notes during the assessment, please note that these notes cannot be taken out of the exam room.\n\n"
        "Learning Objectives\n"
        "Upon achieving an acceptable score on this examination, you will have demonstrated the following key learning objectives of the Pediatric Clerkship Course:\n\n"
        "• Obtain, synthesize, and interpret comprehensive medical histories for newborns, children, and adolescents in various clinical settings. (Patient Care 1.1, EPA 1)\n"
        "• Create an assessment, problem list, differential diagnosis, and management plan for common pediatric complaints. (Patient Care 1.2, EPA 2)\n"
        "• Integrate and adapt knowledge of growth and development to develop individualized pediatric care strategies that address physical, emotional, and psychosocial needs. (Patient Care 1.2)\n"
        "• Analyze social determinants of health, evaluate their impact on pediatric health outcomes, and formulate a plan to address these needs effectively. (Health Humanities 7.1, Systems Based Practice 6.4)\n\n"
        "Good luck, and let this be an opportunity to demonstrate the knowledge, skills, and professionalism you have cultivated throughout your clerkship."
    )
    
    # Add access instructions heading
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run("Practical Examination Test Taking Instructions")
    run.bold = True
    run.underline = True
    doc.add_paragraph(instructions)
    
    return doc

st.title("Pediatric Clerkship Practical Examination Document Generator")

# File uploader widget
uploaded_file = st.file_uploader("Upload your CSV file", type="csv")

if uploaded_file is not None:
    # Read the CSV file into a DataFrame
    df = pd.read_csv(uploaded_file)
    st.write("Data Preview:", df.head())
    
    # Create in-memory ZIP files for each exam set
    zip_buffer_p1 = io.BytesIO()
    zip_buffer_p2 = io.BytesIO()
    
    # Open both ZIP files for writing
    with zipfile.ZipFile(zip_buffer_p1, "a", zipfile.ZIP_DEFLATED, False) as zip_file_p1, \
         zipfile.ZipFile(zip_buffer_p2, "a", zipfile.ZIP_DEFLATED, False) as zip_file_p2:
        # Process each row in the dataset
        for index, row in df.iterrows():
            student_name = row['legal_name']
            code_p1 = row['code_p1']
            code_p2 = row['code_p2']
            
            # Generate document for Practical Examination #1
            doc1 = generate_document(student_name, code_p1, exam_number=1)
            doc1_io = io.BytesIO()
            doc1.save(doc1_io)
            doc1_io.seek(0)
            filename_p1 = f"{student_name}_p1.docx"
            zip_file_p1.writestr(filename_p1, doc1_io.read())
            
            # Generate document for Practical Examination #2
            doc2 = generate_document(student_name, code_p2, exam_number=2)
            doc2_io = io.BytesIO()
            doc2.save(doc2_io)
            doc2_io.seek(0)
            filename_p2 = f"{student_name}_p2.docx"
            zip_file_p2.writestr(filename_p2, doc2_io.read())
    
    # Reset buffers for download
    zip_buffer_p1.seek(0)
    zip_buffer_p2.seek(0)
    
    # Provide download buttons for both ZIP files
    st.download_button(
        label="Download All Exam #1 Documents as ZIP",
        data=zip_buffer_p1,
        file_name="generated_documents_p1.zip",
        mime="application/zip"
    )
    
    st.download_button(
        label="Download All Exam #2 Documents as ZIP",
        data=zip_buffer_p2,
        file_name="generated_documents_p2.zip",
        mime="application/zip"
    )
