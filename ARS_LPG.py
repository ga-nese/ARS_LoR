import streamlit as st
from fpdf import FPDF
from docx import Document
import datetime
import os

# --- CONFIGURE THE PAGE ---
st.set_page_config(page_title="ARS LoR_LP", layout="centered")

st.title("ARS LoR LetterPad Plugin")
st.markdown("""
Upload a **Word Document (.docx)** containing the body of the letter. 
The app will generate a PDF on the official letterhead.
""")

# --- PDF CLASS DEFINITION ---
class PDF(FPDF):
    def header(self):
        if self.page_no() == 1:
            # 1. LOGO AND UNIVERSITY HEADER
            # Note: Ensure "header.png" is in your GitHub repo
            header_path = "header.png" 
            
            if os.path.exists(header_path):
                self.image(header_path, x=10, y=8, w=190)
                self.ln(34) 
            else:
                self.set_text_color(60, 120, 60) 
                self.set_font('Times', 'B', 16)
                self.cell(0, 5, 'KERALA AGRICULTURAL UNIVERSITY', 0, 1, 'C')
                self.ln(10)

            # 2. PROFESSOR DETAILS
            self.set_text_color(0, 0, 0) 
            self.set_font('Times', 'B', 11)
            
            right_margin_start = 110
            self.set_xy(right_margin_start, self.get_y())
            self.cell(0, 5, 'Dr. Archana Raghavan Sathyan', 0, 1, 'R')
            
            self.set_font('Times', 'I', 11)
            self.set_xy(right_margin_start, self.get_y())
            self.cell(0, 5, 'DAAD Research Ambassador', 0, 1, 'R')
            
            self.set_xy(right_margin_start, self.get_y())
            self.cell(0, 5, 'Assistant Professor', 0, 1, 'R')
            
            self.set_xy(right_margin_start, self.get_y())
            self.cell(0, 5, 'Department of Agricultural Extension Education', 0, 1, 'R')

            self.set_xy(right_margin_start, self.get_y())
            self.cell(0, 5, 'Email: archana.rs@kau.in', 0, 1, 'R')
            
            self.ln(2) 
        else:
            self.ln(20)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.set_text_color(0, 0, 0)
        self.cell(0, 10, 'Page ' + str(self.page_no()) + '/{nb}', 0, 0, 'C')

# --- MAIN APP LOGIC ---

uploaded_file = st.file_uploader("Choose your Word file", type="docx")

if uploaded_file is not None:
    # 1. Read the Word Document
    doc = Document(uploaded_file)
    full_text = []
    
    # --- NEW: TEXT CLEANING FUNCTION ---
    # This replaces fancy Word characters with basic ones to prevent crashes
    def clean_text(text):
        replacements = {
            "\u2018": "'",  # Left single quote
            "\u2019": "'",  # Right single quote
            "\u201c": '"',  # Left double quote
            "\u201d": '"',  # Right double quote
            "\u2013": "-",  # En dash
            "\u2014": "-"   # Em dash
        }
        for original, replacement in replacements.items():
            text = text.replace(original, replacement)
        # Final safety net: remove any other unknown characters
        return text.encode('latin-1', 'replace').decode('latin-1')

    for para in doc.paragraphs:
        # Clean the paragraph before adding it
        clean_para = clean_text(para.text)
        full_text.append(clean_para)
    
    # 2. Generate PDF
    pdf = PDF()
    pdf.alias_nb_pages()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Title
    pdf.set_font('Times', 'B', 12)
    pdf.set_text_color(0, 0, 0) 
    pdf.cell(0, 10, 'LETTER OF RECOMMENDATION', 0, 1, 'C')
    pdf.ln(2) 
    
    # Body Text
    pdf.set_font('Times', '', 11)
    
    # Write paragraphs
    for para in full_text:
        if para.strip(): 
            pdf.multi_cell(0, 5, para)
            pdf.ln(1) 
            
    # 3. Dynamic Footer Section
    pdf.ln(5) 
    
    now = datetime.datetime.now()
    date_str = now.strftime("%d-%m-%Y") 
    
    if pdf.get_y() > 250:
        pdf.add_page()
        
    y_pos = pdf.get_y()
    
    # --- LEFT SIDE ---
    pdf.set_xy(10, y_pos)
    pdf.cell(50, 5, 'Vellayani', 0, 1, 'L')
    
    pdf.set_xy(10, y_pos + 6) 
    pdf.cell(50, 5, date_str, 0, 1, 'L')
    
    # --- RIGHT SIDE ---
    pdf.set_xy(10, y_pos) 
    pdf.cell(0, 5, 'Yours sincerely,', 0, 1, 'R')
    
    pdf.set_xy(10, y_pos + 20) 
    pdf.cell(0, 5, 'Dr. Archana Raghavan Sathyan', 0, 1, 'R')

    # 4. Output (Updated to handle encoding errors gracefully)
    # We use 'latin-1' here because that is what FPDF uses internally.
    pdf_output = pdf.output(dest='S').encode('latin-1', 'ignore') 
    
    st.success("PDF Generated Successfully!")
    
    st.download_button(
        label="Download Final PDF",
        data=pdf_output,
        file_name=f"Recommendation_{date_str}.pdf",
        mime="application/pdf"
    )






