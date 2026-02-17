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
        # ONLY PRINT HEADER ON PAGE 1
        if self.page_no() == 1:
            # 1. LOGO AND UNIVERSITY HEADER
            header_path = "C:/Users/Ganesh/Downloads/LPG/header.png"
            
            if os.path.exists(header_path):
                self.image(header_path, x=10, y=8, w=190)
                # --- SPACE SAVER 1: Reduced from 45 to 25 ---
                self.ln(34) 
            else:
                self.set_text_color(60, 120, 60) 
                self.set_font('Times', 'B', 16)
                self.cell(0, 5, 'KERALA AGRICULTURAL UNIVERSITY', 0, 1, 'C')
                self.ln(10)

            # 2. PROFESSOR DETAILS (Top Right)
            self.set_text_color(0, 0, 0) # Reset to black
            self.set_font('Times', 'B', 11)
            
            right_margin_start = 110
            self.set_xy(right_margin_start, self.get_y())
            self.cell(0, 5, 'Dr. Archana Raghavan Sathyan', 0, 1, 'R')
            
            self.set_font('Times', 'I', 11)
            self.set_xy(right_margin_start, self.get_y())
            self.cell(0, 5, 'DAAD Student Ambassador', 0, 1, 'R')
            
            self.set_xy(right_margin_start, self.get_y())
            self.cell(0, 5, 'Assistant Professor', 0, 1, 'R')
            
            self.set_xy(right_margin_start, self.get_y())
            self.cell(0, 5, 'Department of Agricultural Extension Education', 0, 1, 'R')
            
            # Reduce gap before title
            self.ln(2) 
        else:
            # For Page 2 onwards, just add a simple top margin
            self.ln(20)

    def footer(self):
        # Position at 1.5 cm from bottom
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.set_text_color(0, 0, 0)
        # Page number
        self.cell(0, 10, 'Page ' + str(self.page_no()) + '/{nb}', 0, 0, 'C')

# --- MAIN APP LOGIC ---

uploaded_file = st.file_uploader("Choose your Word file", type="docx")

if uploaded_file is not None:
    # 1. Read the Word Document
    doc = Document(uploaded_file)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    
    # 2. Generate PDF
    pdf = PDF()
    pdf.alias_nb_pages()
    pdf.add_page()
    # --- SPACE SAVER 2: Allow text closer to bottom (Margin 15 instead of 25) ---
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Title
    pdf.set_font('Times', 'B', 12)
    pdf.set_text_color(0, 0, 0) 
    pdf.cell(0, 10, 'LETTER OF RECOMMENDATION', 0, 1, 'C')
    pdf.ln(2) # Reduced title gap
    
    # Body Text
    # --- SPACE SAVER 3: Smaller Font (11pt) ---
    pdf.set_font('Times', '', 11)
    
    # Write paragraphs
    for para in full_text:
        if para.strip(): 
            # --- SPACE SAVER 4: Tighter Lines (Height 5 instead of 6) ---
            pdf.multi_cell(0, 5, para)
            # --- SPACE SAVER 5: Smaller Paragraph Gap (2 instead of 4) ---
            pdf.ln(2) 
            
    # 3. Dynamic Footer Section
    pdf.ln(5) # Reduced space before signature
    
    now = datetime.datetime.now()
    date_str = now.strftime("%d-%m-%Y") 
    
    # Check if we need a page break to keep signature together
    # Adjusted check value because we changed margins
    if pdf.get_y() > 250:
        pdf.add_page()
        
    y_pos = pdf.get_y()
    
    # --- LEFT SIDE: Place and Date ---
    pdf.set_xy(10, y_pos)
    pdf.cell(50, 5, 'Vellayani', 0, 1, 'L')
    
    pdf.set_xy(10, y_pos + 6) 
    pdf.cell(50, 5, date_str, 0, 1, 'L')
    
    # --- RIGHT SIDE: Yours sincerely & Name ---
    pdf.set_xy(10, y_pos) 
    pdf.cell(0, 5, 'Yours sincerely,', 0, 1, 'R')
    
    # Move down for Name
    pdf.set_xy(10, y_pos + 12) 
    pdf.cell(0, 5, 'Archana Raghavan Sathyan', 0, 1, 'R')

    # 4. Output
    pdf_output = pdf.output(dest='S').encode('latin-1') 
    
    st.success("PDF Generated Successfully!")
    
    st.download_button(
        label="Download Final PDF",
        data=pdf_output,
        file_name=f"Recommendation_{date_str}.pdf",
        mime="application/pdf"
    )