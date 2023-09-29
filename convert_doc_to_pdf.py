import os
import comtypes.client

# Path to the folder containing docx files
input_folder = r'C:\Users\anast\OneDrive\Documents\College\Computation and Complexitity\exams_docs'

# Initialize Word application
word_app = comtypes.client.CreateObject("Word.Application")
word_app.Visible = 1  # Set to 0 for no visible window

# Loop through docx files in the folder
for filename in os.listdir(input_folder):
    if filename.endswith('.doc'):
        docx_path = os.path.join(input_folder, filename)
        pdf_path = os.path.splitext(docx_path)[0] + '.pdf'

        doc = word_app.Documents.Open(docx_path)
        doc.SaveAs(pdf_path, FileFormat=17)  # 17 represents PDF format
        doc.Close()

# Close Word application
word_app.Quit()
