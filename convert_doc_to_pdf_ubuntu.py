import os
import subprocess

# Path to the folder containing docx files
input_folder = '/home/anastasia.lurye/Documents/colman/123/'

# Loop through docx files in the folder
for filename in os.listdir(input_folder):
    if filename.endswith('.docx'):
        docx_path = os.path.join(input_folder, filename)
        pdf_path = os.path.splitext(docx_path)[0] + '.pdf'

        # Use unoconv to convert docx to pdf
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', docx_path])

# No need to close anything, libreoffice handles the conversion
'''
sudo apt-get install libreoffice
possibly:
sudo apt-get install python3-uno
or 
sudo apt-get install unoconv
'''