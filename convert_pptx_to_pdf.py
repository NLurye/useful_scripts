import os
import comtypes.client

# Path to the folder containing pptx files
input_folder = r'C:\Users\anast\OneDrive\temptemp'

# Initialize PowerPoint application
ppt_app = comtypes.client.CreateObject("PowerPoint.Application")
ppt_app.Visible = 1  # Set to 0 for no visible window

# Loop through pptx files in the folder
for filename in os.listdir(input_folder):
    if filename.endswith('.pptx'):
        ppt_path = os.path.join(input_folder, filename)
        pdf_path = os.path.splitext(ppt_path)[0] + '.pdf'

        presentation = ppt_app.Presentations.Open(ppt_path)
        presentation.SaveAs(pdf_path, 32)  # 32 represents PDF format
        presentation.Close()

# Close PowerPoint application
ppt_app.Quit()
