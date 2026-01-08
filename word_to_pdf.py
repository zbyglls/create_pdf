from docx2pdf import convert
import os

output_folder = "output"
for template_name in os.listdir(output_folder):
    if template_name.endswith('.docx'):
        template_file = os.path.join(output_folder, template_name)
        output_pdf_file = os.path.join(output_folder, f"{template_name.split('.')[0]}.pdf")
        convert(template_file, output_pdf_file)
