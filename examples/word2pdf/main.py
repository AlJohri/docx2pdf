from docx2pdf import convert
import os
### Imports ####

def make_pdf_(input_file_path, open_=False):
    folder_name = os.path.dirname(input_file_path) ## Some Variable making for the functioning of the program
    filename_no_extension = os.path.basename(input_file_path.split('.')[0])
    output_filename = fr'{folder_name}\{filename_no_extension}.pdf'

    convert(input_file_path, output_filename) ## Conversion to PDF done here
    if open_ == True:
        os.system(fr'"{output_filename}"') ## opens the converted PDF file in the default PDF viewer if set to true


