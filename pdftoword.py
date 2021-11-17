# importing required modules
import os
from natsort import natsort_keygen
from pdf2jpg import pdf2jpg
from docx.shared import Inches
from docx import Document
import shutil
import time
import sys

# pip install python-docx pdf2jpg natsort

NATSORT_KEY = natsort_keygen()

PAPER_SIZE = (8.5, 11)  # INCHES


clear = lambda: os.system('cls')
# Print iterations progress
# https://stackoverflow.com/questions/3173320/text-progress-bar-in-the-console
def printProgressBar (iteration, total, prefix = '', suffix = '', decimals = 1, length = 100, fill = '█', printEnd = "\r"):
    # created function to clear terminal
    clear()
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
        printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
    """
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    print(f'\r{prefix} |{bar}| {percent}% {suffix}', end = printEnd)
    # Print New Line on Complete
    if iteration == total: 
        print()
        
def convert_pdf_to_images(pdf_path: str, output_path: str):
    pdf2jpg.convert_pdf2jpg(pdf_path, output_path, pages="ALL")
    
def convert_image_to_document(output_path, name):
    # Compile and Save
    all_file_names = os.listdir(output_path)

    # Sorting files using natural sorting (natsort)
    all_file_names.sort(key=NATSORT_KEY)
    document = Document()
    
    #changing the page margins
    sections = document.sections
    printProgressBar(0, len(sections), prefix = 'Configuring word document:', suffix = 'Complete', length = 50)
    for i, section in enumerate(sections):
        section.top_margin = Inches(0)
        section.bottom_margin = Inches(0)
        section.left_margin = Inches(0)
        section.right_margin = Inches(0)
        printProgressBar(i, len(document.sections), prefix = 'Configuring word document sections:', suffix = 'Complete', length = 50)

    printProgressBar(0, len(all_file_names), prefix = 'Compiling images to word document:', suffix = 'Complete', length = 50)

    # Combining all exported pictures from the pdf into one word document.
    for i, path in enumerate(all_file_names):
        # This if statement is not needed, this is only for personal use case, it shouldn't bother the program.
        if path.endswith('.onetoc2'): continue 
        document.add_picture(output_path + path, width=Inches(PAPER_SIZE[0]))
        printProgressBar(i, len(all_file_names), prefix = 'Compiling images to word document:', suffix = 'Complete', length = 50)
        
    printProgressBar(0, 1, prefix = 'Saving word document:', suffix = 'Complete', length = 50)

    # Save new word doc, reading for editing!
    document.save(name.replace('.pdf', '.docx'))

    printProgressBar(1, 1, prefix = 'Saving word document:', suffix = 'Complete', length = 50)

# The whole conversion process
def convert(pdf_file_name):
    full_pdf_file_path: str = os.getcwd() + '/process/' + pdf_file_name.replace('.pdf', '.pdf_dir/')
    
    printProgressBar(0, 2, prefix = 'Creating temperary directories:', suffix = 'Complete', length = 50)

    # Create directory if it doesn't exist.
    if not os.path.isdir(os.getcwd() + '/process'): 
        os.mkdir(os.getcwd() + '/process')

    # Create another directory if it doesn't exist, which it most likely wont.
    printProgressBar(1, 2, prefix = 'Creating temperary directories:', suffix = 'Complete', length = 50)
    if not os.path.isdir((os.getcwd() + '/process/' + pdf_file_name.replace('.pdf', '.pdf_dir/'))):
        os.mkdir(os.getcwd() + '/process/' + pdf_file_name.replace('.pdf', '.pdf_dir/'))
    printProgressBar(2, 2, prefix = 'Creating temperary directories:', suffix = 'Complete', length = 50)
    
    # Initial call to print 0% progress
    printProgressBar(0, 1, prefix = 'Extracting Images from PDF:', suffix = 'Complete', length = 50)

    # Export images from pdf
    convert_pdf_to_images(os.getcwd() + '/' + pdf_file_name, os.getcwd() + '/process/')

    printProgressBar(1, 1, prefix = 'Extracting Images from PDF:', suffix = 'Complete', length = 50)

    # # Get exported images and compile into a word document
    convert_image_to_document(full_pdf_file_path, pdf_file_name)

    # Do some clean up, since these pictures can get large in size, and this does everything in CWD.
    printProgressBar(0, 1, prefix = 'Cleaning up:', suffix = 'Complete', length = 50)
    shutil.rmtree(os.getcwd() + '/process/')

    printProgressBar(1, 1, prefix = 'Cleaning up:', suffix = 'Complete', length = 50)

    clear()
    
    print('Finished!')
    time.sleep(3)
    sys.exit()

# Get the file name
file_name = sys.argv[-1].split('\\')[-1]
convert(file_name)
