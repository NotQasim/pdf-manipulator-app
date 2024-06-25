from docx2pdf import convert
from pypdf import PdfWriter


def convert_word_to_pdf(word_filename: str) -> None:
    '''Takes in the name of a Word document (.doc or .docx) and converts it into a pdf.
       Document must be in the same directory as main.py.
       Only works if Microsoft 365 is installed.
    '''
    convert(word_filename, f'{word_filename.split('.doc')[0]}.pdf')

    return None


def merge_pdfs(pdf_names: tuple[str]) -> None:
    '''Takes in a list of names of PDF files.
       Merges all files into a single PDF from left to right in the list.
       All files must be in the same directory as main.py.
    '''
    merger = PdfWriter()

    for pdf in pdf_names:
        merger.append(pdf)

    with open('merged_pdf.pdf', 'wb') as f:
        merger.write(f)
    
    return None
