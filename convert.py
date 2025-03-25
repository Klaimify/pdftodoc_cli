from pdf2docx import Converter
import os

def convert_pdf_to_word(pdf_path):
    try:
        # Create output filename by replacing .pdf extension with .docx
        word_path = os.path.splitext(pdf_path)[0] + '.docx'
        
        cv = Converter(pdf_path)
        cv.convert(word_path, start=0, end=None, pdf2docx=True)
        cv.close()
        print(f"Conversion successful: {word_path}")
    except Exception as e:
        print(f"Error converting {pdf_path}: {e}")

def process_directory():
    # Get current directory
    current_dir = os.getcwd()
    
    # Find all PDF files in current directory
    pdf_files = [f for f in os.listdir(current_dir) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        print("No PDF files found in current directory")
        return
    
    print(f"Found {len(pdf_files)} PDF files to convert")
    
    # Process each PDF file
    for pdf_file in pdf_files:
        pdf_path = os.path.join(current_dir, pdf_file)
        convert_pdf_to_word(pdf_path)

if __name__ == "__main__":
    process_directory()