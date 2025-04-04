import pdfplumber
import pandas as pd
import os

def segment_text_by_lines(word_data, line_threshold=3):
    """
    Organize extracted words into lines based on vertical positioning.
    """
    word_data.sort(key=lambda item: (item['top'], item['x0']))
    structured_lines, temp_line = [], []
    reference_top = None
    
    for word in word_data:
        if reference_top is None or abs(word['top'] - reference_top) <= line_threshold:
            temp_line.append(word)
        else:
            structured_lines.append(temp_line)
            temp_line = [word]
        reference_top = word['top']
    
    if temp_line:
        structured_lines.append(temp_line)
    
    return structured_lines

def arrange_text_into_columns(text_segments, column_spacing=5):
    """
    Cluster words into columns using horizontal gaps.
    """
    text_segments.sort(key=lambda segment: segment['x0'])
    column_structure, temp_column = [], []
    prev_x_end = None
    
    for segment in text_segments:
        if prev_x_end is not None and (segment['x0'] - prev_x_end) > column_spacing:
            column_structure.append(" ".join(temp_column))
            temp_column = []
        temp_column.append(segment['text'])
        prev_x_end = segment['x1']
    
    if temp_column:
        column_structure.append(" ".join(temp_column))
    
    return column_structure

def extract_table_from_pdf(file_path):
    """
    Extract tabular text content from a given PDF file.
    """
    extracted_data = []
    with pdfplumber.open(file_path) as pdf_doc:
        for sheet in pdf_doc.pages:
            word_list = sheet.extract_words()
            structured_lines = segment_text_by_lines(word_list)
            for text_line in structured_lines:
                columns = arrange_text_into_columns(text_line)
                extracted_data.append(columns)
    
    return extracted_data

def export_to_excel(data_matrix, save_path):
    """
    Store extracted tabular data in an Excel spreadsheet.
    """
    data_frame = pd.DataFrame(data_matrix)
    data_frame.to_excel(save_path, index=False, header=False)
    print(f"Data successfully saved to: {save_path}")

def batch_process_pdfs(source_directory, destination_directory):
    """
    Process multiple PDF files from a given directory and save extracted content as Excel files.
    """
    os.makedirs(destination_directory, exist_ok=True)
    
    for filename in os.listdir(source_directory):
        if filename.lower().endswith(".pdf"):
            source_pdf = os.path.join(source_directory, filename)
            target_excel = os.path.join(destination_directory, f"{os.path.splitext(filename)[0]}.xlsx")
            print(f"Currently processing: {filename}")
            tabular_content = extract_table_from_pdf(source_pdf)
            export_to_excel(tabular_content, target_excel)

if __name__ == "__main__":
    input_directory = "Input_pdfs"  # Directory holding PDF files
    output_directory = "extracted_tables"  # Directory for output Excel files
    batch_process_pdfs(input_directory, output_directory)