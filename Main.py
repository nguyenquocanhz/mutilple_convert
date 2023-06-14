import os
import win32com.client as win32

def convert_doc_to_docx(doc_file_path, docx_file_path):
    # Create an instance of the Word application
    word = win32.gencache.EnsureDispatch("Word.Application")

    try:
        # Open the DOC file
        doc = word.Documents.Open(doc_file_path)

        # Save the contents of the DOC file as DOCX
        doc.SaveAs2(docx_file_path, FileFormat=16)  # FileFormat 16 corresponds to DOCX format

        # Close the document
        doc.Close()

        print(f"Conversion successful: {doc_file_path}")
    except Exception as e:
        print(f"Error converting file: {doc_file_path}")
        print(f"Error message: {str(e)}")
    finally:
        # Quit the Word application
        word.Quit()

def batch_convert_folder_to_docx(input_folder_path, output_folder_path):
    # Create the output folder if it doesn't exist
    if not os.path.exists(output_folder_path):
        os.makedirs(output_folder_path)

    # Get a list of all files in the input folder
    files = os.listdir(input_folder_path)

    # Iterate over each file in the input folder
    for file_name in files:
        # Check if the file has a .doc extension
        if file_name.endswith(".doc"):
            # Create the input and output file paths
            doc_file_path = os.path.join(input_folder_path, file_name)
            docx_file_path = os.path.join(output_folder_path, f"{os.path.splitext(file_name)[0]}.docx")

            # Convert the file to DOCX
            convert_doc_to_docx(doc_file_path, docx_file_path)

# Specify the path to the input folder containing the files
input_folder_path = input("Path Folder DOC FILE :")

# Specify the path to the output folder for saving the converted files
output_folder_path = input("Path Folder DOCX FILE :")

# Call the batch conversion function
batch_convert_folder_to_docx(input_folder_path, output_folder_path)


# pip install -r requirements.txt
# It's Try Donwload All