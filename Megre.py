from docx import Document
import os

def merge_word_documents(directory, output_file):
    # Create a new Document object for the merged document
    merged_document = Document()

    # Get the list of .docx files in the directory
    docx_files = [doc for doc in os.listdir(directory) if doc.endswith('.docx')]

    # Iterate through each file in the list
    for index, doc in enumerate(docx_files):
        # Create a Document object from the file
        sub_doc = Document(os.path.join(directory, doc))

        # We do not want a page break before the first document
        if index > 0:
            merged_document.add_page_break()

        # Iterate through the element objects in the sub-document
        for element in sub_doc.element.body:
            # Import each element to the merged document
            merged_document.element.body.append(element)

    # Save the merged document to the output file
    merged_document.save(output_file)

# Directory containing the Word documents
directory_path = "C:/Users/mzalqahtani/Desktop/testMerge"

# Output file path
output_path = "C:/Users/mzalqahtani/Desktop/testMerge/output.docx"

merge_word_documents(directory_path, output_path)
