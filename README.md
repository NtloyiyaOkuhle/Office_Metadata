# Office_Metadata
This script extracts metadata from an office document file, such as a Microsoft Word or PowerPoint file. It reads the contents of the file as a ZIP archive and extracts the relevant XML files containing metadata information. It then parses the XML files and outputs the metadata information to the console.

The metadata information includes details such as the document title, subject, author, keywords, description, created and modified dates, and various counts such as page count, word count, and character count. The script uses dictionaries to map XML tag names to the corresponding metadata titles, and then loops through the XML elements to extract the metadata values.

To use this script, you need to have Python installed on your system, and the PyPDF2 and argparse modules. You can run the script from the command line by passing the path to the office document file as an argument, like this:



python script_name.py path/to/office/file.docx
