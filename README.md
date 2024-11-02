Project Title: Outlook Automation with PDF Text Extraction with and without OCR
Automation of Email Processing in Outlook: Extracting Text from PDF Attachments with and without OCR

Motivation
This project addresses the challenge of managing large volumes of emails with PDF attachments in professional environments, such as legal, academic, or corporate settings. Manually sorting and classifying these emails can be time-consuming and error-prone. This script automates the process of reading emails, extracting text from PDF attachments, and organizing emails into specified folders based on the content of the attachments. It significantly reduces manual work and improves efficiency by using text extraction through Optical Character Recognition (OCR) when necessary.

Method and Results
The script uses Microsoft’s Outlook API through the win32com.client library to access and manipulate emails. For PDF attachments, it first attempts to extract text using PyPDF2, a Python library capable of directly reading text from PDFs that are not image-based. If PyPDF2 fails to extract text (common with scanned documents), the script then uses EasyOCR, an OCR tool capable of interpreting text from images.

Repository Overviewa
Below is an overview of the repository structure and key files.

Execution Instructions
To replicate and use this workflow, follow these steps:

Prerequisites:
Microsoft Outlook must be installed and configured on your machine.

Python Installation and Setup Guide
Follow these steps to download and install the .exe on a Windows system:

Extract the zip file to a local folder.
Place all code files (.py) and dist within the same folder as the executable (.exe).
Run the executable.
Distribution on Other Machines
Share the ZIP file.
Alternatively, compress all items into a ZIP format, distribute, and unzip.
Copying a Script from GitHub
To copy a script from GitHub, follow these steps:

Navigate to the GitHub repository you wish to download.
Click the 'Code' button, then select 'Download ZIP' from the dropdown menu.
Once downloaded, extract the ZIP file using your file decompression tool to access the content.
Configuration:
Update outlook_parameters.xlsx in the configuration directory to set up your filters and destination folders according to your needs.

Folder Destination	Filter_1	Filter_2	Filter_3	Filter_4	Filter_5
folder/	LSXXXX	TVA-CH			
folder/sub_folder	LS0000				
Logging:
Refer to logs/outlook_process_log.log for details on script execution and issues encountered.

Additional Resources
For more information on the libraries and tools used in this project:

Part of the pywin32 library, which provides access to the Win32 API, including COM support.
bash
Code kopieren
pip install pywin32
A powerful data analysis and manipulation library for Python.
bash
Code kopieren
pip install pandas
A pure Python library built as a PDF toolkit. It can extract information, split documents page by page, merge documents, crop pages, etc.
bash
Code kopieren
pip install PyPDF2
easyocr An OCR library for Python that supports more than 40 languages.
bash
Code kopieren
pip install easyocr
fitz / PyMuPDF A Python binding for MuPDF – a lightweight PDF and XPS viewer.
bash
Code kopieren
pip install PyMuPDF
numpy A fundamental package for scientific computing with Python. It adds support for large multidimensional arrays and matrices, along with a large collection of high-level mathematical functions.
bash
Code kopieren
pip install numpy
Pillow (PIL Fork) Python Imaging Library adds image processing capabilities to your Python interpreter.
bash
Code kopieren
pip install Pillow
Licenses
Each component of this project relies on software under specific open-source licenses. Here is a list of these dependencies and their licenses:

Jinja2 (3.1.3) - BSD License: Allows commercial use with minimal restrictions.
MarkupSafe (2.1.5) - BSD License: Permissive, commercial-friendly license with minimal obligations.
PyPDF2 (3.0.1) - BSD License: Free to use commercially, no copyleft requirement.
PyYAML (6.0.1) - MIT License: Very permissive, allowing commercial use and modification without significant restrictions.
[List of additional licenses and their details continues]

Disclaimer
By using this software, you, the user, agree to the following:

No Warranty
The user acknowledges that this script is provided "as is" and that the developers give no warranties, express or implied, regarding the functionality of the script or its fitness for a particular purpose. The user assumes all risks associated with operating the script.

Data Processing
The script interacts with Microsoft Outlook and processes potentially sensitive information. The user is solely responsible for ensuring all data handling complies with applicable data privacy and protection laws.

Modifications
Any modifications made by the user to the script that cause system malfunction or data loss will be the responsibility of the user, and the developers disclaim any liability for such issues.

Compliance
The user is responsible for using the script in compliance with all applicable laws, including, but not limited to, data protection and privacy laws.

Limitation of Liability
In no event shall the developers be liable for direct, indirect, incidental, consequential, special, or exemplary damages resulting from the use or inability to use the software, even if advised of the possibility of such damages.

Acknowledgments
By downloading, copying, installing, or otherwise using this software, you acknowledge that you have read, understood, and agree to be bound by the terms of this disclaimer. If you do not agree to these terms, you are not authorized to use the software.
