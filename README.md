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

Installation
To get started, ensure you have Python installed on your system. Follow these steps to install the required dependencies for this project.

Clone the Repository (if applicable):

bash
Code kopieren
git clone https://github.com/your-username/your-repository.git
cd your-repository
Install Dependencies: All necessary dependencies are listed in the requirements.txt file. To install them, use the following command:

bash
Code kopieren
pip install -r requirements.txt
This command will install all required packages specified in the requirements.txt file.

Verify Installation: After installation, you can verify that the dependencies are installed correctly by running:

bash
Code kopieren
python -m pip check


Update Excel-File
Folder Destination	Filter_1	Filter_2	Filter_3	Filter_4	Filter_5
folder/	LSXXXX	TVA-CH			
folder/sub_folder	LS0000				
Logging:
Refer to logs/outlook_process_log.log for details on script execution and issues encountered.

Additional Resources
For more information on the libraries and tools used in this project:


Licenses
Each component of this project relies on software under specific open-source licenses. Here is a list of these dependencies and their licenses:

Jinja2 (3.1.3)

License: BSD License
Usage: Allows commercial use with minimal restrictions. Requires the preservation of copyright notices.
MarkupSafe (2.1.5)

License: BSD License
Usage: Permissive, commercially friendly license with minimal obligations.
PyPDF2 (3.0.1)

License: BSD License
Usage: Free for commercial use. No copyleft requirements.
PyYAML (6.0.1)

License: MIT License
Usage: Very permissive, allows commercial use and modification without significant restrictions.
certifi (2024.2.2)

License: Mozilla Public License 2.0 (MPL 2.0)
Usage: Commercial use permitted. Modifications to licensed files must be shared under the same license.
cffi (1.16.0)

License: MIT License
Usage: May be used commercially with few restrictions.
charset-normalizer (3.3.2)

License: MIT License
Usage: Minimal restrictions on usage, including commercial use. Easy to integrate without legal complexity.
cryptography (42.0.5)

License: Apache Software License; BSD License
Usage: Dual license for flexibility, supporting commercial use with requirements for notice and conditions as stated in the Apache License.
easyocr (1.7.1)

License: Apache License 2.0
Usage: Allows commercial use with conditions regarding modifications and notices.
et-xmlfile (1.1.0)

License: MIT License
Usage: Free for commercial use, including modification and distribution.
filelock (3.13.4)

License: The Unlicense
Usage: Public domain software. No usage restrictions.
fsspec (2024.3.1)

License: BSD License
Usage: Permissive, suitable for commercial projects. Minimal compliance requirements.
idna (3.7)

License: BSD License
Usage: Minimal restrictions; commercial use permitted.
imageio (2.34.0)

License: BSD License
Usage: Commercial-friendly with few restrictions.
lazy_loader (0.4)

License: BSD License
Usage: Commercial use permitted with minimal obligations.
licenses (0.6.1)

License: Public Domain
Usage: No copyright; free for any use.
mpmath (1.3.0)

License: BSD License
Usage: Free to use, modify, and distribute commercially.
networkx (3.3)

License: BSD License
Usage: Commercial use permitted; one of the most permissive licenses.
ninja (1.11.1.1)

License: Apache Software License; BSD License
Usage: Dual licensing offers flexibility for commercial use.
numpy (1.26.4)

License: BSD License
Usage: Commercial use with minimal restrictions; widely used in commercial scientific computing.
opencv-python (4.9.0.80)

License: Apache Software License
Usage: Commercial use permitted with requirements for notice and conditions as stated.
openpyxl (3.1.2)

License: MIT License
Usage: Very permissive for commercial integration and distribution.
packaging (24.0)

License: Apache Software License; BSD License
Usage: Dual licensing, enhancing usage flexibility in commercial applications.
pandas (2.2.2)

License: BSD License
Usage: Permissive, well-suited for commercial use in data analysis.
pdf2image (1.17.0)

License: MIT License
Usage: Very permissive, allowing commercial use and distribution without major restrictions.
pdfminer.six (20231228)

License: MIT License
Usage: Allows extensive modification and commercial distribution.
pdfplumber (0.11.0)

License: MIT License
Usage: Minimal restrictions, suitable for commercial projects involving PDF processing.
pillow (10.3.0)

License: Historical Permission Notice and Disclaimer (HPND)
Usage: Permissive, with historical permissions granting significant freedom for usage.

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
