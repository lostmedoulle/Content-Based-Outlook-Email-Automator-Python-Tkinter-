Automating Email Processing in Outlook: Efficiently Extract Text from PDF Attachments Using OCR


Table of Contents
Introduction
Motivation
Features
Installation
Usage
Excel Configuration
Repository Structure
Logging
Licenses
Support
Disclaimer
Acknowledgments


Introduction
This project automates the processing of emails in Microsoft Outlook by extracting text from PDF attachments, including scanned PDFs, using OCR technology. Users can configure specific keywords for sorting and processing emails based on an Excel file.


Motivation
Handling large email volumes with PDF attachments is time-consuming and error-prone in professional settings. This tool automates the task, reducing human effort and improving accuracy and efficiency.


Features
PDF text extraction with and without OCR.
Automated sorting and classification of emails based on user-defined rules.
Periodic email processing with configurable intervals.
Easy-to-use graphical user interface (GUI) built with Tkinter.


Installation
Prerequisites
Microsoft Outlook installed and configured.
Python 3.x installed on your system.


Setup Instructions
Clone the repository:
git clone https://github.com/your-username/your-repository.git
cd your-repository

Install the required dependencies:
pip install -r requirements.txt

Verify the installation:
python -m pip check

Usage
Steps to Run the Application
Set Up the Excel File:

Open the outlook_parameters_mailbox.xlsx file.
Define keywords in the appropriate columns based on your filtering needs (see Excel Configuration for details).
Run the GUI:
Launch the GUI by running the Outlook_GUI.py file:
python Outlook_GUI.py


This will open a graphical interface for managing email processing.
Load the Configuration:

Use the GUI to load the updated Excel configuration file by selecting it through the interface.
Start Processing:

Click "Exécuter une fois" to process emails once.
Click "Exécuter toutes les 10 minutes" to start periodic processing.
Close Outlook:

When the script runs, Outlook will automatically close to avoid session conflicts. Emails will be processed in the background.
Stop Processing:

Click "Arrêter" to stop periodic processing.
Click "Quitter" to exit the application.


Excel Configuration
The outlook_parameters_mailbox.xlsx file is used to define sorting rules. Configure it as follows:

Folder Destination	Filter_1	Filter_2	Filter_3	Filter_4	Filter_5
folder/	keyword1	keyword2			
folder/sub_folder	keyword3				
Folder Destination: Specify the folder where emails matching the filters should be moved.
Filter Columns: Define keywords for filtering emails based on the content of PDF attachments.

Repository Structure
├── LICENSE                       # License file for the project
├── Outlook_GUI.py                # GUI interface script for the application
├── README.md                     # Project documentation
├── company_logo_client.png       # Company logo used in the GUI
├── core_app.py                   # Core logic of the application
├── outlook_parameters_mailbox.xlsx # Configuration file for mailbox rules
├── outlook_process_log.log       # Log file for email processing
├── requirements.txt              # List of dependencies


Logging
Execution logs are saved in outlook_process_log.log. These logs provide details about:

Successfully processed emails.
Errors encountered during processing.
Debugging information for troubleshooting.


Licenses
This project leverages several open-source libraries under permissive licenses, such as:

PyPDF2: BSD License
EasyOCR: Apache License 2.0
pandas: BSD License
pdfplumber: MIT License
Refer to requirements.txt for the complete list of dependencies and their licenses.
Project Title: Outlook Automation with PDF Text Extraction with and without OCR
Automation of Email Processing in Outlook: Extracting Text from PDF Attachments with and without OCR


--
if you want to support my work 

<a href="https://buymeacoffee.com/lostmedoulle" target="_blank"><img src="https://www.buymeacoffee.com/assets/img/custom_images/orange_img.png" alt="Buy Me A Coffee" style="height: 41px !important;width: 174px !important;box-shadow: 0px 3px 2px 0px rgba(190, 190, 190, 0.5) !important;-webkit-box-shadow: 0px 3px 2px 0px rgba(190, 190, 190, 0.5) !important;" ></a>

--
Disclaimer
This software is provided "as is," without any warranty. Users are responsible for ensuring compliance with data protection laws and handling sensitive information appropriately.

Acknowledgments
Thanks to the developers of the following libraries and tools:

PyPDF2: For PDF text extraction.
EasyOCR: For OCR functionality.
pandas: For data manipulation.
Tkinter: For building the GUI.


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
folder/	abcd	TVA-CH			
folder/sub_folder	abcd				
Logging:
Refer to logs/outlook_process_log.log for details on script execution and issues encountered.

Additional Resources
For more information on the libraries and tools used in this project:

Licenses
Each component of this project relies on software under specific open-source licenses. Here is a list of these dependencies and their licenses:

Acknowledgments
By downloading, copying, installing, or otherwise using this software, you acknowledge that you have read, understood, and agree to be bound by the terms of this disclaimer. If you do not agree to these terms, you are not authorized to use the software.
