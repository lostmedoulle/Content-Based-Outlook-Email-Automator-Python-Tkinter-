**Automating Email Processing in Outlook: Efficiently Extract Text from PDF Attachments with and without OCR**

![Python](https://img.shields.io/badge/Python-3.x-blue) 
![License](https://img.shields.io/badge/License-MIT-green) 
![Dependencies](https://img.shields.io/badge/Dependencies-Up%20to%20Date-brightgreen)
![Build](https://img.shields.io/badge/Build-Passing-brightgreen)
![Last Commit](https://img.shields.io/github/last-commit/your-username/your-repository)

---

## **Table of Contents**
- [Introduction](#introduction)
- [Motivation](#motivation)
- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [Excel Configuration](#excel-configuration)
- [Screenshots](#screenshots)
- [Logging](#logging)
- [Support](#support)
- [Acknowledgments](#acknowledgments)
- [License](#license)

---

## **Introduction**

This project automates email processing in Microsoft Outlook by extracting text from PDF attachments. It supports both native PDF text extraction and OCR-based extraction for scanned documents, reducing manual effort and improving productivity.

---

## **Motivation**

Managing emails with attachments in a professional setting can be time-consuming. This project automates the tedious process of sorting and classifying emails, saving time and increasing accuracy for tasks like legal document processing, research workflows, and corporate operations.

---

## **Features** ✨

- 📄 Extract text from PDFs (including scanned documents using OCR).  
- 📥 Automatically sort emails into folders based on keywords.  
- ⏲️ Process emails once or at periodic intervals (every 10 minutes).  
- 🖥️ Intuitive GUI built with Tkinter.  
- 🔍 Comprehensive logging for all operations.  
- 🧩 Extensible rules: Configure and load dynamic Excel rules for email sorting.

---

## **Installation**

### **Prerequisites**
- Microsoft Outlook installed and configured.
- Python 3.x installed on your system.

### **Setup Instructions**
1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/your-repository.git
   cd your-repository


## **Usage**

1. **Set Up Excel File**:
   - Open the `outlook_parameters_mailbox.xlsx` file.
   - Define your folder destinations and keywords for sorting and processing emails.

   **Example Excel Configuration**:
   | **Folder Destination** | **Filter_1** | **Filter_2** | **Filter_3** | **Filter_4** | **Filter_5** |
   |-------------------------|--------------|--------------|--------------|--------------|--------------|
   | folder/                | invoice      | payment      | contract     |              |              |
   | folder/sub_folder      | tax          | report       |              |              |              |

2. **Launch the Application**:
   - Start the GUI by running the following command:
     ```bash
     python Outlook_GUI.py
     ```
   - In the GUI interface:
     - Use the **Browse** button to load your Excel configuration file.
     - View logs and configurations directly in the main tab.

3. **Processing Modes**:
   - **"Exécuter une fois"**: Processes all emails once based on the configured rules.
   - **"Exécuter toutes les 10 minutes"**: Continuously processes emails every 10 minutes.

4. **Close Outlook**:
   - To ensure a conflict-free environment, the application will automatically close Outlook during processing. Emails are processed in the background.

5. **Stop Processing**:
   - Use the **"Arrêter"** button in the GUI to stop any ongoing processing tasks.

6. **Exit the Application**:
   - Click the **"Quitter"** button to safely close the application.

---

## **Screenshots**

### **GUI Interface**
- The GUI is designed for ease of use, with simple buttons and clear log displays.
![GUI Screenshot](assets/gui_screenshot.png)

### **Email Processing in Action**
- Watch the automation in real-time as the application processes emails and attachments.
![Workflow GIF](assets/email_processing_demo.gif)

### **Repository Structure**
    ```bash
    ├── LICENSE                       # License file for the project
    ├── Outlook_GUI.py                # GUI interface script for the application
    ├── README.md                     # Project documentation
    ├── company_logo_client.png       # Company logo used in the GUI
    ├── core_app.py                   # Core logic of the application
    ├── outlook_parameters_mailbox.xlsx # Configuration file for mailbox rules
    ├── outlook_process_log.log       # Log file for email processing
    ├── requirements.txt              # List of dependencies

---

## **Logging**

Logs are generated for every action, providing transparency and debugging assistance.

**Log File**: `outlook_process_log.log`

**Log Details**:
- Records the status of processed emails.
- Provides details of any errors encountered during execution.
- Includes performance metrics, such as processing times.

---

## **Support my work**

### Buy me a Coffe
<a href="https://buymeacoffee.com/lostmedoulle" target="_blank"><img src="https://www.buymeacoffee.com/assets/img/custom_images/orange_img.png" alt="Buy Me A Coffee" style="height: 41px !important;width: 174px !important;box-shadow: 0px 3px 2px 0px rgba(190, 190, 190, 0.5) !important;-webkit-box-shadow: 0px 3px 2px 0px rgba(190, 190, 190, 0.5) !important;" ></a>

### Donate Cryptocurrency

### **Stablecoin Donations (USDC)**

- **USDC on Ethereum**: `0x87358fF28b29E09037C8068260062742CDeAD671`
- **USDC on Base Chain**: `0x87358fF28b29E09037C8068260062742CDeAD671`
- **SOL** : `Gd4ncC2zXuj7ickNHJuHHtAoEKESTYd5FCJzzQwqANWJ`


