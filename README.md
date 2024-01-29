# Automated Dismissal Letter Generation with Streamlit Interface

This project streamlines the process of generating dismissal letters by automatically filling in the blank spaces with employee information. It saves significant time and reduces errors in the documentation process. The tool now includes a Streamlit-based user interface, enabling users to generate customized dismissal letters in Word format, which are then converted to PDF files. The dismissal letters are automatically filled with data from an Excel file and a predefined Word template.

# Project Purpose

The primary purpose of this project is to automate the time-consuming task of manually entering employee details into dismissal letters. By fetching employee data from an Excel file and populating a Word template, this tool significantly reduces the time and effort required to create individual dismissal letters, ensuring accuracy and consistency in the documentation.

## Requirements

1. Python 3.x
2. Libraries: pandas, python-docx, docx2pdf, logging, datetime, babel, pythoncom, streamlit
3. Place your Excel file in the "excel" folder and rename it to "Employee_data.xlsx". Ensure that the Excel file contains the necessary employee information.
4. Store the Word template for dismissal letters in the "template" folder.

## Installation

1. Clone or download this repository to your local machine.
2. Install the required Python libraries:
   ```
   pip install pandas python-docx docx2pdf streamlit babel pythoncom
   ```
3. Ensure that your environment meets all requirements listed above.

## Usage

1. Run the Streamlit app:
   ```
   streamlit run app.py
   ```
2. Open the provided local URL in your web browser.
3. Enter the employee's email and the dismissal date in the Streamlit interface.
4. Click on "Generate Dismissal Letter" to produce the dismissal letter.

The generated dismissal letter in Word format will be automatically converted to PDF and can be found in the designated output folder.

## Additional Notes

- The script `app.py` serves as the entry point for the Streamlit interface.
- `dismissal_letter_automation.py` contains the core logic for generating the dismissal letters.
- An alternative way to run this application is using Docker. This can provide a more streamlined setup and execution process, especially for deployment environments.

---
