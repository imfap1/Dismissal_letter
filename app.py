import streamlit as st
from scr import dismissal_letter_automation as dla
import pythoncom

pythoncom.CoInitialize()

def main():
    st.title("Dismissal Letter Generator")

    # Adjust these file paths as per your project structure
    template_file_path = './template/template.docx'
    excel_file_path = './excel/employee_data.xlsx'

    email = st.text_input("Enter employee email:")
    dismissal_date_str = st.text_input("Dismissal date (DD-MM-YYYY):")

    if st.button("Generate Dismissal Letter"):
        message = dla.fill_word_document(template_file_path, excel_file_path, email, dismissal_date_str)
        st.write(message)


if __name__ == "__main__":
    main()

pythoncom.CoUninitialize()
