import pandas as pd
from docx import Document
from docx2pdf import convert
import logging
import datetime
from babel.dates import format_date
import pythoncom

pythoncom.CoInitialize()


def validate_email(email, df):
    '''Validate if the email exists in the dataframe.'''
    return not df[df['Work Email'] == email].empty

def validate_date(date_text):
    '''Validate if the input date is in the correct format.'''
    try:
        datetime.datetime.strptime(date_text, '%d-%m-%Y')
        return True
    except ValueError:
        return False

def format_dates_with_month_name(dismissal_date_str):
    '''Formats a date string into English and Spanish formats with month names.'''
    dismissal_date = datetime.datetime.strptime(dismissal_date_str, '%d-%m-%Y')
    dismissal_date_eng = format_date(dismissal_date, format='long', locale='en')
    dismissal_date_esp = format_date(dismissal_date, format='long', locale='es')
    return dismissal_date_eng, dismissal_date_esp

def fill_word_document(template_file_path, excel_file_path, email, dismissal_date_str):
    '''Fill a Word document with information from an Excel file and convert it to PDF.'''
    try:
        df = pd.read_excel(excel_file_path)

        if not validate_email(email, df):
            logging.error("Unknown employee.")
            return "Unknown employee."

        if not validate_date(dismissal_date_str):
            logging.error("Invalid date format.")
            return "Invalid date format."

        dismissal_date_eng, dismissal_date_esp = format_dates_with_month_name(dismissal_date_str)
        employee_data = df[df['Work Email'] == email].iloc[0]
        employee_name = employee_data['Full name']
        start_date = employee_data['Start date'].date().strftime('%d-%m-%Y')
        job_title = employee_data['Job title']

        data_to_insert = {
            'fecha_inicio': start_date,
            'fecha_despido_esp': dismissal_date_esp,
            'fecha_despido_ing': dismissal_date_eng,
            'cargo': job_title,
            'employee_name': employee_name
        }

        doc = Document(template_file_path)
        for paragraph in doc.paragraphs:
            for key, value in data_to_insert.items():
                placeholder = f'({key})'
                if placeholder in paragraph.text:
                    paragraph.text = paragraph.text.replace(placeholder, str(value))

        output_path = f'./Dismissal letter/Dismissal letter {employee_name}.docx'
        doc.save(output_path)
        logging.info("Dismissal letter created.")

        convert(output_path, output_path.replace('.docx', '.pdf'))
        logging.info("Word document successfully converted into PDF.")

        return "Dismissal letter created and converted to PDF."

    except Exception as e:
        logging.error(f"An error occurred: {e}")
        return f"An error occurred: {e}"


def main():
    logging.basicConfig(level=logging.INFO)
    template_file_path = './template/template.docx'
    excel_file_path = './excel/employee_data.xlsx'
    fill_word_document(template_file_path, excel_file_path)

if __name__ == "__main__":
    main()
    
pythoncom.CoUninitialize()

