import os
import openpyxl
from openpyxl.utils import get_column_letter
import docx
import datetime

class CreateCertificates:
    def __init__(self):
        self.today = self.get_date_with_ordinal(datetime.datetime.now())
        self.intro()
        self.check_files_exist()
        self.check_number_rows()
        # Methods above will advise user of error messages.
        self.print_certificates()

    def intro(self):
        print("To proceed we need to find:")
        print("The 'studentcertificates.xlsx' file in the folder 'excelfiles'")
        print("...AND the 'cerfificate.docx' file in the 'wordfiles' folder.")
        print("===Checking folders===")
        print("NOTE: The words you want to replace in the certificate file must match the title columns in the spreadsheet.")

    def check_files_exist(self):
        os.chdir('excelfiles')
        excel_files = os.listdir()
        os.chdir('../wordfiles')
        word_files = os.listdir()

        if (not "studentcertificates.xlsx" in excel_files) or (not "certificate.docx" in word_files):
            print("Sorry, we couldn't find the files above.")
            print("Please check the files exist in the correct folders and try again.")
            quit()
        else:
            try:
                certificate_doc = docx.Document('certificate.docx')
                os.chdir("../excelfiles")
                student_certificates_excel = openpyxl.load_workbook('studentcertificates.xlsx')
            except:
                print("We found the files but there was a problem opening the files. Please check.")
                quit()
        
        # If 'studentcertificates.xlsx' opens ok, assign self.sheet to be used in print_certificates
        self.sheet = student_certificates_excel.active

    def check_number_rows(self):
        if len(self.sheet.max_row) < 2:
            print("Sorry, there aren't any data rows in the database")
            quit()

    '''
    Returns a list of row titles, used for replacing text in the certificate document file
    '''
    def get_row_titles():
        row_titles = []
        for y in range(1, self.sheet.max_column + 1):
            row_titles.append(self.sheet[get_column_letter(y) + "1"].value)

        return row_titles
    
    def print_certificates(self):
        print("CREATING CERTIFICATES")

        os.chdir('../wordfiles')
        row_titles = self.get_row_titles()

        for i in range(1, self.sheet.max_row + 1):
            certificate_doc = docx.Document('certificate.docx')

            for para in certificate_doc.paragraphs:
            if para.text == "":
                continue

            # Get current font settings for paragraph
            para_font_size = para.runs[0].font.size
            para_bold = para.runs[0].bold

            for y in range(1, self.sheet.max_column + 1):
                para.text = para.text.replace(
                    str(row_titles[y-1].lower()), 
                    str(self.sheet[get_column_letter(y) + str(i)].value)
                )

            # ADD todays date
            para.text = para.text.replace("date", str(self.today))

            # Re-apply existing font size and bold settting
            para.runs[0].font.size = para_font_size
            para.runs[0].bold = para_bold
                    
        # Save word doc for each excel row with name
        certificate_doc.save(str(self.sheet["A" + str(i)].value) + str(self.today) + "_certificate.docx")

        print("DONE!")
        print("You can find your certificates for each student created in the 'wordfiles' folder.")

    '''
    DATE, TODAYS DATE METHODS
    '''
    def return_ordinal(self, day):
        ending_one = [1,21,31]
        ending_two = [2,22]
        ending_three = [3,23]
        if int(day) in ending_one:
            return str(day) + "st"
        if int(day) in ending_two:
            return str(day) + "nd"
        if int(day) in ending_three:
            return str(day) + "rd"
        return str(day) + "th"

    def convert_date_to_month_ordinal_year(self,date):
        date_list = date.split(" ")
        # change day to ordinal eg 3 -> 3rd, 21 -> 21st
        date_list[1] = self.return_ordinal(date_list[1])
        return " ".join(date_list)

    def get_date_with_ordinal(self, date):
        todays_date = str(date.strftime('%B %d %Y'))
        return self.convert_date_to_month_ordinal_year(todays_date)