import os
import openpyxl
from openpyxl.utils import get_column_letter
import docx
from datetime import datetime

class CreateCertificates:
    def __init__(self):
        self.intro()
        self.check_files_exist()
        self.check_number_rows()
        # Methods above will advise user of error messages.
        self.create_certificates()

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
        if self.sheet.max_row < 2:
            print("Sorry, there aren't any data rows in the database")
            quit()

    '''
    Returns a list of row titles
    Used in create_certificates() 
    '''
    def get_row_titles(self):
        row_titles = []
        for y in range(1, self.sheet.max_column + 1):
            row_titles.append(self.sheet[get_column_letter(y) + "1"].value)

        return row_titles
    
    def create_certificates(self):
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
                    # Get cell value - convert any dates or date strings to type '26th Jan 2020'
                    cell_value = self.convert_any_dates_to_month_ordinal_year(str(self.sheet[get_column_letter(y) + str(i)].value))

                    # If we find a row title eg . 'first name' in text, replace with corresponding cell value from this row
                    para.text = para.text.replace(
                        str(row_titles[y-1].lower()), 
                        cell_value
                    )

                # ADD todays date
                para.text = para.text.replace("date", str(self.convert_any_dates_to_month_ordinal_year(datetime.now())))

                # Re-apply existing font size and bold settting
                para.runs[0].font.size = para_font_size
                para.runs[0].bold = para_bold
                    
            # Save word doc for each excel row with name and todays date
            certificate_doc.save(str(self.sheet["A" + str(i)].value) + self.convert_any_dates_to_month_ordinal_year(datetime.now()) + "_certificate.docx")

        print("DONE!")
        print("You can find your certificates for each student created in the 'wordfiles' folder.")

    '''
    DATE METHODS
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

    '''
    @ returns a string - either formatted date or original cell
    Block 1 tests cell is DATETIME type
    Block 2 tests is string with '21/2/2001' date format
    Block 3 tests is string with '5/21/2001' date format
    '''
    def convert_any_dates_to_month_ordinal_year(self, cell):
        try:
            day_ordinal = self.return_ordinal(cell.day)
            date_string = "{} {} {}".format(str(cell.strftime('%B')), day_ordinal, str(cell.year))
            return date_string
        except:
            pass

        try:
            date = datetime.strptime(cell, "%d/%m/%Y")
            day_ordinal = self.return_ordinal(date.day)
            date_string = "{} {} {}".format(str(date.strftime('%B')), day_ordinal, str(date.year))
            return date_string
        except:
            pass

        try:
            date = datetime.strptime(cell, "%m/%d/%Y")
            day_ordinal = self.return_ordinal(date.day)
            date_string = "{} {} {}".format(str(date.strftime('%B')), day_ordinal, str(date.year))
            return date_string
        except:
            pass

        return str(cell)

