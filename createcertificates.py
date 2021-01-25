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
        self.rows = self.create_dictionary_for_each_row()
        for row in self.rows:
            print(row.keys())
        '''
        self.test_number_columns()
        self.test_column_names()
        self.print_certificates()
        '''

    def intro(self):
        print("To proceed we need to find:")
        print("The 'studentcertificates.xlsx' file in the folder 'excelfiles'")
        print("...AND the 'cerfificate.docx' file in the 'wordfiles' folder.")
        print("===Checking folders===")

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
    
    '''
    Returns a list of dictionaries
    Each dictionary represents a row with {column title: row value}
    '''
    def create_dictionary_for_each_row(self):
        # Get column titles in a list
        column_titles = []
        for i in range(1, self.sheet.max_column + 1):
            column_titles.append(self.sheet[get_column_letter(i) + "1"].value)

        # List to contain each row dictionary
        row_dictionaries = []
        for i in range(2, self.sheet.max_row + 1):
            row_as_dictionary = {}

            for y in range(1, self.sheet.max_column + 1):
                # add column_titles[y-1] as dictionary key
                row_as_dictionary[column_titles[y-1]] = self.sheet[get_column_letter(y) +  str(i)].value

            # add row_as_dictionary to row_dictionaries
            row_dictionaries.append(row_as_dictionary)

        return row_dictionaries

    def print_certificates(self):
        print("CREATING CERTIFICATES")

        os.chdir('../wordfiles')

        for row in self.rows:
            print(row.keys)
        
        for i in range(2, self.sheet.max_row + 1):
            first_name = str(self.sheet['A' + str(i)].value)
            last_name = str(self.sheet['B' + str(i)].value)
            name = first_name + " " + last_name
            duration = str(self.sheet['C' + str(i)].value)
            start_date = str(self.sheet['D' + str(i)].value)

            # Open the original certificate doc every time
            certificate_doc = docx.Document("certificate.docx")
            for para in certificate_doc.paragraphs:
                if para.text == "":
                    continue
                # Get existing font size and bold setting
                para_font_size = para.runs[0].font.size
                para_bold = para.runs[0].bold
                para.text = para.text.replace("name", name)
                para.text = para.text.replace("duration", duration)
                para.text = para.text.replace("date", self.today)
                # Apply existing font size and bold settting
                para.runs[0].font.size = para_font_size
                para.runs[0].bold = para_bold
                # save with students name
            certificate_doc.save(first_name + "_" + last_name + "_certificate.docx")

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