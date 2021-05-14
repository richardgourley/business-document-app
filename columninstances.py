import os
import openpyxl
from openpyxl.utils import get_column_letter

class CountColumnInstances:
    def __init__(self):
        self.check_excel_files_exist()
        self.display_available_files()
        self.choose_excel_file_and_assign_class_variables()
        self.display_column_letters_and_titles()
        self.choose_and_assign_column_letter()
        self.print_results()

    def check_excel_files_exist(self):
        os.chdir('excelfiles')
        self.excel_files = os.listdir()
        if len(self.excel_files) == 0:
            print('============')
            print("We couldn't find any files in the 'excelfiles' directory. Please add some files and run the program again.")
            print('============')
            quit()

    def display_available_files(self):
        print('Here are the available excel files you can search.')
        print('===============')
        for file in self.excel_files:
            print(file)
        print('===============')

    def choose_excel_file_and_assign_class_variables(self):
        excel_file = None
        while excel_file is None:
            print("Enter the name of a file to search. (Must match a file name shown above)")
            excel_file = input()
            try:
                excel_file = openpyxl.load_workbook(excel_file)
            except:
                excel_file = None
                print("Sorry, the file name you entered doesn't match an available file.")
        
        self.assign_class_variables(excel_file)
        print("===============")

    def assign_class_variables(self, excel_file):
        self.excel_file = excel_file
        self.sheet = excel_file.active
        self.sheet_column_letters = list()
        for i in range(1, self.sheet.max_column + 1):
            self.sheet_column_letters.append(get_column_letter(i))

    def display_column_letters_and_titles(self):
        print('Here are the available columns.')
        print('===============')
        
        for i in range(1, self.sheet.max_column + 1):
            print(get_column_letter(i), self.sheet[get_column_letter(i) + "1"].value)
        print('===============')

    def choose_and_assign_column_letter(self):
        chosen_column_letter = None

        while chosen_column_letter == None:
            print("Enter a column letter from above - A,B,C,D etc.")
            chosen_column_letter = input()

            if chosen_column_letter in self.sheet_column_letters:
                self.chosen_column_letter = chosen_column_letter
            else:
                chosen_column_letter = None

        print("=============")

    def print_results(self):
        self.count_rows_print_row_message()
        results = self.add_row_value_count_to_dict()
        sorted_results = self.sort_results_by_row_value_count(results)
        self.print_sorted_results(sorted_results)

    def count_rows_print_row_message(self):
        count = (self.sheet.max_row -1)
        print("From " + str(count) + " rows, we found:")

    def add_row_value_count_to_dict(self):
        results = {}

        for i in range(2, self.sheet.max_row + 1):
            value = self.sheet[self.chosen_column_letter + str(i)].value
            if not value in results:
                results[value] = 1
            else:
                results[value] += 1

        return results

    def sort_results_by_row_value_count(self, results):
        return sorted( [ (v,k) for k,v in results.items() ], reverse=True )

    def print_sorted_results(self, sorted_results):
        print("NUM INSTANCES,", self.sheet[self.chosen_column_letter + "1"].value)
        for result in sorted_results:
            result_row = ""
            for item in result:
                result_row += str(item) + " "
            print(result_row)