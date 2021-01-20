import os
import openpyxl

class SearchSpreadsheet:
    def __init__(self):
        self.check_excel_files_exist()
        self.display_available_files()
        self.choose_excel_file()
        self.choose_search_term()
        self.print_search_results()

    # Returns a data type of results and result information
    def print_search_results(self):
        count = 0
        search_term_lower = self.search_term.lower()
        for row in list(self.sheet.rows):
            for cell in row:
                if search_term_lower in str(cell.value).lower():
                    count += 1
                    print(", ".join(map(lambda x: str(x.value), row)))
                    break
        print("TOTAL RESULTS = ", count)
