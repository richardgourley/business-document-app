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
        # count will display how many results we find
        count = 0
        # we assigned 'self.search_term' in 'choose_search_term()'
        # use 'self.search_term' lower case - easier to compare with 'cell.value' lowercase
        search_term_lower = self.search_term.lower()

        # Print out column names before results
        print(self.column_names)

        for row in list(self.sheet.rows):
            # loop through each cell, if search_term matches cell value - ..
            # .. convert tuple to string + print and stop comparing (break)
            for cell in row:
                if search_term_lower in str(cell.value).lower():
                    count += 1
                    print(", ".join(map(lambda x: str(x.value), row)))
                    break
        print("TOTAL RESULTS = ", count)
