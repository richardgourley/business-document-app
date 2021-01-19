import os
import openpyxl
from openpyxl.utils import get_column_letter
from searchspreadsheet import SearchSpreadsheet
from countcolumninstances import CountColumnInstances
import docx
from docx.shared import Pt
import openpyxl
import datetime

main_menu = MainMenu()
print("LINE 14: CURRENT DIR IS: ", os.getcwd())

if main_menu.choice == "1":
    #search_spreadsheet = SearchSpreadsheet(spreadsheet_file, search_term)
    print("option is 1")

if main_menu.choice == "2":
    #count_column_instances = CountColumnInstances(spreadsheet_file, column_letter)
    print("option is 2")

if main_menu.choice == "3":
    #create_certificates = CreateCertificates(sheet, datetime.datetime.now())
    print("option is 3")

quit()