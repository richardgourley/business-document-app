import os
import openpyxl
from openpyxl.utils import get_column_letter
from searchspreadsheet import SearchSpreadsheet
from countcolumninstances import CountColumnInstances
import docx
from docx.shared import Pt
import openpyxl
import datetime

print('Hello, what is your name?')
name = input()
print('Hello ' + name + '!!!') 
print('Choose an option from the menu below:')