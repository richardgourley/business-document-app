# Business Document App
This is an written purely in Python which gives the user a menu, where he or she can:
1. Search an excel with any term - returns any rows with a cell match to the search term.
2. Count the number of instances a word, number of date appears in a column. eg. a column named department could return 2 for 'marketing', 3 for 'production' etc.
3. Create certificates - 
   - Opens a certificate Word file with name, course duration and date.
   - Opens an excel file, retrieves information for students in a row containing name and course duration
   - Replaces 'name', 'duration' and 'date' with the excel info for each student.
   - Saves a new Word document with the name of the student in the file name. 

Files ending '.txt' are included in the 'wordfiles' and 'excelfiles' directory as examples of the excel and word files used when making the application.

It utilizes os, openpyxl and docx.

## SKILLS COVERED
For Python students, this application is simple enough to follow but covers a number of key data structures and could be a good reference for a beginner to Python to see some scenarios of using data structures.  It also imports and uses Python libraries.
Some of the topics covered are:

- openpyxl - for working with excel spreadsheets using Python
- docx - for working with Word documents
- os - importing os and changing directories
- Retrieve excel cell data in Python
- Retrieve excel data, open a word document, replace words in the word document and save it as a new file
- While loops for user input validation
- Classes - OOP
- Class methods and properties
- import (how to import your own class modules and python libraries)
- Map 
- Join
- Split
- Lambda
- Combines map,lambda, join and split examples in one line of code within a for loop
- Lists
- Tuples
- Dictionaries - counting instances
- Lists - List comprehension used to help order the values of a dictionary
- Sorting a dictionary
