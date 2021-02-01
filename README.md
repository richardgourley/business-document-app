# Business Document App
This is an app
 written purely in Python which gives the user a menu, where he or she can:
1. Search an excel file with any term - returns any rows with a cell match to the search term.
2. Count the number of instances a unique, word, date or number appears in a column.  For example, a search of a column called 'DEPARTMENT' could return something like:
  DEPARTMENT, NUMBER OF INSTANCES
  Marketing   2
  Production  1
  HR          2
3. Create certificates - 
   - The user can add an excel file with details of employees or students who the user would like to create a certificate for.
   - In this app, there is a Word document with the words first name, last name, start date, duration and today's date ready to be replaced.
   - The app takes each line in the excel file and creates a new certificate as a a Word document by replacing the words with the information in the excel columns. 

There are files included in the 'excelfiles' and 'wordfiles' directory as examples of the data this app would work with.

It utilizes os, openpyxl and docx.

## SKILLS COVERED
For Python students, this application is simple enough to follow but covers a number of key data structure examples.
It also imports and uses Python libraries.
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
