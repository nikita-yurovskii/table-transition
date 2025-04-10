# Excel tables to word 

### Table of Contents
1. Description
2. Installation
3. Usage
4. Credits

## Description
Python3 script for copying specific IDs from a master excel table to a word document.

Tables from excel are copied to word using docx library.
Excel tables need to be a table per sheet and should match up to the word template if possible.
If there's no template, the script will create one based on the table in the excel sheet (only 1 row as header).
This is a test script that may form the basis of some automated reporting in the future.

Library requirements: pandas, docx, os, datetime, numpy

## Installation
To install this program
- install python 3 
- pip install required libraries
- download main.py and place into chosen folder

## Usage
Input requirements:
- An excel master document to look for data, one table per sheet - 'test_table.xlsx' is created if create_test_date() is used.
- An excel input document with a list of IDs in column A without header - 'input_data.xlsx' is created if create_test_data() is used.
- A template word document is optional, this should include tables that match the excel master doc, i.e. sheet1 table should be the first table in the word template, sheet2 should be the 2nd. The number of cols/rows should also be the same.
- If no template is used, the script will create a template based on the excel document although this is limited to only 1 header row.
An excel master document to look for data, one table per sheet - 'test_table.xlsx' is created if create_test_date() is used.
- There are a number of docx table specific functions that are included but not used. These are merge_rows(), cell_colour() and update_cell_text().


To use this program:
- put the master excel document, input data excel document and template in the same folder as the .py file
- run main.py from installation folder
- an ouptut + datetime document will be produced with the excel data in the word file

## Credits
Me :ok_hand:


