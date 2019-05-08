# 1. xl-manipulation
Python script for xlsm manipulation as well as other utilities

#### A. Changes from common parameters
Functions which read column B from build plans "Common Parameters", save values to a list, and replace the proper excel cell values with corrected information in designated sheets

#### B. Changes based on different module types
Functions which check to see vm module type per sheet, check correct information to replace based on module type via a dictionary, and replace improper cell information with correct information by iterating through key, value pairs in designated sheets

#### C. Changes in tag values
Changes in tag values need to be handled specifically so that cell order does not matter per sheet. Functions iterate through all rows in a column, check to find the specific parameter name, and replace adjacent cell with corrected parameter value based on module type and key, value pair

#### Network replacement
Function utilizes 1B. but also check to make sure cells are non-empty. Changes will not occur if cell is empty.

#### Changes in VM name
Changes in vm name follow same procedure as 1B. but also appends to the corrected value the designated file number

#### Changes in availability zones
Changes follow procedure of 1B. but also make sure to append "1" or "0" depending on whether a file's numbering is even or odd

#### Changes in IP's
Changes in IP's follow 1B. but append the proper IP numbering suffix based on module type

#### Utility 1
Allows user to extract cell value based on user input of sheet (0 indexed) and cell index (ex. 'C25')






