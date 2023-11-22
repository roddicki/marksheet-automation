# Create marksheets  

## Overview  

This repo generates .xlsx files from a list of student IDs and names.  
As each file is generated the respective cell in the sheet is updated with the student name and ID.  

## Input
Add student names and IDs in a file called “student-names.csv” (download from grade centre)  
Follow the existing column format:  
```ID | username | name```  

Add a template .xlsx marksheet ```marksheet.xlsx```  

## Output generated:
An .xlsx file for each student. The file will be called ```feedback_studentID``` with the field for their ```studentID``` and ```name``` populated.

## Setup
Install dependency: open excel
```Pip install openpyxl```

Edit or replace  ```student-names.csv``` or  ```marksheet.xlsx```  with appropriate files.
Student names and IDs should be in a file called ```student-names.csv```
This should follow the existing column format:
```ID | username | name ``` 

Edit create-files.py with the correct Sheet name and Row and columns to be populated :
```
sheet_name = 'Sheet1'  
row_number = 5  
col_student_name = 'B'  
col_student_id = 'D'
```  

## Run
```$python3 create-files.py```  

