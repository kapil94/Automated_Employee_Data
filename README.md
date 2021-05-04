# Automated_Employee_Data
Motive is to store employee data in database by automating data from Excel file

Program uses openpyxl to access excel data and sqlite3 to store data in database.

Prerequisites:
Employee table should already be present in Employee database

Algorithm:

1. Check if records are already present in Employee database.

2. Condition 1: Records are already present and last record stored in Employee table is not the last record in Employee.xlsx:
          1. get the data from excel for new records and store in Employee Table.

3. Condition 2: Records are already present and last record stored in Employee table is the last record in Employee.xlsx:
          1. print "Records are already present" and exit.
         
4. Condition 3: No record is present in Employee table:
          1. Traverse and store each record in Employee table from Employee.xlsx
