
The above code works in the command line.
Usage:
python process_excel.py <filepath> [operation] [opfile]

<filepath> : path of the input excel file.This is mandatory parameter.If operations are niot given it will display the contents of excel file.
[operation]
add
del
update

[opfile] : comma separated file for add , del operations and file with dictionary for update values.


Examples :
python process_excel.py employee_details.xlsx add add.txt
python process_excel.py employee_details.xlsx del del.txt
python process_excel.py employee_details.xlsx update update.txt
python process_excel.py employee_details.xlsx
