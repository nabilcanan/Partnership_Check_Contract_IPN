# Excel File Comparator

The Excel File Comparator is a Python program that allows you to compare data between two Excel contracts, the one from last week then the one from the current week. 
Once you select both files you can input the header name you want to compare by. In most cases this will be 'IPN'. Then you can save the data in your new workbook created 
wherever you'd like. 

## How It Works

1. **Run the Program**: Execute the Python script in your terminal or IDE.

2. **Select Excel Files**: The program will prompt you to select two Excel files (.xlsx or .xls) that you want to compare. Click the "Select" buttons to choose the files from your computer.

3. **Compare Files**: Once you have selected the two files, click the "Compare Files" button. The program will then compare the data in the selected Excel files based on the 'IPN' (Item Part Number) column.

4. **View Differences**: If there are any differences between the two files, the program will display them in a table in the GUI. The table will show the rows and columns corresponding to the missing IPNs that are present in the second file but not in the first file.

5. **Scroll the Table**: If the table extends beyond the visible area of the GUI, you can scroll both horizontally and vertically using the scrollbars to view the entire table.

6. **Save the Differences**: If there are missing IPNs, the program will prompt you to save a new Excel file containing the rows corresponding to the missing IPNs. The new file will be saved with the .xlsx extension.

7. **Exit the Program**: You can close the GUI by clicking the "X" button or using the close button of your window manager.

## Requirements

- Python 3.6 or higher
- pandas library
- tkinter library
- openpyxl library


