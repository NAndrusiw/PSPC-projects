VBA Script - Data Highlighting and Summary
This VBA script is designed to highlight rows of data in a specified column and provide a summary of the highlighted cells. It allows the user to select the column to reference and performs calculations based on the values within that column. The script can be used to identify and analyze specific data points in a dataset.

Features:

Automatically identifies the range based on the first and last non-zero numerical values in the selected column.
Randomly selects rows from the top values until reaching the initial target total.
Continues selecting rows until reaching the final target total or the maximum number of cells are selected.
Highlights the selected rows in the worksheet.
Provides a summary with information about the highlighted cells, total sum, target maximum, and status.
Creates a new worksheet for each run to store the summary information.
Getting Started
Prerequisites
Microsoft Excel (version 2010 or later) is required to run the VBA script.
Usage
Open the Excel file containing the dataset.
Press Alt + F11 to open the VBA editor.
Insert a new module and copy the script code into the module.
Modify the script if needed, such as adjusting target percentages or maximum cells.
Save the Excel file and run the script by executing the HighlightRowsAndWriteMessage macro.
Follow the prompts to select the column to reference.
The script will create a copy of the "Master" sheet, highlight rows, and generate a summary sheet with the results.
License
This project is licensed under the MIT License. Feel free to modify and use the script according to your needs.

Contributing
Contributions are welcome! If you have any improvements or bug fixes, please submit a pull request or open an issue to discuss the changes.

Authors
Your Name - Initial work
Acknowledgments
The script was inspired by the need to analyze specific data points in large datasets.
