# VBA Script - Data Highlighting and Summary

This VBA script is designed to highlight rows of data in a specified column and provide a summary of the highlighted cells. It allows the user to select the column to reference and performs calculations based on the values within that column. The script can be used to identify and analyze specific data points in a dataset. This script is designed with the 80/20 rule in mind to optimize a selection.

This script is designed with performance in mind, many coworkers do not have high spec computing devices this represents a very low cost solution achieved through iterative loops and data manipulation through rudementary calculations.

## Updates
  July 19, 2023: Fixed bug with sheet naming reference
        Updated usage instructions
  August 24, 2023: Improved error handling, now requires minimal formatting for same results
        Will now automatically filter out invalid data types provided in the file

## Features

- Automatically identifies the range based on the first and last non-zero numerical values in the selected column.
- Randomly selects rows from the top values until reaching the initial target total.
- Continues selecting rows until reaching the final target total or the maximum number of cells are selected.
- Highlights the selected rows in the worksheet.
- Provides a summary with information about the highlighted cells, total sum, target maximum, and status.
- Creates a new worksheet for each run to store the summary information.

## Getting Started

### Prerequisites

Microsoft Excel (version 2010 or later) is required to run the VBA script.

### Usage
1. Open the 80/20 tool and copy all of the code.
2. Open the Excel file containing the dataset.
3. Hide all columns below and above target to avoid them being used in the calculations (Can unhide after script is run)
4. Press Alt + F11 (Alt + Fn + F8 if not using keyboard with dedicated function row) to open the VBA editor.
5. Insert a new module and paste the script code into the module.
6. Press Alt + F8 (Alt + Fn + F8 if not using keyboard with dedicated function row) to run the script by executing the `HighlightRowsAndWriteMessage` macro.
7. Follow the prompts to input the reference column containing the prices.
8. The script will create a copy of the Master sheet, highlight rows, and generate a summary sheet with the results.
9. Depending on the difference shown in summary, you may need to remove a low cost element and add one that is closest to the value of the difference between the target and actual values.
10. this script can be executed many times on the same sheet by repeating steps 5-8, it is important to ensure this is ONLY run on the original. Copies with your selection will be created as a separate page in the same file. You may pick the iteration which best suites your needs, it is important to remember this tool used randomization to achieve its results so it may not produce a desireable result EVERY time, but it is tuned so that you should find a very high rate of success within a couple iterations at most.

### License

This project is licensed under the MIT License. Feel free to modify and use the script according to your needs.

### Contributing

Contributions are welcome! If you have any improvements or bug fixes, please submit a pull request or open an issue to discuss the changes.

### Authors

Nick Andrusiw - Sole author and contributer.

### Acknowledgments

The script was inspired by the need to analyze specific data points in large datasets, and optimally gather a weighted random selection to meet the objective criteria of 20% of cells amounting to 80% of value.
