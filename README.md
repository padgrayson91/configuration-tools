# configuration-tools
A simple tool for mapping Excel files to .csv templates

This is a very tailored tool: it can only map from .xlsx to .csv (not the other way around)

## usage
The tool is run with the command `python ExcelConverterDialog.py`
Once started, you should do the following steps in order:
 1. Select an Excel file (bottom left)
 2. Select a csv template (bottom left)
 3. Select at least one mapping from an Excel column name to a csv column name (if more are needed, use the "add mapping" button)
 4. Use the "constants" entry widget (top) to enter values for any columns in the csv which will remain the same in every row. Note: at this time, constants can only be deleted manually by modifying the "mappings.json" file, so ensure that a value is correct before hitting the "store" button.  To examine constants for a mapping, hit the "view" button.
 5. Click the "Generate csv" button (bottom right) to transfer the rows from the Excel file to the csv template you selected (the excel data will be appended to the template)

## saving and loading mappings
If you are using the same mapping repeatedly, you can save time by storing the mapping.  Once a mapping is created as described above, you may save the mapping by entering a name into the "Save as" entry widget at the top of the application window.  Note: the application will not warn you if you are overwriting an existing mapping.

To load a saved mapping, first select your excel file and csv template as usual, then select the name of the desired mapping from the drop down menu at the top of the application window and press the "load" button.  This will load the correct mapping configuration as well as any constant values associated with the mapping.
