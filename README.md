# VBA-Fill-Down-for-Stepped-reports
Use VBA within Excel to change a report that is laid out in a stepped manner to have the data filled down onto every row.  
# Example
        A   B   C   D
    1   a   b   c   d
    2           e   f
    3       g   h   i
    4       j   k   l
    5   m   n   o   p
    6       q   r   s

Will turn into

        A   B   C   D
    1   a   b   c   d
    2   a   b   e   f
    3   a   g   h   i
    4   a   j   k   l
    5   m   n   o   p
    6   m   q   r   s


# Directions
* Open this Fill Down Values Macro.xlsm file.
* Enable macros / allow permissions.
* Open/activate the workbook with the data in it that you want to fill down in the foreground.
* Save the workbook with the data in it, so that you do not accidentally mess up the data irreparably, then turn off AutoSave.
* Press Alt+F8 to open the Macros window, or click the Developer tab (which you may have to enable in Excel settings) and click the Macros button.
* Choose 'Fill Down Values Macro.xlsm'!FillDown.FillDown and Run Macro.
* Enter the column letters that you wish to check for empties to fill down in the first box prompt (from the example above, that would be A,B).
* Enter the first row of data after the headers (from the example above, there are no headers, so that would be 1).
* Enter the column that will always be empty when you want to fill down (from the example above, that would be A).
* Confirm your settings on the final message box.
* The data should now be filled down properly.  If you are happy with the changes, turn AutoSave back on and save your data.  If you are not happy with it, close your data file without saving.
