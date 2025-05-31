# Merge Many Excel Files into one VBS

Originally posted here:
https://www.codeproject.com/Tips/5264786/Merge-Many-Excel-Files-into-one-VBS

This VBS will merge Many Excel files into one

## Using the Code
Before you can run the script, you need to setup the configuration (MergeExcel.txt) file. In Windows Explorer, hold shift and right-click on the file you want to merge, select "Copy as path". Paste the path into MergeExcel.txt file. Each file in the file is the path to the Excel file to be merged. The configuration has to reside in the same folder as the VBS script.

```
c:\folder1\Excel1.xlsx
c:\folder1\Excel2.xlsx
c:\folder3\Excel3.xlsx
```

Double click to run MergeExcel.vbs. The script will read MergeExcel.txt file located in the same folder and imports all worksheets into one workbook. The script is using VBA to open Excel and import worksheets.

You can also drag and drop excel files on top of this script file to merge them.

