if WScript.Arguments.Count = 0 then
    MsgBox "Please drag and drop a folder on top of this script file to merge sheets into single sheet."
    WScript.Quit
End If

Set fso = CreateObject("Scripting.FileSystemObject")
sFolderePath = WScript.Arguments(0)

If fso.FolderExists(sFolderePath) = False Then
  MsgBox "Could not find folder: " & sFolderePath 
  WScript.Quit
End If

If MsgBox("Merge worksheets for this folder: " & sFolderePath, vbYesNo + vbQuestion) = vbNo Then
  WScript.Quit
End If

Set oFolder = fso.GetFolder(sFolderePath)

Dim oExcel: Set oExcel = CreateObject("Excel.Application")
oExcel.Visible = True
oExcel.DisplayAlerts = false

Set oMasterWorkbook = oExcel.Workbooks.Add()
Set oCombined = oMasterWorkbook.Worksheets("Sheet1")
iRowOffset = 0

For Each oFile In oFolder.Files
    If oFile.Attributes And 2 Then
        'Hidden
    Else

        Set oWorkbook = oExcel.Workbooks.Open(oFile.Path)
        Set oSheet = oWorkbook.Worksheets(1)
        iRowsCount = GetLastRowWithData(oSheet)

        If iRowOffset = 0 Then 
            iStartRow = 4
        Else 
            iStartRow = 5
        end if

        oSheet.Range(oSheet.Cells(iStartRow, 1), oSheet.Cells(iRowsCount, oSheet.UsedRange.Columns.Count)).Copy
        oCombined.Activate
        oCombined.Cells(iRowOffset + 1, 1).Select
        oCombined.Paste       

        iRowOffset = iRowOffset + iRowsCount - iStartRow + 1
        oWorkbook.Close
    End If
Next

MsgBox "Done!"

Function GetLastRowWithData(oSheet)
    iMaxRow = oSheet.UsedRange.Rows.Count
    If iMaxRow > 500 Then
        iMaxRow = oSheet.Cells.Find("*", oSheet.Cells(1, 1),  -4163, , 1, 2).Row
    End If

    For iRow = iMaxRow to 1 Step -1
         For iCol = 1 to oSheet.UsedRange.Columns.Count
            If Trim(oSheet.Cells(iRow, iCol).Value) <> "" Then
                GetLastRowWithData = iRow
                Exit Function
            End If
         Next
    Next
    GetLastRowWithData = 1
End Function
