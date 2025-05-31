Set fso = CreateObject("Scripting.FileSystemObject")
sConfigFilePath = GetFolderPath() & "\MergeExcel.txt"

if WScript.Arguments.Count > 0 then
    If WScript.Arguments.Count = 1 Then
    MsgBox "Please drag and drop more than one excel file on top of this script file."
    WScript.Quit
    End If
ElseIf fso.FileExists(sConfigFilePath) = False Then
    MsgBox "Could not file configuration file: " & sConfigFilePath & ". You can also drag and drop excel files on top of this script file."
    WScript.Quit
End If

Dim oExcel: Set oExcel = CreateObject("Excel.Application")
oExcel.Visible = True
oExcel.DisplayAlerts = false
Set oMasterWorkbook = oExcel.Workbooks.Add()
Set oMasterSheet = oMasterWorkbook.Worksheets("Sheet1")
oMasterSheet.Name = "temp_delete"

Deletesheet oMasterWorkbook, "Sheet2"
Deletesheet oMasterWorkbook, "Sheet3"

if WScript.Arguments.Count > 0 then
    MergeFromArguments
Else
    MergeFromFile sConfigFilePath
End If

Deletesheet oMasterWorkbook, "temp_delete"
MsgBox "Done"

Sub MergeFromArguments()
    For i = 0 to WScript.Arguments.Count - 1
      sFilePath = WScript.Arguments(i)
  
      If fso.FileExists(sFilePath) Then

        If fso.GetAbsolutePathName(sFilePath) <> sFilePath Then
          sFilePath = fso.GetAbsolutePathName(sFilePath)
        End If

        Set oWorkBook = oExcel.Workbooks.Open(sFilePath)
    
        For Each oSheet in oWorkBook.Worksheets
          oSheet.Copy oMasterSheet
        Next
    
        oWorkBook.Close()
      End If
    Next
End Sub

Sub MergeFromFile(sConfigFilePath)
    Set oFile = fso.OpenTextFile(sConfigFilePath, 1)   
    Do until oFile.AtEndOfStream
      sFilePath = Replace(oFile.ReadLine,"""","")
  
      If fso.FileExists(sFilePath) Then

        If fso.GetAbsolutePathName(sFilePath) <> sFilePath Then
          sFilePath = fso.GetAbsolutePathName(sFilePath)
        End If

        Set oWorkBook = oExcel.Workbooks.Open(sFilePath)
    
        For Each oSheet in oWorkBook.Worksheets
          oSheet.Copy oMasterSheet
        Next
    
        oWorkBook.Close()
      End If
    Loop
    oFile.Close
End Sub

Function GetFolderPath()
	Dim oFile 'As Scripting.File
	Set oFile = fso.GetFile(WScript.ScriptFullName)
	GetFolderPath = oFile.ParentFolder
End Function

Sub Deletesheet(oWorkbook, sSheetName)
  on error resume next
  oWorkbook.Worksheets(sSheetName).Delete
End Sub
