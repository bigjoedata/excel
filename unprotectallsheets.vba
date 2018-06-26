#from https://www.extendoffice.com/documents/excel/1154-excel-unprotect-multiple-sheets.html
Sub unprotect_all_sheets()
On Error Goto booboo
unpass = InputBox("password")
For Each Worksheet In ActiveWorkbook.Worksheets
Worksheet.Unprotect Password:=unpass
Next
Exit Sub
booboo: MsgBox "There is s problem - check your password, capslock, etc."
End Sub
