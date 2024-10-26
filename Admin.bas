Attribute VB_Name = "Admin"
Public Sub deleteunlocksheet(dataSheet As Worksheet)
'unlock the worksheet

On Error Resume Next
 dataSheet.Unprotect Password:="Cerner"

End Sub
Public Sub deletelocksheet(dataSheet As Worksheet)
'lock the worksheet

On Error Resume Next
 dataSheet.Protect Password:="Cerner"

End Sub

Sub saveControlPanelValues(delimiter As String, vendor As String, headerRec As String, rowLength As String)
  Sheets("saved").Cells(1, 3) = delimiter
  Sheets("saved").Cells(2, 3) = vendor
  Sheets("saved").Cells(3, 3) = headerRec
  Sheets("saved").Cells(4, 3) = rowLength
End Sub

Sub resetControlPanelValues()
  Sheets("saved").Cells(1, 3) = ""
  Sheets("saved").Cells(2, 3) = ""
  Sheets("saved").Cells(3, 3) = ""
  Sheets("saved").Cells(4, 3) = ""
End Sub

