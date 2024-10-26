Attribute VB_Name = "ErrorHandler"
Sub errMsg(errNbr As Integer, _
            errDescr As String, _
            codeModule As String)

  MsgBox "Error in Module: " & codeModule & vbCrLf & _
    "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbCritical, "Error"
    
  End
  
End Sub
