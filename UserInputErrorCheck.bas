Attribute VB_Name = "UserInputErrorCheck"
Sub checkUserInput()

On Error GoTo errHandler

  Call checkSheetRow(Sheets("Base Fields"), bFieldNameCol, bFormatCol, bEnumValuesCol, bRegExCol, bDataTypeCol, bLowRangeCol, bHighRangeCol)
  Call checkSheetRow(Sheets("Filtered Fields"), fComputedFieldNameCol, fFormatCol, fEnumValuesCol, fRegExCol, fDataTypeCol, fLowRangeCol, fHighRangeCol)
  Call checkSheetRow(Sheets("Concat Fields"), cFieldName1Col, cFormatCol, cEnumValuesCol, cRegExCol, cDataTypeCol, cLowRangeCol, cHighRangeCol)
 
 Exit Sub
 
errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)
  
End Sub

Sub checkControlPanelInput(optDelimiter As Boolean, _
                            delimiter As String, _
                            optFixedWidth As Boolean, _
                            rowLength As String, _
                            vendor As String, _
                            headerRec As String)
  
On Error GoTo errHandler

  If optDelimiter = True Then
  
    If delimiter = "" Then
      MsgBox "Unable to determine File Delimiter", vbCritical
      userInputErrChk = True
      Exit Sub
    End If
  
  Else
    
    If rowLength = "" Then
      MsgBox "Unable to determine Row Length", vbCritical
      userInputErrChk = True
      Exit Sub
    End If
    
    If InStr(1, headerRec, "|") = 0 Then
        MsgBox "Please use PIPE delimiter in Header Row", vbCritical
        userInputErrChk = True
        Exit Sub
    End If
  
  End If
    
  If vendor = "" Then
    MsgBox "Unable to determine Vendor", vbCritical
    userInputErrChk = True
    Exit Sub
  End If
  
  If headerRec = "" Then
    MsgBox "Unable to determine Header Record", vbCritical
    userInputErrChk = True
    Exit Sub
  End If
  
  Exit Sub
 
errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)
  
End Sub

Sub checkSheetRow(ds As Worksheet, _
                fieldNameCol As Integer, _
                Optional formatCol As Integer, _
                Optional enumValuesCol As Integer, _
                Optional regExCol As Integer, _
                Optional dataTypeCol As Integer, _
                Optional lowRangeCol As Integer, _
                Optional highRangeCol As Integer)

On Error GoTo errHandler

  Dim i As Integer
  
  i = 0
  While ds.Cells(i + startHdrFldRow, fieldNameCol) <> ""
    With ds
        .Select
        'check for use of format, enum, & regex use at same time
        If WorksheetFunction.CountA(.Range(.Cells(i + startHdrFldRow, formatCol), .Cells(i + startHdrFldRow, regExCol))) > 1 Then
            MsgBox "Can only use 1 of these at a time: " & vbCrLf & "Format, Enum Value, or RegEx" & vbCrLf & vbCrLf & "Please check sheet: " & ds.Name & vbCrLf & "Row: " & i + startHdrFldRow, vbCritical
            userInputErrChk = True
            Exit Sub
        End If
        'check for range without format or datatype set
        If WorksheetFunction.CountA(.Range(.Cells(i + startHdrFldRow, lowRangeCol), .Cells(i + startHdrFldRow, highRangeCol))) > 1 And ds.Cells(i + startHdrFldRow, dataTypeCol) = "" And ds.Cells(i + startHdrFldRow, formatCol) = "" Then
            MsgBox "Unable to determine range data type " & vbCrLf & "Either select data type or specify date format" & vbCrLf & vbCrLf & "Please check sheet: " & ds.Name & vbCrLf & "Row: " & i + startHdrFldRow, vbCritical
            userInputErrChk = True
            Exit Sub
        End If
    End With
    i = i + 1
  Wend

 Exit Sub
 
errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)
  
End Sub
