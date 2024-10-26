Attribute VB_Name = "ProcessHeaderRecord"
Sub processHdrRec(dataSheet As Worksheet, _
                startColumn As String, _
                endColumn As String, _
                fieldNameCol As Integer, _
                fixedWidth As Boolean, _
                dataTypeCol As Integer)
  
On Error GoTo errHandler
  
  Dim hdrRecord As String
  Dim hdrRecordSplit() As String

  Dim i As Integer
  
  hdrRecord = frmControlPanel.txtHdrRec.text
  
  Select Case frmControlPanel.cbxDelimiter.text
    Case "PIPE"
        hdrRecordSplit() = Split(hdrRecord, "|")
    Case "TAB"
        hdrRecordSplit() = Split(hdrRecord, vbTab)
    Case "CSV"
        hdrRecordSplit() = Split(hdrRecord, ",")
    Case "FULLCSV"
        hdrRecordSplit() = Split(hdrRecord, ",")
    Case Else
        hdrRecordSplit() = Split(hdrRecord, "|")
  End Select
  
  'populate all rule tabs with header fields
  For i = LBound(hdrRecordSplit) To UBound(hdrRecordSplit)
    Call fmtHeaderToSheets(startColumn, endColumn, i + startHdrFldRow, hdrRecordSplit(i), dataSheet, fieldNameCol, fixedWidth, dataTypeCol)
  Next i
  
  Exit Sub
  
errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)
  
End Sub

Sub fmtHeaderToSheets(startColumn As String, _
                    endColumn As String, _
                    startRow As Integer, _
                    hdrValue As String, _
                    dataSheet As Worksheet, _
                    fieldNameCol As Integer, _
                    fixedWidth As Boolean, _
                    dataTypeCol As Integer)
  
On Error GoTo errHandler
  
  Dim formatRange As String

  formatRange = startColumn & startRow & ":" & endColumn & startRow
  dataSheet.Cells(startRow, fieldNameCol) = Replace(hdrValue, """", "")
  If fixedWidth = True Then
    dataSheet.Cells(startRow, dataTypeCol) = "IGNORED"
  End If
  dataSheet.Select
  Range(formatRange).Select
  With Selection.Interior
    .Pattern = xlNone
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With

  Exit Sub
  
errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)
  
End Sub

Sub clearSheets(dataSheet As Worksheet, _
                fieldNameColumn As Integer, _
                startColumn As Integer, _
                endColumn As Integer)
  
On Error GoTo errHandler
  
  Dim i As Integer
  
  i = 0
  dataSheet.Select
  While dataSheet.Cells(i + startHdrFldRow, fieldNameColumn) <> ""
    formatRange = "A" & i + startHdrFldRow & ":Q" & i + startHdrFldRow
    With dataSheet
        .Range(.Cells(i + startHdrFldRow, startColumn), .Cells(i + startHdrFldRow, endColumn)).Select
        Selection.ClearContents
    End With
    i = i + 1
  Wend
  
  dataSheet.Cells(1, 1).Select
  
  Exit Sub
  
errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)
  
End Sub
