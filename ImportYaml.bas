Attribute VB_Name = "ImportYaml"
Dim dataTypeCol As Integer
Dim dataTypeLvlCol As Integer
Dim requiredCol As Integer
Dim requiredLvlCol As Integer
Dim uniqueCol As Integer
Dim uniqueLvlCol As Integer
Dim formatCol As Integer
Dim enumValuesCol As Integer
Dim regExCol As Integer
Dim formatLvlCol As Integer
Dim lowRangeCol As Integer
Dim highRangeCol As Integer
Dim rangeLvlCol As Integer
Dim fieldLengthCol As Integer

Sub ImportYamlFile(yamlFullPath As String)
  
On Error GoTo errHandler

  If right(yamlFullPath, 5) = ".yaml" Then

    Dim myFile As String
    Dim text As String
    Dim textLine As String
    Dim prevTextLine As String
    Dim i As Integer
    Dim ds As Worksheet
    Dim fieldType As String
    Dim baseRowCnt As Integer
    Dim filterRowCnt As Integer
    Dim concatRowCnt As Integer
    Dim codeRowCnt As Integer
    Dim concatFieldSplit() As String
    Dim headerRec() As String
    Dim splitDelimitor As String
    Dim headerMismatch As Boolean
    
    fieldType = ""
    baseRowCnt = 0
    filterRowCnt = 0
    concatRowCnt = 0
    codeRowCnt = 0
    headerMismatch = False
    
    i = 1
    myFile = yamlFullPath
    
    Open myFile For Input As #1
    
    'import yaml for to display in preview on home sheet
    Do Until EOF(1)
        Line Input #1, textLine
        text = text & textLine & vbCrLf
        i = i + 1
    Loop
    Close #1
    Sheets("Home").Cells(4, 4) = text

    'reset variables
    i = 1
    text = ""
    textLine = ""
    
    Open myFile For Input As #1

    Do Until EOF(1)
        Line Input #1, textLine
        textLine = Trim(textLine)
        text = textLine
        
        If InStr(textLine, "filetype: ") > 0 Then
            'set delimiter in control panel
            frmControlPanel.cbxDelimiter.text = removeConfigText(text, textLine, "filetype: ")
            Sheets("saved").Cells(1, 3) = removeConfigText(text, textLine, "filetype: ")
            frmControlPanel.optDelimited.Value = True
            Select Case Sheets("saved").Cells(1, 3)
                Case "PIPE"
                    splitDelimiter = "|"
                Case "TAB"
                    splitDelimiter = vbTab
                Case "CSV"
                    splitDelimiter = ","
                Case "FULLCSV"
                    splitDelimiter = ","
                Case "FIXEDWIDTH"
                    splitDelimiter = "|"
                    frmControlPanel.cbxDelimiter.text = ""
                    frmControlPanel.optDelimited.Value = False
                    frmControlPanel.optFixedWidth.Value = True
                    Sheets("saved").Cells(1, 3) = ""
            End Select
        ElseIf InStr(textLine, "vendor: ") > 0 Then
            'set vendor in control panel
            frmControlPanel.cbxVendor.text = removeConfigText(text, textLine, "vendor: ")
            Sheets("saved").Cells(2, 3) = removeConfigText(text, textLine, "vendor: ")
         ElseIf InStr(textLine, "rowlength: ") > 0 Then
            'set rowlength in control panel
            frmControlPanel.txtRowLength.text = removeConfigText(text, textLine, "rowlength: ")
            Sheets("saved").Cells(4, 3) = removeConfigText(text, textLine, "rowlength: ")
        ElseIf InStr(textLine, "header: ") > 0 Then
            'set header record in control panel
            frmControlPanel.txtHdrRec.text = Mid(Trim(Replace(text, "header: ", "")), 2, (Len(Trim(Replace(text, "header: ", ""))) - 2))
            Sheets("saved").Cells(3, 3) = Mid(Trim(Replace(text, "header: ", "")), 2, (Len(Trim(Replace(text, "header: ", ""))) - 2))
            
            Sheets("Base Fields").Select
            headerRec() = Split(Replace(removeQuotes(Trim(text)), "header: ", ""), splitDelimiter)

            For i = LBound(headerRec()) To UBound(headerRec())
                Sheets("Base Fields").Cells(i + startHdrFldRow, bFieldNameCol) = headerRec(i)
            Next i
            
        ElseIf InStr(textLine, "FieldMapping.Config:") > 0 Then
            'find field name
            Line Input #1, textLine
            'set to field name
            prevTextLine = Trim(textLine)
            
            Do Until EOF(1)
                Line Input #1, textLine
                If textLine = "       type: IGNORED" Or textLine = "       type: IGNORE" Then
                    Set ds = Sheets("Base Fields")
                    fieldType = "IGNORED"
                ElseIf textLine = "       type: SIMPLE" Then
                    Set ds = Sheets("Base Fields")
                    fieldType = "SIMPLE"
                ElseIf textLine = "       type: FILTER" Then
                    Set ds = Sheets("Filtered Fields")
                    fieldType = "FILTER"
                ElseIf textLine = "       type: CONCAT" Then
                    Set ds = Sheets("Concat Fields")
                    fieldType = "CONCAT"
                ElseIf textLine = "       type: CODED" Then
                    Set ds = Sheets("Coded Fields")
                    fieldType = "CODED"
                End If
                
                Call setSheetColumnValues(fieldType)
                
                Select Case fieldType
                    Case "IGNORED"
                        ds.Select
                        Do Until ds.Cells(baseRowCnt + startHdrFldRow, bFieldNameCol) = removeQuotes(Replace(prevTextLine, ":", ""))
                            baseRowCnt = baseRowCnt + 1
                        Loop
                        
                        ds.Cells(baseRowCnt + startHdrFldRow, bFieldNameCol) = removeQuotes(Replace(prevTextLine, ":", ""))
                        ds.Cells(baseRowCnt + startHdrFldRow, bDataTypeCol) = "IGNORED"
                        
                        filePost = Seek(1)
                        
                        Line Input #1, textLine
                        text = textLine
                            
                        If InStr(textLine, "fixed:") > 0 Then
                            ds.Cells(baseRowCnt + startHdrFldRow, bEndIndexCol) = removeConfigText(text, textLine, "length: ")
                        Else
                            Seek 1, filePost
                        End If
                        
                        fieldType = ""
                        baseRowCnt = baseRowCnt + 1
                    Case "SIMPLE"
                        ds.Select
                        Do Until ds.Cells(baseRowCnt + startHdrFldRow, bFieldNameCol) = removeQuotes(Replace(prevTextLine, ":", ""))
                            If ds.Cells(baseRowCnt + startHdrFldRow, bFieldNameCol) = "" Then
                                headerMismatch = True
                                Exit Do
                            Else
                                baseRowCnt = baseRowCnt + 1
                            End If
                        Loop
                            If ds.Cells(baseRowCnt + startHdrFldRow, bFieldNameCol) = removeQuotes(Replace(prevTextLine, ":", "")) Then
                                ds.Cells(baseRowCnt + startHdrFldRow, bFieldNameCol) = removeQuotes(Replace(prevTextLine, ":", ""))
                                Do Until fieldType <> "SIMPLE" Or EOF(1)
                                    Line Input #1, textLine
                                    text = textLine
                                    
                                    If InStr(textLine, "fixed:") > 0 Then
                                        ds.Cells(baseRowCnt + startHdrFldRow, bEndIndexCol) = removeConfigText(text, textLine, "length: ")
                                    ElseIf InStr(textLine, "histogram:") > 0 Then
                                        ds.Cells(baseRowCnt + startHdrFldRow, bHistoCol) = removeQuotes(Replace(Trim(text), "histogram: ", ""))
                                    ElseIf InStr(textLine, "validation:") > 0 Then
                                        Do Until textLine = "          ]"
                                            Line Input #1, textLine
                                            text = textLine
                                            
                                            'validation types
                                            Call populateValidation(ds, baseRowCnt + startHdrFldRow, text)
                                        Loop
                                        fieldType = ""
                                    End If
                                    
                                    filePost = Seek(1)

                                    Line Input #1, textLine
                                    text = textLine
                                    If InStr(textLine, "fixed:") = 0 And InStr(textLine, "validation:") = 0 Then
                                        fieldType = ""
                                    End If
                                        
                                    Seek 1, filePost
                                Loop
                            End If
                    Case "FILTER"
                        ds.Select
                        Do Until fieldType <> "FILTER" Or EOF(1)
                            Line Input #1, textLine
                            text = textLine
                                                        
                            If InStr(textLine, "filter:") > 0 Then
                                ds.Cells(filterRowCnt + startHdrFldRow, fFilterFieldCol) = Replace(Replace(Mid(text, InStrRev(textLine, "field: "), ((InStr(textLine, ",") - InStrRev(textLine, "field: "))) - 1), """", ""), "'", "")
                                ds.Cells(filterRowCnt + startHdrFldRow, fFilterCol) = Replace(Replace(Mid(text, InStrRev(textLine, "condition: "), ((InStrRev(textLine, ",") - InStrRev(textLine, "condition: "))) - 1), """", ""), "'", "")
                                ds.Cells(filterRowCnt + startHdrFldRow, fFilterValueCol) = Replace(Replace(Mid(text, InStrRev(textLine, "value: "), ((InStrRev(textLine, "}") - InStrRev(textLine, "value: "))) - 1), """", ""), "'", "")
                            ElseIf InStr(textLine, "histogram:") > 0 Then
                                ds.Cells(filterRowCnt + startHdrFldRow, fHistoCol) = removeQuotes(Replace(Trim(text), "histogram: ", ""))
                            ElseIf InStr(textLine, "validation:") > 0 Then
                                Do Until textLine = "          ]"
                                    Line Input #1, textLine
                                    text = textLine
                                    
                                    'validation types
                                    Call populateValidation(ds, filterRowCnt + startHdrFldRow, text)
                                Loop
                                fieldType = ""
                            End If

                            filePost = Seek(1)

                            Line Input #1, textLine
                            text = textLine
                            If InStr(textLine, "histogram:") = 0 And InStr(textLine, "filter:") = 0 And InStr(textLine, "validation:") = 0 Then
                                fieldType = ""
                                filterRowCnt = filterRowCnt + 1
                            End If
                                        
                            Seek 1, filePost
                        Loop
                        
                    Case "CONCAT"
                        
                        ds.Select
                        Do Until fieldType <> "CONCAT" Or EOF(1)
                            Line Input #1, textLine
                            text = textLine
                            
                            If InStr(textLine, "fields:") > 0 Then
                                concatFieldSplit() = Split(Replace(Mid(text, InStrRev(textLine, "fields: "), ((InStr(textLine, ",") - InStrRev(textLine, "fields: "))) - 1), """", ""), "|")
                                For x = LBound(concatFieldSplit()) To UBound(concatFieldSplit())
                                    ds.Cells(concatRowCnt + startHdrFldRow, cFieldName1Col + x) = concatFieldSplit(x)
                                Next x
                                ds.Cells(concatRowCnt + startHdrFldRow, cOutputDelimiterCol) = Replace(Replace(Mid(text, InStrRev(textLine, "delimiter: "), ((InStrRev(textLine, "}") - InStrRev(textLine, "delimiter: "))) - 1), """", ""), "'", "")
                            ElseIf InStr(textLine, "histogram:") > 0 Then
                                ds.Cells(concatRowCnt + startHdrFldRow, cHistoCol) = removeQuotes(Replace(Trim(text), "histogram: ", ""))
                            ElseIf InStr(textLine, "validation:") > 0 Then
                                Do Until textLine = "          ]"
                                    Line Input #1, textLine
                                    text = textLine
                                    
                                    'validation types
                                    Call populateValidation(ds, concatRowCnt + startHdrFldRow, text)
                                Loop
                                fieldType = ""
                            End If
                            
                            filePost = Seek(1)

                            Line Input #1, textLine
                            text = textLine
                            If InStr(textLine, "histogram:") = 0 And InStr(textLine, "fields:") = 0 And InStr(textLine, "validation:") = 0 Then
                                fieldType = ""
                                concatRowCnt = concatRowCnt + 1
                            End If
                                        
                            Seek 1, filePost
                        Loop
                    Case "CODED"
                        ds.Select
                        'ds.Cells(codeRowCnt + startHdrFldRow, cdConfigFieldNameCol) = Replace(prevTextLine, ":", "")
                        Do Until fieldType <> "CODED"
                            If Not EOF(1) Then
                                Line Input #1, textLine
                                text = textLine
                            End If
                            text = textLine
                            
                            If InStr(textLine, "code_id_field:") > 0 Then
                                ds.Cells(codeRowCnt + startHdrFldRow, cdCodeIDCol) = removeConfigText(text, textLine, "code_id_field: ")
                            ElseIf InStr(textLine, "code_system_id_field:") > 0 Then
                                ds.Cells(codeRowCnt + startHdrFldRow, cdCodeSysIDCol) = removeConfigText(text, textLine, "code_system_id_field: ")
                            ElseIf InStr(textLine, "code_display_field:") > 0 Then
                                ds.Cells(codeRowCnt + startHdrFldRow, cdCodeDisplayCol) = removeConfigText(text, textLine, "code_display_field: ")
                            ElseIf InStr(textLine, "concept:") > 0 Then
                                ds.Cells(codeRowCnt + startHdrFldRow, cdConceptCol) = Replace(removeConfigText(text, textLine, "concept: "), frmControlPanel.cbxVendor & ":", "")
                            ElseIf InStr(textLine, "default_coding_system_id:") > 0 Then
                                
                                With Sheets("saved").Range("M:N") 'Mid(text, (InStrRev(textLine, "default_coding_system_id: ")), ((Len(textLine) + 1) - InStrRev(textLine, "default_coding_system_id: "))),
                                    Set Rng = .Find(What:=removeConfigText(text, textLine, "default_coding_system_id: "), _
                                                    After:=.Cells(.Cells.Count), _
                                                    LookIn:=xlValues, _
                                                    LookAt:=xlWhole, _
                                                    SearchOrder:=xlByRows, _
                                                    SearchDirection:=xlNext, _
                                                    MatchCase:=False)
                                    If Not Rng Is Nothing Then
                                        ds.Cells(codeRowCnt + startHdrFldRow, cdDefaultCodeSysCol) = Sheets("saved").Cells(Rng.Row, 13)
                                    End If
                                End With
                            Else
                                fieldType = ""
                                codeRowCnt = codeRowCnt + 1
                            End If
                        Loop
                End Select
                prevTextLine = Trim(textLine)
            Loop
        End If
        
        i = i + 1
    Loop
    Close #1
    
    Sheets("Home").Select
    
    If headerMismatch = False Then
        frmControlPanel.Hide
        MsgBox yamlFullPath & vbCrLf & vbCrLf & " Imported Successfully", vbExclamation
    Else
        MsgBox "Header Record and validation fields do not match", vbCritical, "Header Mismatch"
    End If
  Else
    MsgBox "Not Valid Yaml file", vbInformation
  End If
  
  Exit Sub
  
errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)

End Sub

Sub populateValidation(ds As Worksheet, _
                        rowNbr As Integer, _
                        textLine As String)
  
On Error GoTo errHandler
  
  Dim text As String
  
  text = textLine
  If InStr(textLine, "dateformat") > 0 Then
    If InStrRev(textLine, "format: ") > 0 Then
        ds.Cells(rowNbr, formatCol) = removeValidationText(text, textLine, "format: ")
        ds.Cells(rowNbr, formatLvlCol) = removeValidationText(text, textLine, "severity: ")
    Else
        ds.Cells(rowNbr, formatCol) = "ISO"
        ds.Cells(rowNbr, formatLvlCol) = removeValidationText(text, textLine, "severity: ")
    End If
  ElseIf InStr(textLine, "numbertype") > 0 Then
    ds.Cells(rowNbr, dataTypeCol) = removeValidationText(text, textLine, "value: ")
    ds.Cells(rowNbr, dataTypeLvlCol) = removeValidationText(text, textLine, "severity: ")
  ElseIf InStr(textLine, "required") > 0 Then
    ds.Cells(rowNbr, requiredCol) = "Y"
    ds.Cells(rowNbr, requiredLvlCol) = removeValidationText(text, textLine, "severity: ")
  ElseIf InStr(textLine, "regex") > 0 Then
    ds.Cells(rowNbr, regExCol) = removeValidationText(text, textLine, "value: ")
    ds.Cells(rowNbr, formatLvlCol) = removeValidationText(text, textLine, "severity: ")
  ElseIf InStr(textLine, "unique") > 0 Then
    ds.Cells(rowNbr, uniqueCol) = "Y"
    ds.Cells(rowNbr, uniqueLvlCol) = removeValidationText(text, textLine, "severity: ")
  ElseIf InStr(textLine, "daterange") > 0 Then
    If InStrRev(textLine, "low: ") > 0 Then
        ds.Cells(rowNbr, lowRangeCol) = removeValidationText(text, textLine, "low: ")
    End If
    If InStrRev(textLine, "high: ") > 0 Then
        ds.Cells(rowNbr, highRangeCol) = removeValidationText(text, textLine, "high: ")
    End If
    ds.Cells(rowNbr, rangeLvlCol) = removeValidationText(text, textLine, "severity: ")
  ElseIf InStr(textLine, "numberrange") > 0 Then
    ds.Cells(rowNbr, lowRangeCol) = removeValidationText(text, textLine, "low: ")
    ds.Cells(rowNbr, highRangeCol) = removeValidationText(text, textLine, "high: ")
    ds.Cells(rowNbr, rangeLvlCol) = removeValidationText(text, textLine, "severity: ")
  ElseIf InStr(textLine, "allowedvalue") > 0 Then
    ds.Cells(rowNbr, enumValuesCol) = removeValidationText(text, textLine, "value: ")
    ds.Cells(rowNbr, formatLvlCol) = removeValidationText(text, textLine, "severity: ")
  End If

  Exit Sub
  
errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)

End Sub

Sub setSheetColumnValues(fieldType As String)

On Error GoTo errHandler

  fieldLengthCol = bEndIndexCol

  Select Case fieldType
    Case "SIMPLE"
        dataTypeCol = bDataTypeCol
        dataTypeLvlCol = bDataTypeLvlCol
        requiredCol = bRequiredCol
        requiredLvlCol = bRequiredLvlCol
        uniqueCol = bUniqueCol
        uniqueLvlCol = bUniqueLvlCol
        formatCol = bFormatCol
        enumValuesCol = bEnumValuesCol
        regExCol = bRegExCol
        formatLvlCol = bFormatLvlCol
        lowRangeCol = bLowRangeCol
        highRangeCol = bHighRangeCol
        rangeLvlCol = bRangeLvlCol
    Case "FILTER"
        dataTypeCol = fDataTypeCol
        dataTypeLvlCol = fDataTypeLvlCol
        requiredCol = fRequiredCol
        requiredLvlCol = fRequiredLvlCol
        uniqueCol = fUniqueCol
        uniqueLvlCol = fUniqueLvlCol
        formatCol = fFormatCol
        enumValuesCol = fEnumValuesCol
        regExCol = fRegExCol
        formatLvlCol = fFormatLvlCol
        lowRangeCol = fLowRangeCol
        highRangeCol = fHighRangeCol
        rangeLvlCol = fRangeLvlCol
    Case "CONCAT"
        dataTypeCol = cDataTypeCol
        dataTypeLvlCol = cDataTypeLvlCol
        requiredCol = cRequiredCol
        requiredLvlCol = cRequiredLvlCol
        uniqueCol = cUniqueCol
        uniqueLvlCol = cUniqueLvlCol
        formatCol = cFormatCol
        enumValuesCol = cEnumValuesCol
        regExCol = cRegExCol
        formatLvlCol = cFormatLvlCol
        lowRangeCol = cLowRangeCol
        highRangeCol = cHighRangeCol
        rangeLvlCol = cRangeLvlCol
  End Select

  Exit Sub
  
errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)

End Sub

Public Function InStrRev(String1 As String, String2 As String, _
   Optional MatchCase As Boolean = False, _
   Optional ReverseCount As Boolean = False) _
   As Long

On Error GoTo errHandler

  'Test if the function must match the string case
  'If true, converts all strings to upper case
  If MatchCase = False Then
    String1 = UCase(String1)
    String2 = UCase(String2)
  End If

  'Initiate a loop for each character
  'Test from the position i with second string length
  'If it matches, then store the position as the last position found
  For i = 1 To Len(String1)
    If Mid(String1, i, Len(String2)) = String2 Then
        Last_Found = i + Len(String2)
    End If
    'MsgBox Mid(String1, i, Len(String2))
  Next

  If ReverseCount And Last_Found > 0 Then
 'Inverts counting to backwards
    'InStrRev = Len(String1) - Last_Found + 1
    InStrRev = Last_Found
  Else
 'If ReverseCount was not requested and/or the
  'last position found is equal to zero
  'Passes the found position as is
    InStrRev = Last_Found
  End If

  Exit Function
  
errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)

End Function

Public Function removeQuotes(text As String)

On Error GoTo errHandler

  removeQuotes = Replace(Replace(Replace(Replace(text, """", ""), "'", ""), "}", ""), "{", "")
  
  Exit Function
  
errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)

End Function

Public Function removeConfigText(text As String, textLine As String, configText As String)

On Error GoTo errHandler

  removeConfigText = removeQuotes(Mid(text, InStrRev(textLine, configText), ((Len(textLine) + 1) - InStrRev(textLine, configText))))

  Exit Function
  
errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)

End Function

Public Function removeValidationText(text As String, textLine As String, validationText As String)

On Error GoTo legacyYaml

  removeValidationText = removeQuotes(Mid(text, InStrRev(textLine, validationText), (InStr((InStrRev(textLine, validationText) + 2), textLine, "'") - InStrRev(textLine, validationText))))

  Exit Function
  
legacyYaml:

On Error GoTo errHandler
    
  removeValidationText = removeQuotes(Mid(text, InStrRev(textLine, validationText), (InStr((InStrRev(textLine, validationText) + 2), textLine, """") - InStrRev(textLine, validationText))))

  Exit Function

errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)

End Function
