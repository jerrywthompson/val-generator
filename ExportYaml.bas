Attribute VB_Name = "ExportYaml"
Dim yamlText As String

Sub ExportYamlFile(previewMode As Boolean)
  
On Error GoTo errHandler

  Dim fileType As String
  Dim vendor As String
  Dim headerRec As String
  Dim rowLength As String
  
  fileType = frmControlPanel.cbxDelimiter
  vendor = frmControlPanel.cbxVendor
  headerRec = frmControlPanel.txtHdrRec
  rowLength = frmControlPanel.txtRowLength
  
  yamlText = createFileMeta(gblGeneratorVersion, fileType, vendor, headerRec, rowLength) & "FieldMapping.Config:" & vbCrLf
  
  'loop through and build validation for Base Fields
  Call buildFieldRules(Sheets("Base Fields"), _
                        baseFieldRangeRuleStart, baseFieldRangeEnd, _
                        bFieldNameCol, bDataTypeCol, bDataTypeLvlCol, bRequiredCol, bRequiredLvlCol, _
                        bUniqueCol, bUniqueLvlCol, bFormatCol, bEnumValuesCol, bRegExCol, _
                        bFormatLvlCol, bLowRangeCol, bHighRangeCol, bRangeLvlCol, bHistoCol, _
                        , , , , , , , , , _
                        bEndIndexCol)

  'loop through and build validation for Filtered Fields
  Call buildFieldRules(Sheets("Filtered Fields"), _
                        filterFieldRangeRuleStart, filterFieldRangeEnd, _
                        fComputedFieldNameCol, fDataTypeCol, fDataTypeLvlCol, fRequiredCol, fRequiredLvlCol, _
                        fUniqueCol, fUniqueLvlCol, fFormatCol, fEnumValuesCol, fRegExCol, _
                        fFormatLvlCol, fLowRangeCol, fHighRangeCol, fRangeLvlCol, fHistoCol, _
                        fFilterFieldCol, fFilterCol, fFilterValueCol, _
                        , , , , , , _
                        100)

  'loop through and build validation for Concatenated Fields
  Call buildFieldRules(Sheets("Concat Fields"), _
                        concatFieldRangeRuleStart, concatFieldRangeEnd, _
                        cFieldName1Col, cDataTypeCol, cDataTypeLvlCol, cRequiredCol, cRequiredLvlCol, _
                        cUniqueCol, cUniqueLvlCol, cFormatCol, cEnumValuesCol, cRegExCol, _
                        cFormatLvlCol, cLowRangeCol, cHighRangeCol, cRangeLvlCol, cHistoCol, _
                        , , , _
                        cFieldName1Col, cFieldName2Col, cFieldName3Col, cFieldName4Col, cFieldName5Col, _
                        cOutputDelimiterCol, _
                        100)
    
  'loop through and build Coded Fields
  Call buildCodedFieldRules(Sheets("Coded Fields"), _
                        cdConfigFieldNameCol, _
                        cdConceptCol, _
                        cdCodeIDCol, _
                        cdCodeSysIDCol, _
                        cdCodeDisplayCol, _
                        cdDefaultCodeSysCol, _
                        vendor)
  
  Sheets("Home").Select
  
  If previewMode = True Then
    Sheets("Home").Cells(4, 4) = yamlText
  Else
      Dim varResult As Variant
      'displays the save file dialog
      varResult = Application.GetSaveAsFilename(InitialFileName:="qualityReports", _
        fileFilter:="Yaml Files (*.yaml), *.yaml")
    
      'checks to make sure the user hasn't canceled the dialog
      If varResult <> False Then
        Call saveFile(varResult, yamlText)
      End If
  End If

  Exit Sub
  
errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)
  
End Sub

Sub saveFile(strPath As Variant, _
            content As String)

On Error GoTo errHandler

  Dim fso As Object
  Set fso = CreateObject("Scripting.FileSystemObject")
  Dim oFile As Object
  Set oFile = fso.CreateTextFile(strPath)
  oFile.WriteLine content
  oFile.Close
  Set fso = Nothing
  Set oFile = Nothing

  Exit Sub
  
errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)
  
End Sub

Sub buildCodedFieldRules(ds As Worksheet, _
                        configFieldNameCol As Integer, _
                        conceptCol As Integer, _
                        codeIDCol As Integer, _
                        codeSysIDCol As Integer, _
                        codeDisplayCol As Integer, _
                        defaultCodeSysCol As Integer, _
                        vendor As String)

On Error GoTo errHandler
  
  Dim i As Integer

  i = 0
  While ds.Cells(i + startHdrFldRow, configFieldNameCol) <> ""
    ds.Select
    yamlText = yamlText & "     " & ds.Cells(i + startHdrFldRow, configFieldNameCol) & ":" & vbCrLf & _
               "       type: CODED" & vbCrLf & _
               "       code_id_field: " & ds.Cells(i + startHdrFldRow, codeIDCol) & vbCrLf
    If ds.Cells(i + startHdrFldRow, codeSysIDCol) <> "" Then
        yamlText = yamlText & "       code_system_id_field: " & ds.Cells(i + startHdrFldRow, codeSysIDCol) & vbCrLf
    End If
    If ds.Cells(i + startHdrFldRow, codeDisplayCol) <> "" Then
        yamlText = yamlText & "       code_display_field: " & ds.Cells(i + startHdrFldRow, codeDisplayCol) & vbCrLf
    End If
    If ds.Cells(i + startHdrFldRow, conceptCol) <> "" Then
        yamlText = yamlText & "       concept: " & vendor & ":" & ds.Cells(i + startHdrFldRow, conceptCol) & vbCrLf
    End If
    If ds.Cells(i + startHdrFldRow, defaultCodeSysCol) <> "" Then
        With Sheets("saved").Range("M:N")
            Set Rng = .Find(What:=ds.Cells(i + startHdrFldRow, defaultCodeSysCol), _
                            After:=.Cells(.Cells.Count), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
            If Not Rng Is Nothing Then
                yamlText = yamlText & "       default_code_system_id: " & Sheets("saved").Cells(Rng.Row, 14) & vbCrLf
            End If
        End With
    End If
    i = i + 1
  Wend
  
  Exit Sub
  
errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)
  
End Sub
Sub buildFieldRules(ds As Worksheet, _
                    fieldRangeRuleStart As String, fieldRangeEnd As String, _
                    fieldNameCol As Integer, _
                    dataTypeCol As Integer, dataTypeLvlCol As Integer, _
                    requiredCol As Integer, requiredLvlCol As Integer, _
                    uniqueCol As Integer, uniqueLvlCol As Integer, _
                    formatCol As Integer, enumValuesCol As Integer, regExCol As Integer, formatLvlCol As Integer, _
                    lowRangeCol As Integer, highRangeCol As Integer, rangeLvlCol As Integer, _
                    histoCol As Integer, _
                    Optional filterFieldCol As Integer, _
                    Optional filterCol As Integer, _
                    Optional filterValueCol As Integer, _
                    Optional concatFieldName1Col As Integer, _
                    Optional concatFieldName2Col As Integer, _
                    Optional concatFieldName3Col As Integer, _
                    Optional concatFieldName4Col As Integer, _
                    Optional concatFieldName5Col As Integer, _
                    Optional concatOutputDelimiterCol As Integer, _
                    Optional fixedFieldLengthCol As Integer)
  
On Error GoTo errHandler

  Dim arrConcat As Variant
  Dim varConcat As String
  Dim i As Integer
  Dim ii As Integer
  
  i = 0
  ii = 0
  While ds.Cells(i + startHdrFldRow, fieldNameCol) <> ""
    ds.Select
    If ds.Cells(i + startHdrFldRow, dataTypeCol) = "IGNORE" Or ds.Cells(i + startHdrFldRow, dataTypeCol) = "IGNORED" Then
        'sets field to ignore is selected on base fields sheet
        yamlText = yamlText & "     " & ds.Cells(i + startHdrFldRow, fieldNameCol) & ":" & vbCrLf & _
                              "       type: IGNORED" & vbCrLf
        
        'build fixed width field length
        If ds.Cells(i + startHdrFldRow, fixedFieldLengthCol) <> "" And frmControlPanel.optFixedWidth.Value = True Then
            yamlText = yamlText & "       fixed: {length: " & Chr(39) & ds.Cells(i + startHdrFldRow, fixedFieldLengthCol) & Chr(39) & "}" & vbCrLf
        End If
    Else
        If WorksheetFunction.CountA(Range(fieldRangeRuleStart & (i + startHdrFldRow) & ":" & fieldRangeEnd & (i + startHdrFldRow))) > 0 Then
            'only used for concat sheet
            If concatFieldName1Col + concatFieldName2Col + concatFieldName3Col + concatFieldName4Col + concatFieldName5Col > 0 Then
                varConcat = ""
                'define array:
                arrConcat = Array(ds.Cells(i + startHdrFldRow, concatFieldName1Col), ds.Cells(i + startHdrFldRow, concatFieldName2Col), ds.Cells(i + startHdrFldRow, concatFieldName3Col), ds.Cells(i + startHdrFldRow, concatFieldName4Col), ds.Cells(i + startHdrFldRow, concatFieldName5Col))
                'concatenate each element of the array:
                For ii = 0 To UBound(arrConcat)
                'Concatenate with |:
                    If arrConcat(ii) <> "" Then
                        varConcat = varConcat & "|" & arrConcat(ii)
                        arrConcat(ii) = ""
                    End If
                Next ii
                varConcat = Mid(varConcat, 2)
            End If
            
            'call createFieldType to build the field type & field attributes
            yamlText = yamlText & createFieldType(ds, i + startHdrFldRow, fieldNameCol, filterFieldCol, filterCol, filterValueCol, varConcat, concatOutputDelimiterCol)
            
            'build histogram if selected
            If ds.Cells(i + startHdrFldRow, histoCol) <> "" Then
                yamlText = yamlText & "       histogram: " & Chr(39) & "Y" & Chr(39) & vbCrLf
            End If
            
            'build fixed width field length
            If ds.Cells(i + startHdrFldRow, fixedFieldLengthCol) <> "" And frmControlPanel.optFixedWidth.Value = True Then
                yamlText = yamlText & "       fixed: {length: " & Chr(39) & ds.Cells(i + startHdrFldRow, fixedFieldLengthCol) & Chr(39) & "}" & vbCrLf
            End If
            
            'check to see if any validation rules are defined
            With ds
                If WorksheetFunction.CountA(.Range(.Cells(i + startHdrFldRow, dataTypeCol), .Cells(i + startHdrFldRow, rangeLvlCol))) > 0 Then
                    'build all validation rules
                    yamlText = yamlText & "       validation: [" & vbCrLf
                    yamlText = yamlText & createValidation(ds, formatCol, i + startHdrFldRow, "dateformat", formatLvlCol)
                    yamlText = yamlText & createValidation(ds, dataTypeCol, i + startHdrFldRow, "numbertype", dataTypeLvlCol)
                    yamlText = yamlText & createValidation(ds, requiredCol, i + startHdrFldRow, "required", requiredLvlCol)
                    yamlText = yamlText & createValidation(ds, regExCol, i + startHdrFldRow, "regex", formatLvlCol)
                    yamlText = yamlText & createValidation(ds, uniqueCol, i + startHdrFldRow, "unique", uniqueLvlCol)
                    yamlText = yamlText & createValidation(ds, enumValuesCol, i + startHdrFldRow, "allowedvalue", formatLvlCol)
                    
                    'determine datatype for validation range rules
                    If (IsEmpty(ds.Cells(i + startHdrFldRow, lowRangeCol)) = False Or _
                        IsEmpty(ds.Cells(i + startHdrFldRow, highRangeCol)) = False) And _
                        (ds.Cells(i + startHdrFldRow, dataTypeCol) = "INT" Or _
                        ds.Cells(i + startHdrFldRow, dataTypeCol) = "DOUBLE") Then
                            yamlText = yamlText & createValidation(ds, fieldNameCol, i + startHdrFldRow, "numberrange", rangeLvlCol, lowRangeCol, highRangeCol)
                    ElseIf (IsEmpty(ds.Cells(i + startHdrFldRow, lowRangeCol)) = False Or _
                        IsEmpty(ds.Cells(i + startHdrFldRow, highRangeCol)) = False) Then
                            yamlText = yamlText & createValidation(ds, fieldNameCol, i + startHdrFldRow, "daterange", rangeLvlCol, lowRangeCol, highRangeCol)
                    End If
                    
                    'remove trailing comma and carriage return
                    yamlText = left(yamlText, (Len(yamlText) - 3))
                    'syntax to close validation rules
                    yamlText = yamlText & vbCrLf & "          ]" & vbCrLf
                End If
            End With
        End If
    End If
    i = i + 1
  Wend

  Exit Sub
  
errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)

End Sub

Function createFieldType(ds As Worksheet, _
                        startHdrFldRow As Integer, _
                        fieldNameCol As Integer, _
                        Optional filterFieldCol As Integer, _
                        Optional filterCol As Integer, _
                        Optional filterValueCol As Integer, _
                        Optional concatField As String, _
                        Optional concatOutputDelimiterCol As Integer)
  
On Error GoTo errHandler
  
  'determine field type and create syntax
  Select Case ds.Name
    Case "Base Fields"
        createFieldType = "     " & ds.Cells(startHdrFldRow, fieldNameCol) & ":" & vbCrLf & _
                          "       type: SIMPLE" & vbCrLf
    Case "Filtered Fields"
        createFieldType = "     " & ds.Cells(startHdrFldRow, fieldNameCol) & ":" & vbCrLf & _
                          "       type: FILTER" & vbCrLf & _
                          "       filter: {field: " & Chr(39) & ds.Cells(startHdrFldRow, filterFieldCol) & Chr(39) & ", condition: " & Chr(39) & ds.Cells(startHdrFldRow, filterCol) & Chr(39) & ", value: " & Chr(39) & ds.Cells(startHdrFldRow, filterValueCol) & Chr(39) & "}" & vbCrLf
    Case "Concat Fields"
        createFieldType = "     " & Replace(concatField, "|", "_") & ":" & vbCrLf & _
                          "       type: CONCAT" & vbCrLf & _
                          "       concat: {fields: " & Chr(39) & concatField & Chr(39) & ", delimiter: " & Chr(39) & ds.Cells(startHdrFldRow, concatOutputDelimiterCol) & Chr(39) & "}" & vbCrLf
    Case "Coded Fields"
  End Select

  Exit Function
  
errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)

End Function

Function createFileMeta(version As String, _
                        fileType As String, _
                        vendor As String, _
                        headerRec As String, _
                        Optional rowLength As String)

On Error GoTo errHandler
  
  Dim fileMeta As String
  
  fileMeta = "FileMeta.Config:" & vbCrLf & _
                   "    generatorversion: " & Chr(39) & version & Chr(39) & vbCrLf & _
                   "    created: " & Chr(39) & Format(Now(), "yyyy-MM-dd hh:mm:ss") & Chr(39) & vbCrLf
                   
  If rowLength = "" Then
    fileMeta = fileMeta & "    filetype: " & Chr(39) & fileType & Chr(39) & vbCrLf
  Else
    fileMeta = fileMeta & "    filetype: 'FIXEDWIDTH'" & vbCrLf & _
                                "    rowlength: " & rowLength & vbCrLf
  End If
  
  createFileMeta = fileMeta & "    vendor: " & Chr(39) & vendor & Chr(39) & vbCrLf & _
                        "    header: " & Chr(39) & headerRec & Chr(39) & vbCrLf
  
  Exit Function
  
errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)
  
End Function

Function createValidation(dataSheet As Worksheet, _
                        columnLocation As Integer, _
                        rowLocation As Integer, _
                        validationType As String, _
                        severityColumn As Integer, _
                        Optional lowRangeColumn As Integer, _
                        Optional highRangeColumn As Integer)
  
On Error GoTo errHandler
  
  'determine validation type and creates syntax
  If dataSheet.Cells(rowLocation, columnLocation) <> "" Then
    Select Case validationType
      Case "dateformat"
        If dataSheet.Cells(rowLocation, columnLocation) = "ISO" Then
            createValidation = "          {type: " & Chr(39) & "dateformat" & Chr(39) & ", severity: " & Chr(39) & setSeverity(dataSheet.Cells(rowLocation, severityColumn)) & Chr(39) & "}," & vbCrLf
        Else
            createValidation = "          {type: " & Chr(39) & "dateformat" & Chr(39) & ", format: " & Chr(39) & dataSheet.Cells(rowLocation, columnLocation) & Chr(39) & ", severity: " & Chr(39) & setSeverity(dataSheet.Cells(rowLocation, severityColumn)) & Chr(39) & "}," & vbCrLf
        End If
      Case "numbertype"
        createValidation = "          {type: " & Chr(39) & "numbertype" & Chr(39) & ", value: " & Chr(39) & dataSheet.Cells(rowLocation, columnLocation) & Chr(39) & ", severity: " & Chr(39) & setSeverity(dataSheet.Cells(rowLocation, severityColumn)) & Chr(39) & "}," & vbCrLf
      Case "required"
        createValidation = "          {type: " & Chr(39) & "required" & Chr(39) & ", severity: " & Chr(39) & setSeverity(dataSheet.Cells(rowLocation, severityColumn)) & Chr(39) & "}," & vbCrLf
      Case "regex"
        createValidation = "          {type: " & Chr(39) & "regex" & Chr(39) & ", value: " & Chr(39) & dataSheet.Cells(rowLocation, columnLocation) & Chr(39) & ", severity: " & Chr(39) & setSeverity(dataSheet.Cells(rowLocation, severityColumn)) & Chr(39) & "}," & vbCrLf
      Case "unique"
        createValidation = "          {type: " & Chr(39) & "unique" & Chr(39) & ", severity: " & Chr(39) & setSeverity(dataSheet.Cells(rowLocation, severityColumn)) & Chr(39) & "}," & vbCrLf
      Case "allowedvalue"
        createValidation = "          {type: " & Chr(39) & "allowedvalue" & Chr(39) & ", value: " & Chr(39) & dataSheet.Cells(rowLocation, columnLocation) & Chr(39) & ", severity: " & Chr(39) & setSeverity(dataSheet.Cells(rowLocation, severityColumn)) & Chr(39) & "}," & vbCrLf
      Case "daterange"
        createValidation = "          {type: " & Chr(39) & "daterange" & Chr(39) & ", severity: " & Chr(39) & setSeverity(dataSheet.Cells(rowLocation, severityColumn)) & Chr(39) & ", " & setRangeValue(dataSheet.Cells(rowLocation, lowRangeColumn), dataSheet.Cells(rowLocation, highRangeColumn)) & "}," & vbCrLf
      Case "numberrange"
        createValidation = "          {type: " & Chr(39) & "numberrange" & Chr(39) & ", severity: " & Chr(39) & setSeverity(dataSheet.Cells(rowLocation, severityColumn)) & Chr(39) & ", " & setRangeValue(dataSheet.Cells(rowLocation, lowRangeColumn), dataSheet.Cells(rowLocation, highRangeColumn)) & "}," & vbCrLf
    End Select
  End If
  
  Exit Function
  
errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)
  
End Function

Function setSeverity(severity As String)
  If severity <> "" Then
    setSeverity = severity
  Else
    setSeverity = "INFO"
  End If
End Function

Function setRangeValue(Optional lowRangeValue As String, _
                      Optional highRangeValue As String)

On Error GoTo errHandler
  
  If lowRangeValue <> "" And highRangeValue <> "" Then
    setRangeValue = "low: " & Chr(39) & lowRangeValue & Chr(39) & ", high: " & Chr(39) & highRangeValue & Chr(39)
  ElseIf lowRangeValue <> "" Then
    setRangeValue = "low: " & Chr(39) & lowRangeValue & Chr(39)
  ElseIf highRangeValue <> "" Then
    setRangeValue = "high: " & Chr(39) & highRangeValue & Chr(39)
  End If
  
  Exit Function
  
errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)
  
End Function
