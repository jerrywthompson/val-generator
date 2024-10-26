Attribute VB_Name = "VersionControl"
Sub ControlPanelLoad()
 Dim html As String
 
On Error GoTo errHandler
  
  If updateCode = True Then
    updateCode = False
    html = GetHTTPResult("http://github.cerner.com/dataworksdev/profiling-config-generator/raw/master/Code/Modules/GlobalVars.bas")
    If InStr(html, "gblGeneratorVersion = """ & gblGeneratorVersion & """") = 0 Then
        MsgBox "A newer Generator Version is available: " & Replace(Mid(html, InStr(html, "gblGeneratorVersion = ") + 23, (InStr(html, "'") - (InStr(html, "gblGeneratorVersion = ") + 23))), """", "") & vbCrLf & vbCrLf & _
                "Please goto: http://github.cerner.com/dataworksdev/profiling-config-generator/tree/master/generator" & vbCrLf & _
                "To download this version", vbInformation, "Update"
    End If
  End If
  
  frmControlPanel.Show
  
  Exit Sub

errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)

End Sub
Sub SaveCodeModules()
    Dim intResult As Integer
    Dim strPath As String
    Dim i As Integer
    Dim sName As String
    Dim bExport As Boolean

On Error GoTo errHandler
  
  'the dialog is displayed to the user
  intResult = Application.FileDialog(msoFileDialogFolderPicker).Show
  'checks if user has cancled the dialog
  If intResult <> 0 Then
    With ThisWorkbook.VBProject
        For i = 1 To .VBComponents.Count
            sName = .VBComponents(i).codeModule.Name
            If .VBComponents(i).codeModule.CountOfLines > 0 Then
                bExport = True
                sName = .VBComponents(i).codeModule.Name
                Select Case .VBComponents(i).Type
                    Case vbext_ct_ClassModule
                        sName = sName & ".cls"
                    Case vbext_ct_MSForm
                        sName = sName & ".frm"
                    Case vbext_ct_StdModule
                        sName = sName & ".bas"
                    Case vbext_ct_Document
                        sName = sName & ".cls"
                        'bExport = False
                End Select
                    
                If bExport Then
                    .VBComponents(i).Export Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1) & "\" & sName
                End If
            End If
        Next i
    End With
  End If
  
  Exit Sub

errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)
  
End Sub
Sub ImportCodeModules()

On Error GoTo errHandler
  
  Call replaceModules("http://github.cerner.com/dataworksdev/profiling-config-generator/raw/master/Code/Modules/GlobalVars.bas", "GlobalVars")
                       'http://github.cerner.com/dataworksdev/profiling-config-generator/raw/master/Code/Modules/GlobalVars.bas

  updateCode = False

  Exit Sub

errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)

End Sub
Function replaceModules(url As String, _
                    moduleName As String)

On Error GoTo errHandler

  Dim newmod As Object
  Dim html As String

  html = GetHTTPResult(url)
  With ThisWorkbook.VBProject
    .VBComponents.Remove .VBComponents(moduleName)
    Set newmod = .VBComponents.Add(vbext_ct_StdModule)
    newmod.Name = moduleName
    newmod.codeModule.AddFromString (html)
  End With
    
  Exit Function
  
errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)
    
End Function
Function GetHTTPResult(sURL As String) As String
  
On Error GoTo errHandler
  
  Dim XMLHTTP As Variant, sResult As String

  Set XMLHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
  XMLHTTP.Open "GET", sURL, False
  XMLHTTP.send
  sResult = XMLHTTP.responseText
  Set XMLHTTP = Nothing
  GetHTTPResult = sResult
  
  Exit Function
  
errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)
  
End Function
