VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmControlPanel 
   Caption         =   "Control Panel"
   ClientHeight    =   6630
   ClientLeft      =   50
   ClientTop       =   370
   ClientWidth     =   5410
   OleObjectBlob   =   "frmControlPanel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmControlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCancel_Click()
  Call saveControlPanelValues(frmControlPanel.cbxDelimiter, frmControlPanel.cbxVendor, frmControlPanel.txtHdrRec, frmControlPanel.txtRowLength)
  frmControlPanel.Hide
End Sub

Private Sub btnClearSaved_Click()
  Call resetControlPanelValues
End Sub

Private Sub btnConfigLocation_Click()
  On Error GoTo configlocationerr
  Dim intResult As Integer
        
  'the dialog is displayed to the user
  intResult = Application.FileDialog(msoFileDialogFolderPicker).Show
  'checks if user has cancled the dialog
  If intResult <> 0 Then
    frmControlPanel.lblConfigLocation.Caption = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
    Call CreateFolderTreeView(trFileDir, Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1), False, True, True)
    frmControlPanel.lblClickImport.Visible = True
    frmControlPanel.trFileDir.Enabled = True
  End If
  Exit Sub

configlocationerr:
  MsgBox Err.Description, vbCritical
End Sub

Private Sub btnConvertHdrRow_Click()
  'turn screen updating on
  Application.ScreenUpdating = False
  
  Call checkControlPanelInput(frmControlPanel.optDelimited, frmControlPanel.cbxDelimiter, frmControlPanel.optFixedWidth.Value, frmControlPanel.txtRowLength, frmControlPanel.cbxVendor, frmControlPanel.txtHdrRec)
  
  If userInputErrChk = False Then
    Call processHdrRec(Sheets("Base Fields"), "B", "Q", bFieldNameCol, frmControlPanel.optFixedWidth.Value, bDataTypeCol)
    
    MsgBox "Successfully Parsed Header Record", vbInformation
    frmControlPanel.Hide
  End If

  Sheets("Home").Select
  'turn screen updating on
  Application.ScreenUpdating = True
End Sub

Private Sub btnLoadConfigLocation_Click()
  If frmControlPanel.txtConfigLocation.text <> "" Then
    Call CreateFolderTreeView(trFileDir, Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1), False, True, True)
  Else
    MsgBox "Please select config file location", vbCritical
  End If
End Sub

Private Sub btnExportCode_Click()
  Call SaveCodeModules
End Sub

Private Sub btnExportConfig_Click()
  'turn screen updating off
  Application.ScreenUpdating = False
  
  Call exportPreview(False)
  
  'turn screen updating on
  Application.ScreenUpdating = True
End Sub

Private Sub btnPreview_Click()
  'turn screen updating off
  Application.ScreenUpdating = False
  
  Call exportPreview(True)
  
  'turn screen updating on
  Application.ScreenUpdating = True
End Sub

Sub exportPreview(preview As Boolean)
  userInputErrChk = False
  
  Call saveControlPanelValues(frmControlPanel.cbxDelimiter, frmControlPanel.cbxVendor, frmControlPanel.txtHdrRec, frmControlPanel.txtRowLength)
  Call checkControlPanelInput(frmControlPanel.optDelimited, frmControlPanel.cbxDelimiter, frmControlPanel.optFixedWidth.Value, frmControlPanel.txtRowLength, frmControlPanel.cbxVendor, frmControlPanel.txtHdrRec)
  
  If userInputErrChk = False Then
    Call checkUserInput
  End If
  
  If userInputErrChk = False Then
    Call ExportYamlFile(preview)
    frmControlPanel.Hide
  End If
End Sub

Private Sub btnResetHdrRow_Click()
  'turn screen updating off
  Application.ScreenUpdating = False
  
  Call clearSheets(Sheets("Base Fields"), bFieldNameCol, bEndIndexCol, bHistoCol)
  Call clearSheets(Sheets("Filtered Fields"), fFilterFieldCol, fFilterFieldCol, fHistoCol)
  Call clearSheets(Sheets("Concat Fields"), cFieldName1Col, cFieldName1Col, cHistoCol)
  Call clearSheets(Sheets("Coded Fields"), cdCodeIDCol, cdConceptCol, cdDefaultCodeSysCol)
  
  frmControlPanel.cbxDelimiter = ""
  frmControlPanel.cbxVendor = ""
  frmControlPanel.txtHdrRec = ""
  frmControlPanel.txtRowLength = ""
  frmControlPanel.optDelimited.Value = False
  frmControlPanel.optFixedWidth.Value = False
  
  Sheets("saved").Cells(1, 3) = ""
  Sheets("saved").Cells(2, 3) = ""
  Sheets("saved").Cells(3, 3) = ""
  Sheets("saved").Cells(4, 3) = ""
  
  Sheets("Home").Cells(4, 4) = ""
  
  Sheets("Home").Select
  
  'turn screen updating on
  Application.ScreenUpdating = True
  
  MsgBox "Successfully Reset Header Record from Rules Sheets", vbInformation
  frmControlPanel.Hide
End Sub

Private Sub optDelimited_Click()
  frmControlPanel.cbxDelimiter.Enabled = True
  frmControlPanel.txtRowLength.Enabled = False
End Sub

Private Sub optFixedWidth_Click()
  frmControlPanel.cbxDelimiter.Enabled = False
  frmControlPanel.txtRowLength.Enabled = True
End Sub

Private Sub trFileDir_DblClick()
  'turn screen updating off
  Application.ScreenUpdating = False
    
  'clear and existing values in sheets
  Call clearSheets(Sheets("Base Fields"), bFieldNameCol, bEndIndexCol, bHistoCol)
  Call clearSheets(Sheets("Filtered Fields"), fFilterFieldCol, fFilterFieldCol, fHistoCol)
  Call clearSheets(Sheets("Concat Fields"), cFieldName1Col, cFieldName1Col, cHistoCol)
  Call clearSheets(Sheets("Coded Fields"), cdCodeIDCol, cdConceptCol, cdDefaultCodeSysCol)
  
  'import yaml file
  Call ImportYamlFile(frmControlPanel.trFileDir.SelectedItem.Tag)
  
  'turn screen updating on
  Application.ScreenUpdating = True
End Sub

Private Sub UserForm_Initialize()
  Dim delimiterArray() As String
  Dim vendorArray() As String
  Dim i As Integer
  
  'load values for delimiter combo box
  delimiterArray() = Split(delimiterValue, ";")
  For i = LBound(delimiterArray()) To UBound(delimiterArray())
    With Me.cbxDelimiter
      .AddItem delimiterArray(i)
    End With
  Next i
  
  'load values for vendor combo box
  vendorArray() = Split(vendorValue, ";")
  For i = LBound(vendorArray()) To UBound(vendorArray())
    With Me.cbxVendor
      .AddItem vendorArray(i)
    End With
  Next i
  
  frmControlPanel.lblGeneratorVersion = "Generator Version: " & gblGeneratorVersion
    
  Me.cbxDelimiter.text = Sheets("saved").Cells(1, 3)
  Me.cbxVendor.text = Sheets("saved").Cells(2, 3)
  Me.txtHdrRec.text = Sheets("saved").Cells(3, 3)
  Me.txtRowLength.text = Sheets("saved").Cells(4, 3)
End Sub
