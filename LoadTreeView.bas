Attribute VB_Name = "LoadTreeView"
Public Sub CreateFolderTreeView(TVX As MSComctlLib.TreeView, _
                            TopFolderName As String, _
                            ShowFullPaths As Boolean, _
                            ShowFiles As Boolean, _
                            ItemDescriptionToTag As Boolean)

' This procedure loads the TreeView with file and folder elements.
' Parameters:
' ------------
' TVX: The TreeView control that will be loaded.
'
' TopFolderName: The fully-qualified folder name whose contents are to be listed.
'
' ShowFullPaths: If True, the items in the tree will display the fully qualified folder or file name. If False, only the name, with not path information, will be displayed.
'
' ShowFiles: If True, files within a folder will be listed. If False, only folders, not files, will appear in the Tree listing.
'
' ItemDescriptionToTag: If True, information about the file or folder is stored in the Tag property of the Node.
' This information is either the word "FOLDER" or "FILE" followed by a vbCrLf followed by the fully qualified name of the folder or file.
'
' This code can reside in a standard code module - it need not be in the UserForm's code module.
'

On Error GoTo errHandler

  Dim TopFolder As Scripting.Folder
  Dim OneFile As Scripting.File
  Dim fso As Scripting.FileSystemObject
  Dim TopNode As MSComctlLib.node
  Dim S As String
  Dim FileNode As MSComctlLib.node


  Set fso = New Scripting.FileSystemObject
  ' Clear the tree

  TVX.nodes.clear

  Set TopFolder = fso.GetFolder(folderpath:=TopFolderName)

  ' Create the top node of the TreeView.
  If ShowFullPaths = True Then
    S = TopFolder.Path
  Else
    S = TopFolder.Name
  End If
    
  Set TopNode = TVX.nodes.Add(text:=S)

  If ItemDescriptionToTag = True Then
    TopNode.Tag = "FOLDER" & vbCrLf & TopFolder.Path
  End If


  ' Call CreateNodes. This procedure creates all the nodes of the tree using a recursive technique -- that is, the code
  ' calls itself to create child nodes, child of child nodes, and so on.
  CreateNodes TVX:=TVX, _
            fso:=fso, _
            ParentNode:=TopNode, _
            FolderObject:=TopFolder, _
            ShowFullPaths:=ShowFullPaths, _
            ShowFiles:=ShowFiles, _
            ItemDescriptionToTag:=ItemDescriptionToTag

  ' After all the subfolders are built, we need to add the folders that are directly below the TopFolder.
  If ShowFiles = True Then
    For Each OneFile In TopFolder.Files
        If ShowFullPaths = True Then
            S = OneFile.Path
        Else
            S = OneFile.Name
        End If
        Set FileNode = TVX.nodes.Add(relative:=TopNode, relationship:=tvwChild, text:=S)
        If ItemDescriptionToTag = True Then
            'FileNode.Tag = "FILE" & vbCrLf & OneFile.Path
            FileNode.Tag = OneFile.Path
        End If
    Next OneFile
  End If

  ' Finally, now that everything has been added to the TreeView, expand the top node.
  With TVX.nodes
    If .Count > 0 Then
        .Item(1).Expanded = True
    End If
  End With
    
  Exit Sub
  
errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)
  
End Sub


Private Sub CreateNodes(TVX As MSComctlLib.TreeView, _
    fso As Scripting.FileSystemObject, _
    ParentNode As MSComctlLib.node, _
    FolderObject As Scripting.Folder, _
    ShowFullPaths As Boolean, _
    ShowFiles As Boolean, _
    ItemDescriptionToTag As Boolean)

' This procedure creates the nodes on the Tree. This procedure can reside in a standard code module --
' it need not reside in the UserForm's code module. This procedure uses a recursive technique to build
' child nodes. That is, it calls itself to build the children, children of children, and so on.
'
' Parameters:
' -----------
' TVX: The TreeView controls to which the nodes will be added.
'
' FSO: A FileSystemObject used to enumerate subfolders and files of FolderObject.
'
' ParentNode: The Node that will be the parent of all nodes created by the current iteration of this procedure.
'
' FolderObject: The Folder object whose contents we are going to list. The TreeView element for FolderObject has already been added to the tree.
'
' ShowFullPaths: If True, elements in the tree will display the fully-qualified folder or file name. If
' False, only the name of the folder or file will appear in the tree. No path information will be displayed.
'
' ShowFiles: If True, files within a folder are listed in the tree. If False, no files are listed and only Folders will appear in the tree.
'
' ItemDescriptionToTag: If True, information about the folder or file is stored in the Tag property of the node. The infomation is the word "FOLDER" or "FILE"
' followed by a vbCrLF followed by the fully qualified name of the folder or file. If False, no information is stored in the Tag property.

On Error GoTo errHandler

  Dim SubNode As MSComctlLib.node
  Dim FileNode As MSComctlLib.node
  Dim OneSubFolder As Scripting.Folder
  Dim OneFile As Scripting.File
  Dim S As String

  ' Process each folder directly under FolderObject.
  For Each OneSubFolder In FolderObject.SubFolders
    If ShowFullPaths = True Then
        S = OneSubFolder.Path
    Else
        S = OneSubFolder.Name
    End If
    
    ' Add the node to the tree.
    Set SubNode = TVX.nodes.Add(relative:=ParentNode, relationship:=tvwChild, text:=S)
    If ItemDescriptionToTag = True Then
        'SubNode.Tag = "FOLDER" & vbCrLf & OneSubFolder.Path
        SubNode.Tag = OneSubFolder.Path
    End If
    
    ' Recursive code. CreateNodes calls itself to add sub nodes to the tree. This recursion creates the children, children of children, and so on.
    CreateNodes TVX:=TVX, _
                fso:=fso, _
                ParentNode:=SubNode, _
                FolderObject:=OneSubFolder, _
                ShowFullPaths:=ShowFullPaths, _
                ShowFiles:=ShowFiles, _
                ItemDescriptionToTag:=ItemDescriptionToTag

    ' If ShowFiles is True, add nodes for all the files.
    If ShowFiles = True Then
        For Each OneFile In OneSubFolder.Files
            If ShowFullPaths = True Then
                S = OneFile.Path
            Else
                S = OneFile.Name
            End If
            Set FileNode = TVX.nodes.Add(relative:=SubNode, relationship:=tvwChild, text:=S)
            If ItemDescriptionToTag = True Then
                'FileNode.Tag = "FILE" & vbCrLf & OneFile.Path
                FileNode.Tag = OneFile.Path
            End If
        Next OneFile
    
    End If
    
  Next OneSubFolder

  Exit Sub
  
errHandler:
  Call errMsg(Err.Number, Err.Description, Application.VBE.ActiveCodePane.codeModule.Name)
  
End Sub
