VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form 
   Caption         =   "CodeExport"
   ClientHeight    =   5520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8190
   OleObjectBlob   =   "form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public repo As cls_Repo
Public repo_col As Collection

'Private Const ColorBlue& = &HC00000
'Private Const ColorNormal& = &H80000012

Private LastItem$


Public Function ExportQueue()
  For Each repo In repo_col
    If repo.InQueue Then repo.Export
  Next
End Function

Private Sub UserForm_Initialize()

  Dim vbp As Object 'VBProject
  
  Set repo_col = New Collection
  
  For Each vbp In Application.VBE.VBProjects
    Set repo = New cls_Repo
    Set repo.VBProject = vbp
    list.AddItem repo.Name
    repo_col.Add Item:=repo, Key:=repo.Name
  Next
  
  cbQueue.Value = False
  labLocked.Visible = False
  AllVisible False
  
  LastItem = ""
  
End Sub

Private Sub UserForm_Terminate()
  SaveLastItem
  'Set vb = Nothing
  Set repo = Nothing
  Set repo_col = Nothing
End Sub

Private Sub btnBrowse_Click()
  Dim f$
  f = BrowseForFolderDlg(tbPath.Value, "Select Source Folder", 0)
  If f <> "" Then tbPath.Value = f
End Sub

Private Sub btnExp_Click()
  SaveLastItem
  repo_col(LastItem).Export
End Sub

Private Sub btnOK_Click()
  SaveLastItem
  Unload Me
End Sub

Private Sub list_Change()
  
  SaveLastItem
  
  LastItem = list.list(list.ListIndex)
  
  With repo_col(LastItem)
    labLocked.Visible = .Locked
    If Not .Locked Then
      AllVisible True
      cbQueue.Value = .InQueue
      tbPath.Value = .ExportPath
      cbSub.Value = .InSubfolder
      tbSub.Value = .Subfolder
      cbGMS.Value = .IsCopy
    Else
      AllVisible False
    End If
  End With
  
End Sub

Private Function SaveLastItem()
  If LastItem <> "" Then
    'Debug.Print LastItem
    With repo_col(LastItem)
      If Not .Locked Then
        .InQueue = cbQueue.Value
        .ExportPath = tbPath.Value
        .InSubfolder = cbSub.Value
        .Subfolder = tbSub.Value
        .IsCopy = cbGMS.Value
        .Save
      End If
    End With
  End If
End Function

Private Function AllVisible(v As Boolean)
  cbQueue.Visible = v
  labFolder.Visible = v
  tbPath.Visible = v
  btnBrowse.Visible = v
  cbSub.Visible = v
  tbSub.Visible = v
  cbGMS.Visible = v
End Function
