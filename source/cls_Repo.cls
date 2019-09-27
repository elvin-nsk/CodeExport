VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Repo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=======================================================================================
' cls_Repo
' Settings, file and ini functions for project repository
'=======================================================================================

Option Explicit

'=======================================================================================
' public vars
'=======================================================================================

Public ExportPath$
Public Subfolder$


'=======================================================================================
' private vars
'=======================================================================================

Private Enum cfgAction
  cfgRead = 1
  cfgWrite = 2
  cfgReset = 3
End Enum

Private Const INI_FILENAME$ = "CodeExport.ini"

Private vbp As Object
Private ini_File$

Private InQueueStr$
Private InSubfolderStr$
Private IsCopyStr$


'=======================================================================================
' public properties
'=======================================================================================

Property Set VBProject(v As Object)
  Set vbp = v
  init
End Property

Property Get VBProject() As Object
  Set VBProject = vbp
End Property

Property Get Name() As String
  Name = vbp.Name
End Property

Property Get Locked() As Boolean
  Locked = vbp.Protection
End Property

Property Let InQueue(q As Boolean)
  If q Then InQueueStr = "True" Else InQueueStr = "False"
End Property

Property Get InQueue() As Boolean
  If InQueueStr = "True" Then InQueue = True Else InQueue = False
End Property

Property Let InSubfolder(s As Boolean)
  If s Then InSubfolderStr = "True" Else InSubfolderStr = "False"
End Property

Property Get InSubfolder() As Boolean
  If InSubfolderStr = "True" Then InSubfolder = True Else InSubfolder = False
End Property

Property Let IsCopy(c As Boolean)
  If c Then IsCopyStr = "True" Else IsCopyStr = "False"
End Property

Property Get IsCopy() As Boolean
  If IsCopyStr = "True" Then IsCopy = True Else IsCopy = False
End Property


'=======================================================================================
' public functions
'=======================================================================================

Public Function Load()
  cfgAssign cfgRead
End Function

Public Function Save()
  cfgAssign cfgWrite
End Function

Public Function Reset()
  cfgAssign cfgReset
End Function

Public Function Export()

  Dim vc As Object 'VBComponent
  Dim SrcPath$
  Dim FileName$
    
  If Len(ExportPath) = 0 Or Locked Then Exit Function
  
  If InSubfolder Then
    SrcPath = ExportPath & "\" & Subfolder
  Else
    SrcPath = ExportPath
  End If
  CreateDir SrcPath
  
  For Each vc In vbp.VBComponents
    FileName = vc.Name

    Select Case vc.Type
      Case 1 'vbext_ct_StdModule
        FileName = FileName & ".bas"
      Case 2, 100 'vbext_ct_ClassModule, vbext_ct_Document
        FileName = FileName & ".cls"
      Case 3 'vbext_ct_MSForm
        FileName = FileName & ".frm"
    End Select

    vc.Export SrcPath & "\" & FileName
  Next

  If IsCopy Then FileSystem.FileCopy vbp.FileName, ExportPath & "\" & Name & ".gms"
  'MsgBox "Done!"

End Function


'=======================================================================================
' private functions
'=======================================================================================

Private Sub Class_Initialize()
End Sub

Private Function init()
  Select Case True
  Case FileExist(Application.GMSManager.UserGMSPath & INI_FILENAME)
    ini_File = Application.GMSManager.UserGMSPath & INI_FILENAME
    Load
  Case FileExist(Application.GMSManager.GMSPath & INI_FILENAME)
    ini_File = Application.GMSManager.GMSPath & INI_FILENAME
    Load
  Case Else
    ini_File = Application.GMSManager.UserGMSPath & INI_FILENAME
    Reset
  End Select
End Function

Private Function cfgAssign(Action As cfgAction)
  
  iniProcess Action, Name, "InQueue", InQueueStr, "False"
  iniProcess Action, Name, "ExportPath", ExportPath, "c:\"
  iniProcess Action, Name, "InSubfolder", InSubfolderStr, "True"
  iniProcess Action, Name, "Subfolder", Subfolder, "source"
  iniProcess Action, Name, "IsCopy", IsCopyStr, "True"
  
End Function

Private Function iniProcess(Action As cfgAction, ByVal Section$, ByVal Key$, ByRef Value$, Optional ByVal DefaultValue$)
  Select Case Action
  Case cfgWrite
    sManageSectionEntry iniWrite, Section, Key, ini_File, Value
  Case cfgRead
    Value = sManageSectionEntry(iniRead, Section, Key, ini_File)
    If Value = "" Then Value = DefaultValue
  Case cfgReset
    Value = DefaultValue
  End Select
End Function

Private Function CreateDir(path)
  If Trim(Dir(path, vbDirectory)) = "" Then FileSystem.MkDir path
End Function

'---------------------------------------------------------------------------------------
' Procedure : FileExist
' DateTime  : 2007-Mar-06 13:51
' Author    : CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Test for the existance of a file; Returns True/False
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' strFile - name of the file to be tested for including full path
'---------------------------------------------------------------------------------------
Private Function FileExist(strFile As String) As Boolean
    On Error GoTo Err_Handler
 
    FileExist = False
    If Len(Dir(strFile)) > 0 Then
        FileExist = True
    End If
 
Exit_Err_Handler:
    Exit Function
 
Err_Handler:
    MsgBox "The following error has occured." & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: FileExist" & vbCrLf & _
           "Error Description: " & Err.Description, _
           vbCritical, "An Error has Occured!"
    GoTo Exit_Err_Handler
End Function
