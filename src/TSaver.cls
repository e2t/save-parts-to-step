VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TSaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim CurrentFolder As String
Dim FromCurrentFolderMask As String
Dim IsPipeRegex As RegExp

Private Sub Class_Initialize()

  Set IsPipeRegex = New RegExp
  IsPipeRegex.Pattern = ".*труба.*"
  IsPipeRegex.IgnoreCase = True
  
End Sub

Public Sub Init(ACurrentFolder As String)

  CurrentFolder = ACurrentFolder
  FromCurrentFolderMask = LCase(gFSO.BuildPath(ACurrentFolder, "*"))

End Sub

Public Function TrySaveToSTEP( _
  Doc As ModelDoc2, Conf As String, NeedTranslit As Boolean) As String

  Dim Property As String
  Dim NewName As String

  TrySaveToSTEP = ""
  Property = GetProperty(KeyPrpBlank, Conf, Doc.Extension)
  If IsPipeRegex.Test(Property) Then
    NewName = GetNewName(Doc, Conf, CurrentFolder, "STEP", NeedTranslit)
    SaveToSTEP Doc, Conf, NewName
    TrySaveToSTEP = NewName
  End If

End Function

Public Function TrySaveToDWG( _
  Doc As ModelDoc2, Conf As String, NeedTranslit As Boolean) As String
  
  Dim NewName As String

  TrySaveToDWG = ""
  If LCase(Doc.GetPathName) Like FromCurrentFolderMask Then
    If IsSheetMetal(Doc) Then
      NewName = GetNewName(Doc, Conf, CurrentFolder, "DWG", NeedTranslit)
      SaveToDWG Doc, Conf, NewName
      TrySaveToDWG = NewName
    End If
  End If

End Function
