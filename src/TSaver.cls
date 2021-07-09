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
Dim PipeFolder As String
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
  PipeFolder = gFSO.BuildPath(CurrentFolder, "STEP для труб")

End Sub

Public Function IfPipeSaveToSTEP( _
  Doc As ModelDoc2, Conf As String, NeedTranslit As Boolean, NeedCreateEmptyPDF As Boolean) As String

  Dim Property As String
  Dim NewName As String
  Dim BaseName As String
  Dim Suffix As String
  
  IfPipeSaveToSTEP = ""
  Property = GetProperty(KeyPrpBlank, Conf, Doc.Extension)
  If IsPipeRegex.Test(Property) Then
    BaseName = GetNewBaseName(Doc, Conf, PipeFolder, NeedTranslit)
    NewName = PrepareNewName(BaseName, PipeFolder, "STEP")
    SaveToSTEP Doc, Conf, NewName
    IfPipeSaveToSTEP = NewName
    
    If NeedCreateEmptyPDF Then
      Suffix = IIf(NeedTranslit, "EMPTY", "ПУСТОЙ")
      CreateEmptyPDF PrepareNewName(BaseName + " " + Suffix, PipeFolder, "PDF")
    End If
  End If

End Function

Public Function IfSheetSaveToDWG( _
  Doc As ModelDoc2, Conf As String, NeedTranslit As Boolean) As String
  
  Dim NewName As String
  Dim BaseName As String

  IfSheetSaveToDWG = ""
  If LCase(Doc.GetPathName) Like FromCurrentFolderMask Then
    If IsSheetMetal(Doc) Then
      BaseName = GetNewBaseName(Doc, Conf, CurrentFolder, NeedTranslit)
      NewName = PrepareNewName(BaseName, CurrentFolder, "DWG")
      SaveToDWG Doc, Conf, NewName
      IfSheetSaveToDWG = NewName
    End If
  End If

End Function
