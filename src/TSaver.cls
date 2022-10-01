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
Dim KeyPrpName As String
Dim NeedTranslit As Boolean

Private Sub Class_Initialize()

  Set IsPipeRegex = New RegExp
  IsPipeRegex.Pattern = ".*(труба|pipe|tube).*"
  IsPipeRegex.IgnoreCase = True
  
End Sub

Public Sub Init(ACurrentFolder As String, aNeedTranslit As Boolean, IsNameEn As Boolean)

  CurrentFolder = ACurrentFolder
  FromCurrentFolderMask = LCase(gFSO.BuildPath(ACurrentFolder, "*"))
  PipeFolder = gFSO.BuildPath(CurrentFolder, "STEP for pipes")
  NeedTranslit = aNeedTranslit
  KeyPrpName = IIf(IsNameEn, KeyPrpNameEN, KeyPrpNameRU)

End Sub

Public Function IfPipeSaveToSTEP( _
  Doc As ModelDoc2, Conf As String, NeedCreateEmptyPDF As Boolean) As String

  Const Suffix = "EMPTY"
  
  Dim SpecName As String
  Dim Blank As String
  Dim NewName As String
  Dim BaseName As String
  
  IfPipeSaveToSTEP = ""
  SpecName = GetProperty(KeyPrpName, Conf, Doc.Extension)
  Blank = GetProperty(KeyPrpBlank, Conf, Doc.Extension)
  If IsPipeRegex.Test(SpecName) Or IsPipeRegex.Test(Blank) Then
    BaseName = GetNewBaseName(Doc, Conf, PipeFolder)
    NewName = PrepareNewName(BaseName, PipeFolder, "STEP")
    SaveToSTEP Doc, Conf, NewName
    IfPipeSaveToSTEP = NewName
    
    If NeedCreateEmptyPDF Then
      CreateEmptyPDF PrepareNewName(BaseName + " " + Suffix, PipeFolder, "PDF")
    End If
  End If

End Function

Public Function IfSheetSaveToDWG( _
  Doc As ModelDoc2, Conf As String) As String
  
  Dim NewName As String
  Dim BaseName As String

  IfSheetSaveToDWG = ""
  If LCase(Doc.GetPathName) Like FromCurrentFolderMask Then
    If IsSheetMetal(Doc) Then
      BaseName = GetNewBaseName(Doc, Conf, CurrentFolder)
      NewName = PrepareNewName(BaseName, CurrentFolder, "DWG")
      SaveToDWG Doc, Conf, NewName
      IfSheetSaveToDWG = NewName
    End If
  End If

End Function

Private Function GetNewBaseName( _
  Doc As ModelDoc2, Conf As String, FolderName As String) As String

  Dim PrpDesignation As String
  Dim PrpName As String
  Dim ChangeNumber As Integer

  PrpDesignation = GetProperty(KeyPrpDesignation, Conf, Doc.Extension)
  
  ChangeNumber = 0
  If Not GetChangeNumber(Doc, PrpDesignation, PrpName, ChangeNumber) Then
    GetChangeNumber Doc, GetBaseDesignation(PrpDesignation), PrpName, ChangeNumber
  End If
  
  If NeedTranslit Then
    GetNewBaseName = Translit(PrpDesignation) + " " + Translit(PrpName)
  Else
    GetNewBaseName = PrpDesignation + " " + PrpName
  End If
  GetNewBaseName = Trim(GetNewBaseName)
  
  If ChangeNumber > 0 Then
    GetNewBaseName = GetNewBaseName + " (rev." + Format(ChangeNumber, "00") + ")"
  End If

End Function

