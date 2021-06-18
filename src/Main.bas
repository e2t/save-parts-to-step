Attribute VB_Name = "Main"
Option Explicit

Enum TSaveAction
  
  SavePipeToSTEP
  SaveSheetToDWG
  
End Enum

Public Const KeyPrpName As String = "Наименование"
Public Const KeyPrpDesignation As String = "Обозначение"
Const KeyPrpBlank As String = "Заготовка"

Public gFSO As FileSystemObject
Public swApp As Object
Dim IsPipeRegex As RegExp
Dim CurrentDoc As ModelDoc2

Sub Main()
  
  Set swApp = Application.SldWorks
  Set CurrentDoc = swApp.ActiveDoc
  If CurrentDoc Is Nothing Then Exit Sub
  If CurrentDoc.GetType <> swDocASSEMBLY Then
    MsgBox "Только для сборок.", vbCritical
    Exit Sub
  End If
  
  Set gFSO = New FileSystemObject
  Set IsPipeRegex = New RegExp
  IsPipeRegex.Pattern = ".*труба.*"
  IsPipeRegex.IgnoreCase = True
  
  MainForm.Show
  
End Sub

Sub Run(SaveAction As TSaveAction)

  Dim CurrentFolder As String
  Dim Asm As AssemblyDoc
  Dim CompArray As Variant
  Dim Comp_ As Variant
  Dim Comp As Component2
  Dim KeyComp As String
  Dim Doc As ModelDoc2
  Dim DocConf As String
  Dim Key_ As Variant
  Dim SavedPartCount As Long
  Dim HighIndex As Long
  Dim Msg As String
  Dim UniqueComp As Dictionary
  Dim SavedFiles() As String
  Dim NewName As String
  Dim WasSaved As Boolean
  
  Set UniqueComp = New Dictionary
  CurrentFolder = gFSO.GetParentFolderName(CurrentDoc.GetPathName)
  Set Asm = CurrentDoc
  CompArray = Asm.GetComponents(False)
  For Each Comp_ In CompArray
    Set Comp = Comp_
    KeyComp = GetKeyComponent(Comp)
    If Not UniqueComp.Exists(KeyComp) Then
      Set Doc = Comp.GetModelDoc2
      If Not Doc Is Nothing Then
        DocConf = Comp.ReferencedConfiguration
        
        Select Case SaveAction
          Case SavePipeToSTEP
            WasSaved = TrySaveComponentToSTEP(Doc, DocConf, NewName, CurrentFolder)
          Case SaveSheetToDWG
            WasSaved = TrySaveComponentToDWG(Doc, DocConf, NewName, CurrentFolder)
        End Select

        If WasSaved Then
          UniqueComp.Add KeyComp, NewName
        End If
      End If
    End If
  Next
  
  ReDim SavedFiles(UniqueComp.Count)
  HighIndex = -1
  For Each Key_ In UniqueComp.Items
    If Key_ <> "" Then
      HighIndex = HighIndex + 1
      SavedFiles(HighIndex) = Key_
    End If
  Next
  SavedPartCount = HighIndex + 1
      
  Msg = "Сохранено компонентов:" + Str(SavedPartCount) + "."
  If SavedPartCount > 0 Then
    If MsgBox(Msg + vbNewLine + "Показать?", vbYesNo) = vbYes Then
      ReDim Preserve SavedFiles(HighIndex)
      QuickSort SavedFiles, 0, HighIndex
      Shell "explorer /select,""" & SavedFiles(0) & """", vbNormalFocus
      End If
  Else
    MsgBox Msg
  End If

End Sub

Function TrySaveComponentToSTEP( _
  Doc As ModelDoc2, DocConf As String, ByRef NewName As String, CurrentFolder As String) As Boolean

  Dim Property As String

  Property = GetProperty(KeyPrpBlank, DocConf, Doc.Extension)
  TrySaveComponentToSTEP = IsPipeRegex.Test(Property)
  If TrySaveComponentToSTEP Then
    NewName = GetNewName(Doc, DocConf, CurrentFolder, "STEP", True)
    SaveToSTEP Doc, DocConf, NewName
  End If

End Function

Function TrySaveComponentToDWG( _
  Doc As ModelDoc2, DocConf As String, ByRef NewName As String, CurrentFolder As String) As Boolean

  TrySaveComponentToDWG = IsSheetMetal(Doc)
  If TrySaveComponentToDWG Then
    NewName = GetNewName(Doc, DocConf, CurrentFolder, "DWG", False)
    SaveToDWG Doc, DocConf, NewName
  End If

End Function

Function ExitApp() 'hide

  Unload MainForm
  End

End Function
