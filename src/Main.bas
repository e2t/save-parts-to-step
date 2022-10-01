Attribute VB_Name = "Main"
Option Explicit

Enum TSaveAction
  SavePipeToSTEP
  SaveSheetToDWG
End Enum

Public Const KeyPrpNameRU = "Наименование"
Public Const KeyPrpNameEN = "Наименование EN"
Public Const KeyPrpDesignation = "Обозначение"
Public Const KeyPrpBlank = "Заготовка"

Public gFSO As FileSystemObject
Public swApp As Object
Dim CurrentDoc As ModelDoc2

Sub Main()

  Set swApp = Application.SldWorks
  Set CurrentDoc = swApp.ActiveDoc
  If CurrentDoc Is Nothing Then Exit Sub
  If CurrentDoc.GetType <> swDocASSEMBLY Then
    MsgBox "For assemblies only.", vbCritical
    Exit Sub
  End If
  
  Set gFSO = New FileSystemObject
  
  MainForm.Show
  
End Sub

Sub Run(SaveAction As TSaveAction, NeedTranslit As Boolean, IsNameEn As Boolean)
    
  Dim Asm As AssemblyDoc
  Dim Comp_ As Variant
  Dim Comp As Component2
  Dim KeyComp As String
  Dim Doc As ModelDoc2
  Dim Conf As String
  Dim UniqueComp As Dictionary
  Dim NewName As String
  Dim Saver As TSaver
  
  Set Saver = New TSaver
  Saver.Init gFSO.GetParentFolderName(CurrentDoc.GetPathName), NeedTranslit, IsNameEn
  
  Set UniqueComp = New Dictionary
  Set Asm = CurrentDoc
  
  For Each Comp_ In Asm.GetComponents(False)
    Set Comp = Comp_
    Set Doc = Comp.GetModelDoc2
    If Doc Is Nothing Then GoTo NextComp
    If Doc.GetType <> swDocPART Then GoTo NextComp
    Conf = Comp.ReferencedConfiguration
    KeyComp = GetKeyComponent(Comp.GetPathName, Conf)
    If UniqueComp.Exists(KeyComp) Then GoTo NextComp
    
    Select Case SaveAction
      Case SavePipeToSTEP
        NewName = Saver.IfPipeSaveToSTEP(Doc, Conf, True)
      Case SaveSheetToDWG
        NewName = Saver.IfSheetSaveToDWG(Doc, Conf)
    End Select
    If NewName <> "" Then
      UniqueComp.Add KeyComp, NewName
    End If
    
NextComp:
  Next
      
  ShowResults UniqueComp

End Sub

Sub ShowResults(UniqueComp As Dictionary)

  Dim I As Integer
  Dim Msg As String
  Dim LowerFileName As String

  Msg = "Saved:" + Str(UniqueComp.Count) + "."
  If UniqueComp.Count > 0 Then
    If MsgBox(Msg + vbNewLine + "Show?", vbYesNo) = vbYes Then
      
      LowerFileName = UniqueComp.Items(0)
      For I = 1 To UniqueComp.Count - 1
        If StrComp(LowerFileName, UniqueComp.Items(I), vbTextCompare) = 1 Then
          LowerFileName = UniqueComp.Items(I)
        End If
      Next
      
      Shell "explorer /select,""" & LowerFileName & """", vbNormalFocus
      End If
  Else
    MsgBox Msg
  End If

End Sub

Function ExitApp() 'hide

  Unload MainForm
  End

End Function
