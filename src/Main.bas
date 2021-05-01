Attribute VB_Name = "Main"
Option Explicit

Const KeyPrpBlank As String = "Заготовка"
Const KeyPrpName As String = "Наименование"
Const KeyPrpDesignation As String = "Обозначение"

Dim swApp As Object
Dim gFSO As FileSystemObject
Dim CurrentFolder As String

Sub Main()

  Dim CurrentDoc As ModelDoc2
  Dim Asm As AssemblyDoc
  Dim CompArray As Variant
  Dim Comp_ As Variant
  Dim Comp As Component2
  Dim DocConf As String
  Dim Doc As ModelDoc2
  Dim Property As String
  Dim Regex As RegExp
  Dim KeyComp As String
  Dim Key_ As Variant
  Dim SavedPartCount As Long
  Dim HighIndex As Long
  Dim Msg As String
  Dim UniqueComp As Dictionary
  Dim StepFiles() As String
  Dim StepName As String
  
  Set swApp = Application.SldWorks
  Set CurrentDoc = swApp.ActiveDoc
  If CurrentDoc Is Nothing Then Exit Sub
  If CurrentDoc.GetType <> swDocASSEMBLY Then
    MsgBox "Только для сборок.", vbCritical
    Exit Sub
  End If
  
  Set gFSO = New FileSystemObject
  Set UniqueComp = New Dictionary
  
  Set Regex = New RegExp
  Regex.Pattern = ".*труба.*"
  Regex.Global = True
  Regex.IgnoreCase = True
  
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
        Property = GetProperty(KeyPrpBlank, DocConf, Doc.Extension)
        If Regex.Test(Property) Then
          StepName = SaveToSTEP(Doc, DocConf)
        Else
          StepName = ""
        End If
        UniqueComp.Add KeyComp, StepName
      End If
    End If
  Next
  
  ReDim StepFiles(UniqueComp.Count)
  HighIndex = -1
  For Each Key_ In UniqueComp.Items
    If Key_ <> "" Then
      HighIndex = HighIndex + 1
      StepFiles(HighIndex) = Key_
    End If
  Next
  SavedPartCount = HighIndex + 1
      
  Msg = "Сохранено компонентов:" + Str(SavedPartCount) + "."
  If SavedPartCount > 0 Then
    If MsgBox(Msg + vbNewLine + "Показать?", vbYesNo) = vbYes Then
      ReDim Preserve StepFiles(HighIndex)
      QuickSort StepFiles, 0, HighIndex
      Shell "explorer /select,""" & StepFiles(0) & """", vbNormalFocus
    End If
  Else
    MsgBox Msg
  End If
  
End Sub

Function GetKeyComponent(Comp As Component2) As String

  Dim BaseName As String
  
  BaseName = gFSO.GetBaseName(Comp.GetPathName)
  GetKeyComponent = BaseName + "@" + Comp.ReferencedConfiguration
    
End Function

Function GetProperty(Property As String, Conf As String, DocExt As ModelDocExtension) As String

  Dim Value As String
  Dim RawValue As String
  Dim WasResolved As Boolean
  Dim GetPrpResult As swCustomInfoGetResult_e
  
  Value = ""
  GetPrpResult = DocExt.CustomPropertyManager(Conf).Get5(Property, False, RawValue, Value, WasResolved)
  If GetPrpResult = swCustomInfoGetResult_NotPresent Then
    DocExt.CustomPropertyManager("").Get5 Property, False, RawValue, Value, WasResolved
  End If
  GetProperty = Trim(Value)
    
End Function

Sub ActivatePartConfiguration(Doc As ModelDoc2, Conf As String)

  Dim Errors As swActivateDocError_e
  
  swApp.ActivateDoc3 Doc.GetPathName, False, swDontRebuildActiveDoc, Errors
  Doc.ShowConfiguration2 Conf
    
End Sub

Function SaveToSTEP(Doc As ModelDoc2, Conf As String) As String

  Dim DocExt As ModelDocExtension
  Dim NewName As String
  Dim Errors As swFileSaveError_e
  Dim Warnings As swFileSaveWarning_e
  
  Set DocExt = Doc.Extension
  NewName = GetNewName(DocExt, Conf)
  ActivatePartConfiguration Doc, Conf
  DocExt.SaveAs NewName, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, Errors, Warnings
  swApp.QuitDoc Doc.GetPathName
  SaveToSTEP = NewName
    
End Function

Function GetNewName(DocExt As ModelDocExtension, Conf As String) As String

  Dim PrpDesignation As String
  Dim PrpName As String
  
  PrpDesignation = GetProperty(KeyPrpDesignation, Conf, DocExt)
  PrpName = GetProperty(KeyPrpName, Conf, DocExt)
  GetNewName = CurrentFolder + "\" + Trim(Translit(PrpDesignation) + " " + Translit(PrpName)) + ".STEP"
    
End Function
