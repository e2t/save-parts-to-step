Attribute VB_Name = "Main"
Option Explicit

Const KeyPrpBlank As String = "Заготовка"
Const KeyPrpName As String = "Наименование"
Const KeyPrpDesignation As String = "Обозначение"

Dim swApp As Object
Dim fso As FileSystemObject
Dim currentFolder As String

Sub Main()
    Dim currentDoc As ModelDoc2
    Dim asm As AssemblyDoc
    Dim compArray As Variant
    Dim comp_ As Variant
    Dim comp As Component2
    Dim docconf As String
    Dim doc As ModelDoc2
    Dim docext As ModelDocExtension
    Dim property As String
    Dim regex As RegExp
    Dim keyComp As String
    Dim key_ As Variant
    Dim savedPartCount As Long
    Dim highIndex As Long
    Dim msg As String
    Dim uniqueComp As Dictionary
    Dim stepFiles() As String
    Dim stepName As String
    
    Set swApp = Application.SldWorks
    Set currentDoc = swApp.ActiveDoc
    If currentDoc Is Nothing Then Exit Sub
    If currentDoc.GetType <> swDocASSEMBLY Then
        MsgBox "Только для сборок.", vbCritical
        Exit Sub
    End If
    
    Set fso = New FileSystemObject
    Set uniqueComp = New Dictionary
    
    Set regex = New RegExp
    regex.Pattern = ".*труба.*"
    regex.Global = True
    regex.IgnoreCase = True
    
    currentFolder = fso.GetParentFolderName(currentDoc.GetPathName)
    Set asm = currentDoc
    compArray = asm.GetComponents(False)
    For Each comp_ In compArray
        Set comp = comp_
        keyComp = GetKeyComponent(comp)
        
        If Not uniqueComp.Exists(keyComp) Then
            Set doc = comp.GetModelDoc2
            If Not doc Is Nothing Then
                Set docext = doc.Extension
                docconf = comp.ReferencedConfiguration
                property = GetProperty(KeyPrpBlank, docconf, docext)
                If regex.Test(property) Then
                    stepName = SaveToSTEP(docext, docconf)
                Else
                    stepName = ""
                End If
                uniqueComp.Add keyComp, stepName
            End If
        End If
    Next
    
    ReDim stepFiles(uniqueComp.Count)
    highIndex = -1
    For Each key_ In uniqueComp.Items
        If key_ <> "" Then
            highIndex = highIndex + 1
            stepFiles(highIndex) = key_
        End If
    Next
    savedPartCount = highIndex + 1
        
    msg = "Сохранено компонентов:" + Str(savedPartCount) + "."
    If savedPartCount > 0 Then
        If MsgBox(msg + vbNewLine + "Показать?", vbYesNo) = vbYes Then
            ReDim Preserve stepFiles(highIndex)
            QuickSort stepFiles, 0, highIndex
            Shell "explorer /select,""" & stepFiles(0) & """", vbNormalFocus
        End If
    Else
        MsgBox msg
    End If
End Sub

Function GetKeyComponent(comp As Component2) As String
    Dim basename As String
    
    basename = fso.GetBaseName(comp.GetPathName)
    GetKeyComponent = basename + "@" + comp.ReferencedConfiguration
End Function

Function GetProperty(property As String, conf As String, docext As ModelDocExtension) As String
    Dim value As String
    Dim rawValue As String
    Dim wasResolved As Boolean
    Dim getPrpResult As swCustomInfoGetResult_e
    
    value = ""
    getPrpResult = docext.CustomPropertyManager(conf).Get5(property, False, rawValue, value, wasResolved)
    If getPrpResult = swCustomInfoGetResult_NotPresent Then
        docext.CustomPropertyManager("").Get5 property, False, rawValue, value, wasResolved
    End If
    GetProperty = value
End Function

Function SaveToSTEP(docext As ModelDocExtension, conf As String) As String
    Dim newname As String
    Dim errors As swFileSaveError_e
    Dim warnings As swFileSaveWarning_e
    
    newname = GetNewName(docext, conf)
    docext.SaveAs newname, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, errors, warnings
    SaveToSTEP = newname
End Function

Function GetNewName(docext As ModelDocExtension, conf As String) As String
    Dim prpDesignation As String
    Dim prpName As String
    
    prpDesignation = GetProperty(KeyPrpDesignation, conf, docext)
    prpName = GetProperty(KeyPrpName, conf, docext)
    GetNewName = currentFolder + "\" + prpDesignation + " " + prpName + ".STEP"
End Function
