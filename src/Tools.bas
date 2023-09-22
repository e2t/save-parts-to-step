Attribute VB_Name = "Tools"
Option Explicit

Function GetKeyComponent(DocPath As String, Conf As String) As String
    GetKeyComponent = gFSO.GetBaseName(DocPath) + "@" + Conf
End Function

Function GetProperty(Property As String, Conf As String, DocExt As ModelDocExtension) As String
    Const UseCached = False
    Dim Value As String
    Dim RawValue As String
    Dim WasResolved As Boolean
    Dim GetPrpResult As swCustomInfoGetResult_e
    
    Value = ""
    GetPrpResult = DocExt.CustomPropertyManager(Conf).Get5(Property, UseCached, RawValue, Value, WasResolved)
    If GetPrpResult = swCustomInfoGetResult_NotPresent Then
        DocExt.CustomPropertyManager("").Get5 Property, UseCached, RawValue, Value, WasResolved
    End If
    GetProperty = Trim(Value)
End Function

Sub ActivatePartConfiguration(Doc As ModelDoc2, Conf As String)
    Dim Errors As swActivateDocError_e
    
    swApp.ActivateDoc3 Doc.GetPathName, False, swDontRebuildActiveDoc, Errors
    Doc.ShowConfiguration2 Conf
End Sub

Sub CreateEmptyPDF(NewName As String)
    Const EmptyPDFBody = _
        "%PDF-1.4" + vbNewLine + _
        "1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj" + vbNewLine + _
        "2 0 obj<</Type/Pages/Count 1/Kids[3 0 R]>>endobj" + vbNewLine + _
        "3 0 obj<</Type/Page/MediaBox[0 0 595 842]/Parent 2 0 R/Resources<<>>>>endobj" + vbNewLine + _
        "xref" + vbNewLine + _
        "0 4" + vbNewLine + _
        "0000000000 65535 f" + vbNewLine + _
        "0000000009 00000 n" + vbNewLine + _
        "0000000052 00000 n" + vbNewLine + _
        "0000000101 00000 n" + vbNewLine + _
        "trailer<</Size 4/Root 1 0 R>>" + vbNewLine + _
        "startxref" + vbNewLine + _
        "178" + vbNewLine + _
        "%%EOF"
      
    Dim TextFile As TextStream
    
    Set TextFile = gFSO.CreateTextFile(NewName)
    TextFile.Write EmptyPDFBody
    TextFile.Close
End Sub

Sub SaveToDWG(Doc As ModelDoc2, Conf As String, NewName As String)
    Dim Part As PartDoc
    Dim SheetMetalOptions As Integer
    Dim DataAlignment(11) As Double
    
    ActivatePartConfiguration Doc, Conf
    
    Set Part = Doc
    SheetMetalOptions = 1
    DataAlignment(0) = 0#
    DataAlignment(1) = 0#
    DataAlignment(2) = 0#
    DataAlignment(3) = 1#
    DataAlignment(4) = 0#
    DataAlignment(5) = 0#
    DataAlignment(6) = 0#
    DataAlignment(7) = 1#
    DataAlignment(8) = 0#
    DataAlignment(9) = 0#
    DataAlignment(10) = 0#
    DataAlignment(11) = 1#
    Part.ExportToDWG2 NewName, Doc.GetPathName, swExportToDWG_ExportSheetMetal, True, DataAlignment, False, False, SheetMetalOptions, Null
    
    swApp.QuitDoc Doc.GetPathName
End Sub

Sub SaveToSTEP(Doc As ModelDoc2, Conf As String, NewName As String)
    Dim Errors As swFileSaveError_e
    Dim Warnings As swFileSaveWarning_e
    
    ActivatePartConfiguration Doc, Conf
    Doc.Extension.SaveAs NewName, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, Errors, Warnings
    swApp.QuitDoc Doc.GetPathName
End Sub

Function PrepareNewName(BaseName As String, FolderName As String, ext As String) As String
    PrepareNewName = gFSO.BuildPath(FolderName, BaseName + "." + ext)
    If Not gFSO.FolderExists(FolderName) Then
        gFSO.CreateFolder FolderName
    End If
End Function

Function FindFeatureThisType(TypeName As String, Model As ModelDoc2) As Feature
    Dim Feat As Feature
    
    Set Feat = Model.FirstFeature
    Do While Not Feat Is Nothing
        If Feat.GetTypeName2 = TypeName Then
            Set FindFeatureThisType = Feat
            Exit Do
        End If
        Set Feat = Feat.GetNextFeature
    Loop
End Function

Function IsSheetMetal(Doc As ModelDoc2) As Boolean
    Dim swSheetMetalFolder As SheetMetalFolder
    Dim SheetMetalFeat As Feature
    
    Set swSheetMetalFolder = Doc.FeatureManager.GetSheetMetalFolder
    If swSheetMetalFolder Is Nothing Then  'for models created in SolidWorks 2012 and earlier
        Set SheetMetalFeat = FindFeatureThisType("SheetMetal", Doc)
    Else
        Set SheetMetalFeat = swSheetMetalFolder.GetFeature
    End If
    IsSheetMetal = Not (SheetMetalFeat Is Nothing)
End Function

Function ConvertStringToChangeNumber(ChangeNumberProperty As String) As Integer
    ConvertStringToChangeNumber = 0
    On Error Resume Next
    ConvertStringToChangeNumber = CInt(ChangeNumberProperty)
End Function

Function GetChangeNumber(Doc As ModelDoc2, Designation As String, Name As String, ByRef Number As Integer) As Boolean
    Dim DrawingName As String
    Dim Model As ModelDoc2
    Dim Errors As swFileLoadError_e
    Dim Errors2 As swActivateDocError_e
    Dim Warnings As swFileLoadWarning_e
    Dim DocFolder As String
    
    GetChangeNumber = False
    DocFolder = gFSO.GetParentFolderName(Doc.GetPathName)
    DrawingName = gFSO.BuildPath(DocFolder, Designation + " " + Name + ".SLDDRW")
    If gFSO.FileExists(DrawingName) Then
        Set Model = swApp.OpenDoc6(DrawingName, swDocDRAWING, _
            swOpenDocOptions_Silent + swOpenDocOptions_ViewOnly + swOpenDocOptions_ReadOnly, _
            "", Errors, Warnings)
        Number = ConvertStringToChangeNumber(GetProperty("Изменение", "", Model.Extension))
        GetChangeNumber = True
        swApp.QuitDoc DrawingName
        swApp.ActivateDoc3 Doc.GetPathName, False, swDontRebuildActiveDoc, Errors2
    End If
End Function

Function GetBaseDesignation(Designation As String) As String
    Dim LastFullstopPosition As Integer
    Dim FirstHyphenPosition As Integer
    
    GetBaseDesignation = Designation
    LastFullstopPosition = InStrRev(Designation, ".")
    If LastFullstopPosition > 0 Then
        FirstHyphenPosition = InStr(LastFullstopPosition, Designation, "-")
        If FirstHyphenPosition > 0 Then
            GetBaseDesignation = Left(Designation, FirstHyphenPosition - 1)
        End If
    End If
End Function
