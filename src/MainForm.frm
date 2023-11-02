VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "SavePartsToSTEP v23.2"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5745
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const KeyEmptyPDF = "EmptyPDF"

Private Sub CommandButtonClose_Click()
    ExitApp
End Sub

Private Sub CommandButtonRun_Click()
    Dim Action As TSaveAction
    Dim NeedTranslit As Boolean
    Dim IsNameEn As Boolean
    Dim NeedCreateEmptyPDF As Boolean
    
    Me.Hide
    If Me.OptionButtonPipeToSTEP.value Then
        Action = SavePipeToSTEP
        NeedTranslit = Me.CheckBoxPipeToSTEPTranslit.value
    ElseIf Me.OptionButtonSheetToDWG.value Then
        Action = SaveSheetToDWG
        NeedTranslit = False
    Else
        Action = SaveSheetToDXF
        NeedTranslit = False
    End If
    IsNameEn = Me.ChkNameEn.value
    NeedCreateEmptyPDF = Me.emptyPdfChk.value
    
    Run Action, NeedTranslit, IsNameEn, NeedCreateEmptyPDF
    ExitApp
End Sub

Private Sub emptyPdfChk_Change()
    SaveBoolSetting KeyEmptyPDF, Me.emptyPdfChk.value
End Sub

Private Sub UserForm_Initialize()
    Me.emptyPdfChk.value = GetBoolSetting(KeyEmptyPDF, True)
End Sub
