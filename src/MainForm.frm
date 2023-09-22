VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Export parts to other formats"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6465
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButtonClose_Click()
    ExitApp
End Sub

Private Sub CommandButtonRun_Click()
    Dim Action As TSaveAction
    Dim NeedTranslit As Boolean
    Dim IsNameEn As Boolean
    
    Me.Hide
    If Me.OptionButtonPipeToSTEP.Value Then
        Action = SavePipeToSTEP
        NeedTranslit = Me.CheckBoxPipeToSTEPTranslit.Value
    ElseIf Me.OptionButtonSheetToDWG.Value Then
        Action = SaveSheetToDWG
        NeedTranslit = False
    Else
        Action = SaveSheetToDXF
        NeedTranslit = False
    End If
    IsNameEn = Me.ChkNameEn.Value
    
    Run Action, NeedTranslit, IsNameEn
    ExitApp
End Sub
