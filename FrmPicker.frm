VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmPicker 
   Caption         =   "DataPicker"
   ClientHeight    =   5730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6075
   OleObjectBlob   =   "FrmPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' Module FrmPicker
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 07 Oct 22
'===============================================================

 Option Explicit

Private Const StrMODULE As String = "FrmPicker"
Public Event SearchTextChanged()
Public Event ResultsListSelect()
Public Event ItemSelected()
Public Event CreateNew()

Private Sub BtnClose_Click()
    Unload Me
End Sub

Private Sub BtnNew_Click()
    RaiseEvent CreateNew
    Unload Me
End Sub

Private Sub BtnSelect_Click()
    RaiseEvent ItemSelected
End Sub

Private Sub LstResults_Click()
    RaiseEvent ResultsListSelect
End Sub

Private Sub TxtSearch_Change()
    RaiseEvent SearchTextChanged
End Sub

