VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmInputBox 
   Caption         =   "Enter Time"
   ClientHeight    =   2295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4395
   OleObjectBlob   =   "FrmInputBox.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "FrmInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




'===============================================================
' Module FrmInputBox
' displays form to select time
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 06 Jul 20
'===============================================================

Private Const StrMODULE As String = "FrmInputBox"

Option Explicit

Public ReturnValue As String


Private Sub BtnEnter_Click()
    ReturnValue = TxtInput
    Hide
End Sub

Private Sub UserForm_Initialize()
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    

End Sub
