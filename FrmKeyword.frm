VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmKeyword 
   Caption         =   "Keywords"
   ClientHeight    =   9975.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12765
   OleObjectBlob   =   "FrmKeyword.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "FrmKeyword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' Module FrmKeyword
' Displays keywords
'---------------------------------------------------------------
' Created by Julian Turner
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 29 Nov 20
'===============================================================
Option Explicit
Private EventHandlers As Collection

'===============================================================
' BtnClose_Click
'---------------------------------------------------------------
Private Sub BtnClose_Click()
    Unload Me
End Sub

' ===============================================================
' UserForm_Initialize
' ---------------------------------------------------------------
Private Sub UserForm_Initialize()
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
End Sub
