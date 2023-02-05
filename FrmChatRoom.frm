VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmChatRoom 
   Caption         =   "Project Chat Room"
   ClientHeight    =   8370.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   OleObjectBlob   =   "FrmChatRoom.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmChatRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===============================================================
' Module FrmChatRoom
' Chat room form
'---------------------------------------------------------------
' Created by Julian Turner
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 10 Oct 22
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmChatRoom"

Public ProjectNo As Integer
Public Event SendMessage()

'===============================================================
' ClearForm
'---------------------------------------------------------------
Public Sub ClearForm()
    xTxtAllMessages = ""
    xTxtNewMessage = ""
    
End Sub

'===============================================================
' BtnClose_Click
'---------------------------------------------------------------
Private Sub BtnClose_Click()
    Hide
End Sub

'===============================================================
' UserForm_Activate
'---------------------------------------------------------------
Private Sub UserForm_Activate()
    Ttl = "Project " & ProjectNo & " Chat Room"
End Sub

'===============================================================
' UserForm_Click
'---------------------------------------------------------------
Private Sub xBtnSend_Click()
    RaiseEvent SendMessage
End Sub
