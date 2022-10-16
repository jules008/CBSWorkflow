VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsChatEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsChatEntry
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 16 Oct 22
'===============================================================
' Methods
'---------------------------------------------------------------
'===============================================================
' Index - ProjectNo
'---------------------------------------------------------------
Option Explicit

Private pProjectNo As Integer
Private pUserName As String
Private pEntryDateTime As Date
Private pMessage As String

'---------------------------------------------------------------
Public Property Get ProjectNo() As Integer
    ProjectNo = pProjectNo
End Property

Public Property Let ProjectNo(ByVal vNewValue As Integer)
    pProjectNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get UserName() As String
    UserName = pUserName
End Property

Public Property Let UserName(ByVal vNewValue As String)
    pUserName = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get EntryDateTime() As Date
    EntryDateTime = pEntryDateTime
End Property

Public Property Let EntryDateTime(ByVal vNewValue As Date)
    pEntryDateTime = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Message() As String
    Message = pMessage
End Property

Public Property Let Message(ByVal vNewValue As String)
    pMessage = vNewValue
End Property

' ===============================================================

Private Sub Class_Initialize()

End Sub
' ===============================================================
Private Sub Class_Terminate()
End Sub
' ===============================================================
