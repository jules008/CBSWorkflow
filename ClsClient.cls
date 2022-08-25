VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
End
Attribute VB_Name = "ClsClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsClient
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 20 Aug 22
'===============================================================
' Methods
'---------------------------------------------------------------
' DBGet ( As Integer) - Gets class from Database
' DBSave - Saves class to Database
' DBDelete(Optional FullDelete As Boolean) - Marks record as deleted or fully deletes
'===============================================================

Option Explicit
Private pDeleted as Date
Private pName as String

'---------------------------------------------------------------
Public Property Get Deleted() As Date
    Deleted = pDeleted
End Property

Public Property Let Deleted(ByVal vNewValue As Date)
    pDeleted = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal vNewValue As String)
    pName = vNewValue
End Property

'---------------------------------------------------------------
' ===============================================================
' Method DBGet ( As Integer)
' Gets class from Database
'---------------------------------------------------------------
Public Sub DBGet ( As Integer)

End Sub

' ===============================================================
' Method DBSave
' Saves class to Database
'---------------------------------------------------------------
Public Sub DBSave

End Sub

' ===============================================================
' Method DBDelete(Optional FullDelete As Boolean)
' Marks record as deleted or fully deletes
'---------------------------------------------------------------
Public Sub DBDelete(Optional FullDelete As Boolean)

End Sub


' ===============================================================
