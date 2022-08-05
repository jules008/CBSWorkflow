VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsEmails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsEmails
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 11 May 20
'===============================================================

Option Explicit

Private pEmails As Collection

'---------------------------------------------------------------
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pEmails.[_NewEnum]
End Function

' ===============================================================
' Method FindItem
' Finds Item from index
'---------------------------------------------------------------
Public Function FindItem(EmailNo As Variant) As ClsEmail
Attribute FindItem.VB_UserMemId = 0
    On Error Resume Next
    Set FindItem = pEmails.Item(EmailNo)
End Function

' ===============================================================
' Method AddItem
' Add item to Collection
'---------------------------------------------------------------
Public Sub AddItem(Email As ClsEmail)
    If pEmails Is Nothing Then
        Set pEmails = New ClsUIFrames
        pEmails.Initialize Me
    End If
    pEmails.Add Email, Key:=CStr(Email.EmailNo)
End Sub

' ===============================================================
' Method RemoveItem
' Deletes item from collection
'---------------------------------------------------------------
Public Sub RemoveItem(EmailNo As Variant)
    pEmails.Remove EmailNo
End Sub

' ===============================================================
' Method ItemCount
' Returns number of items in collection
'---------------------------------------------------------------
Public Function Count() As Integer
    Count = pEmails.Count
End Function

' ===============================================================
' Method GetCollection
' Gets all items in collection
'---------------------------------------------------------------
Public Sub GetCollection()
    Dim RstEmail As Recordset
    Dim Email As ClsEmail
    Dim i As Integer

    Set RstEmail = ModDatabase.SQLQuery("SELECT * FROM TblEmail WHERE Deleted IS NULL")
    
    If RstEmail.RecordCount = 0 Then Exit Sub
    
    With RstEmail
        .MoveLast
        .MoveFirst
        For i = 1 To .RecordCount
            Set Email = New ClsEmail
            Email.DBGet !EmailNo
            Me.AddItem Email
            .MoveNext
        Next
    End With
End Sub

' ===============================================================
' Method DeleteCollection
' Deletes all items in collection
'---------------------------------------------------------------
Public Sub DeleteCollection()
    Dim Email As ClsEmail
    For Each Email In pEmails
        pEmails.Remove CStr(Email.EmailNo)
    Next
End Sub

' ===============================================================
Private Sub Class_Initialize()
    GetCollection
End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
End Sub
'---------------------------------------------------------------
