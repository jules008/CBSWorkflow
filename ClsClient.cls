VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
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
' Date - 09 Sep 22
'===============================================================
' Methods
'---------------------------------------------------------------
' DBGet - Gets class from Database
' DBSave - Saves class to Database
' DBDelete - Deletes record
'===============================================================
' Index - ClientNo
'---------------------------------------------------------------
Option Explicit

Private pName As String
Private pUrl As String
Private pPhoneNo As String
Private pClientNo As Integer

'---------------------------------------------------------------
Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal vNewValue As String)
    pName = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Url() As String
    Url = pUrl
End Property

Public Property Let Url(ByVal vNewValue As String)
    pUrl = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get PhoneNo() As String
    PhoneNo = pPhoneNo
End Property

Public Property Let PhoneNo(ByVal vNewValue As String)
    pPhoneNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ClientNo() As Integer
    ClientNo = pClientNo
End Property

Public Property Let ClientNo(ByVal vNewValue As Integer)
    pClientNo = vNewValue
End Property

' ===============================================================
' Method DBGet(Key as Integer)
' Gets class from Database
' ---------------------------------------------------------------
Public Sub DBGet(ClientNo As Integer)
    Dim RstClient As Recordset

    Set RstClient = ModDatabase.SQLQuery("SELECT * FROM TblClient WHERE ClientNo = " & ClientNo & " AND Deleted IS NULL")
    With RstClient
        If .RecordCount > 0 Then
            If Not IsNull(!Name) Then pName = !Name
            If Not IsNull(!Url) Then pUrl = !Url
            If Not IsNull(!PhoneNo) Then pPhoneNo = !PhoneNo
            If Not IsNull(!ClientNo) Then pClientNo = !ClientNo
        End If
    End With
    Set RstClient = Nothing
End Sub

' ===============================================================
' Method DBSave
' Saves class to Database
' ---------------------------------------------------------------
Public Sub DBSave()
    Dim RstClient As Recordset
    Dim RstMaxNo As Recordset
    Dim LastNo As Integer

    Set RstClient = ModDatabase.SQLQuery("SELECT * FROM TblClient WHERE ClientNo = " & pClientNo & " AND Deleted IS NULL")
    Set RstMaxNo = ModDatabase.SQLQuery("SELECT MAX(ClientNo) FROM TblClient ")

    If RstMaxNo.Fields(0).Value <> 0 Then
        LastNo = RstMaxNo.Fields(0).Value
    Else
        LastNo = 0
    End If

    With RstClient
        If .RecordCount = 0 Then
            .addnew
            pClientNo = LastNo + 1
        Else
            .edit
        End If

        !Url = pUrl
        !PhoneNo = pPhoneNo
        !ClientNo = pClientNo
        .update
    End With

    Set RstClient = Nothing
    Set RstMaxNo = Nothing
End Sub

' ===============================================================
' Method DBDelete
' Deletes Record
' ---------------------------------------------------------------
Public Sub DBDelete(Optional FullDelete As Boolean)

    DB.Execute "DELETE * FROM TblClient WHERE ClientNo = " & pClientNo

End Sub
' ===============================================================

' ===============================================================
Private Sub Class_Terminate()
End Sub
' ===============================================================
