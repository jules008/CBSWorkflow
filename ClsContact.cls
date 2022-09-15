VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsContact
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 13 Sep 22
'===============================================================
' Methods
'---------------------------------------------------------------
' DBGet - Gets class from Database
' DBSave - Saves class to Database
' DBDelete - Deletes record
'===============================================================
' Index - ContactNo
'---------------------------------------------------------------
Option Explicit

Private pContactNo As Integer
Private pPosition As String
Private pContactName As String
Private pAddress As String
Private pPhone1 As String
Private pPhone2 As String

'---------------------------------------------------------------
Public Property Get ContactNo() As Integer
    ContactNo = pContactNo
End Property

Public Property Let ContactNo(ByVal vNewValue As Integer)
    pContactNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Position() As String
    Position = pPosition
End Property

Public Property Let Position(ByVal vNewValue As String)
    pPosition = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ContactName() As String
    ContactName = pContactName
End Property

Public Property Let ContactName(ByVal vNewValue As String)
    pContactName = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Address() As String
    Address = pAddress
End Property

Public Property Let Address(ByVal vNewValue As String)
    pAddress = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Phone1() As String
    Phone1 = pPhone1
End Property

Public Property Let Phone1(ByVal vNewValue As String)
    pPhone1 = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Phone2() As String
    Phone2 = pPhone2
End Property

Public Property Let Phone2(ByVal vNewValue As String)
    pPhone2 = vNewValue
End Property

' ===============================================================
' Method DBGet(Key as Integer)
' Gets class from Database
' ---------------------------------------------------------------
Public Sub DBGet(ContactNo As Integer)
    Dim RstContact As Recordset

    Set RstContact = ModDatabase.SQLQuery("SELECT * FROM TblContact WHERE ContactNo = " & ContactNo)
    With RstContact
        If .RecordCount > 0 Then
            If Not IsNull(!ContactNo) Then pContactNo = !ContactNo
            If Not IsNull(!Position) Then pPosition = !Position
            If Not IsNull(!ContactName) Then pContactName = !ContactName
            If Not IsNull(!Address) Then pAddress = !Address
            If Not IsNull(!Phone1) Then pPhone1 = !Phone1
            If Not IsNull(!Phone2) Then pPhone2 = !Phone2
        End If
    End With
    Set RstContact = Nothing
End Sub

' ===============================================================
' Method DBSave
' Saves class to Database
' ---------------------------------------------------------------
Public Sub DBSave()
    Dim RstContact As Recordset
    Dim RstMaxNo As Recordset
    Dim LastNo As Integer

    Set RstContact = ModDatabase.SQLQuery("SELECT * FROM TblContact WHERE ContactNo = " & pContactNo)
    Set RstMaxNo = ModDatabase.SQLQuery("SELECT MAX(ContactNo) FROM TblContact ")

    If RstMaxNo.Fields(0).Value <> 0 Then
        LastNo = RstMaxNo.Fields(0).Value
    Else
        LastNo = 0
    End If

    With RstContact
        If .RecordCount = 0 Then
            .AddNew
            pContactNo = LastNo + 1
        Else
            .Edit
        End If

        If CStr(pContactNo) = "" Then !ContactNo = Null Else !ContactNo = pContactNo
        If CStr(pPosition) = "" Then !Position = Null Else !Position = pPosition
        If CStr(pContactName) = "" Then !ContactName = Null Else !ContactName = pContactName
        If CStr(pAddress) = "" Then !Address = Null Else !Address = pAddress
        If CStr(pPhone1) = "" Then !Phone1 = Null Else !Phone1 = pPhone1
        If CStr(pPhone2) = "" Then !Phone2 = Null Else !Phone2 = pPhone2

        .Update
    End With

    Set RstContact = Nothing
    Set RstMaxNo = Nothing
End Sub

' ===============================================================
' Method DBDelete
' Deletes Record
' ---------------------------------------------------------------
Public Sub DBDelete(Optional FullDelete As Boolean)

    DB.Execute "DELETE * FROM TblContact WHERE ContactNo = " & pContactNo

End Sub
' ===============================================================

' ===============================================================
Private Sub Class_Terminate()
End Sub
' ===============================================================
