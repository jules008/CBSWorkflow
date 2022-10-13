VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsLender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsLender
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 13 Oct 22
'===============================================================
' Methods
'---------------------------------------------------------------
' DBGet - Gets class from Database
' DBSave - Saves class to Database
' DBDelete - Deletes record
'===============================================================
' Index - LenderNo
'---------------------------------------------------------------
Option Explicit

Private pLenderNo As Integer
Private pName As String
Private pPhoneNo As String
Private pLenderType As String
Private pAddress As String
Private pContacts As ClsContacts
Private WithEvents pLenderForm As FrmLenderForm
Attribute pLenderForm.VB_VarHelpID = -1

'---------------------------------------------------------------
Public Property Get LenderNo() As Integer
    LenderNo = pLenderNo
End Property

Public Property Let LenderNo(ByVal vNewValue As Integer)
    pLenderNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal vNewValue As String)
    pName = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get PhoneNo() As String
    PhoneNo = pPhoneNo
End Property

Public Property Let PhoneNo(ByVal vNewValue As String)
    pPhoneNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get LenderType() As String
    LenderType = pLenderType
End Property

Public Property Let LenderType(ByVal vNewValue As String)
    pLenderType = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Address() As String
    Address = pAddress
End Property

Public Property Let Address(ByVal vNewValue As String)
    pAddress = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Contacts() As ClsContacts
    Set Contacts = pContacts
End Property

Public Property Let Contacts(ByVal vNewValue As ClsContacts)
    Set pContacts = vNewValue
End Property

' ===============================================================
' Method DBNew
' displays new SPV form
' ---------------------------------------------------------------
Public Sub DBNew()
    pAddress = ""
    pPhoneNo = ""
    pName = ""
    pLenderType = ""
    
    With pLenderForm
        .Tag = "New"
        .ClearForm
        .Show
    End With
End Sub

' ===============================================================
' Method DisplayForm
' displays client form
' ---------------------------------------------------------------
Public Sub DisplayForm()
    With pLenderForm
        .Tag = "Update"
        .TxtAddress = pAddress
        .TxtLenderNo = pLenderNo
        .TxtName = pName
        .TxtPhoneNo = pPhoneNo
        .CmoLenderType = ""
        .LstContacts = ""
        .Show
    End With
End Sub

' ===============================================================
' Method DBGet(Key as Integer)
' Gets class from Database
' ---------------------------------------------------------------
Public Sub DBGet(LenderNo As Integer)
    Dim RstLender As Recordset

    Set RstLender = ModDatabase.SQLQuery("SELECT * FROM TblLender WHERE LenderNo = " & LenderNo)
    With RstLender
        If .RecordCount > 0 Then
            If Not IsNull(!LenderNo) Then pLenderNo = !LenderNo
            If Not IsNull(!Name) Then pName = !Name
            If Not IsNull(!PhoneNo) Then pPhoneNo = !PhoneNo
            If Not IsNull(!LenderType) Then pLenderType = !LenderType
            If Not IsNull(!Address) Then pAddress = !Address
        pContacts.GetCollection
        End If
    End With
    Set RstLender = Nothing
End Sub

' ===============================================================
' Method DBSave
' Saves class to Database
' ---------------------------------------------------------------
Public Sub DBSave()
    Dim RstLender As Recordset
    Dim RstMaxNo As Recordset
    Dim LastNo As Integer

    Set RstLender = ModDatabase.SQLQuery("SELECT * FROM TblLender WHERE LenderNo = " & pLenderNo)
    Set RstMaxNo = ModDatabase.SQLQuery("SELECT MAX(LenderNo) FROM TblLender ")

    If RstMaxNo.Fields(0).Value <> 0 Then
        LastNo = RstMaxNo.Fields(0).Value
    Else
        LastNo = 0
    End If

    With RstLender
        If .RecordCount = 0 Then
            .AddNew
            pLenderNo = LastNo + 1
        Else
            .Edit
        End If

        If CStr(pLenderNo) = "" Then !LenderNo = Null Else !LenderNo = pLenderNo
        If CStr(pName) = "" Then !Name = Null Else !Name = pName
        If CStr(pPhoneNo) = "" Then !PhoneNo = Null Else !PhoneNo = pPhoneNo
        If CStr(pLenderType) = "" Then !LenderType = Null Else !LenderType = pLenderType
        If CStr(pAddress) = "" Then !Address = Null Else !Address = pAddress

        pContacts.SaveCollection
        .Update
    End With

    Set RstLender = Nothing
    Set RstMaxNo = Nothing
End Sub

' ===============================================================
' Method DBDelete
' Deletes Record
' ---------------------------------------------------------------
Public Sub DBDelete(Optional FullDelete As Boolean)

    DB.Execute "DELETE * FROM TblLender WHERE LenderNo = " & pLenderNo

End Sub

' ===============================================================
' Event pContactForm_Update
' Event from client form to update Contact
' ---------------------------------------------------------------
Private Sub pLenderForm_Update()
    With pLenderForm
        pAddress = .TxtAddress
        pName = .TxtName
        pPhoneNo = .TxtPhoneNo
        pLenderType = .CmoLenderType
    End With
    DBSave
End Sub

' ===============================================================
' Event pContactForm_CreateNew
' Event from SPV form to create new Contact
' ---------------------------------------------------------------
Private Sub pLenderForm_CreateNew()
    With pLenderForm
        .Hide
    End With
    DBNew
End Sub

' ===============================================================
' Event pContactForm_Delete
' Deletes Contact
' ---------------------------------------------------------------
Private Sub pLenderForm_Delete()
    DBDelete
End Sub

' ===============================================================
Private Sub Class_Initialize()
    Set pContacts = New ClsContacts
    Set pLenderForm = New FrmLenderForm
End Sub
' ===============================================================
Private Sub Class_Terminate()
    Set pContacts = Nothing
    Set pLenderForm = Nothing
End Sub
' ===============================================================

