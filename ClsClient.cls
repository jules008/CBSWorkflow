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
' Date - 13 Sep 22
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
Private pSPVs As ClsSPVs
Private pContacts As ClsContacts
Private WithEvents pClientForm As FrmClientForm
Attribute pClientForm.VB_VarHelpID = -1

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

'---------------------------------------------------------------
Public Property Get SPVs() As ClsSPVs
    Set SPVs = pSPVs
End Property

Public Property Let SPVs(ByRef vNewValue As ClsSPVs)
    Set pSPVs = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Contacts() As ClsContacts
    Set Contacts = pContacts
End Property

Public Property Let Contacts(ByRef vNewValue As ClsContacts)
    Set pContacts = vNewValue
End Property

' ===============================================================
' Method DBNew
' displays new client form
' ---------------------------------------------------------------
Public Sub DBNew()
    pClientNo = 0
    pName = ""
    pPhoneNo = ""
    pUrl = ""
    
    With pClientForm
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
    With pClientForm
        .Tag = "Update"
        .TxtClientNo = pClientNo
        .TxtName = pName
        .TxtPhoneNo = pPhoneNo
        .TxtUrl = pUrl
        .Show
    End With
End Sub

' ===============================================================
' Method DBGet(Key as variant)
' Gets class from Database
' ---------------------------------------------------------------
Public Sub DBGet(ClientNameNo As Variant)
    Dim RstClient As Recordset

    If IsNumeric(ClientNameNo) Then
        Set RstClient = ModDatabase.SQLQuery("SELECT * FROM TblClient WHERE ClientNo = " & ClientNameNo)
    Else
        Set RstClient = ModDatabase.SQLQuery("SELECT * FROM TblClient WHERE Name = '" & ClientNameNo & "'")
    End If
    
    With RstClient
        If .RecordCount > 0 Then
            If Not IsNull(!Name) Then pName = !Name
            If Not IsNull(!Url) Then pUrl = !Url
            If Not IsNull(!PhoneNo) Then pPhoneNo = !PhoneNo
            If Not IsNull(!ClientNo) Then pClientNo = !ClientNo
        pContacts.GetCollection "Client", pName
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

    Set RstClient = ModDatabase.SQLQuery("SELECT * FROM TblClient WHERE ClientNo = " & pClientNo)
    Set RstMaxNo = ModDatabase.SQLQuery("SELECT MAX(ClientNo) FROM TblClient ")

    If RstMaxNo.Fields(0).Value <> 0 Then
        LastNo = RstMaxNo.Fields(0).Value
    Else
        LastNo = 0
    End If

    With RstClient
        If .RecordCount = 0 Then
            .AddNew
            pClientNo = LastNo + 1
        Else
            .Edit
        End If

        If CStr(pName) = "" Then !Name = Null Else !Name = pName
        If CStr(pUrl) = "" Then !Url = Null Else !Url = pUrl
        If CStr(pPhoneNo) = "" Then !PhoneNo = Null Else !PhoneNo = pPhoneNo
        If CStr(pClientNo) = "" Then !ClientNo = Null Else !ClientNo = pClientNo

        pSPVs.SaveCollection
        pContacts.SaveCollection
        .Update
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
' Event pClientForm_Update
' Event from client form to update client
' ---------------------------------------------------------------
Private Sub pClientForm_Update()
    With pClientForm
        pName = .TxtName
        pPhoneNo = .TxtPhoneNo
        pUrl = .TxtUrl
    End With
    DBSave
End Sub

' ===============================================================
' Event pClientForm_CreateNew
' Event from client form to create new client
' ---------------------------------------------------------------
Private Sub pClientForm_CreateNew()
    With pClientForm
        .Hide
    End With
    DBNew
End Sub

' ===============================================================
' Event pClientForm_Delete
' Deletes client
' ---------------------------------------------------------------
Private Sub pClientForm_Delete()
    DBDelete
End Sub
' ===============================================================

Private Sub Class_Initialize()
    Set pSPVs = New ClsSPVs
    Set pContacts = New ClsContacts
    Set pClientForm = New FrmClientForm

    pSPVs.Initialize Me
End Sub
' ===============================================================
Private Sub Class_Terminate()
    Set pSPVs = Nothing
    Set pContacts = Nothing
    Set pClientForm = Nothing
End Sub
' ===============================================================
