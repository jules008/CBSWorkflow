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
Private pContactType As String
Private pOrganisation As String
Private pPrimaryContact As Boolean
Private pEmailAddress As String
Private pOptOut As Boolean
Private pLastComm As Date
Private pComFrq As Integer

Private WithEvents pContactForm As FrmContactForm
Attribute pContactForm.VB_VarHelpID = -1
Private WithEvents pCommForm As FrmClientComms
Attribute pCommForm.VB_VarHelpID = -1

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

'---------------------------------------------------------------
Public Property Get LastComm() As Date
    LastComm = pLastComm
End Property

Public Property Let LastComm(ByVal vNewValue As Date)
    pLastComm = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get PrimaryContact() As Boolean
    PrimaryContact = pPrimaryContact
End Property

Public Property Let PrimaryContact(ByVal vNewValue As Boolean)
    pPrimaryContact = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get EmailAddress() As String
    EmailAddress = pEmailAddress
End Property

Public Property Let EmailAddress(ByVal vNewValue As String)
    pEmailAddress = vNewValue
End Property

' ===============================================================
' Method DBNew
' displays new SPV form
' ---------------------------------------------------------------
Public Sub DBNew()
    pContactNo = 0
    pContactName = ""
    pContactType = 0
    pOrganisation = ""
    pAddress = ""
    pPhone1 = ""
    pPhone2 = ""
    pPosition = ""
    pEmailAddress = ""
    pPrimaryContact = False
    pOptOut = False
    pLastComm = ""
    pComFrq = 7
    
    With pContactForm
        .Tag = "New"
        .ClearForm
        .BtnUpdate.Caption = "Create"
        .Show
    End With
End Sub

' ===============================================================
' Method DisplayForm
' displays client form
' ---------------------------------------------------------------
Public Sub DisplayForm()
    With pContactForm
        .Tag = "Update"
        .TxtAddress1 = pAddress
        .TxtContactName = pContactName
        .TxtContactNo = pContactNo
        .CmoContactType = pContactType
        .CmoOrganisation = pOrganisation
        .TxtPhone1 = pPhone1
        .TxtPhone2 = pPhone2
        .TxtPosition = pPosition
        .TxtEmailAdd = pEmailAddress
        .ChkPrimaryContact = pPrimaryContact
        .ChkOptOut = pOptOut
        .CmoCommFreq = pComFrq
        If pLastComm = 0 Then .TxtLastComm = "" Else .TxtLastComm = pLastComm
        .Show
    End With
End Sub

' ===============================================================
' Method DisplayCommForm
' displays client Communication form
' ---------------------------------------------------------------
Public Sub DisplayCommForm()
    Dim RstCommsList As Recordset
    
    Set RstCommsList = GetCommsList
    
    With pCommForm
        .AddFields RstCommsList
        .Show
    End With
    
    Set RstCommsList = Nothing
End Sub

' ===============================================================
' Method CommsNo
' Returns the number of communications for the todo button
' ---------------------------------------------------------------
Public Function CommsNo()
    Dim RstCommsList As Recordset
    
    Set RstCommsList = ModDatabase.SQLQuery("SELECT Count(ContactNo)AS [CommsNo] FROM TblContact " _
                                            & "WHERE (ContactType='Lead' OR ContactType='Client') AND LastComm<=Date()-[comfrq] AND OptOut=False ")
    
    CommsNo = RstCommsList!CommsNo
End Function

' ===============================================================
' Method DBGet(Key as Integer)
' Gets class from Database
' ---------------------------------------------------------------
Public Sub DBGet(ContactNo As Integer)
    Dim RstContact As Recordset
    Dim RstNoCont As Recordset

    Set RstContact = ModDatabase.SQLQuery("SELECT * FROM TblContact WHERE ContactNo = " & ContactNo)
    Set RstNoCont = ModDatabase.SQLQuery("SELECT COUNT (ContactNo) FROM TblContact " _
                                            & "WHERE ContactType = '" & pContactType & "' and Organisation = '" & pOrganisation & "'")
    With RstContact
        If .RecordCount > 0 Then
            If Not IsNull(!ContactNo) Then pContactNo = !ContactNo
            If Not IsNull(!Position) Then pPosition = !Position
            If Not IsNull(!ContactName) Then pContactName = !ContactName
            If Not IsNull(!Address) Then pAddress = !Address
            If Not IsNull(!Organisation) Then pOrganisation = !Organisation
            If Not IsNull(!ContactType) Then pContactType = !ContactType
            If Not IsNull(!Phone1) Then pPhone1 = !Phone1
            If Not IsNull(!Phone2) Then pPhone2 = !Phone2
            If Not IsNull(!PrimaryContact) Then pPrimaryContact = !PrimaryContact
            If Not IsNull(!EmailAddress) Then pEmailAddress = !EmailAddress
            If Not IsNull(!OptOut) Then pOptOut = !OptOut
            If Not IsNull(!ComFrq) Then pComFrq = !ComFrq
            If Not IsNull(!LastComm) Then pLastComm = !LastComm
            If RstNoCont.Fields(0) = 0 Then
                pPrimaryContact = True
                DB.Execute "UPDATE TblContact SET PrimaryContact = TRUE WHERE ContactType = '" & pContactType & "' and Organisation = '" & pOrganisation & "'"
            End If
        End If
    End With
    Set RstContact = Nothing
    Set RstNoCont = Nothing
End Sub

' ===============================================================
' Method DBSave
' Saves class to Database
' ---------------------------------------------------------------
Public Sub DBSave()
    Dim RstContact As Recordset
    Dim RstMaxNo As Recordset
    Dim RstNoCont As Recordset
    
    Dim LastNo As Integer

    Set RstContact = ModDatabase.SQLQuery("SELECT * FROM TblContact WHERE ContactNo = " & pContactNo)
    Set RstMaxNo = ModDatabase.SQLQuery("SELECT MAX(ContactNo) FROM TblContact ")
    Set RstNoCont = ModDatabase.SQLQuery("SELECT COUNT (ContactNo) FROM TblContact " _
                                            & "WHERE ContactType = '" & pContactType & "' and Organisation = '" & pOrganisation & "'")

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

        'set to primary contact if this is the only contact for the organisation
        If RstNoCont.Fields(0) = 0 Then pPrimaryContact = True
        
        'if it has been flagged as a primary contact, clear all others for that organsiation
        If pPrimaryContact Then DB.Execute "UPDATE TblContact SET PrimaryContact = False WHERE ContactType = '" & pContactType & "' and Organisation = '" & pOrganisation & "'"
        
        !ContactNo = CleanSQLText(pContactNo, True)
        !Position = CleanSQLText(pPosition, True)
        !ContactName = CleanSQLText(pContactName, True)
        !Address = CleanSQLText(pAddress, True)
        !ContactType = CleanSQLText(pContactType, True)
        !Organisation = CleanSQLText(pOrganisation, True)
        !Phone1 = CleanSQLText(pPhone1, True)
        !Phone2 = CleanSQLText(pPhone2, True)
        !PrimaryContact = CleanSQLText(pPrimaryContact, True)
        !OptOut = CleanSQLText(pOptOut, True)
        !LastComm = CleanSQLText(pLastComm, True)
        !ComFrq = CleanSQLText(pComFrq, True)

        .Update
    End With

    Set RstContact = Nothing
    Set RstMaxNo = Nothing
    Set RstNoCont = Nothing
End Sub

' ===============================================================
' Method DBDelete
' Deletes Record
' ---------------------------------------------------------------
Public Sub DBDelete(Optional FullDelete As Boolean)

    DB.Execute "DELETE * FROM TblContact WHERE ContactNo = " & pContactNo

End Sub

' ===============================================================
' Method GetCommsList
' Gets all contacts that are eligible for communication today
' ---------------------------------------------------------------
Private Function GetCommsList() As Recordset
    Dim RstCommsList As Recordset
    
    Set RstCommsList = ModDatabase.SQLQuery("SELECT ContactNo, ContactName, EmailAddress, ContactType, Organisation " _
                                            & "FROM TblContact " _
                                            & "WHERE (ContactType='Lead' OR ContactType='Client') AND LastComm<=Date()-[comfrq] AND OptOut=False ")
    
    
    Set GetCommsList = RstCommsList
    
    Set RstCommsList = Nothing
End Function

' ===============================================================
' Event pContactForm_Update
' Event from client form to update Contact
' ---------------------------------------------------------------
Private Sub pContactForm_Update()
    With pContactForm
        pAddress = CleanString(.TxtAddress1)
        pContactName = CleanString(.TxtContactName)
        pPhone1 = CleanString(.TxtPhone1)
        pPhone2 = CleanString(.TxtPhone2)
        pPosition = CleanString(.TxtPosition)
        pContactType = CleanString(.CmoContactType)
        pOrganisation = CleanString(.CmoOrganisation)
        pEmailAddress = Trim(.TxtEmailAdd)
        pPrimaryContact = CleanString(.ChkPrimaryContact)
        pOptOut = CleanString(.ChkOptOut)
        If .TxtLastComm = "" Then pLastComm = 0 Else pLastComm = .TxtLastComm
        pComFrq = CleanString(.CmoCommFreq)
    End With
    DBSave
    
    Select Case pContactType
        Case "Client"
            Set ActiveClient = New ClsClient
            With ActiveClient
                .DBGet pOrganisation
                .Contacts.Add Me
                .DBSave
            End With
        Case "Lender"
            Set ActiveLender = New ClsLender
            With ActiveLender
                .DBGet pOrganisation
                .Contacts.Add Me
                .DBSave
            End With
        Case "Project"
            Set ActiveProject = New ClsProject
            With ActiveProject
                .DBGet pOrganisation
                .Contacts.Add Me
                .DBSave
            End With
        Case "SPV"
            Set ActiveSPV = New ClsSPV
            With ActiveSPV
                .DBGet pOrganisation
                .Contacts.Add Me
                .DBSave
            End With
    End Select
End Sub

' ===============================================================
' Event pContactForm_CreateNew
' Event from SPV form to create new Contact
' ---------------------------------------------------------------
Private Sub pContactForm_CreateNew()
    With pContactForm
        .Hide
    End With
    DBNew
End Sub

' ===============================================================
' Event pContactForm_Delete
' Deletes Contact
' ---------------------------------------------------------------
Private Sub pContactForm_Delete()
    DBDelete
End Sub

' ===============================================================
' Event pCommForm_MailSent
' Marks contact as communication sent
' ---------------------------------------------------------------
Private Sub pCommForm_MailSent(ContactNo As Integer)
    DB.Execute "UPDATE TblContact SET LastComm = Now WHERE ContactNo = " & ContactNo
End Sub

' ===============================================================
Private Sub Class_Initialize()
    Set pContactForm = New FrmContactForm
    Set pCommForm = New FrmClientComms
End Sub

' ===============================================================
Private Sub Class_Terminate()
    Set pContactForm = Nothing
    Set pCommForm = Nothing
End Sub
