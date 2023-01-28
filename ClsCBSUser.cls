VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsCBSUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsCBSUser
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 10 Oct 22
'===============================================================
' Methods
'---------------------------------------------------------------
' DBGet - Gets class from Database
' DBSave - Saves class to Database
' DBDelete - Deletes record
'===============================================================
' Index - CBSUserNo
'---------------------------------------------------------------
Option Explicit

Private pCBSUserNo As Integer
Private pUserName As String
Private pPosition As String
Private pPhoneNo As String
Private pUserLvl As EnUserLvl
Private pSupervisor As ClsCBSUser

Private WithEvents pCBSUserForm As FrmAdnCBSUser
Attribute pCBSUserForm.VB_VarHelpID = -1

'---------------------------------------------------------------
Public Property Get CBSUserNo() As Integer
    CBSUserNo = pCBSUserNo
End Property

Public Property Let CBSUserNo(ByVal vNewValue As Integer)
    pCBSUserNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get UserName() As String
    UserName = pUserName
End Property

Public Property Let UserName(ByVal vNewValue As String)
    pUserName = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Position() As String
    Position = pPosition
End Property

Public Property Let Position(ByVal vNewValue As String)
    pPosition = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get PhoneNo() As String
    PhoneNo = pPhoneNo
End Property

Public Property Let PhoneNo(ByVal vNewValue As String)
    pPhoneNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get UserLvl() As EnUserLvl
    UserLvl = pUserLvl
End Property

Public Property Let UserLvl(ByVal vNewValue As EnUserLvl)
    pUserLvl = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Supervisor() As ClsCBSUser
    Set Supervisor = pSupervisor
End Property

Public Property Let Supervisor(ByVal Supervisor As ClsCBSUser)
    Set pSupervisor = Supervisor
End Property

' ===============================================================
' Method DBNew
' displays new User form
' ---------------------------------------------------------------
Public Sub DBNew()
    pUserName = ""
    pPosition = ""
    pPhoneNo = ""
    pUserLvl = ""
    Set pSupervisor = Nothing
    
    With pCBSUserForm
        .Tag = "New"
        .Show
    End With
End Sub

' ===============================================================
' Method DisplayForm
' displays client form
' ---------------------------------------------------------------
Public Sub DisplayForm()
    With pCBSUserForm
        .Tag = "Update"
        .TxtCBSUserNo = pCBSUserNo
        .TxtPhoneNo = pPhoneNo
        .TxtPosition = pPosition
        .TxtUserName = pUserName
        If pUserLvl > 0 Then .CmoUserLvl = pUserLvl
        If Not pSupervisor Is Nothing Then .CmoSupervisor = pSupervisor.UserName
        .Show
    End With
End Sub

' ===============================================================
' Method DBGet(Key as Integer)
' Gets class from Database
' ---------------------------------------------------------------
Public Sub DBGet(CBSUserNameOrNo As Variant)
    Dim RstCBSUser As Recordset

    If IsNumeric(CBSUserNameOrNo) Then
        Set RstCBSUser = ModDatabase.SQLQuery("SELECT * FROM TblCBSUser WHERE CBSUserNo = " & CBSUserNameOrNo)
    Else
        Set RstCBSUser = ModDatabase.SQLQuery("SELECT * FROM TblCBSUser WHERE UserName = '" & CBSUserNameOrNo & "'")
    End If
    
    With RstCBSUser
        If .RecordCount > 0 Then
            If Not IsNull(!CBSUserNo) Then pCBSUserNo = !CBSUserNo
            If Not IsNull(!UserName) Then pUserName = !UserName
            If Not IsNull(!Position) Then pPosition = !Position
            If Not IsNull(!PhoneNo) Then pPhoneNo = !PhoneNo
            If Not IsNull(!UserLvl) Then pUserLvl = !UserLvl
            If Not IsNull(!Supervisor) Then
                Set pSupervisor = New ClsCBSUser
                pSupervisor.DBGet !Supervisor
            End If
        End If
    End With
    Set RstCBSUser = Nothing
End Sub

' ===============================================================
' Method DBSave
' Saves class to Database
' ---------------------------------------------------------------
Public Sub DBSave()
    Dim RstCBSUser As Recordset
    Dim RstMaxNo As Recordset
    Dim LastNo As Integer

    Set RstCBSUser = ModDatabase.SQLQuery("SELECT * FROM TblCBSUser WHERE CBSUserNo = " & pCBSUserNo)
    Set RstMaxNo = ModDatabase.SQLQuery("SELECT MAX(CBSUserNo) FROM TblCBSUser ")

    If RstMaxNo.Fields(0).Value <> 0 Then
        LastNo = RstMaxNo.Fields(0).Value
    Else
        LastNo = 0
    End If

    With RstCBSUser
        If .RecordCount = 0 Then
            .AddNew
            pCBSUserNo = LastNo + 1
        Else
            .Edit
        End If

        !CBSUserNo = CleanSQLText(pCBSUserNo, True)
        !UserName = CleanSQLText(pUserName, True)
        !Position = CleanSQLText(pPosition, True)
        !PhoneNo = CleanSQLText(pPhoneNo, True)
        !UserLvl = CleanSQLText(pUserLvl, True)
        If Not pSupervisor Is Nothing Then !Supervisor = CleanSQLText(pSupervisor.CBSUserNo, True)

        .Update
    End With

    Set RstCBSUser = Nothing
    Set RstMaxNo = Nothing
End Sub

' ===============================================================
' Method DBDelete
' Deletes Record
' ---------------------------------------------------------------
Public Sub DBDelete(Optional FullDelete As Boolean)

    DB.Execute "DELETE * FROM TblCBSUser WHERE CBSUserNo = " & pCBSUserNo

End Sub

' ===============================================================
' Event pCBSUserForm_Update
' Event from client form to update CBS User
' ---------------------------------------------------------------
Private Sub pCBSUserForm_Update()
    With pCBSUserForm
        pUserName = CleanSQLText(.TxtUserName)
        pPosition = CleanSQLText(.TxtPosition)
        pPhoneNo = CleanSQLText(.TxtPhoneNo)
        pUserLvl = .CmoUserLvl
        
        If .CmoSupervisor.ListIndex <> -1 Then
            Set pSupervisor = New ClsCBSUser
            pSupervisor.DBGet .CmoSupervisor
        End If
        
    End With
    
    If pCBSUserNo = CurrentUser.CBSUserNo Then
        CurrentUser.UserLvl = pUserLvl
    End If
    
    DBSave
End Sub

' ===============================================================
' Event pCBSUserForm_CreateNew
' Event from user form to create new SPV
' ---------------------------------------------------------------
Private Sub pCBSUserForm_CreateNew()
    With pCBSUserForm
        .ClearForm
        .Hide
    End With
    DBNew
End Sub

' ===============================================================
' Event pCBSUserForm_Delete
' Deletes Contact
' ---------------------------------------------------------------
Private Sub pCBSUserForm_Delete()
    DBDelete
End Sub

' ===============================================================
Private Sub Class_Initialize()
    Set pCBSUserForm = New FrmAdnCBSUser
    
End Sub

' ===============================================================
Private Sub Class_Terminate()
    Set pCBSUserForm = Nothing
    If Not pSupervisor Is Nothing Then Set pSupervisor = Nothing
End Sub
' ===============================================================
