VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmAdnCBSUser 
   ClientHeight    =   3765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11745
   OleObjectBlob   =   "FrmAdnCBSUser.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmAdnCBSUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===============================================================
' Module FrmCBSUser
' Admin form for members
'---------------------------------------------------------------
' Created by Julian Turner
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 10 Oct 22
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmCBSUser"
Public Event CreateNew()
Public Event Update()
Public Event Delete()
Public Event ShowAccessFrm()

' ===============================================================
' BtnAccessLvl_Click
' Mnages access levels for user
' ---------------------------------------------------------------
Private Sub BtnAccessLvl_Click()
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "BtnAccessLvl_Click()"

    On Error GoTo ErrorHandler

Restart:

    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART
    
    RaiseEvent ShowAccessFrm
    
GracefulExit:


Exit Sub

ErrorExit:

    '***CleanUpCode***

Exit Sub

ErrorHandler:
    If Err.Number >= 2000 And Err.Number <= 2500 Then
        ErrNo = Err.Number
        CustomErrorHandler (Err.Number)
        If ErrNo = SYSTEM_RESTART Then Resume Restart Else Resume GracefulExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

'===============================================================
' BtnClose_Click
'---------------------------------------------------------------
Private Sub BtnClose_Click()
    Me.Hide
End Sub

' ===============================================================
' BtnDelete_Click
' Marks member as deleted
' ---------------------------------------------------------------
Private Sub BtnDelete_Click()
    Dim Response As Integer
    
    Response = MsgBox("Are you sure you want to delete the CBS User from the database?", vbYesNo + vbExclamation, APP_NAME)
    
    If Response = 6 Then
        RaiseEvent Delete
        Me.Hide
    End If
End Sub

' ===============================================================
' BtnNew_Click
' Creates new Contact
' ---------------------------------------------------------------
Private Sub BtnNew_Click()
    RaiseEvent CreateNew
End Sub

' ===============================================================
' ClearForm
' Clears form
' ---------------------------------------------------------------
Public Sub ClearForm()
    TxtCBSUserNo = ""
    TxtPhoneNo = ""
    TxtPosition = ""
    TxtUserName = ""
    CmoUserLvl = ""
    CmoSupervisor = ""
End Sub

' ===============================================================
' BtnUpdate_Click
' update changes and close
' ---------------------------------------------------------------
Private Sub BtnUpdate_Click()
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "BtnUpdate_Click()"

    On Error GoTo ErrorHandler

Restart:

    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART
    
    Select Case ValidateForm

        Case Is = enFunctionalError
            Err.Raise HANDLED_ERROR
        
        Case Is = enValidationError
            
        Case Is = enFormOK
            
            RaiseEvent Update
            Me.Hide
    End Select
    
GracefulExit:

Exit Sub

ErrorExit:

Exit Sub

ErrorHandler:
    If Err.Number >= 2000 And Err.Number <= 2500 Then
        ErrNo = Err.Number
        CustomErrorHandler (Err.Number)
        If ErrNo = SYSTEM_RESTART Then Resume Restart Else Resume GracefulExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' TxtCBSUserNo_Change
' ---------------------------------------------------------------
Private Sub TxtCBSUserNo_Change()
    TxtCBSUserNo.BackColor = COL_WHITE
End Sub

' ===============================================================
' TxtPosition_Change
' ---------------------------------------------------------------
Private Sub TxtPosition_Change()
    TxtPosition.BackColor = COL_WHITE
End Sub

' ===============================================================
' TxtUserName_Change
' ---------------------------------------------------------------
Private Sub TxtUserName_Change()
    TxtUserName.BackColor = COL_WHITE
End Sub

' ===============================================================
' CmoUserLvl_Change
' ---------------------------------------------------------------
Private Sub CmoUserLvl_Change()
    If CmoUserLvl = 3 Then BtnAccessLvl.Visible = True Else BtnAccessLvl.Visible = False
    CmoUserLvl.BackColor = COL_WHITE
End Sub

' ===============================================================
' UserForm_Initialize
' Initialises form controls
' ---------------------------------------------------------------
Private Sub UserForm_Initialize()
    Dim UserLvl As EnUserLvl
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    If Me.Tag = "New" Then
        BtnNew.Visible = False
        BtnUpdate.Caption = "Create"
    End If
    
    With CmoUserLvl
        .Clear
        .AddItem
        .List(0, 0) = 1
        .List(0, 1) = "Admin"
        .AddItem
        .List(1, 0) = 2
        .List(1, 1) = "Senior Manager"
        .AddItem
        .List(2, 0) = 3
        .List(2, 1) = "Case Manager"
    End With
    
    Dim RstSupervisor As Recordset
    Dim x As Integer
    UserLvl = enSenMgr
    Set RstSupervisor = ModDatabase.SQLQuery("SELECT CBSUserNo, UserName FROM TblCBSUser WHERE UserLvl = " & UserLvl)
    
    With CmoSupervisor
        .Clear
        Do While Not RstSupervisor.EOF
            .AddItem
            If Not IsNull(RstSupervisor!CBSUserNo) Then .List(x, 0) = RstSupervisor!CBSUserNo
            If Not IsNull(RstSupervisor!UserName) Then .List(x, 1) = RstSupervisor!UserName
            RstSupervisor.MoveNext
            x = x + 1
        Loop
    End With
    ClearForm
    
    Set RstSupervisor = Nothing
End Sub

' ===============================================================
' ValidateForm
' Ensures the form is filled out correctly before moving on
' ---------------------------------------------------------------
Private Function ValidateForm() As enFormValidation
    
    Const StrPROCEDURE As String = "ValidateForm()"

    On Error GoTo ErrorHandler
           
    With TxtUserName
        If .Value = "" Then
            .BackColor = COL_AMBER
            ValidateForm = enValidationError
        End If
    End With
           
    With CmoUserLvl
      If .ListIndex = -1 Then
        .BackColor = COL_AMBER
        ValidateForm = enValidationError
      End If
    End With
                     
    If ValidateForm = enValidationError Then
        Err.Raise FORM_INPUT_EMPTY
    Else
        ValidateForm = enFormOK
    End If
    
Exit Function

enValidationError:
    
    ValidateForm = enValidationError
Exit Function

ErrorExit:

    ValidateForm = enFunctionalError
Exit Function

ErrorHandler:
    
    If Err.Number >= 2000 And Err.Number <= 2500 Then
        CustomErrorHandler Err.Number
        Resume enValidationError:
    End If

If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================

