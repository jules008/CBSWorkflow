VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmContactForm 
   Caption         =   "Contact"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12180
   OleObjectBlob   =   "FrmContactForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmContactForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===============================================================
' Module FrmContactForm
' Admin form for members
'---------------------------------------------------------------
' Created by Julian Turner
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 10 Oct 22
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmContactForm"
Public Event CreateNew()
Public Event Update()
Public Event Delete()

'===============================================================
' BtnClose_Click
'---------------------------------------------------------------
Private Sub BtnClose_Click()
    Unload Me
End Sub

' ===============================================================
' BtnDelete_Click
' Marks member as deleted
' ---------------------------------------------------------------
Private Sub BtnDelete_Click()
    Dim Response As Integer
    
    On Error GoTo ErrorHandler
    
    If CurrentUser.UserLvl <> enAdmin Then Err.Raise ACCESS_DENIED
    
    Response = MsgBox("Are you sure you want to delete the Contact from the database?", vbYesNo + vbExclamation, APP_NAME)
    
    If Response = 6 Then
        RaiseEvent Delete
        Unload Me
    End If
    
ErrorHandler:
    Dim ErrNo As Integer
    If Err.Number >= 2000 And Err.Number <= 2500 Then
        ErrNo = Err.Number
        CustomErrorHandler (Err.Number)
    End If
End Sub

' ===============================================================
' ClearForm
' Clears form
' ---------------------------------------------------------------
Public Sub ClearForm()
    TxtAddress1 = ""
    TxtAddress2 = ""
    TxtContactName = ""
    TxtEmailAdd = ""
    TxtContactNo = ""
    TxtPhone1 = ""
    TxtPhone2 = ""
    TxtPosition = ""
    TxtLastComm = ""
    CmoCommFreq = ""
    CmoContactType = ""
    TxtOrganisation = ""
    ChkOptOut = False
End Sub

' ===============================================================
' BtnNew_Click
' Creates new Contact
' ---------------------------------------------------------------
Private Sub BtnNew_Click()
    On Error GoTo ErrorHandler

    If CurrentUser.UserLvl <> enAdmin Or CurrentUser.UserLvl <> enCaseMgr Then Err.Raise ACCESS_DENIED

    RaiseEvent CreateNew
ErrorHandler:
    Dim ErrNo As Integer
    If Err.Number >= 2000 And Err.Number <= 2500 Then
        ErrNo = Err.Number
        CustomErrorHandler (Err.Number)
    End If
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
    If CurrentUser.UserLvl <> enAdmin Then Err.Raise ACCESS_DENIED
    
    Select Case ValidateForm

        Case Is = enFunctionalError
            Err.Raise HANDLED_ERROR
        
        Case Is = enValidationError
            
        Case Is = enFormOK
            
            RaiseEvent Update
            Unload Me
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
' xBtnSent_Click
' ---------------------------------------------------------------
Private Sub xBtnSent_Click()
    TxtLastComm = Format(Now, "dd mmm yy")
End Sub

' ===============================================================
' ChkOptOut_Click
' ---------------------------------------------------------------
Private Sub ChkOptOut_Click()
    If ChkOptOut Then
        TxtLastComm.Enabled = False
        CmoCommFreq.Enabled = False
    Else
        TxtLastComm.Enabled = True
        CmoCommFreq.Enabled = True
    End If
End Sub

' ===============================================================
' CmoCommFreq_Change
' ---------------------------------------------------------------
Private Sub CmoCommFreq_Change()
    CmoCommFreq.BackColor = COL_WHITE
End Sub

' ===============================================================
' TxtContactName_Change
' ---------------------------------------------------------------
Private Sub TxtContactName_Change()
    TxtContactName.BackColor = COL_WHITE
End Sub

' ===============================================================
' TxtContactNo_Change
' ---------------------------------------------------------------
Private Sub TxtContactNo_Change()
    TxtContactNo.BackColor = COL_WHITE
End Sub

' ===============================================================
' TxtEmailAdd_Change
' ---------------------------------------------------------------
Private Sub TxtEmailAdd_Change()
    TxtEmailAdd.BackColor = COL_WHITE
End Sub

' ===============================================================
' TxtLastComm_Change
' ---------------------------------------------------------------
Private Sub TxtLastComm_Change()
    TxtLastComm.BackColor = COL_WHITE
End Sub

' ===============================================================
' TxtPhone1_Change
' ---------------------------------------------------------------
Private Sub TxtPhone1_Change()
    TxtPhone1.BackColor = COL_WHITE
End Sub

' ===============================================================
' TxtPhone2_Change
' ---------------------------------------------------------------
Private Sub TxtPhone2_Change()
    TxtPhone2.BackColor = COL_WHITE
End Sub

' ===============================================================
' TxtPosition_Change
' ---------------------------------------------------------------
Private Sub TxtPosition_Change()
    TxtPosition.BackColor = COL_WHITE
End Sub

' ===============================================================
' UserForm_Initialize
' Initialises form controls
' ---------------------------------------------------------------
Private Sub UserForm_Initialize()
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    If Me.Tag = "New" Then
        BtnNew.Visible = False
        BtnUpdate.Caption = "Create"
    End If
    ClearForm
    
    With CmoContactType
        .Clear
        .AddItem
        .List(0, 0) = 0
        .List(0, 1) = "Lender"
        .AddItem
        .List(1, 0) = 1
        .List(1, 1) = "Project"
        .AddItem
        .List(2, 0) = 2
        .List(2, 1) = "SPV"
        .AddItem
        .List(3, 0) = 3
        .List(3, 1) = "Client"
        .AddItem
        .List(4, 0) = 4
        .List(4, 1) = "Lead"
        .AddItem
        .List(4, 0) = 5
        .List(4, 1) = "MS"
        .AddItem
        .List(4, 0) = 6
        .List(4, 1) = "Valuer"
    End With
    
    Dim i As Integer
    
    With CmoCommFreq
        .Clear
        For i = 0 To 31
            .AddItem i
            i = i + 1
        Next
    End With
    
End Sub

' ===============================================================
' ValidateForm
' Ensures the form is filled out correctly before moving on
' ---------------------------------------------------------------
Private Function ValidateForm() As enFormValidation
    
    Const StrPROCEDURE As String = "ValidateForm()"

    On Error GoTo ErrorHandler
           
    With TxtContactName
        If .Value = "" Then
            .BackColor = COL_AMBER
            ValidateForm = enValidationError
        End If
    End With
           
    With CmoContactType
        If .ListIndex = -1 Then
            .BackColor = COL_AMBER
            ValidateForm = enValidationError
        End If
    End With
    
    If TxtEmailAdd <> "" Then
        If Not ChkOptOut Then
            With CmoCommFreq
                If .ListIndex = -1 Then
                    .BackColor = COL_AMBER
                    ValidateForm = enValidationError
                End If
            End With
        End If
    End If
    
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

