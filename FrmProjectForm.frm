VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmProjectForm 
   Caption         =   "CRM - Project"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   OleObjectBlob   =   "FrmProjectForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmProjectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===============================================================
' Module FrmProjectForm
' Admin form for Projects
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 19 Oct 22
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmProjectForm"

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
    
    If CurrentUser.UserLvl <> "Admin" Then Err.Raise ACCESS_DENIED
    
    Response = MsgBox("Are you sure you want to delete the Project from the database?", vbYesNo + vbExclamation, APP_NAME)
    
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
    TxtCBSComPC = 0
    TxtLoanTerm = 0
    TxtProjectNo = 0
    CmoCaseManager = ""
    CmoClientNo = ""
    CmoSPVNo = ""
    ChkExitFee = False
End Sub

' ===============================================================
' BtnNew_Click
' Creates new Contact
' ---------------------------------------------------------------
Private Sub BtnNew_Click()
    On Error GoTo ErrorHandler
    
    If CurrentUser.UserLvl <> "Admin" Or CurrentUser.UserLvl <> "Case Manager" Then Err.Raise ACCESS_DENIED
    
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
    If CurrentUser.UserLvl <> "Admin" Then Err.Raise ACCESS_DENIED
    
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
' CmoCaseManager_Change
' ---------------------------------------------------------------
Private Sub CmoCaseManager_Change()
    CmoCaseManager.BackColor = COL_WHITE
End Sub

' ===============================================================
' CmoClientNo_Change
' ---------------------------------------------------------------
Private Sub CmoClientNo_Change()
    CmoClientNo.BackColor = COL_WHITE
End Sub

' ===============================================================
' CmoClientNo_Change
' ---------------------------------------------------------------
Private Sub CmoSPVNo_Change()
    CmoSPVNo.BackColor = COL_WHITE
End Sub

' ===============================================================
' TxtCBSComPC_Change
' ---------------------------------------------------------------
Private Sub TxtCBSComPC_Change()
    TxtCBSComPC.BackColor = COL_WHITE
End Sub

' ===============================================================
' TxtLoanTerm_Change
' ---------------------------------------------------------------
Private Sub TxtLoanTerm_Change()
    TxtLoanTerm.BackColor = COL_WHITE
End Sub

' ===============================================================
' TxtProjName_Change
' ---------------------------------------------------------------
Private Sub TxtProjName_Change()
    TxtProjName.BackColor = COL_WHITE
End Sub

' ===============================================================
' UserForm_Initialize
' Initialises form controls
' ---------------------------------------------------------------
Private Sub UserForm_Initialize()
    Dim RstSource As Recordset
    Dim i As Integer
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    If Me.Tag = "New" Then
        BtnNew.Visible = False
        BtnUpdate.Caption = "Create"
    End If
    ClearForm
    
    Set RstSource = ModDatabase.SQLQuery("SELECT CBSUserNo, UserName FROM TblCBSUser")
    
    i = 0
    With CmoCaseManager
        .Clear
        Do While Not RstSource.EOF
            .AddItem
            If Not IsNull(RstSource!CBSUserNo) Then .List(i, 0) = RstSource!CBSUserNo
            If Not IsNull(RstSource!UserName) Then .List(i, 1) = RstSource!UserName
            RstSource.MoveNext
            i = i + 1
        Loop
    End With
    
    Set RstSource = ModDatabase.SQLQuery("SELECT SPVNo, Name FROM TblSPV")
    
    i = 0
    With CmoSPVNo
        .Clear
        Do While Not RstSource.EOF
            .AddItem
            If Not IsNull(RstSource!SPVNo) Then .List(i, 0) = RstSource!SPVNo
            If Not IsNull(RstSource!Name) Then .List(i, 1) = RstSource!Name
            RstSource.MoveNext
            i = i + 1
        Loop
    End With
    
    Set RstSource = ModDatabase.SQLQuery("SELECT ClientNo, Name FROM TblClient")
    
    i = 0
    With CmoClientNo
        .Clear
        Do While Not RstSource.EOF
            .AddItem
            If Not IsNull(RstSource!ClientNo) Then .List(i, 0) = RstSource!ClientNo
            If Not IsNull(RstSource!Name) Then .List(i, 1) = RstSource!Name
            RstSource.MoveNext
            i = i + 1
        Loop
    End With
End Sub

' ===============================================================
' TxtLoanTerm_KeyPress
' ---------------------------------------------------------------
Private Sub TxtLoanTerm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

' ===============================================================
' ValidateForm
' Ensures the form is filled out correctly before moving on
' ---------------------------------------------------------------
Private Function ValidateForm() As enFormValidation
    
    Const StrPROCEDURE As String = "ValidateForm()"

    On Error GoTo ErrorHandler
           
    With TxtProjName
        If CleanString(.Value) = "" Then
            .BackColor = COL_AMBER
            ValidateForm = enValidationError
        End If
    End With
    
    With TxtCBSComPC
        If .Value = "" Or Not IsNumeric(.Value) Then
            .BackColor = COL_AMBER
            ValidateForm = enValidationError
        End If
    End With
           
    With TxtLoanTerm
        If .Value = "" Or Not IsNumeric(.Value) Then
            .BackColor = COL_AMBER
            ValidateForm = enValidationError
        End If
    End With
           
    With CmoCaseManager
        If .ListIndex = -1 Then
            .BackColor = COL_AMBER
            ValidateForm = enValidationError
        End If
    End With
           
    With CmoClientNo
        If .ListIndex = -1 Then
            .BackColor = COL_AMBER
            ValidateForm = enValidationError
        End If
    End With
           
    With CmoSPVNo
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



