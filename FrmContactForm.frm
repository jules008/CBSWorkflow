VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmContactForm 
   Caption         =   "Contact"
   ClientHeight    =   3930
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
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
    
    Response = MsgBox("Are you sure you want to delete the Contact from the database?", vbYesNo + vbExclamation, APP_NAME)
    
    If Response = 6 Then
    End If
End Sub

' ===============================================================
' ClearForm
' Clears form
' ---------------------------------------------------------------
Private Sub ClearForm()
    TxtAddress = ""
    TxtContactNo = ""
    TxtPhone1 = ""
    TxtPhone2 = ""
    TxtPosition = ""
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
            
            RaiseEvent CreateNew
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
' TxtAddress_Change
' ---------------------------------------------------------------
Private Sub TxtAddress_Change()
    TxtAddress.BackColor = COL_OFF_WHITE
End Sub

' ===============================================================
' TxtContactNo_Change
' ---------------------------------------------------------------
Private Sub TxtContactNo_Change()
    TxtContactNo.BackColor = COL_OFF_WHITE
End Sub

' ===============================================================
' TxtPhone1_Change
' ---------------------------------------------------------------
Private Sub TxtPhone1_Change()
    TxtPhone1.BackColor = COL_OFF_WHITE
End Sub

' ===============================================================
' TxtPhone2_Change
' ---------------------------------------------------------------
Private Sub TxtPhone2_Change()
    TxtPhone2.BackColor = COL_OFF_WHITE
End Sub

' ===============================================================
' TxtPosition_Change
' ---------------------------------------------------------------
Private Sub TxtPosition_Change()
    TxtPosition.BackColor = COL_OFF_WHITE
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
End Sub

' ===============================================================
' ValidateForm
' Ensures the form is filled out correctly before moving on
' ---------------------------------------------------------------
Private Function ValidateForm() As enFormValidation
    
    Const StrPROCEDURE As String = "ValidateForm()"

    On Error GoTo ErrorHandler
           
    With TxtAddress
        If .Value = "" Then
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

