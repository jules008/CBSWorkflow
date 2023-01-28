VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmAccessCntrl 
   Caption         =   "User Access Control Form"
   ClientHeight    =   7005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8490.001
   OleObjectBlob   =   "FrmAccessCntrl.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmAccessCntrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' Module FrmAccessCntrl
' controls access to lenders and clients
'---------------------------------------------------------------
' Created by Julian Turner
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 23 Jan 23
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmAccessCntrl"

' ===============================================================
' Form_Intialise
' Form Inititialise
' ---------------------------------------------------------------
Private Function Form_Intialise() As Boolean
    Dim RstSource As Recordset
    
    Const StrPROCEDURE As String = "Form_Intialise()"
    On Error GoTo ErrorHandler
    
    HdgClients.AddItem "Clients"
    HdgLenders.AddItem "Lenders"
    
    Form_Intialise = True
    
    Set RstSource = Nothing
    
Exit Function

ErrorExit:

    Set RstSource = Nothing
    
    Form_Intialise = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

'===============================================================
' BtnClose_Click
'---------------------------------------------------------------
Private Sub BtnClose_Click()
    Hide
End Sub

' ===============================================================
' ValidateForm
' Ensures the form is filled out correctly before moving on
' ---------------------------------------------------------------
Private Function ValidateForm() As enFormValidation
    
    Const StrPROCEDURE As String = "ValidateForm()"

    On Error GoTo ErrorHandler
           
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
' ClearForm
' Clears form
' ---------------------------------------------------------------
Public Sub ClearForm()
End Sub

' ===============================================================
' UserForm_Initialize
' ---------------------------------------------------------------
Private Sub UserForm_Initialize()
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "UserForm_Initialize()"

    On Error GoTo ErrorHandler

Restart:

    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART

    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    If Not Form_Intialise Then Err.Raise HANDLED_ERROR

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
' xBtnAddClient_Click
' Adds client to the access list
' ---------------------------------------------------------------
Private Sub xBtnAddClient_Click()
    Dim Picker As ClsFrmPicker
    
    Set Picker = New ClsFrmPicker
    
    With Picker
    
    
    
    Set Picker = Nothing
End Sub
