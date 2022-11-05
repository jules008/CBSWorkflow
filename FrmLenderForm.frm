VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmLenderForm 
   Caption         =   "Lender"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   OleObjectBlob   =   "FrmLenderForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmLenderForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===============================================================
' Module FrmLenderForm
' Admin form for Lenders
'---------------------------------------------------------------
' Created by Julian Turner
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 10 Oct 22
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmLenderForm"
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
    
    Response = MsgBox("Are you sure you want to delete the Lender from the database?", vbYesNo + vbExclamation, APP_NAME)
    
    If Response = 6 Then
        RaiseEvent Delete
        Unload Me
    End If
End Sub

' ===============================================================
' ClearForm
' Clears form
' ---------------------------------------------------------------
Public Sub ClearForm()
    TxtAddress = ""
    TxtLenderNo = ""
    TxtName = ""
    TxtPhoneNo = ""
    CmoLenderType = ""
    LstContacts = ""
End Sub

' ===============================================================
' BtnNew_Click
' Creates new Contact
' ---------------------------------------------------------------
Private Sub BtnNew_Click()
    RaiseEvent CreateNew
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
' CmoLenderType_Change
' ---------------------------------------------------------------
Private Sub CmoLenderType_Change()
    CmoLenderType.BackColor = COL_WHITE
End Sub

' ===============================================================
' TxtName_Change
' ---------------------------------------------------------------
Private Sub TxtName_Change()
    TxtName.BackColor = COL_WHITE
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
    
    Set RstSource = ModDatabase.SQLQuery("TblLenderType")
    
    i = 0
    With CmoLenderType
        .Clear
        Do While Not RstSource.EOF
            .AddItem
            If Not IsNull(RstSource!TypeNo) Then .List(i, 0) = RstSource!TypeNo
            If Not IsNull(RstSource!LenderType) Then .List(i, 1) = RstSource!LenderType
            RstSource.MoveNext
            i = i + 1
        Loop
    End With
    
    If Me.Tag = "New" Then
        BtnNew.Visible = False
        BtnUpdate.Caption = "Create"
    End If
    ClearForm
    Set RstSource = Nothing
End Sub

' ===============================================================
' ValidateForm
' Ensures the form is filled out correctly before moving on
' ---------------------------------------------------------------
Private Function ValidateForm() As enFormValidation
    
    Const StrPROCEDURE As String = "ValidateForm()"

    On Error GoTo ErrorHandler
           
    With TxtName
        If .Value = "" Then
            .BackColor = COL_AMBER
            ValidateForm = enValidationError
        End If
    End With
                     
    With CmoLenderType
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


