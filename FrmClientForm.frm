VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmClientForm 
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11820
   OleObjectBlob   =   "FrmClientForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmClientForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' Module FrmClientForm
' Admin form for members
'---------------------------------------------------------------
' Created by Julian Turner
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 10 Oct 22
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmClientForm"
Public Event CreateNew()
Public Event Update()
Public Event Delete()

'===============================================================
' BtnClose_Click
'---------------------------------------------------------------
Private Sub BtnClose_Click()
    Hide
End Sub

' ===============================================================
' BtnDelete_Click
' Marks member as deleted
' ---------------------------------------------------------------
Private Sub BtnDelete_Click()
    Dim Response As Integer
    
    Response = MsgBox("Are you sure you want to delete the Client from the database?", vbYesNo + vbExclamation, APP_NAME)
    
    If Response = 6 Then
        RaiseEvent Delete
    End If
    Hide
End Sub

' ===============================================================
' ClearForm
' Clears form
' ---------------------------------------------------------------
Public Sub ClearForm()
    TxtClientNo = ""
    TxtName = ""
    TxtPhoneNo = ""
    TxtUrl = ""
    OptDevelopment = False
    OptCommercial = False
    OptBridgeExit = False
    ChkSenior = False
    ChkMezzanine = False
    ChkEquity = False
    ChkVAT = False
    ChkSDLT = False
    Chk1stChargeCM = False
    Chk2ndCharge = False
    ChkFirstCharge = False
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
            Hide
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
' OptBridgeExit_Click
' ---------------------------------------------------------------
Private Sub OptBridgeExit_Click()
    With ChkSenior
        .Visible = False
        .Value = False
    End With
    
    With ChkMezzanine
        .Visible = False
        .Value = False
    End With
    
    With ChkEquity
        .Visible = False
        .Value = False
   End With
    
    With ChkVAT
        .Visible = False
        .Value = False
    End With
    
    With ChkSDLT
        .Visible = False
        .Value = False
    End With
    
    With ChkFirstCharge
        .Visible = False
        .Value = False
    End With
    
    With Chk2ndCharge
        .Visible = False
        .Value = False
    End With
    
    With Chk1stChargeCM
        .Visible = True
        .Top = 12
        .Left = 387
    End With
End Sub

' ===============================================================
' OptCommercial_Click
' ---------------------------------------------------------------
Private Sub OptCommercial_Click()
    With ChkSenior
        .Visible = False
        .Value = False
    End With
    
    With ChkMezzanine
        .Visible = False
        .Value = False
    End With
    
    With ChkEquity
        .Visible = False
        .Value = False
   End With
    
    With ChkVAT
        .Visible = False
        .Value = False
    End With
    
    With ChkSDLT
        .Visible = False
        .Value = False
    End With
    
    With ChkFirstCharge
        .Visible = True
        .Top = 12
        .Left = 387
    End With
    
    With Chk2ndCharge
        .Visible = True
        .Top = 32
        .Left = 387
    End With
    
    With Chk1stChargeCM
        .Visible = False
        .Value = False
    End With
End Sub

' ===============================================================
' OptDevelopment_Click
' ---------------------------------------------------------------
Private Sub OptDevelopment_Click()
    ChkSenior.Visible = True
    ChkMezzanine.Visible = True
    ChkEquity.Visible = True
    ChkVAT.Visible = True
    ChkSDLT.Visible = True
    
    With ChkFirstCharge
        .Visible = False
        .Value = False
    End With
    
    With Chk2ndCharge
        .Visible = False
        .Value = False
    End With
    
    With Chk1stChargeCM
        .Visible = False
        .Value = False
    End With
End Sub

' ===============================================================
' CmoContract_Change
' ---------------------------------------------------------------
Private Sub TxtClientNo_Change()
    TxtClientNo.BackColor = COL_WHITE
End Sub

' ===============================================================
' TxtName_Change
' ---------------------------------------------------------------
Private Sub TxtName_Change()
    TxtName.BackColor = COL_WHITE
End Sub

' ===============================================================
' TxtPhoneNo_Change
' ---------------------------------------------------------------
Private Sub TxtPhoneNo_Change()
    TxtPhoneNo.BackColor = COL_WHITE
End Sub

' ===============================================================
' TxtUrl_Change
' ---------------------------------------------------------------
Private Sub TxtUrl_Change()
    TxtUrl.BackColor = COL_WHITE
End Sub

Private Sub UserForm_Activate()
    If Me.Tag = "New" Then
        TtlTop.Caption = "Create New Client"
        BtnUpdate.Caption = "Create"
    ElseIf Me.Tag = "Update" Then
        TtlTop.Caption = "Update Client"
        BtnUpdate.Caption = "Update"
    End If
End Sub

' ===============================================================
' UserForm_Initialize
' Initialises form controls
' ---------------------------------------------------------------
Private Sub UserForm_Initialize()
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    ChkSenior.Visible = False
    ChkMezzanine.Visible = False
    ChkEquity.Visible = False
    ChkVAT.Visible = False
    ChkSDLT.Visible = False
    ChkFirstCharge.Visible = False
    Chk2ndCharge.Visible = False
    Chk1stChargeCM.Visible = False
    
    If Me.Tag = "New" Then
        TtlTop.Caption = "Create New Client"
        BtnUpdate.Caption = "Create"
    ElseIf Me.Tag = "Update" Then
        TtlTop.Caption = "Update Client"
        BtnUpdate.Caption = "Update"
    End If
    
    With CmoCBS
        .Clear
        .AddItem "CBS"
        .AddItem "HP"
    End With
    ClearForm
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
' GetClientNeed
' Returns the Client need in binary form
' ---------------------------------------------------------------
Public Function GetClientNeed() As Byte
    Const BIN_SENIOR As Byte = 1
    Const BIN_MEZZ  As Byte = 2
    Const BIN_EQUITY  As Byte = 4
    Const BIN_VAT  As Byte = 8
    Const BIN_SDLT As Byte = 16
    Const BIN_IST As Byte = 32
    Const BIN_2ND  As Byte = 64
    Const BIN1STCM  As Byte = 128
    
    Dim ClientNeeds As Byte
    
    If ChkSenior = True Then ClientNeeds = ClientNeeds + BIN_SENIOR
    If ChkMezzanine = True Then ClientNeeds = ClientNeeds + BIN_MEZZ
    If ChkEquity = True Then ClientNeeds = ClientNeeds + BIN_EQUITY
    If ChkVAT = True Then ClientNeeds = ClientNeeds + BIN_VAT
    If ChkSDLT = True Then ClientNeeds = ClientNeeds + BIN_SDLT
    If ChkFirstCharge = True Then ClientNeeds = ClientNeeds + BIN_IST
    If Chk2ndCharge = True Then ClientNeeds = ClientNeeds + BIN_2ND
    If Chk1stChargeCM = True Then ClientNeeds = ClientNeeds + BIN1STCM

    GetClientNeed = ClientNeeds
End Function

' ===============================================================
' SetClientNeed
' sets the Client need in binary form
' ---------------------------------------------------------------
Public Sub SetClientNeed(ClientNeeds As Byte)
    Const BIN_SENIOR  As Byte = 1
    Const BIN_MEZZ  As Byte = 2
    Const BIN_EQUITY  As Byte = 4
    Const BIN_VAT As Byte = 8
    Const BIN_SDLT As Byte = 16
    Const BIN_IST As Byte = 32
    Const BIN_2ND  As Byte = 64
    Const BIN1STCM  As Byte = 128
    
    If ClientNeeds > 0 Then
        Select Case ClientNeeds
            Case Is < 32
                OptDevelopment.Value = True
            Case Is < 128
                OptCommercial.Value = True
            Case Else
                OptBridgeExit.Value = True
        End Select
        
        If (ClientNeeds And BIN_SENIOR) <> 0 Then ChkSenior = True Else ChkSenior = False
        If (ClientNeeds And BIN_MEZZ) <> 0 Then ChkMezzanine = True Else ChkMezzanine = False
        If (ClientNeeds And BIN_EQUITY) <> 0 Then ChkEquity = True Else ChkEquity = False
        If (ClientNeeds And BIN_VAT) <> 0 Then ChkVAT = True Else ChkVAT = False
        If (ClientNeeds And BIN_SDLT) <> 0 Then ChkSDLT = True Else ChkSDLT = False
        If (ClientNeeds And BIN_IST) <> 0 Then ChkFirstCharge = True Else ChkFirstCharge = False
        If (ClientNeeds And BIN_2ND) <> 0 Then Chk2ndCharge = True Else Chk2ndCharge = False
        If (ClientNeeds And BIN1STCM) <> 0 Then Chk1stChargeCM = True Else Chk1stChargeCM = False
    End If
End Sub

