VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmAdnWFStep 
   Caption         =   "Workflow Administration"
   ClientHeight    =   9165.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11265
   OleObjectBlob   =   "FrmAdnWFStep.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "FrmAdnWFStep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===============================================================
' Module FrmAdnWFStep
' Admin form for workflow steps
'---------------------------------------------------------------
' Created by Julian Turner
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 15 Dec 20
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmAdnWFStep"

Private StepTmplt As ClsStep
Private Steps As ClsSteps

Public Event CreateNew()
Public Event Update()
Public Event Delete()

' ===============================================================
' Form_Intialise
' Form Inititialise
' ---------------------------------------------------------------
Private Function Form_Intialise() As Boolean
    Dim RstSource As Recordset
    Dim i As Integer
    
    Const StrPROCEDURE As String = "Form_Intialise()"

    On Error GoTo ErrorHandler

    With CmoStepType
        .Clear
        .Value = ""
        .AddItem enStepTypeDisp(enYesNo)
        .AddItem enStepTypeDisp(enStep)
        .AddItem enStepTypeDisp(enDataInput)
    End With
    
    i = 0
    Set RstSource = ModDatabase.SQLQuery("SELECT * FROM TblEmail ORDER BY EmailNo")
    With RstSource
        If .RecordCount > 0 Then
            With CmoEmail
                .Clear
                Do While Not RstSource.EOF
                    .AddItem
                    .List(i, 0) = RstSource!EmailNo
                    .List(i, 1) = RstSource!TemplateName
                    RstSource.MoveNext
                    i = i + 1
                Loop
            End With
            
            i = 0
            .MoveFirst
            With CmoAltEmail
                .Clear
                Do While Not RstSource.EOF
                    .AddItem
                    .List(i, 0) = RstSource!EmailNo
                    .List(i, 1) = RstSource!TemplateName
                    RstSource.MoveNext
                    i = i + 1
                Loop
            End With
           
        End If
    End With
    
    i = 0
    Set RstSource = ModDatabase.SQLQuery("TblDataFormats")
    With RstSource
        If .RecordCount > 0 Then
            With CmoDataFormat
                .Clear
                Do While Not RstSource.EOF
                    .AddItem
                    .List(i, 0) = RstSource!FormCode
                    .List(i, 1) = RstSource!Format
                    RstSource.MoveNext
                    i = i + 1
                Loop
            End With
        End If
    End With
    
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

' ===============================================================
' BtnHelp_Click
' Shows help file
' ---------------------------------------------------------------
Private Sub BtnHelp_Click()
    FrmWflowHelp.Show
End Sub

' ===============================================================
' xBtnOpenEmail_Click
' Opens admin form for the selected email
' ---------------------------------------------------------------
Private Sub xBtnOpenEmail_Click()
    Dim ErrNo As Integer
    Dim Email As ClsEmail
    Dim EmailNo As String
    
    Const StrPROCEDURE As String = "BtnOpenEmail_Click()"

    On Error GoTo ErrorHandler

Restart:

    Set Email = New ClsEmail
    
    With CmoEmail
        If .ListIndex = -1 Then Exit Sub
        EmailNo = .List(.ListIndex, 0)
    End With
    
    If EmailNo = "" Then Err.Raise HANDLED_ERROR, , "No Email found"
    
    With Email
        .DBGet CInt(EmailNo)
        .DisplayForm
    End With
    
GracefulExit:
    Set Email = Nothing

Exit Sub

ErrorExit:
    Set Email = Nothing
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
' xBtnOpenAltEmail_Click
' Opens admin form for the selected AltEmail
' ---------------------------------------------------------------
Private Sub xBtnOpenAltEmail_Click()
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "BtnOpenAltEmail_Click()"

    On Error GoTo ErrorHandler

Restart:
    
    Dim AltEmailNo As String
    
    With CmoAltEmail
        If .ListIndex = -1 Then Exit Sub
        AltEmailNo = .List(.ListIndex, 0)
    End With
    
    If AltEmailNo = "" Then Err.Raise HANDLED_ERROR, , "No Email found"
    
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
' CmoStepType_Change
' adapts the form for each steptype
' ---------------------------------------------------------------
Private Sub CmoStepType_Change()
    Dim StepType As enStepType
    
    If CmoStepType.ListIndex = -1 Then Exit Sub
    
    StepType = CmoStepType.ListIndex
    Select Case StepType
        Case enAltBranch
            CmoDataFormat.Enabled = False
            CmoDataFormat.BackStyle = fmBackStyleTransparent
            TxtDataDest.Enabled = False
            TxtDataDest.BackStyle = fmBackStyleTransparent
            CmoAltEmail.Visible = True
            LblAltEmail.Visible = True
            xBtnOpenAltEmail.Visible = True
            ChkWait.Enabled = False
            ChkWait.BackStyle = fmBackStyleTransparent
            
        Case enDataInput
            CmoDataFormat.Enabled = True
            CmoDataFormat.BackStyle = fmBackStyleOpaque
            TxtDataDest.Enabled = True
            TxtDataDest.BackStyle = fmBackStyleOpaque
            CmoAltEmail.Visible = False
            LblAltEmail.Visible = False
            xBtnOpenAltEmail.Visible = False
            ChkWait.Enabled = False
            ChkWait.BackStyle = fmBackStyleTransparent
            
        Case enStep
            CmoDataFormat.Enabled = False
            CmoDataFormat.BackStyle = fmBackStyleTransparent
            TxtDataDest.Enabled = False
            TxtDataDest.BackStyle = fmBackStyleTransparent
            CmoAltEmail.Visible = False
            LblAltEmail.Visible = False
            xBtnOpenAltEmail.Visible = False
            ChkWait.Enabled = True
            ChkWait.BackStyle = fmBackStyleTransparent
        
        Case enYesNo
            CmoDataFormat.Enabled = False
            CmoDataFormat.BackStyle = fmBackStyleTransparent
            TxtDataDest.Enabled = False
            TxtDataDest.BackStyle = fmBackStyleTransparent
            CmoAltEmail.Visible = True
            LblAltEmail.Visible = True
            xBtnOpenAltEmail.Visible = True
            ChkWait.Enabled = True
            ChkWait.BackStyle = fmBackStyleTransparent
        
    End Select
End Sub

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
           
    With TxtAmberThresh
        If .Value = "" Then
            .BackColor = COL_AMBER
            ValidateForm = enValidationError
        End If
        
        If Not IsNumeric(.Value) Then
            .BackColor = COL_AMBER
            ValidateForm = enValidationError
        End If
    End With
    
    With TxtDataDest
        If .Value = "" And CmoDataFormat.ListIndex <> -1 Then
            .BackColor = COL_AMBER
            ValidateForm = enValidationError
        End If
    End With
    
    With TxtNextStep
        If .Value = "" Then
            .BackColor = COL_AMBER
            ValidateForm = enValidationError
        End If
    End With
    
    With TxtRedThresh
        If .Value = "" Then
            .BackColor = COL_AMBER
            ValidateForm = enValidationError
        End If
        
        If Not IsNumeric(.Value) Then
            .BackColor = COL_AMBER
            ValidateForm = enValidationError
        End If
        
    End With
    
    With xTxtStepAction
        If .Value = "" Then
            .BackColor = COL_AMBER
            ValidateForm = enValidationError
        End If
    End With
    
    With TxtStepName
        If .Value = "" Then
            .BackColor = COL_AMBER
            ValidateForm = enValidationError
        End If
    End With
    
    With TxtStepNo
        If .Value = "" Then
            .BackColor = COL_AMBER
            ValidateForm = enValidationError
        End If
    End With
      
    With CmoDataFormat
        If .ListIndex = -1 And TxtDataDest <> "" Then
            .BackColor = COL_AMBER
            ValidateForm = enValidationError
        End If
    End With
      
    With CmoStepType
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
' ClearForm
' Clears form
' ---------------------------------------------------------------
Public Sub ClearForm()
    TxtAltStep = ""
    TxtAmberThresh = ""
    TxtDataDest = ""
    TxtNextStep = ""
    TxtRedThresh = ""
    xTxtStepAction = ""
    TxtStepName = ""
    TxtStepNo = ""
    CmoStepType = ""
    ChkWait = False
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


