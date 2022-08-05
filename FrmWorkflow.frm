VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmWorkflow 
   Caption         =   "Workflow"
   ClientHeight    =   5535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15960
   OleObjectBlob   =   "FrmWorkflow.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "FrmWorkflow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===============================================================
' Module FrmWorkflow
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 25 Jun 20
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmWorkflow"
 
 #If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hwnd As LongPtr) As Long
 #Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
 #End If
 
Private FormClosing As Boolean

' ===============================================================
' ShowForm
'Shows form
' ---------------------------------------------------------------
Public Function ShowForm() As Boolean

    Const StrPROCEDURE As String = "ShowForm()"

    On Error GoTo ErrorHandler
    
Restart:
    
    If ActiveWorkFlow.ActiveStep Is Nothing Then Err.Raise HANDLED_ERROR, Description:="No activeworkflow"
    
    With ActiveWorkFlow
        .ActiveStep.Start
        .DBSave
'        .Steps.OpenNewEmails 'disabled as emails move from inbox
    End With
    
    If Not PopulateForm Then Err.Raise HANDLED_ERROR
    Me.Show

GracefulExit:
    
    ShowForm = True
Exit Function

ErrorExit:
    
    ShowForm = False
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
' PopulateForm
' Fills form with data
' ---------------------------------------------------------------
Private Function PopulateForm() As Boolean
    Dim AmberTime As Date
    Dim RedTime As Date
    Dim TimeToAmber As Integer
    Dim TimeToRed As Integer
    Dim ProgPC As Single
    Dim TimeNow As Date
    Dim Step As ClsStep
    Dim TmpWorkflow As ClsWorkflow
    Dim i As Integer
    
    Const StrPROCEDURE As String = "PopulateForm()"

    On Error GoTo ErrorHandler
    
    ProgPC = ActiveWorkFlow.Steps.PCComplete
    Progress ProgPC
    
    With ActiveWorkFlow
        TxtWorkflowNo = .WorkflowNo
        TxtWFName = .Name
    End With
    
    With ActiveWorkFlow.ActiveStep
        TxtStepName = .StepNo & " - " & .StepName
        TxtAction = .StepAction
    End With

    Select Case ActiveWorkFlow.ActiveStep.StepType
        Case enYesNo
            TxtDataInput.Visible = False
            BtnNo.Visible = True
            
            With BtnComplete
                .Visible = True
                .Caption = "Yes"
            End With
            
        Case enStep
            TxtDataInput.Visible = False
            BtnNo.Visible = False
            
            With BtnComplete
                .Visible = True
                .Caption = "Step Complete"
            End With
            
        Case enDataInput
            With TxtDataInput
                .Visible = True
                .Value = ""
            End With
            BtnNo.Visible = False
            
            With BtnComplete
                .Visible = True
                .Caption = "Step Complete"
            End With
       
            With FrmCalPicker
                If ActiveWorkFlow.ActiveStep.DataFormat = "Date" And TxtDataInput = "" Then
                    Set TmpWorkflow = ActiveWorkFlow
                    TxtDataInput = Format(.ShowForm, "dd mmm yy")
                    Set ActiveWorkFlow = TmpWorkflow
                End If
            End With
            
'            With FrmTimePicker
'                If ActiveWorkFlow.ActiveStep.DataFormat = "Time" And TxtDataInput = "" Then
'                    Set TmpWorkflow = ActiveWorkFlow
'                    .Show
'                    TxtDataInput = Format(.ReturnValue, "hh:mm")
'                    Set ActiveWorkFlow = TmpWorkflow
'                End If
'            End With
                
        Case enAltBranch
            TxtDataInput.Visible = False
            BtnCopyText.Visible = False
            BtnNo.Visible = True
            
            With BtnComplete
                .Visible = True
                .Caption = "Yes"
            End With
            
    End Select
    
    If ActiveWorkFlow.ActiveStep.CopyTextName <> "" Then
        With BtnCopyText
            .Visible = True
            .Caption = ActiveWorkFlow.ActiveStep.CopyTextName
        End With
    Else
        BtnCopyText.Visible = False
    End If
            
    Set TmpWorkflow = Nothing
    Set Step = Nothing
    
    PopulateForm = True

Exit Function

ErrorExit:
    Set TmpWorkflow = Nothing
    Set Step = Nothing
    
    PopulateForm = False

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
' BtnClose_Click
' ---------------------------------------------------------------
Private Sub BtnClose_Click()
    SaveWorkflow
    Me.Hide
End Sub

' ===============================================================
' BtnComplete_Click
' ---------------------------------------------------------------
Private Sub BtnComplete_Click()
    
    If TxtDataInput <> "" Then
        ActiveWorkFlow.ActiveStep.DataItem = TxtDataInput
        TxtDataInput = ""
    End If
    
    With ActiveWorkFlow
        .MoveToNextStep
    
    With ActiveWorkFlow.ActiveStep
       
            If .LastStep Then
                Unload Me
            Else
                If .Wait = True Then
            SaveWorkflow
            Me.Hide
        Else
            PopulateForm
            If Not Me.Visible Then Me.Show
        End If
            End If
        End With
    End With
    
End Sub

' ===============================================================
' BtnCopyText_Click
'
' ---------------------------------------------------------------
Private Sub BtnCopyText_Click()
    Dim ErrNo As Integer
    Dim StrText As String

    Const StrPROCEDURE As String = "BtnCopyText_Click()"

    On Error GoTo ErrorHandler

Restart:

'    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART
    
'    StrText = ModWorkflow.ReplaceKeyWords(ActiveWorkFlow.ActiveStep.CopyText, ActiveWorkFlow)
'    ModLibrary.CopyTextToClipboard StrText

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
' BtnDelete_Click
' ---------------------------------------------------------------
Private Sub BtnDelete_Click()
    Dim Response As Integer
    
    With ActiveWorkFlow
        Response = MsgBox("The following workflow will be marked as deleted " _
                        & vbCr _
                        & vbCr _
                        & vbCr _
                        & vbCr & "Are you sure you wish to continue?", vbYesNo + vbExclamation, APP_NAME)
                        
        If Response = 6 Then
            .DBDelete
            .Deleted = Now
        End If
    End With
    
    Me.Hide
End Sub

' ===============================================================
' BtnNo_Click
' ---------------------------------------------------------------
Private Sub BtnNo_Click()
    ActiveWorkFlow.MoveToAltStep
    SaveWorkflow
    Me.Hide
End Sub

' ===============================================================
' BtnPause_Click
' ---------------------------------------------------------------
Private Sub BtnPause_Click()
    ActiveWorkFlow.Pause
End Sub

' ===============================================================
' BtnPrevStep_Click
' ---------------------------------------------------------------
Private Sub BtnPrevStep_Click()
    
    With ActiveWorkFlow
        .MoveToPrevStep
        .ActiveStep.Start
        .DBSave
    End With
    
    TxtDataInput = ""
    
    If Not PopulateForm Then Err.Raise HANDLED_ERROR

End Sub

' ===============================================================
' BtnReset_Click
' ---------------------------------------------------------------
Private Sub BtnReset_Click()
    ActiveWorkFlow.Reset
    
    TxtDataInput = ""
    
    If Not PopulateForm Then Err.Raise HANDLED_ERROR
    
End Sub

' ===============================================================
' SaveWorkflow
'
' ---------------------------------------------------------------
Private Function SaveWorkflow() As Boolean
    Const StrPROCEDURE As String = "SaveWorkflow()"

    On Error GoTo ErrorHandler

    FormClosing = True
        
    ActiveWorkFlow.DBSave

    SaveWorkflow = True


Exit Function

ErrorExit:

    '***CleanUpCode***
    SaveWorkflow = False

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
' progress
' Updates progress bar
' ---------------------------------------------------------------
Sub Progress(pctCompl As Single)

    LblText.Caption = Format(pctCompl, "0") & "%"
    LblBar.Width = Frame7.Width / 100 * pctCompl
    
End Sub

' ===============================================================
' Refresh
' Refreshes form with existing data
' ---------------------------------------------------------------
Function Refresh() As Boolean
    Const StrPROCEDURE As String = "Refresh()"

    On Error GoTo ErrorHandler

    If Not PopulateForm Then Err.Raise HANDLED_ERROR
    
    Refresh = True


Exit Function

ErrorExit:

    '***CleanUpCode***
    Refresh = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

Public Sub BringToFront()
    Dim hwnd As Long
    hwnd = FindWindow("ThunderDFrame", Me.Caption)
    BringWindowToTop (hwnd)
End Sub

Private Sub UserForm_Initialize()
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
End Sub
