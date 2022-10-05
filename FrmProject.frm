VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmProject 
   Caption         =   "New Project Workflow"
   ClientHeight    =   9885.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15960
   OleObjectBlob   =   "FrmProject.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "FrmProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===============================================================
' Module FrmProject
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

Private Const StrMODULE As String = "FrmProject"
 
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hwnd As LongPtr) As Long
 
Private FormClosing As Boolean

' ===============================================================
' ShowForm
'Shows form
' ---------------------------------------------------------------
Public Function ShowForm() As Boolean

    Const StrPROCEDURE As String = "ShowForm()"

    On Error GoTo ErrorHandler
    
Restart:
    
    If ActiveProject Is Nothing Then Err.Raise HANDLED_ERROR, Description:="No Active Project"
    
    With ActiveProject.ProjectWorkflow
        .ActiveStep.Start
        .DBSave
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
    Dim ProgPC As Single
    Dim Step As ClsStep
    Dim TmpWorkflow As ClsWorkflow
    
    Const StrPROCEDURE As String = "PopulateForm()"

    On Error GoTo ErrorHandler
    
    ProgPC = ActiveProject.ProjectWorkflow.Steps.PCClosed
    Progress ProgPC
    
    With ActiveProject
        TxtProjectNo = .ProjectNo
        TxtLoanTerm = .LoanTerm
        TxtCommision = .CBSComPC
        ChkExitFee = .ExitFee
    End With
        
    With ActiveProject.ProjectWorkflow.ActiveStep
        TxtStepName = .StepNo & " - " & .StepName
        TxtAction = .StepAction
    End With

    Select Case ActiveProject.ProjectWorkflow.ActiveStep.StepType
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
                .Caption = "Step Closed"
            End With
            
        Case enDataInput
            With TxtDataInput
                .Visible = True
                .Value = ""
            End With
            BtnNo.Visible = False
            
            With BtnComplete
                .Visible = True
                .Caption = "Step Closed"
            End With
       
            With FrmCalPicker
                If ActiveProject.ProjectWorkflow.ActiveStep.DataFormat = "Date" And TxtDataInput = "" Then
                    Set TmpWorkflow = ActiveProject.ProjectWorkflow
                    TxtDataInput = Format(.ShowForm, "dd mmm yy")
                    Set ActiveProject.ProjectWorkflow = TmpWorkflow
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
    
    If ActiveProject.ProjectWorkflow.ActiveStep.CopyTextName <> "" Then
        With BtnCopyText
            .Visible = True
            .Caption = ActiveProject.ProjectWorkflow.ActiveStep.CopyTextName
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
        ActiveProject.ProjectWorkflow.ActiveStep.DataItem = TxtDataInput
        TxtDataInput = ""
    End If
    
    With ActiveProject.ProjectWorkflow
        .MoveToNextStep
    
    With ActiveProject.ProjectWorkflow.ActiveStep
       
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
' BtnNo_Click
' ---------------------------------------------------------------
Private Sub BtnNo_Click()
    
    With ActiveProject.ProjectWorkflow
        .MoveToAltStep
    
    With ActiveProject.ProjectWorkflow.ActiveStep
       
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
' BtnPrevStep_Click
' ---------------------------------------------------------------
Private Sub BtnPrevStep_Click()
    
    With ActiveProject.ProjectWorkflow
        .MoveToPrevStep
        .ActiveStep.Start
        .DBSave
    End With
    
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
        
    ActiveProject.ProjectWorkflow.DBSave

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
    Dim hwnd As LongPtr
    hwnd = FindWindow("ThunderDFrame", Me.Caption)
    BringWindowToTop (hwnd)
End Sub

Private Sub UserForm_Initialize()
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
End Sub


