VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmSelectWorkFlow 
   Caption         =   "Select Workflow"
   ClientHeight    =   3345
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7065
   OleObjectBlob   =   "FrmSelectWorkFlow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmSelectWorkFlow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'===============================================================
' Module FrmSelectWorkFlow
' Select new workflow form
'---------------------------------------------------------------
' Created by Julian Turner
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 03 Feb 21
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmSelectWorkFlow"
Private RstWorkflow As Recordset

' ===============================================================
' ShowForm
' Shows form
' ---------------------------------------------------------------
Public Function ShowForm() As String
    Dim SelWorkflow As String
    
    Const StrPROCEDURE As String = "ShowForm()"

    On Error GoTo ErrorHandler
    
Restart:
    
    Show
      
    Set RstWorkflow = ModDatabase.SQLQuery("SELECT * FROM TblWorkflowName WHERE Deleted IS NULL")
    
    If CmoWorkflow.ListIndex <> -1 Then
         With RstWorkflow
             .FindFirst "WFDispName = '" & CmoWorkflow.Value & "'"
             ShowForm = !WFName
        End With
    Else
        ShowForm = ""
    End If
    
GracefulExit:
   
   Set RstWorkflow = Nothing
   
Exit Function

ErrorExit:
    
   Set RstWorkflow = Nothing
   ShowForm = "Error"

Exit Function

ErrorHandler:
    
    If Err.Number >= 2000 And Err.Number <= 2500 Then
        If CustomErrorHandler(Err.Number) = SYSTEM_RESTART Then
            Resume Restart
        Else
            Resume GracefulExit
        End If
    End If
    
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
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
    Unload Me
End Sub

' ===============================================================
' BtnSelect_Click
' Select button
' ---------------------------------------------------------------
Private Sub BtnSelect_Click()
    Dim ErrNo As Integer
    Dim Table As String
    Dim i As Integer
    Dim WFNameNo As String
    Dim WFName As String
    Dim WFDispName As String
    Dim Description As String
    
    Const StrPROCEDURE As String = "BtnSelect_Click()"

    On Error GoTo ErrorHandler

Restart:
            
    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART
        
    If CmoWorkflow.ListIndex <> -1 Then Unload Me

    
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
' CmoWorkflow_Change
' ---------------------------------------------------------------
Private Sub CmoWorkflow_Change()
    If CmoWorkflow.ListIndex <> -1 Then
        If Not PopulateForm Then Err.Raise HANDLED_ERROR
    End If
End Sub

' ===============================================================
' UserForm_Initialize
' ---------------------------------------------------------------
Private Sub UserForm_Initialize()
    Dim Workflow As String
    
    Set RstWorkflow = ModDatabase.SQLQuery("SELECT * FROM TblWorkflowName WHERE Deleted IS NULL")
    
    With CmoWorkflow
        .Value = ""
        .Clear
        With RstWorkflow
            Do While Not .EOF
                CmoWorkflow.AddItem !WFDispName
                .MoveNext
            Loop
        End With
    End With
    
    LblDesc = ""
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    

End Sub

' ===============================================================
' PopulateForm
' Fills form with data
' ---------------------------------------------------------------
Private Function PopulateForm() As Boolean
    Dim i As Integer
    Dim StepNo As String
    Dim LastStepNo As String
    Dim CntrlName As String
    
    Const StrPROCEDURE As String = "PopulateForm()"

    On Error GoTo ErrorHandler
    
    With RstWorkflow
        Debug.Print .RecordCount
        .MoveFirst
        .FindFirst "WFDispName = '" & CmoWorkflow & "'"
        LblDesc = !Description
    
    End With
    
    PopulateForm = True
  
Exit Function

ErrorExit:
    
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


