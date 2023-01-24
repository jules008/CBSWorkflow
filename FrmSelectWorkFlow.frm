VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmSelectWorkFlow 
   Caption         =   "Select Workflow"
   ClientHeight    =   3795
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6450
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
Public Event Update()
Public Event CloseFrm()

'===============================================================
' BtnClose_Click
'---------------------------------------------------------------
Private Sub BtnClose_Click()
    RaiseEvent CloseFrm
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
        
    If CmoSecondTier.ListIndex <> -1 Then

        RaiseEvent Update
        Unload Me
    End If
    
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
' CmoLoanType_Change
' ---------------------------------------------------------------
Private Sub CmoLoanType_Change()
    Dim Workflow As String
    
    With CmoSecondTier
        If CmoLoanType.ListIndex <> -1 Then
            .Enabled = True
    
            Set RstWorkflow = ModDatabase.SQLQuery("SELECT SecondTier FROM TblWorkflowTable WHERE LoanType = '" & CmoLoanType & "'")
            
            With CmoSecondTier
                .Value = ""
                .Clear
                With RstWorkflow
                    Do While Not .EOF
                        CmoSecondTier.AddItem !SecondTier
                        .MoveNext
                    Loop
                End With
            End With
            
            Set RstWorkflow = Nothing
        Else
            .Enabled = False
            .Clear
    End If
    End With
End Sub

' ===============================================================
' UserForm_Initialize
' ---------------------------------------------------------------
Private Sub UserForm_Initialize()
    Dim Workflow As String
    
    Set RstWorkflow = ModDatabase.SQLQuery("SELECT DISTINCT LoanType FROM TblWorkflowTable")
    
    With CmoLoanType
        .Value = ""
        .Clear
        With RstWorkflow
            Do While Not .EOF
                CmoLoanType.AddItem !LoanType
                .MoveNext
            Loop
        End With
    End With
    
    CmoSecondTier.Enabled = False
    
    Set RstWorkflow = Nothing
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    

End Sub
