VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmWorkflow 
   Caption         =   "New Project Workflow"
   ClientHeight    =   9885.001
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
 
Private FormClosing As Boolean

' ===============================================================
' ===============================================================
' BtnClose_Click
' ---------------------------------------------------------------
Private Sub BtnClose_Click()
                Unload Me
End Sub


Private Sub UserForm_Initialize()
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    TxtAction = "Has the valuation report been received?"
    TxtClientContactName = "Percy Brown"
    TxtClientContactPhone = "01673 234871"
    TxtClientManager = "Emma Flindell"
    TxtClientName = "Fresco Ltd"
    TxtLoanTerm = "48 Months"
    TxtProjectNo = "2"
    TxtSPVContactName = "Maria Cooper"
    TxtSPVContactPhone = "08712 2844322"
    TxtSPVName = "SPV 2"
    TxtStartDate = "14 Sep 22"
    TxtStepName = "Valuation Report Received"
    ChkExitFee = True
    TxtCommision = "15%"
End Sub
