VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmProjectWF 
   ClientHeight    =   10815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15990
   OleObjectBlob   =   "FrmProjectWF.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "FrmProjectWF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===============================================================
' Module FrmProjectWF
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

Private Const StrMODULE As String = "FrmProjectWF"
 
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hwnd As LongPtr) As Long
 
Private FormClosing As Boolean
Public ScreenAdjusted As Boolean

Public Event StartChat()
Public Event StepComplete()
Public Event PrevStep()
Public Event ClickNo()
Public Event CloseForm()
Public Event UpdateLoan()

' ===============================================================
' xBtnUpdateLoan_Click
' ---------------------------------------------------------------
Private Sub xBtnUpdateLoan_Click()
    RaiseEvent UpdateLoan
End Sub

' ===============================================================
' BtnChat_Click
' ---------------------------------------------------------------
Private Sub xBtnChat_Click()
    RaiseEvent StartChat
End Sub

' ===============================================================
' BtnClose_Click
' ---------------------------------------------------------------
Private Sub BtnClose_Click()
    RaiseEvent CloseForm
End Sub

' ===============================================================
' BtnComplete_Click
' ---------------------------------------------------------------
Private Sub BtnComplete_Click()
    RaiseEvent StepComplete
End Sub

' ===============================================================
' BtnHelp_Click
' ---------------------------------------------------------------
Private Sub xBtnHelp_Click()
    With ActiveProject.ProjectWorkflow.ActiveStep
        .DisplayHelpForm
    End With
End Sub

' ===============================================================
' BtnNo_Click
' ---------------------------------------------------------------
Private Sub BtnNo_Click()
    RaiseEvent ClickNo
End Sub

' ===============================================================
' BtnPrevStep_Click
' ---------------------------------------------------------------
Private Sub BtnPrevStep_Click()
    RaiseEvent PrevStep
End Sub

' ===============================================================
' progress
' Updates progress bar
' ---------------------------------------------------------------
Sub Progress(pctCompl As Single)

    LblText.Caption = Format(pctCompl, "0") & "%"
    xLblBar.Width = Frame7.Width / 100 * pctCompl
    
End Sub

Public Sub BringToFront()
    Dim hwnd As LongPtr
    hwnd = FindWindow("ThunderDFrame", Me.Caption)
    BringWindowToTop (hwnd)
End Sub


Private Sub UserForm_Initialize()
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    xLblBar.ZOrder (1)
    
End Sub


