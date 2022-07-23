VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmProgressBar 
   Caption         =   "UserForm1"
   ClientHeight    =   1230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7065
   OleObjectBlob   =   "FrmProgressBar.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "FrmProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'===============================================================
' Module FrmProgressBar
' v0.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 07 Jan 21
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmProgressBar"

' ===============================================================
' UserForm_Initialize
' Form initialisation
' ---------------------------------------------------------------
Private Sub UserForm_Initialize()
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    HideTitleBar Me
End Sub

' ===============================================================
' Progress
' Updates progress bar
' ---------------------------------------------------------------
Sub Progress(MessTxt As String, pctCompl As Single)
    If Not Me.Visible Then Me.Show
    LblMessage = MessTxt
    LblText = Format(pctCompl, "0") & "%"
    LblProgress.Width = FrmProgBar.Width / 100 * pctCompl
    Repaint
    Application.Wait (Now + TimeValue("00:00:01"))
    DoEvents
End Sub

' ===============================================================
' Start
' Starts progress bar
' ---------------------------------------------------------------
Public Sub Start(ProcName As String)
    Application.Run ProcName
    Unload Me
End Sub
