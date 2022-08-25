VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmStartBanner 
   Caption         =   "UserForm1"
   ClientHeight    =   5520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8685.001
   OleObjectBlob   =   "FrmStartBanner.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "FrmStartBanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




'===============================================================
' Module FrmStartBanner
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 07 Jan 21
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmStartBanner"

' ===============================================================
' FormActivate
' Form activate
' ---------------------------------------------------------------
Private Sub FormActivate()
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "FormActivate()"

    On Error GoTo ErrorHandler

Restart:

    If Not ModStartUp.Initialise Then Err.Raise HANDLED_ERROR
    Unload Me

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
' UserForm_Activate
' Form activate
' ---------------------------------------------------------------
Private Sub UserForm_Activate()
    FormActivate
End Sub

' ===============================================================
' UserForm_Initialize
' Form initialisation
' ---------------------------------------------------------------
Private Sub UserForm_Initialize()
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    HideTitleBar Me
    LblVer = "System:  " & VERSION
    LblDBVer = "DB:  " & DB_VER
    LblDate = "Date:  " & VER_DATE
    LblCopyright = Chr(169) & " Copyright 2021"
End Sub

' ===============================================================
' Progress
' Updates progress bar
' ---------------------------------------------------------------
Sub Progress(MessTxt As String, pctCompl As Single)
    LblMessage = MessTxt
    LblText = Format(pctCompl, "0") & "%"
    LblProgress.Width = FrmProgBar.Width / 100 * pctCompl
    Repaint
    Sleep 500
    DoEvents
End Sub

