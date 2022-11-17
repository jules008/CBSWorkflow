VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim ButArray() As New Class2
        
Private Sub UserForm_Initialize()
    Dim ctlbut As MSForms.CommandButton
    
    Dim butTop As Long, i As Long

    '~~> Decide on the .Top for the 1st TextBox
    butTop = 30

    For i = 1 To 10
        Set ctlbut = Me.Controls.Add("Forms.CommandButton.1", "butTest" & i)

        '~~> Define the TextBox .Top and the .Left property here
        ctlbut.Top = butTop: ctlbut.Left = 50
        ctlbut.Caption = Cells(i, 7).Value
        '~~> Increment the .Top for the next TextBox
        butTop = butTop + 20

        ReDim Preserve ButArray(1 To i)
        Set ButArray(i).butEvents = ctlbut
    Next
End Sub

