VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsCalButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsCalButton
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 30 Apr 20
'===============================================================

Option Explicit

Public WithEvents CmdBtnGroup As MSForms.CommandButton
Attribute CmdBtnGroup.VB_VarHelpID = -1

Sub CmdBtnGroup_Click()
    If G_FORM Then
        GoTo on_Form
    Else
        GoTo AddDate
    End If
AddDate:

    FrmCalPicker.ReturnDate = CDate(CmdBtnGroup.Tag)

    GoTo chg_month
on_Form:
    G_DATE = CmdBtnGroup.Tag
chg_month:
    With FrmCalPicker.CmoMonth
        .ListIndex = Month(CmdBtnGroup.Tag) - 1
    End With
    Unload FrmCalPicker
End Sub
