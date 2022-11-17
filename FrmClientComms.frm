VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmClientComms 
   Caption         =   "Communication"
   ClientHeight    =   8115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9735.001
   OleObjectBlob   =   "FrmClientComms.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmClientComms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' Module FrmClientComms
' Chat room form
'---------------------------------------------------------------
' Created by Julian Turner
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 17 Nov 22
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmClientComms"

Dim ClsChkBoxes() As New ClsChkBox

' ---------------------------------------------------------------
' AddFields
' routine for adding new field
' ---------------------------------------------------------------
Public Sub AddFields(CommsList As Recordset)
    Dim i As Integer
    Dim Chkbox As MSForms.CheckBox
    Dim IntExt As String
    Dim OrgStr As String
    
    Debug.Assert CommsList.RecordCount > 0
    
    i = 1
    With CommsList
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            Debug.Print !ContactName
            If !ContactType = "Client" Then IntExt = "Internal" Else IntExt = "External"
            If Not IsNull(!Organisation) Then OrgStr = " at " & !Organisation Else OrgStr = ""
            
            Set Chkbox = FrmBoxes.Controls.Add("Forms.CheckBox.1")
            
            With Chkbox
                .Name = "ChkBox" & i
                .Top = (i * 15) + 10
                .Caption = "Send " & IntExt & " Communication to " & CommsList!ContactName & OrgStr
                .Left = 25
                .Width = 400
                .Height = 15
                .Font.Size = 8
                .Font.Name = "Tahoma"
                .BackStyle = fmBackStyleTransparent
                .SpecialEffect = fmSpecialEffectFlat
                .Visible = True
            End With
            ReDim Preserve ClsChkBoxes(1 To i)
            Set ClsChkBoxes(i).Chkbox = Chkbox
                    
            i = i + 1
            DoEvents
            .MoveNext
        Loop
    End With
    
    If i < 5 Then i = 5
    
    FrmBoxes.Height = (i * 15) + 20
    Me.Height = (i * 15) + 120
End Sub

' ---------------------------------------------------------------
' BtnClose_Click
' ---------------------------------------------------------------
Private Sub BtnClose_Click()
    Unload Me
End Sub

' ---------------------------------------------------------------
' ChkSelectAll_Click
' Selects all check boxes
' ---------------------------------------------------------------
Private Sub ChkSelectAll_Click()
    Dim Cntrl As MSForms.CheckBox
    
    For Each Cntrl In FrmBoxes.Controls
        Cntrl.Value = ChkSelectAll.Value
    Next
End Sub

' ---------------------------------------------------------------
' UserForm_Terminate
' ---------------------------------------------------------------
Private Sub UserForm_Terminate()
    Dim Chkbox As Variant
    
    For Each Chkbox In ClsChkBoxes
        Set Chkbox = Nothing
    Next
End Sub
