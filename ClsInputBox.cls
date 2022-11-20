VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsInputBox
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 05 Nov 22
'===============================================================
' Methods
'---------------------------------------------------------------
'===============================================================
Option Explicit

Private WithEvents pInputForm As FrmInputBox
Attribute pInputForm.VB_VarHelpID = -1
Private pTitle As String
Private pInstructions As String
Private pReturnValue As String

'---------------------------------------------------------------
Public Property Let Title(ByVal Title As String)
    pInputForm.TtlMessage = Title
End Property

'---------------------------------------------------------------
Public Property Let Instructions(ByVal Instructions As String)
    pInputForm.LblInstructions = Instructions
End Property

'---------------------------------------------------------------
Public Property Get ReturnValue() As String
    ReturnValue = pInputForm.xTxtInput
End Property

'---------------------------------------------------------------
'Public Property Get InputForm() As String
'    InputForm = pInputForm
'End Property
'
'Public Property Let InputForm(ByVal vNewValue As String)
'    pInputForm = vNewValue
'End Property

' ===============================================================
' Method Show
' Shows form
'---------------------------------------------------------------
Public Sub Show()
    pInputForm.Show
End Sub

' ===============================================================
' Method Hide
' Hides form
'---------------------------------------------------------------
Public Sub Hide()
    pInputForm.Hide
End Sub

' ===============================================================
' Method ClearForm
' Clears form
'---------------------------------------------------------------
Public Sub ClearForm()
    pInputForm.xTxtInput = ""
End Sub

' ===============================================================
' Event pInputForm_EnterValue
' Event input being entered
'---------------------------------------------------------------
Private Sub pInputForm_EnterValue()
    pReturnValue = pInputForm.xTxtInput
    Unload pInputForm
End Sub

' ===============================================================
Private Sub Class_Initialize()
    Set pInputForm = New FrmInputBox
End Sub
' ===============================================================
Private Sub Class_Terminate()
    Set pInputForm = Nothing
End Sub
' ===============================================================




