VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUIScreen
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 16 Apr 20
'===============================================================

Option Explicit
Private pUIFrames As ClsUIFrames
'---------------------------------------------------------------
Public Property Get Frames() As ClsUIFrames
    If pUIFrames Is Nothing Then
        Set pUIFrames = New ClsUIFrames
        pUIFrames.Initialize Me
    End If
    Set Frames = pUIFrames
End Property

Public Property Let Frames(ByVal vNewValue As ClsUIFrames)
    Set pUIFrames = vNewValue
End Property

' ===============================================================
Private Sub Class_Initialize()
End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    
End Sub

'---------------------------------------------------------------


