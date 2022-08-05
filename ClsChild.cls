VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUIFrames
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 09 Feb 17
'===============================================================

Option Explicit
Private pUIFrames As Collection
Private pParent As ClsUIScreen

'---------------------------------------------------------------
Public Property Get Parent() As ClsUIScreen
    Set Parent = pParent
End Property

Public Property Set Parent(obj As ClsUIScreen)
    Set pParent = obj
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pUIFrames.[_NewEnum]
End Function

' ===============================================================
Public Sub Initialize(obj As ClsUIScreen)
    Set Me.Parent = obj
    Set pUIFrames = New Collection

End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
End Sub

'---------------------------------------------------------------
