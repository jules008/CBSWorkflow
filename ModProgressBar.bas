Attribute VB_Name = "ModProgressBar"
Option Explicit

Option Private Module

Public Const GWL_STYLE = -16
Public Const WS_CAPTION = &HC00000
    Public Declare Function GetWindowLong _
                           Lib "user32" Alias "GetWindowLongA" ( _
                           ByVal hwnd As Long, _
                           ByVal nIndex As Long) As Long
    Public Declare Function SetWindowLong _
                           Lib "user32" Alias "SetWindowLongA" ( _
                           ByVal hwnd As Long, _
                           ByVal nIndex As Long, _
                           ByVal dwNewLong As Long) As Long
    Public Declare Function DrawMenuBar _
                           Lib "user32" ( _
                           ByVal hwnd As Long) As Long
    Public Declare Function FindWindowA _
                           Lib "user32" (ByVal lpClassName As String, _
                           ByVal lpWindowName As String) As Long
Sub HideTitleBar(Frm As Object)
    Dim lFrmHdl As Long
    Dim lngWindow As Long
    
    lFrmHdl = FindWindowA(vbNullString, Frm.Caption)
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    
    Call SetWindowLong(lFrmHdl, GWL_STYLE, lngWindow)
    Call DrawMenuBar(lFrmHdl)
End Sub

