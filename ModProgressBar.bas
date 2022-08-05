Attribute VB_Name = "ModProgressBar"
'===============================================================
' Module ModProgressBar
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 23 Jul 22
'===============================================================
Private Const StrMODULE As String = "ModProgressBar"

Option Explicit

Option Private Module

Public Const GWL_STYLE = -16
Public Const WS_CAPTION = &HC00000
 #If VBA7 Then
    Declare PtrSafe Function GetWindowLongPtr _
                            Lib "user32" Alias "GetWindowLongPtrA" ( _
                            ByVal hwnd As LongPtr, _
                            ByVal nIndex As Long) As LongPtr
    Declare PtrSafe Function SetWindowLongPtr _
                            Lib "user32" Alias "SetWindowLongPtrA" ( _
                            ByVal hwnd As LongPtr, _
                            ByVal nIndex As Long, _
                            ByVal dwNewLong As LongPtr) As LongPtr
    Declare PtrSafe Function DrawMenuBar _
                            Lib "user32" ( _
                            ByVal hwnd As LongPtr) As Long
    Declare PtrSafe Function FindWindowA _
                           Lib "user32" (ByVal lpClassName As String, _
                            ByVal lpWindowName As String) As LongPtr

 #Else
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
 #End If
 
Sub HideTitleBar(Frm As Object)
    Dim lFrmHdl As Long
    Dim lngWindow As Long
    
    
    #If VBA7 Then
        lFrmHdl = FindWindowA(vbNullString, Frm.Caption)
        lngWindow = GetWindowLongPtr(lFrmHdl, GWL_STYLE)
        lngWindow = lngWindow And (Not WS_CAPTION)
        Call SetWindowLongPtr(lFrmHdl, GWL_STYLE, lngWindow)
        Call DrawMenuBar(lFrmHdl)
    #Else
    lFrmHdl = FindWindowA(vbNullString, Frm.Caption)
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    
    Call SetWindowLong(lFrmHdl, GWL_STYLE, lngWindow)
    Call DrawMenuBar(lFrmHdl)
    #End If
End Sub

