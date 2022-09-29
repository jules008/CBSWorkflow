VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUIStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsStyle
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 23 Sep 22
'===============================================================
' Methods
'---------------------------------------------------------------
'===============================================================
'---------------------------------------------------------------
Option Explicit

Private pForeColour As Long
Private pBorderColour As Long
Private pBorderWidth As Single
Private pFontStyle As String
Private pFontBold As Boolean
Private pFontSize As Integer
Private pFontColour As Long
Private pFontXJust As XLHAlign
Private pFontVJust As XlVAlign
Private pFill1 As Long
Private pFill2 As Long
Private pShadow As MsoShadowType
Private pTextDir As MsoTextOrientation
Private pName As String

'---------------------------------------------------------------
Public Property Get ForeColour() As Long
    ForeColour = pForeColour
End Property

Public Property Let ForeColour(ByVal vNewValue As Long)
    pForeColour = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get BorderColour() As Long
    BorderColour = pBorderColour
End Property

Public Property Let BorderColour(ByVal vNewValue As Long)
    pBorderColour = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get BorderWidth() As Single
    BorderWidth = pBorderWidth
End Property

Public Property Let BorderWidth(ByVal vNewValue As Single)
    pBorderWidth = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get FontStyle() As String
    FontStyle = pFontStyle
End Property

Public Property Let FontStyle(ByVal vNewValue As String)
    pFontStyle = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get FontBold() As Boolean
    FontBold = pFontBold
End Property

Public Property Let FontBold(ByVal vNewValue As Boolean)
    pFontBold = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get FontSize() As Integer
    FontSize = pFontSize
End Property

Public Property Let FontSize(ByVal vNewValue As Integer)
    pFontSize = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get FontColour() As Long
    FontColour = pFontColour
End Property

Public Property Let FontColour(ByVal vNewValue As Long)
    pFontColour = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get FontXJust() As XLHAlign
    FontXJust = pFontXJust
End Property

Public Property Let FontXJust(ByVal vNewValue As XLHAlign)
    pFontXJust = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get FontVJust() As XlVAlign
    FontVJust = pFontVJust
End Property

Public Property Let FontVJust(ByVal vNewValue As XlVAlign)
    pFontVJust = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Fill1() As Long
    Fill1 = pFill1
End Property

Public Property Let Fill1(ByVal vNewValue As Long)
    pFill1 = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Fill2() As Long
    Fill2 = pFill2
End Property

Public Property Let Fill2(ByVal vNewValue As Long)
    pFill2 = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Shadow() As MsoShadowType
    Shadow = pShadow
End Property

Public Property Let Shadow(ByVal vNewValue As MsoShadowType)
    pShadow = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get TextDir() As MsoTextOrientation
    TextDir = pTextDir
End Property

Public Property Let TextDir(ByVal vNewValue As MsoTextOrientation)
    pTextDir = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal vNewValue As String)
    pName = vNewValue
End Property

' ===============================================================

Private Sub Class_Initialize()

End Sub
' ===============================================================
Private Sub Class_Terminate()
End Sub
' ===============================================================
