Attribute VB_Name = "ModUIButtonHandler"
'===============================================================
' Module ModUIButtonHandler
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 02 Oct 22
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModUIButtonHandler"

' ===============================================================
' ProcessBtnClicks
' Processes all button presses in application
' ---------------------------------------------------------------
Public Sub ProcessBtnClicks(ButtonNo As String)
    Dim ErrNo As Integer
    Dim AryBtn() As String
    Dim Picker As ClsFrmPicker
    Dim BtnNo As EnumBtnNo
    Dim BtnIndex As Integer
    Dim ScreenPage As enScreenPage
    
    Const StrPROCEDURE As String = "ProcessBtnClicks()"
    
    On Error GoTo ErrorHandler

Restart:

    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART

    AryBtn = Split(ButtonNo, ":")
    If UBound(AryBtn) > 0 Then ScreenPage = AryBtn(0)
    
    If UBound(AryBtn) = 1 Then
        If Not IsNull(BtnIndex) Then BtnIndex = AryBtn(1)
    Else
        If Not IsNull(AryBtn(1)) Then BtnNo = CInt(AryBtn(1))
    End If
    
    If UBound(AryBtn) = 2 Then
        If Not IsNull(CInt(AryBtn(2))) Then BtnIndex = CInt(AryBtn(2))
    End If
    
    Select Case BtnNo
    
        Case enBtnProjectNew
            
            If Not BtnProjectNewWFClick(ScreenPage) Then Err.Raise HANDLED_ERROR
        
        Case enBtnLenderNewWF
        
            If Not BtnLenderNewWFClick(ScreenPage) Then Err.Raise HANDLED_ERROR

        Case enBtnProjectOpen

            BtnProjectOpenWFClick ScreenPage, BtnIndex
            
        Case enBtnLenderOpenWF

            BtnLenderOpenWFClick ScreenPage, BtnIndex

        Case enBtnCRMOpenItem
            
            BtnCRMOpenItem ScreenPage, BtnIndex
            
        Case enBtnCRMNewItem
            
            BtnCRMOpenItem ScreenPage
            
        Case enBtnCRMContCalImport
            
            BtnCRMContCalImport ScreenPage
            
        Case enBtnCRMContShwLeads
            
            BtnCRMContShwLeads ScreenPage
            
    End Select

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
