Attribute VB_Name = "ModEnums"
'===============================================================
' Module ModEnums
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 23 Jul 22
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModEnums"

' ===============================================================
' Enum Return Functions
' ===============================================================

' ===============================================================
' EnumTriStateVal
' Returns integer value from Enum string
' ---------------------------------------------------------------
Public Function EnumTriStateVal(EnumValue As String) As EnumTriState
    Select Case EnumValue
        Case "xTrue"
            EnumTriStateVal = 0
        Case "xFalse"
            EnumTriStateVal = 1
        Case "xError"
            EnumTriStateVal = 2
    End Select
End Function

' ===============================================================
' EnumBtnNoVal
' Returns integer value from Enum string
' ---------------------------------------------------------------
Public Function EnumBtnNoVal(EnumValue As String) As EnumBtnNo
    Select Case EnumValue
        Case "enBtnForAction"
            EnumBtnNoVal = 1
        Case "enBtnActive"
            EnumBtnNoVal = 2
        Case "enBtnClosed"
            EnumBtnNoVal = 3
        Case "enBtnExit"
            EnumBtnNoVal = 4
        Case "enBtnSupport"
            EnumBtnNoVal = 5
        Case "enBtnNewProjectWF"
            EnumBtnNoVal = 6
    End Select
End Function

' ===============================================================
' enStatusVal
' Returns integer value from Enum string
' ---------------------------------------------------------------
Public Function enStatusVal(EnumValue As String) As enStatus
    Select Case EnumValue
        Case "enNotStarted"
            enStatusVal = 1
        Case "enActionReqd"
            enStatusVal = 2
        Case "enWaiting"
            enStatusVal = 3
        Case "enClosed"
            enStatusVal = 4
    End Select
End Function

' ===============================================================
' enFormValidationVal
' Returns integer value from Enum string
' ---------------------------------------------------------------
Public Function enFormValidationVal(EnumValue As String) As enFormValidation
    Select Case EnumValue
        Case "enFormOK"
            enFormValidationVal = 2
        Case "enValidationError"
            enFormValidationVal = 1
        Case "enFunctionalError"
            enFormValidationVal = 0
    End Select
End Function

' ===============================================================
' enRAGVal
' Returns integer value from Enum string
' ---------------------------------------------------------------
Public Function enRAGVal(EnumValue As String) As enRAG
    Select Case EnumValue
        Case "en3Green"
            enRAGVal = 0
        Case "en2Amber"
            enRAGVal = 1
        Case "en1Red"
            enRAGVal = 2
    End Select
End Function

' ===============================================================
' enStepTypeVal
' Returns integer value from Enum string
' ---------------------------------------------------------------
Public Function enStepTypeVal(EnumValue As String) As enStepType
    Select Case EnumValue
        Case "enYesNo"
            enStepTypeVal = 0
        Case "enStep"
            enStepTypeVal = 1
        Case "enDataInput"
            enStepTypeVal = 2
        Case "enAltBranch"
            enStepTypeVal = 3
    End Select
End Function

' ===============================================================
' EnUserLvlVal
' Returns integer value from Enum string
' ---------------------------------------------------------------
Public Function EnUserLvlVal(EnumValue As String) As EnUserLvl
    Select Case EnumValue
        Case "enBasic"
            EnUserLvlVal = 0
        Case "enAdmin"
            EnUserLvlVal = 1
    End Select
End Function

' ===============================================================
' EnumFormValidationVal
' Returns integer value from Enum string
' ---------------------------------------------------------------
Public Function EnumFormValidationVal(EnumValue As String) As EnumFormValidation
    Select Case EnumValue
        Case "FormOK"
            EnumFormValidationVal = 2
        Case "ValidationError"
            EnumFormValidationVal = 1
        Case "FunctionalError"
            EnumFormValidationVal = 0
    End Select
End Function

' ===============================================================
' EnumTriStateStr
' Returns enum string from Enum integer value
' ---------------------------------------------------------------
Public Function EnumTriStateStr(EnumValue As EnumTriState) As String
    Select Case EnumValue
        Case 0
            EnumTriStateStr = "xTrue"
        Case 1
            EnumTriStateStr = "xFalse"
        Case 2
            EnumTriStateStr = "xError"
    End Select
End Function

' ===============================================================
' EnumBtnNoStr
' Returns enum string from Enum integer value
' ---------------------------------------------------------------
Public Function EnumBtnNoStr(EnumValue As EnumBtnNo) As String
    Select Case EnumValue
        Case 1
            EnumBtnNoStr = "enBtnForAction"
        Case 2
            EnumBtnNoStr = "enBtnActive"
        Case 3
            EnumBtnNoStr = "enBtnClosed"
        Case 4
            EnumBtnNoStr = "enBtnExit"
        Case 5
            EnumBtnNoStr = "enBtnSupport"
        Case 6
            EnumBtnNoStr = "enBtnNewProjectWF"
    End Select
End Function

' ===============================================================
' enStatusStr
' Returns enum string from Enum integer value
' ---------------------------------------------------------------
Public Function enStatusStr(EnumValue As enStatus) As String
    Select Case EnumValue
        Case 1
            enStatusStr = "enNotStarted"
        Case 2
            enStatusStr = "enActionReqd"
        Case 3
            enStatusStr = "enWaiting"
        Case 4
            enStatusStr = "enClosed"
    End Select
End Function

' ===============================================================
' enFormValidationStr
' Returns enum string from Enum integer value
' ---------------------------------------------------------------
Public Function enFormValidationStr(EnumValue As enFormValidation) As String
    Select Case EnumValue
        Case 2
            enFormValidationStr = "enFormOK"
        Case 1
            enFormValidationStr = "enValidationError"
        Case 0
            enFormValidationStr = "enFunctionalError"
    End Select
End Function

' ===============================================================
' enRAGStr
' Returns enum string from Enum integer value
' ---------------------------------------------------------------
Public Function enRAGStr(EnumValue As enRAG) As String
    Select Case EnumValue
        Case 0
            enRAGStr = "en3Green"
        Case 1
            enRAGStr = "en2Amber"
        Case 2
            enRAGStr = "en1Red"
    End Select
End Function

' ===============================================================
' enStepTypeStr
' Returns enum string from Enum integer value
' ---------------------------------------------------------------
Public Function enStepTypeStr(EnumValue As enStepType) As String
    Select Case EnumValue
        Case 0
            enStepTypeStr = "enYesNo"
        Case 1
            enStepTypeStr = "enStep"
        Case 2
            enStepTypeStr = "enDataInput"
        Case 3
            enStepTypeStr = "enAltBranch"
    End Select
End Function

' ===============================================================
' EnUserLvlStr
' Returns enum string from Enum integer value
' ---------------------------------------------------------------
Public Function EnUserLvlStr(EnumValue As EnUserLvl) As String
    Select Case EnumValue
        Case 0
            EnUserLvlStr = "enBasic"
        Case 1
            EnUserLvlStr = "enAdmin"
    End Select
End Function

' ===============================================================
' EnumFormValidationStr
' Returns enum string from Enum integer value
' ---------------------------------------------------------------
Public Function EnumFormValidationStr(EnumValue As EnumFormValidation) As String
    Select Case EnumValue
        Case 2
            EnumFormValidationStr = "FormOK"
        Case 1
            EnumFormValidationStr = "ValidationError"
        Case 0
            EnumFormValidationStr = "FunctionalError"
    End Select
End Function

' ===============================================================
' enStatusDisp
' Returns display string from Enum integer value
' ---------------------------------------------------------------
Public Function enStatusDisp(EnumValue As enStatus) As String
    Select Case EnumValue
        Case enNotStarted
            enStatusDisp = "Not Started"
        Case enActionReqd
            enStatusDisp = "Action Req'd"
        Case enWaiting
            enStatusDisp = "Waiting"
        Case enClosed
            enStatusDisp = "Closed"
    End Select
End Function

' ===============================================================
' enRAGDisp
' Returns display string from Enum integer value
' ---------------------------------------------------------------
Public Function enRAGDisp(EnumValue As enRAG) As String
    Select Case EnumValue
        Case en3Green
            enRAGDisp = "Green"
        Case en2Amber
            enRAGDisp = "Amber"
        Case en1Red
            enRAGDisp = "Red"
    End Select
End Function

' ===============================================================
' enStepTypeDisp
' Returns display string from Enum integer value
' ---------------------------------------------------------------
Public Function enStepTypeDisp(EnumValue As enStepType) As String
    Select Case EnumValue
        Case enYesNo
            enStepTypeDisp = "Yes/No Decision"
        Case enStep
            enStepTypeDisp = "Normal Step"
        Case enDataInput
            enStepTypeDisp = "Data Input"
        Case enAltBranch
            enStepTypeDisp = "Alt. Branch"
    End Select
End Function

' ===============================================================
' EnUserLvlDisp
' Returns display string from Enum integer value
' ---------------------------------------------------------------
Public Function EnUserLvlDisp(EnumValue As EnUserLvl) As String
    Select Case EnumValue
        Case enBasic
            EnUserLvlDisp = "Basic"
        Case enAdmin
            EnUserLvlDisp = "Admin"
    End Select
End Function

