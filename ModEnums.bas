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
' Date - 24 Oct 22
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
' enWorkflowTypeVal
' Returns integer value from Enum string
' ---------------------------------------------------------------
Public Function enWorkflowTypeVal(EnumValue As String) As enWorkflowType
    Select Case EnumValue
        Case "enProject"
            enWorkflowTypeVal = 0
        Case "enLender"
            enWorkflowTypeVal = 1
    End Select
End Function

' ===============================================================
' EnumObjTypeVal
' Returns integer value from Enum string
' ---------------------------------------------------------------
Public Function EnumObjTypeVal(EnumValue As String) As EnumObjType
    Select Case EnumValue
        Case "ObjImage"
            EnumObjTypeVal = 1
        Case "ObjChart"
            EnumObjTypeVal = 2
    End Select
End Function

' ===============================================================
' EnMenuBtnNoVal
' Returns integer value from Enum string
' ---------------------------------------------------------------
Public Function EnMenuBtnNoVal(EnumValue As String) As EnMenuBtnNo
    Select Case EnumValue
        Case "enBtnForAction"
            EnMenuBtnNoVal = 1
        Case "enBtnProjectsActive"
            EnMenuBtnNoVal = 2
        Case "enBtnProjectsClosed"
            EnMenuBtnNoVal = 2
        Case "enBtnCRMClient"
            EnMenuBtnNoVal = 3
        Case "enBtnCRMSPV"
            EnMenuBtnNoVal = 3
        Case "enBtnCRMContacts"
            EnMenuBtnNoVal = 3
        Case "enBtnCRMProjects"
            EnMenuBtnNoVal = 3
        Case "enBtnCRMLenders"
            EnMenuBtnNoVal = 3
        Case "enbtnDashboard"
            EnMenuBtnNoVal = 4
        Case "enBtnReports"
            EnMenuBtnNoVal = 5
        Case "enBtnAdminUsers"
            EnMenuBtnNoVal = 6
        Case "enBtnAdminEmailTs"
            EnMenuBtnNoVal = 6
        Case "enBtnAdminDocuments"
            EnMenuBtnNoVal = 6
        Case "enBtnAdminWorkflows"
            EnMenuBtnNoVal = 6
        Case "enBtnAdminWFTypes"
            EnMenuBtnNoVal = 6
        Case "enBtnAdminLists"
            EnMenuBtnNoVal = 6
        Case "enBtnAdminRoles"
            EnMenuBtnNoVal = 6
        Case "enBtnExit"
            EnMenuBtnNoVal = 7
    End Select
End Function

' ===============================================================
' EnumBtnNoVal
' Returns integer value from Enum string
' ---------------------------------------------------------------
Public Function EnumBtnNoVal(EnumValue As String) As EnumBtnNo
    Select Case EnumValue
        Case "enBtnProjectNew"
            EnumBtnNoVal = 0
        Case "enBtnProjectOpen"
            EnumBtnNoVal = 1
        Case "enBtnLenderNewWF"
            EnumBtnNoVal = 2
        Case "enBtnCRMOpenItem"
            EnumBtnNoVal = 3
        Case "enBtnCRMLenderOpen"
            EnumBtnNoVal = 4
        Case "enBtnCRMLenderNew"
            EnumBtnNoVal = 5
        Case "enBtnCRMNewItem"
            EnumBtnNoVal = 6
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
' enScreenPageVal
' Returns integer value from Enum string
' ---------------------------------------------------------------
Public Function enScreenPageVal(EnumValue As String) As enScreenPage
    Select Case EnumValue
        Case "enScrProjForAction"
            enScreenPageVal = 0
        Case "enScrProjActive"
            enScreenPageVal = 1
        Case "enScrProjComplete"
            enScreenPageVal = 2
        Case "enScrCRMClient"
            enScreenPageVal = 3
        Case "enScrCRMSPV"
            enScreenPageVal = 4
        Case "enScrCRMContact"
            enScreenPageVal = 5
        Case "enScrCRMProject"
            enScreenPageVal = 6
        Case "enScrCRMLender"
            enScreenPageVal = 7
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
' enWorkflowTypeStr
' Returns enum string from Enum integer value
' ---------------------------------------------------------------
Public Function enWorkflowTypeStr(EnumValue As enWorkflowType) As String
    Select Case EnumValue
        Case 0
            enWorkflowTypeStr = "enProject"
        Case 1
            enWorkflowTypeStr = "enLender"
    End Select
End Function

' ===============================================================
' EnumObjTypeStr
' Returns enum string from Enum integer value
' ---------------------------------------------------------------
Public Function EnumObjTypeStr(EnumValue As EnumObjType) As String
    Select Case EnumValue
        Case 1
            EnumObjTypeStr = "ObjImage"
        Case 2
            EnumObjTypeStr = "ObjChart"
    End Select
End Function

' ===============================================================
' EnMenuBtnNoStr
' Returns enum string from Enum integer value
' ---------------------------------------------------------------
Public Function EnMenuBtnNoStr(EnumValue As EnMenuBtnNo) As String
    Select Case EnumValue
        Case 1
            EnMenuBtnNoStr = "enBtnForAction"
        Case 2
            EnMenuBtnNoStr = "enBtnProjectsActive"
        Case 2
            EnMenuBtnNoStr = "enBtnProjectsClosed"
        Case 3
            EnMenuBtnNoStr = "enBtnCRMClient"
        Case 3
            EnMenuBtnNoStr = "enBtnCRMSPV"
        Case 3
            EnMenuBtnNoStr = "enBtnCRMContacts"
        Case 3
            EnMenuBtnNoStr = "enBtnCRMProjects"
        Case 3
            EnMenuBtnNoStr = "enBtnCRMLenders"
        Case 4
            EnMenuBtnNoStr = "enbtnDashboard"
        Case 5
            EnMenuBtnNoStr = "enBtnReports"
        Case 6
            EnMenuBtnNoStr = "enBtnAdminUsers"
        Case 6
            EnMenuBtnNoStr = "enBtnAdminEmailTs"
        Case 6
            EnMenuBtnNoStr = "enBtnAdminDocuments"
        Case 6
            EnMenuBtnNoStr = "enBtnAdminWorkflows"
        Case 6
            EnMenuBtnNoStr = "enBtnAdminWFTypes"
        Case 6
            EnMenuBtnNoStr = "enBtnAdminLists"
        Case 6
            EnMenuBtnNoStr = "enBtnAdminRoles"
        Case 7
            EnMenuBtnNoStr = "enBtnExit"
    End Select
End Function

' ===============================================================
' EnumBtnNoStr
' Returns enum string from Enum integer value
' ---------------------------------------------------------------
Public Function EnumBtnNoStr(EnumValue As EnumBtnNo) As String
    Select Case EnumValue
        Case 0
            EnumBtnNoStr = "enBtnProjectNew"
        Case 1
            EnumBtnNoStr = "enBtnProjectOpen"
        Case 2
            EnumBtnNoStr = "enBtnLenderNewWF"
        Case 3
            EnumBtnNoStr = "enBtnCRMOpenItem"
        Case 4
            EnumBtnNoStr = "enBtnCRMLenderOpen"
        Case 5
            EnumBtnNoStr = "enBtnCRMLenderNew"
        Case 6
            EnumBtnNoStr = "enBtnCRMNewItem"
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
' enScreenPageStr
' Returns enum string from Enum integer value
' ---------------------------------------------------------------
Public Function enScreenPageStr(EnumValue As enScreenPage) As String
    Select Case EnumValue
        Case 0
            enScreenPageStr = "enScrProjForAction"
        Case 1
            enScreenPageStr = "enScrProjActive"
        Case 2
            enScreenPageStr = "enScrProjComplete"
        Case 3
            enScreenPageStr = "enScrCRMClient"
        Case 4
            enScreenPageStr = "enScrCRMSPV"
        Case 5
            enScreenPageStr = "enScrCRMContact"
        Case 6
            enScreenPageStr = "enScrCRMProject"
        Case 7
            enScreenPageStr = "enScrCRMLender"
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

