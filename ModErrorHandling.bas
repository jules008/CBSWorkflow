Attribute VB_Name = "ModErrorHandling"
'===============================================================
' Module ModErrorHandling
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

Private Const StrMODULE As String = "ModErrorHandling"

' ===============================================================
' CentralErrorHandler
' Handles all system errors
' ---------------------------------------------------------------
Public Function CentralErrorHandler( _
            ByVal ErrModule As String, _
            ByVal ErrProc As String, _
            Optional ByVal EntryPoint As Boolean, _
            Optional ByVal CustomData As String) As Boolean

    Static ErrMsg As String
    Static FirstRunDone As Boolean
    
    Dim iFile As Integer
    Dim ErrNum As Long
    Dim ErrFile As String
    Dim ErrHeader As String
    Dim LogText As String
    
    ErrNum = Err.Number
    ErrMsg = Err.Description
    
    If Len(ErrMsg) = 0 Then ErrMsg = Err.Description
                
    On Error Resume Next
    
    ErrFile = ThisWorkbook.Name
    
    If Right$(SYS_PATH, 1) <> "\" Then SYS_PATH = SYS_PATH & "\"
    
    ErrHeader = "[" & Application.Name & "]" & "[" & ErrFile & "]" & ErrModule & "." & ErrProc

    LogText = "  " & ErrHeader & ", Error " & CStr(ErrNum) & ": " & ErrMsg
    
    iFile = FreeFile()
    Open GetDocLocalPath(ThisWorkbook.Path) & ERROR_PATH & FILE_ERROR_LOG For Append As #iFile
    If Not FirstRunDone Then
        Print #iFile,
        Print #iFile, "---------------------------------------------------------------------------------------------------------------------------------------"
        Print #iFile, "User: " & Application.UserName
        Print #iFile, "Project: " & ActiveProject.ProjectNo
        Print #iFile, "Workflow: " & ActiveWorkFlow.WorkflowNo
        Print #iFile, "Step: " & ActiveWorkFlow.CurrentStep
        Print #iFile, "Custom Data: " & CustomData
        FirstRunDone = True
    End If
    Print #iFile, Format$(Now(), "mm/dd/yy hh:mm:ss"); LogText
    Close #iFile
            
    Debug.Print Format$(Now(), "mm/dd/yy hh:mm:ss"); LogText
    
    If EntryPoint Then
        Debug.Print
        ModLibrary.PerfSettingsOff
        
        If Not DEV_MODE And SEND_ERR_MSG Then SendErrMessage
        ErrMsg = vbNullString
        FirstRunDone = False
    End If
    
    CentralErrorHandler = DEV_MODE
    
    ModLibrary.PerfSettingsOff
End Function

' ===============================================================
' CustomErrorHandler
' Handles system custom errors 2000 - 2500
' ---------------------------------------------------------------
Public Function CustomErrorHandler(ErrorCode As Long, Optional Message As String) As Long
    Dim MailSubject As String
    Dim MailBody As String
    
    Const StrPROCEDURE As String = "CustomErrorHandler()"

    On Error Resume Next

    Select Case ErrorCode
        Case UNKNOWN_USER
            
        Case NO_DATABASE_FOUND
            FaultCount1008 = FaultCount1008 + 1
            Debug.Print "Trying to connect to Database....Attempt " & FaultCount1008
            
            If FaultCount1008 <= 3 Then
            
                Application.DisplayStatusBar = True
                Application.StatusBar = "Trying to connect to Database....Attempt " & FaultCount1008
                Application.Wait (Now + TimeValue("0:00:02"))
                Debug.Print FaultCount1008
            Else
                FaultCount1008 = 0
                Application.StatusBar = "System Failed - No Database"
                End
            End If
        
        Case SYSTEM_RESTART
            Debug.Print Now & " - System Restart"
            FaultCount1002 = FaultCount1002 + 1

            If FaultCount1002 <= 3 Then
                Initialize
                Application.DisplayStatusBar = True
                Application.StatusBar = "System failed...Restarting Attempt " & FaultCount1002
                Application.Wait (Now + TimeValue("0:00:02"))
            Else
                FaultCount1002 = 0
                Application.StatusBar = "Sysetm Failed"
                End
            End If
            
        Case ACCESS_DENIED
            MsgBox "You do not have the correct access rights to carry out this action.  Please contact an Administrator for access", vbOKOnly + vbExclamation
            Application.StatusBar = "Access Denied"
            Sleep 2000
            Application.StatusBar = ""
        
        Case NO_INI_FILE
            MsgBox "No INI file has been found, so system cannot continue. This can occur if the file " _
                    & "is copied from its location on the T Drive.  Please delete file and create a shortcut instead", vbCritical, APP_NAME
            Application.StatusBar = "System Failed - No INI File"
            End
        
        Case DB_WRONG_VER
            MsgBox "Incorrect Version Database - System cannot continue", vbCritical + vbOKOnly, APP_NAME
            Application.StatusBar = "System Failed - Wrong DB Version"
            End
        
        Case FORM_INPUT_EMPTY
           MsgBox "Please complete all highlighted fields", vbExclamation, APP_NAME
       
        Case FORM_INPUT_ERROR
           MsgBox "please check highlighted boxes for errors", vbExclamation, APP_NAME
       
        Case NO_USER_SELECTED
            MsgBox "please select a Person", vbExclamation, APP_NAME
        
        Case SYS_ACCESS_DENIED
            MsgBox "To gain access to FIRES, please send a message to an administrator", vbCritical, APP_NAME
            ModCloseDown.Terminate
            ThisWorkbook.Close False
            
        Case ERROR_MSG
            MsgBox Message, vbOKOnly + vbInformation, APP_NAME
    End Select
    
    CustomErrorHandler = ErrorCode
End Function

' ===============================================================
' SendErrMessage
' Sends an email log file
' ---------------------------------------------------------------
Private Sub SendErrMessage()
    
    On Error Resume Next
    
    If MailSystem Is Nothing Then Set MailSystem = New ClsMailSystem

    If Not ModLibrary.OutlookRunning Then
        Shell "Outlook.exe"
    End If

    With MailSystem
        .MailItem.To = "julian.turner@onesheet.co.uk"
        .MailItem.Subject = "Debug Report - " & APP_NAME
        .MailItem.Importance = olImportanceHigh
        .MailItem.Attachments.Add SYS_PATH & FILE_ERROR_LOG
        If SEND_EMAILS Then .SendEmail Else .DisplayEmail
    End With

End Sub
