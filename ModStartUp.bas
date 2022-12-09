Attribute VB_Name = "ModStartUp"
'===============================================================
' Module ModStartUp
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

Private Const StrMODULE As String = "ModStartUp"

' ===============================================================
' Initialize
' Creates the environment for system start up
' ---------------------------------------------------------------
Public Function Initialize() As Boolean
    Dim UserName As String
    Dim Response As String
    
    Const StrPROCEDURE As String = "Initialize()"

    On Error GoTo ErrorHandler
    
    TimeStart = MicroTimer

    ModLibrary.PerfSettingsOn
    
    ShtMain.Unprotect PROTECT_KEY
        
    SYSTEM_CLOSING = False
    
    Application.DisplayStatusBar = True
    
    FrmStartBanner.Progress "Reading INI File.....", 1 / 7 * 100
    
    If Not ReadINIFile Then Err.Raise HANDLED_ERROR
    
    Terminate
    
    FrmStartBanner.Progress "Connecting to DB.....", 2 / 7 * 100
    
    If Not ModDatabase.DBConnect Then Err.Raise HANDLED_ERROR
    
    FrmStartBanner.Progress "Checking DB Version.....", 3 / 7 * 100
    
    If ModDatabase.GetDBVer = OLD_DB_VER Then
    FrmStartBanner.Progress "Performing database update.....", 3.5 / 7 * 100
        ModDeploy.UpdateDBScript
    End If

    If ShtSettings.ChkUpdateDB Then ModDeploy.UpdateTableData

    If ModDatabase.GetDBVer <> DB_VER Then
        Err.Raise DB_WRONG_VER
        Debug.Print Now & "Globals Ver: ", DB_VER, "DB Ver: ", ModDatabase.GetDBVer
    End If
    
    If Now - ModDatabase.GetBackUpDate > BACKUP_INT Then
    
        FrmStartBanner.Progress "Backing Up Database.....", 3.75 / 7 * 100
        
        ModDatabase.BackupFiles
        DB.Execute "UPDATE TblDBVersion SET lastbackup = now"
    End If
    
    If Not SetGlobalClasses Then Err.Raise HANDLED_ERROR
    
    FrmStartBanner.Progress "Logging User On.....", 4 / 7 * 100

    UserName = GetUserName
    
    If UserName = "Error" Then Err.Raise HANDLED_ERROR
    
    If Not ModSecurity.LogUserOn(UserName) Then Err.Raise HANDLED_ERROR
    
    
    'build styles
    FrmStartBanner.Progress "Building Styles.....", 5 / 7 * 100
    
    If Not ModUIStyles.BuildScreenStyles Then Err.Raise HANDLED_ERROR
    
    Windows(ThisWorkbook.Name).Visible = True
    
    FrmStartBanner.Progress "Buidling UI.....", 6 / 7 * 100
        
    'Build menu and backdrop
    If Not ModUIMenu.BuildMenu Then Err.Raise HANDLED_ERROR
            
    If Not HideTabs Then Err.Raise HANDLED_ERROR
    
    If Not ModUIMenu.ButtonClickEvent("1") Then Err.Raise HANDLED_ERROR
    
'    MenuBar.Menu(1).BadgeText = Workflows.CountForAction
    
    FrmStartBanner.Progress "", 7 / 7 * 100
   
    Application.Wait (Now + TimeValue("00:00:01"))
    
   ModLibrary.PerfSettingsOff
   
    Initialize = True

Exit Function

ErrorExit:

    Initialize = False
    ModLibrary.PerfSettingsOff
    
Exit Function

ErrorHandler:
        
    If Err.Number >= 2000 And Err.Number <= 2500 Then
        CustomErrorHandler Err.Number
        Resume Next
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' GetUserName
' gets username from windows, or test user if in test mode
' ---------------------------------------------------------------
Public Function GetUserName() As String
    Dim UserName As String
    Dim CharPos As Integer
    
    Const StrPROCEDURE As String = "GetUserName()"

    On Error GoTo ErrorHandler
    
    If Not UpdateUsername Then Err.Raise HANDLED_ERROR
    
    UserName = Application.UserName
    
    If UserName = "" Then Err.Raise UNKNOWN_USER

    GetUserName = Replace(UserName, "'", "")
        
GracefulExit:
    
Exit Function

ErrorExit:

    GetUserName = "Error"

Exit Function

ErrorHandler:
        
    If Err.Number >= 2000 And Err.Number <= 2500 Then
        CustomErrorHandler Err.Number
        Resume GracefulExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' ReadINIFile
' Gets start up variables from ini file
' ---------------------------------------------------------------
Public Function ReadINIFile() As Boolean
    Dim DebugMode() As String
    Dim EnablePrint() As String
    Dim DBPath() As String
    Dim SendEmails() As String
    Dim DevMode() As String
    Dim StopOnStart() As String
    
    Dim INIFile As Integer
    Dim Line1 As String
    Dim Line2 As String
    Dim Line3 As String
    Dim Line4 As String
    Dim Line5 As String
    Dim Line6 As String
    
    Const StrPROCEDURE As String = "ReadINIFile()"

    On Error GoTo ErrorHandler
       
    INIFile = FreeFile()
    
    SYS_PATH = GetDocLocalPath(ThisWorkbook.Path) & INI_FILE_PATH

    If Dir(SYS_PATH & INI_FILE_NAME) = "" Then Err.Raise NO_INI_FILE
    
    Open SYS_PATH & INI_FILE_NAME For Input As #INIFile
    
    Line Input #INIFile, Line1
    Line Input #INIFile, Line2
    Line Input #INIFile, Line3
    Line Input #INIFile, Line4
    Line Input #INIFile, Line5
    Line Input #INIFile, Line6
    
    Close #INIFile
    DebugMode = Split(Line1, ":")
    SendEmails = Split(Line2, ":")
    EnablePrint = Split(Line3, ":")
    DBPath = Split(Line4, ":")
    DevMode = Split(Line5, ":")
    StopOnStart = Split(Line6, ":")
    
    Line1 = Trim(DebugMode(1))
    Line2 = Trim(SendEmails(1))
    Line3 = Trim(EnablePrint(1))
    Line4 = Trim(DBPath(1))
    Line5 = Trim(DevMode(1))
    Line6 = Trim(StopOnStart(1))
    
    DEBUG_MODE = CBool(Line1)
    SEND_EMAILS = CBool(Line2)
    ENABLE_PRINT = CBool(Line3)
    DB_PATH = Line4
    DEV_MODE = CBool(Line5)
    STOP_FLAG = CBool(Line6)
    
    If STOP_FLAG = True Then Stop
    
    If MAINT_MSG <> "" Then
        MsgBox MAINT_MSG, vbExclamation, APP_NAME
        Application.DisplayAlerts = False
        ActiveWorkbook.Close
        Application.DisplayAlerts = True
        
    End If
    
    
GracefulExit:
    
    ReadINIFile = True
    Application.DisplayAlerts = True

Exit Function

ErrorExit:

    ReadINIFile = False
    Application.DisplayAlerts = True

Exit Function

ErrorHandler:
    
    If Err.Number >= 2000 And Err.Number <= 2500 Then
        CustomErrorHandler Err.Number
        Resume ErrorExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' MessageCheck
' Checks to see if the user message has been read
' ---------------------------------------------------------------
Public Function MessageCheck() As Boolean
    Dim StrMessage As String
    Dim RstMessage As Recordset
    
    Const StrPROCEDURE As String = "MessageCheck()"

    On Error GoTo ErrorHandler
    
'    If CurrentUser.AccessLvl >= BasicLvl_1 Then
'        If Not CurrentUser.MessageRead Then
'
'            Set RstMessage = SQLQuery("TblMessage")
'
'            If RstMessage.RecordCount > 0 Then StrMessage = RstMessage.Fields(0)
'            MsgBox StrMessage, vbOKOnly + vbInformation, APP_NAME
'            CurrentUser.MessageRead = True
'            CurrentUser.DBSave
'
'        End If
'    End If
    
    Set RstMessage = Nothing
    
    MessageCheck = True

Exit Function

ErrorExit:
    Set RstMessage = Nothing
'    ***CleanUpCode***
    MessageCheck = False

Exit Function

ErrorHandler:
    
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' UpdateUsername
' Checks to see whether username needs to be changed and then updates
' ---------------------------------------------------------------
Private Function UpdateUsername() As Boolean
    Const StrPROCEDURE As String = "UpdateUsername()"

    On Error GoTo ErrorHandler
    
    UpdateUsername = True

Exit Function

ErrorExit:

'    ***CleanUpCode***
    UpdateUsername = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


' ===============================================================
' SetGlobalClasses
' Initializes or terminates all global classes
' ---------------------------------------------------------------
Private Function SetGlobalClasses() As Boolean
    
    Const StrPROCEDURE As String = "SetGlobalClasses()"

    On Error GoTo ErrorHandler

    Set CurrentUser = New ClsCBSUser
    Set ActiveWorkflows = New ClsWorkflows
    Set MailSystem = New ClsMailSystem
    Set MailInbox = New ClsMailInbox
    Set CodeTimer = New ClsCodeTimer
    ActiveWorkflows.UpdateRAGs

    SetGlobalClasses = True

Exit Function

ErrorExit:

    '***CleanUpCode***
    SetGlobalClasses = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
