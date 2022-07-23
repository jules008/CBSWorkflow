Attribute VB_Name = "ModSecurity"
'===============================================================
' Module ModSecurity
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 23 Jul 22
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModSecurity"

' ===============================================================
' LogUserOn
' Logs on user and assigns access level.  Terminates if user is not known
' ---------------------------------------------------------------
Public Function LogUserOn(UserName As String) As Boolean
    Const StrPROCEDURE As String = "LogUserOn()"

    On Error GoTo ErrorHandler

    If UserName = "" Then Err.Raise HANDLED_ERROR, , "Username blank"
'
'    CurrentUser.DBGet UserName
'
'    If CurrentUser.UserLvl = enAdmin Then
'        Debug.Print Now & " - " & UserName & " Logged in as Admin User"
'    Else
'        Debug.Print Now & " - " & UserName & " Logged in as Basic User"
'    End If
    
GracefulExit:

    LogUserOn = True

Exit Function

ErrorExit:

    '***CleanUpCode***
    LogUserOn = False

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
' HideTabs
' Hides tabs for all except dev
' ---------------------------------------------------------------
Public Function HideTabs() As Boolean
    Const StrPROCEDURE As String = "HideTabs()"

    On Error GoTo ErrorHandler

    If DEV_MODE Then
'        ShtColours.Visible = xlSheetVisible
    Else
'        ShtColours.Visible = xlSheetVeryHidden
    End If

    HideTabs = True

Exit Function

ErrorExit:

    '***CleanUpCode***
    HideTabs = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
