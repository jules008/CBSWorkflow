Attribute VB_Name = "ModAPICalls"
'===============================================================
' Module ModAPICalls
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 26 Jul 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModAPICalls"

' ===============================================================
' ShellExecute
' Executes shell commands
' ---------------------------------------------------------------
 #If VBA7 Then
    Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As LongPtr, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As LongPtr
#Else
    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
 #End If

' ===============================================================
' CopyMemory
' Copies blocks of memory from one location to another
' ---------------------------------------------------------------
 #If VBA7 Then
    Public Declare PtrSafe Sub CopyMemory _
    Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As LongPtr)
 #Else
    Public Declare Sub CopyMemory _
    Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
 #End If

' ===============================================================
' Sleep
' Pauses execution for a defined number of milliseconds
' ---------------------------------------------------------------
 #If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
 #Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
 #End If

' ===============================================================
' GetSystemMetrics
' Gets the screen height from the API
' ---------------------------------------------------------------
 #If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
    Const SM_CXSCREEN = 0
    Const SM_CYSCREEN = 1
 #Else
    Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
    Const SM_CXSCREEN = 0
    Const SM_CYSCREEN = 1
 #End If

' ===============================================================
' GetScreenHeight
' Gets the screen height from the API
' ---------------------------------------------------------------
Public Function GetScreenHeight() As Integer
    Const StrPROCEDURE As String = "GetScreenHeight()"

    On Error GoTo ErrorHandler

    Dim X  As Long
    Dim y  As Long
   
    X = GetSystemMetrics(SM_CXSCREEN)
    y = GetSystemMetrics(SM_CYSCREEN)

    GetScreenHeight = y

    GetScreenHeight = True

Exit Function

ErrorExit:

    GetScreenHeight = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function



