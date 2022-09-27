VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsCodeTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsCodeTimer
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 01 Jan 21
'===============================================================
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare PtrSafe Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare PtrSafe Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Private Declare PtrSafe Function getTickCount Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
Private Const sCPURegKey = "HARDWARE\DESCRIPTION\System\CentralProcessor\0"
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private pStartingTime As Double

' ===============================================================
' StartTimer
' Starts CTimer
' ---------------------------------------------------------------
Public Sub StartTimer()
    pStartingTime = MarkTime("Start")
End Sub

' ===============================================================
' MarkTime
' marks the time
' ---------------------------------------------------------------
Public Function MarkTime(ProcName As String) As Long
    Dim cyTicks1 As Currency
    Dim MicroTimer As Double
    
    Static cyFrequency As Currency
    '
    MicroTimer = 0
    If cyFrequency = 0 Then getFrequency cyFrequency            ' get ticks/sec
    getTickCount cyTicks1                                       ' get ticks
    If cyFrequency Then MicroTimer = cyTicks1 / cyFrequency ' calc seconds
    If ProcName = "Start" Then
        Debug.Print "Procedure: "; ProcName, "TimeStamp: "; Format(0, "0.000")
    Else
        Debug.Print "Procedure: "; ProcName, "TimeStamp: "; Format((MicroTimer - pStartingTime), "0.000")
    End If
    MarkTime = MicroTimer
End Function

