Attribute VB_Name = "ModGlobals"
'===============================================================
' Module ModGlobals
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 23 Jul 22
'===============================================================
Private Const StrMODULE As String = "ModGlobals"

Option Explicit

' ===============================================================
' Global Constants
' ---------------------------------------------------------------
Public Const PROJECT_FILE_NAME As String = "CBS Workflow"
Public Const APP_NAME As String = "CBSWorkflow"
Public Const DB_FILE_NAME As String = "CBSWorkflowDB"
Public Const INI_FILE_PATH As String = "\System Files\"
Public Const ERROR_PATH As String = "\System Files\"
Public Const TMP_FILES As String = "\System Files\Tmp\"
Public Const BAK_FILES As String = "\System Files\Backups\"
Public Const TEMPLATE_STORE As String = "\System Files\Templates\"
Public Const INI_FILE_NAME As String = "System.ini"
Public Const PROTECT_ON As Boolean = True
Public Const PROTECT_KEY As String = "03383396"
Public Const STOP_FLAG As Boolean = False
Public Const MAINT_MSG As String = ""
Public Const SEND_ERR_MSG As Boolean = False
Public Const TEST_PREFIX As String = "TEST - "
Public Const BACKUP_INT As Integer = 5
Public Const FILE_ERROR_LOG As String = "Error.log"
Public Const OLD_DB_VER = "V0.0.0"
Public Const DB_VER = "V0.0.1"
Public Const VERSION = "V0.0.1"
Public Const VER_DATE = "23 Jul 22"
' ===============================================================
' Error Constants
' ---------------------------------------------------------------
Public Const HANDLED_ERROR As Long = 9999
Public Const UNKNOWN_USER As Long = 2000
Public Const SYSTEM_RESTART As Long = 2001
Public Const NO_DATABASE_FOUND As Long = 2002
Public Const ACCESS_DENIED As Long = 2003
Public Const NO_INI_FILE As Long = 2004
Public Const DB_WRONG_VER As Long = 2005
Public Const GENERIC_ERROR As Long = 2006
Public Const FORM_INPUT_EMPTY As Long = 2007
Public Const NO_USER_SELECTED As Long = 2008
Public Const GRACEFUL_EXIT As Long = 2009
Public Const FORM_INPUT_ERROR As Long = 2010
Public Const ERROR_MSG As Long = 2011
Public Const SYS_ACCESS_DENIED As Long = 2012

' ===============================================================
' Error Variables
' ---------------------------------------------------------------
Public FaultCount1002 As Integer
Public FaultCount1008 As Integer

' ===============================================================
' Global Variables
' ---------------------------------------------------------------
Global DEBUG_MODE As Boolean
Global SYSTEM_CLOSING As Boolean
Global SEND_EMAILS As Boolean
Global ENABLE_PRINT As Boolean
Global DB_PATH As String
Global DEV_MODE As Boolean
Global SYS_PATH As String
Global CURRENT_USER As String
Global MENU_ITEM_SEL As Integer
Global G_DATE As String
Global G_FORM As Boolean

' ===============================================================
' Global Class Declarations
' ---------------------------------------------------------------
Public ActiveWorkFlow As ClsWorkflow

' ===============================================================
' Global UI Class Declarations
' ---------------------------------------------------------------

' ===============================================================
' Colours
' ---------------------------------------------------------------
Public Const COLOUR_1 As Long = &HBBA994
Public Const COLOUR_2 As Long = &H9B6C47
Public Const COLOUR_3 As Long = &H8B991B
Public Const COLOUR_4 As Long = &H151DFF
Public Const COLOUR_5 As Long = &H47453D
Public Const COLOUR_6 As Long = &HD0E528
Public Const COLOUR_7 As Long = &HFFFFFF
Public Const COLOUR_8 As Long = &HEBF7F6
Public Const COLOUR_9 As Long = &HBDBE06
Public Const COLOUR_10 As Long = &HC33E&
Public Const COLOUR_11 As Long = &H31FCFF
Public Const COLOUR_12 As Long = &HEE4A&
Public Const COLOUR_13 As Long = &HB4FF&
Public Const COLOUR_14 As Long = &H696DFF
Public Const COLOUR_15 As Long = &HB4FF&
Public Const COLOUR_16 As Long = &HD9FC&
Public Const COLOUR_17 As Long = &H0&

' ===============================================================
' Type Declarations
' ---------------------------------------------------------------

' ===============================================================
' Enum Declarations
' ---------------------------------------------------------------
Enum EnumTriState
    xTrue
    xFalse
    xError
End Enum

Enum EnumBtnNo
    enBtnForAction = 1
    enBtnActive
    enBtnComplete
    enBtnExit
    enBtnSupport
    enBtnNewWorkflow
End Enum

Enum enStatus           'Status
    enNotStarted = 1    'Not Started
    enActionReqd        'Action Req'd
    enWaiting           'Waiting
    enComplete          'Complete
End Enum                '

Enum enFormValidation
    enFormOK = 2
    enValidationError = 1
    enFunctionalError = 0
End Enum

Enum enRAG              'RAG
    en3Green            'Green
    en2Amber            'Amber
    en1Red              'Red
End Enum                '

Enum enStepType         'Step Type
    enYesNo             'Yes/No Decision
    enStep              'Normal Step
    enDataInput         'Data Input
    enAltBranch         'Alt. Branch
End Enum                '

Enum EnUserLvl          'User Level
    enBasic             'Basic
    enAdmin             'Admin
End Enum                '
 
Enum EnumFormValidation
    FormOK = 2
    ValidationError = 1
    FunctionalError = 0
End Enum

