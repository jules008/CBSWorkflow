Attribute VB_Name = "ModUISettings"
'===============================================================
' Module ModUISettings
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 18 May 20
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModUISettings"

' ===============================================================
' Global Constants
' ---------------------------------------------------------------
' Main Frame
' ---------------------------------------------------------------
Public Const SCREEN_HEIGHT As Integer = 2000
Public Const SCREEN_WIDTH As Integer = 1825
Public Const MAIN_FRAME_TOP As Integer = 90
Public Const MAIN_FRAME_LEFT As Integer = 170
Public Const MAIN_FRAME_WIDTH As Integer = 750
Public Const MAIN_FRAME_HEIGHT As Integer = 360

' ---------------------------------------------------------------
' Button Frame
' ---------------------------------------------------------------
Public Const BUTTON_FRAME_TOP As Integer = 5
Public Const BUTTON_FRAME_LEFT As Integer = 615
Public Const BUTTON_FRAME_WIDTH As Integer = 275
Public Const BUTTON_FRAME_HEIGHT As Integer = 60

' ---------------------------------------------------------------
' Generic Button
' ---------------------------------------------------------------
Public Const GENERIC_BUTTON_WIDTH As Integer = 100
Public Const GENERIC_BUTTON_HEIGHT As Integer = 40

' ---------------------------------------------------------------
' Generic Table Settings
' ---------------------------------------------------------------
Public Const GENERIC_TABLE_ROW_HEIGHT As Integer = 20
Public Const GENERIC_TABLE_WIDTH As Integer = 550
Public Const GENERIC_TABLE_LEFT As Integer = 0
Public Const GENERIC_TABLE_TOP As Integer = 25
Public Const GENERIC_TABLE_ROWOFFSET As Integer = 0
Public Const GENERIC_TABLE_COLOFFSET As Integer = 0
Public Const GENERIC_TABLE_HEADING_HEIGHT As Integer = 20
Public Const GENERIC_TABLE_EXPAND_ICON As String = "Expand.png"
' ---------------------------------------------------------------
' Menu Bar
' ---------------------------------------------------------------
Public Const BUTTON_HEIGHT As Integer = 31
Public Const MENUBAR_HEIGHT As Integer = 2000
Public Const MENUBAR_WIDTH As Integer = 155
Public Const MENUBAR_TOP As Integer = 0
Public Const MENUBAR_LEFT As Integer = 0
Public Const MENU_TOP As Integer = 150
Public Const MENU_LEFT As Integer = 5
Public Const BUTTON_WIDTH As Integer = 150
Public Const BUTTON_COUNT As Integer = 21
Public Const BUTTON_TEXT = "For Action:Projects:Active:Closed:CRM:Clients:SPVs:Contacts:Projects:Lenders:Dashboard:Reports:Admin:Users:Email Templates:Documents:Workflows:Workflow Types:Lists:Roles:Exit"
Public Const BUTTON_INDEX = "1:2:2.1:2.2:3:3.1:3.2:3.3:3.4:3.5:4:5:6:6.1:6.2:6.3:6.4:6.5:6.6:6.7:7"
Public Const LOGO_FILE As String = "Logo.jpg"
Public Const LOGO_TOP As Integer = 13
Public Const LOGO_LEFT As Integer = 5
Public Const LOGO_WIDTH As Integer = 139
Public Const LOGO_HEIGHT As Integer = 90
Public Const HEADER_HEIGHT As Integer = 25
Public Const HEADER_ICON_TOP As Integer = 2
Public Const HEADER_ICON_RIGHT As Integer = 10

' ---------------------------------------------------------------
' Generic Project Screen Settings
' ---------------------------------------------------------------
Public Const PROJECT_TABLE_COL_WIDTHS As String = "50:50:100:100:100:50:200:100"
Public Const PROJECT_CELL_ROW_HEIGHTS As String = "50:50"
Public Const PROJECT_TABLE_STYLES As String = "GENERIC_TABLE"
Public Const PROJECT_TABLE_TITLES As String = "Expand:Project No:Client Name:SPV Name:Case Manager:Step No:Step Name:Status"
Public Const PROJECT_MAX_LINES As Integer = 150
Public Const PROJECT_BTN_MAIN_1_LEFT As Integer = 30
Public Const PROJECT_BTN_MAIN_1_TOP As Integer = 20
Public Const PROJECT_BTN_MAIN_2_LEFT As Integer = 155
Public Const PROJECT_BTN_MAIN_2_TOP As Integer = 20

' ---------------------------------------------------------------
' Project Sub Table Settings
' ---------------------------------------------------------------
Public Const PROJECT_SUB_TABLE_COL_WIDTHS As String = "100:200:100:200:100"
Public Const PROJECT_SUB_TABLE_TITLES As String = "Workflow No:Lender Name:Step No:Step Name:Status"

' ---------------------------------------------------------------
' Project For Action Screen Settings
' ---------------------------------------------------------------
Public Const PROJ_FOR_ACTION_HEADER_TEXT As String = "Project For Action"

' ---------------------------------------------------------------
' Project Active Screen Settings
' ---------------------------------------------------------------
Public Const PROJ_ACTIVE_HEADER_TEXT As String = "Active Projects"

' ---------------------------------------------------------------
' Project Closed Screen Settings
' ---------------------------------------------------------------
Public Const PROJ_CLOSED_HEADER_TEXT As String = "Closed Projects"

' ---------------------------------------------------------------
' Generic CRM Screen Settings
' ---------------------------------------------------------------
Public Const CRM_BTN_MAIN_1_LEFT As Integer = 155
Public Const CRM_BTN_MAIN_1_TOP As Integer = 20
Public Const CRM_CELL_ROW_HEIGHTS As String = "50:50"
Public Const CRM_TABLE_STYLES As String = "GENERIC_TABLE"

' ---------------------------------------------------------------
' CRM Clients Screen
' ---------------------------------------------------------------
Public Const CRM_CLIENT_TABLE_COL_WIDTHS As String = "100:250:150:200"
Public Const CRM_CLIENT_TABLE_TITLES As String = "Client No:Client Name:Phone No:Url"
Public Const CRM_CLIENT_MAX_LINES As Integer = 150

' ---------------------------------------------------------------
' CRM SPV Screen
' ---------------------------------------------------------------
Public Const CRM_SPV_TABLE_COL_WIDTHS As String = "100:300:300"
Public Const CRM_SPV_TABLE_TITLES As String = "SPV No:SPV Name: "
Public Const CRM_SPV_MAX_LINES As Integer = 150

' ---------------------------------------------------------------
' CRM Contact Screen
' ---------------------------------------------------------------
Public Const CRM_CONTACT_TABLE_COL_WIDTHS As String = "100:200:200:150:50"
Public Const CRM_CONTACT_TABLE_TITLES As String = "Contact No:Contact Name:Position:Phone No: "
Public Const CRM_CONTACT_MAX_LINES As Integer = 150

' ---------------------------------------------------------------
' CRM Projects Screen
' ---------------------------------------------------------------
Public Const CRM_PROJECT_TABLE_COL_WIDTHS As String = "100:250:200:200"
Public Const CRM_PROJECT_TABLE_TITLES As String = "Project No:Client Name:SPV:Case Manager"
Public Const CRM_PROJECT_MAX_LINES As Integer = 150

' ---------------------------------------------------------------
' CRM Lenders Screen
' ---------------------------------------------------------------
Public Const CRM_LENDER_TABLE_COL_WIDTHS As String = "100:200:100:100:200"
Public Const CRM_LENDER_TABLE_TITLES As String = "Lender No:Name:Phone No:Lender Type:Address"
Public Const CRM_LENDER_MAX_LINES As Integer = 150

' ===============================================================
' Style Declarations
' ---------------------------------------------------------------
' Main Screen
' ---------------------------------------------------------------
Public SCREEN_STYLE As ClsUIStyle
Public MENUBAR_STYLE As ClsUIStyle
Public BUTTON_SET_STYLE As ClsUIStyle
Public BUTTON_UNSET_STYLE As ClsUIStyle
Public MAIN_FRAME_STYLE As ClsUIStyle
Public BUTTON_FRAME_STYLE As ClsUIStyle
Public HEADER_STYLE As ClsUIStyle
Public BTN_MAIN_STYLE As ClsUIStyle
Public GENERIC_BUTTON As ClsUIStyle
Public GENERIC_TABLE As ClsUIStyle
Public GREEN_CELL As ClsUIStyle
Public AMBER_CELL As ClsUIStyle
Public RED_CELL As ClsUIStyle
Public GENERIC_TABLE_HEADER As ClsUIStyle
Public SUB_TABLE_HEADER As ClsUIStyle

' ---------------------------------------------------------------
' New Order Workflow
' ---------------------------------------------------------------
'Public WF_MAINSCREEN_STYLE As ClsUIStyle

' ===============================================================
' Style Definitions
' ===============================================================
' Generic Styles
' ---------------------------------------------------------------
' Buttons
' ---------------------------------------------------------------
Public Const GENERIC_BUTTON_BORDER_WIDTH As Single = 0
Public Const GENERIC_BUTTON_FILL_1 As Long = COL_BLUE
Public Const GENERIC_BUTTON_FILL_2 As Long = COL_BLUE
Public Const GENERIC_BUTTON_SHADOW As Long = msoShadow21
Public Const GENERIC_BUTTON_FONT_STYLE As String = "Eras Medium ITC"
Public Const GENERIC_BUTTON_FONT_SIZE As Integer = 12
Public Const GENERIC_BUTTON_FONT_COLOUR As Long = COL_WHITE
Public Const GENERIC_BUTTON_FONT_BOLD As Boolean = False
Public Const GENERIC_BUTTON_FONT_X_JUST As Integer = xlHAlignCenter
Public Const GENERIC_BUTTON_FONT_Y_JUST As Integer = xlVAlignCenter

' ---------------------------------------------------------------
' Cells
' ---------------------------------------------------------------
Public Const GENERIC_TABLE_BORDER_WIDTH As Single = 0
Public Const GENERIC_TABLE_FILL_1 As Long = COL_WHITE
Public Const GENERIC_TABLE_FILL_2 As Long = COL_WHITE
Public Const GENERIC_TABLE_SHADOW As Long = 0
Public Const GENERIC_TABLE_FONT_STYLE As String = "Eras Medium ITC"
Public Const GENERIC_TABLE_FONT_SIZE As Integer = 10
Public Const GENERIC_TABLE_FONT_COLOUR As Long = COL_DRK_GREY
Public Const GENERIC_TABLE_FONT_BOLD As Boolean = False
Public Const GENERIC_TABLE_FONT_X_JUST As Integer = xlHAlignCenter
Public Const GENERIC_TABLE_FONT_Y_JUST As Integer = xlVAlignBottom

' ---------------------------------------------------------------
' Cell Headers
' ---------------------------------------------------------------
Public Const GENERIC_TABLE_HEADER_BORDER_WIDTH As Single = 0
Public Const GENERIC_TABLE_HEADER_FILL_1 As Long = COL_BLUE
Public Const GENERIC_TABLE_HEADER_FILL_2 As Long = COL_BLUE
Public Const GENERIC_TABLE_HEADER_SHADOW As Long = 0
Public Const GENERIC_TABLE_HEADER_FONT_STYLE As String = "Calibri"
Public Const GENERIC_TABLE_HEADER_FONT_SIZE As Integer = 10
Public Const GENERIC_TABLE_HEADER_FONT_COLOUR As Long = COL_WHITE
Public Const GENERIC_TABLE_HEADER_FONT_BOLD As Boolean = False
Public Const GENERIC_TABLE_HEADER_FONT_X_JUST As Integer = xlHAlignCenter
Public Const GENERIC_TABLE_HEADER_FONT_Y_JUST As Integer = xlVAlignBottom

' ---------------------------------------------------------------
' Sub Table Header
' ---------------------------------------------------------------
Public Const SUB_TABLE_HEADER_FILL_1 As Long = COL_AQUA
Public Const SUB_TABLE_HEADER_FILL_2 As Long = COL_AQUA

' ---------------------------------------------------------------
' Text Box
' ---------------------------------------------------------------
Public Const TRANSPARENT_TEXT_BOX_BORDER_WIDTH As Single = 0
Public Const TRANSPARENT_TEXT_BOX_FILL_1 As Long = COL_BLUE
Public Const TRANSPARENT_TEXT_BOX_FILL_2 As Long = COL_BLUE
Public Const TRANSPARENT_TEXT_BOX_SHADOW As Long = 0
Public Const TRANSPARENT_TEXT_BOX_FONT_STYLE As String = "Eras Medium ITC"
Public Const TRANSPARENT_TEXT_BOX_FONT_SIZE As Integer = 10
Public Const TRANSPARENT_TEXT_BOX_FONT_COLOUR As Long = COL_PINK
Public Const TRANSPARENT_TEXT_BOX_FONT_BOLD As Boolean = False
Public Const TRANSPARENT_TEXT_BOX_FONT_X_JUST As Integer = xlHAlignLeft
Public Const TRANSPARENT_TEXT_BOX_FONT_Y_JUST As Integer = xlVAlignCenter

' ===============================================================
' Special Styles
' ===============================================================
' Cell Green
' ---------------------------------------------------------------
Public Const GREEN_CELL_BORDER_WIDTH As Single = 0
Public Const GREEN_CELL_FILL_1 As Long = COL_GREEN
Public Const GREEN_CELL_FILL_2 As Long = COL_GREEN
Public Const GREEN_CELL_SHADOW As Long = 0
Public Const GREEN_CELL_FONT_STYLE As String = "Eras Medium ITC"
Public Const GREEN_CELL_FONT_SIZE As Integer = 11
Public Const GREEN_CELL_FONT_COLOUR As Long = COL_DRK_GREY
Public Const GREEN_CELL_FONT_BOLD As Boolean = False
Public Const GREEN_CELL_FONT_X_JUST As Integer = xlHAlignCenter
Public Const GREEN_CELL_FONT_Y_JUST As Integer = xlVAlignCenter

' ---------------------------------------------------------------
' Cell Amber
' ---------------------------------------------------------------
Public Const AMBER_CELL_BORDER_WIDTH As Single = 0
Public Const AMBER_CELL_FILL_1 As Long = COL_AMBER
Public Const AMBER_CELL_FILL_2 As Long = COL_AMBER
Public Const AMBER_CELL_SHADOW As Long = 0
Public Const AMBER_CELL_FONT_STYLE As String = "Eras Medium ITC"
Public Const AMBER_CELL_FONT_SIZE As Integer = 11
Public Const AMBER_CELL_FONT_COLOUR As Long = COL_DRK_GREY
Public Const AMBER_CELL_FONT_BOLD As Boolean = False
Public Const AMBER_CELL_FONT_X_JUST As Integer = xlHAlignCenter
Public Const AMBER_CELL_FONT_Y_JUST As Integer = xlVAlignCenter

' ---------------------------------------------------------------
' Cell Red
' ---------------------------------------------------------------
Public Const RED_CELL_BORDER_WIDTH As Single = 0
Public Const RED_CELL_FILL_1 As Long = COL_RED
Public Const RED_CELL_FILL_2 As Long = COL_RED
Public Const RED_CELL_SHADOW As Long = 0
Public Const RED_CELL_FONT_STYLE As String = "Eras Medium ITC"
Public Const RED_CELL_FONT_SIZE As Integer = 11
Public Const RED_CELL_FONT_COLOUR As Long = COL_WHITE
Public Const RED_CELL_FONT_BOLD As Boolean = False
Public Const RED_CELL_FONT_X_JUST As Integer = xlHAlignCenter
Public Const RED_CELL_FONT_Y_JUST As Integer = xlVAlignCenter

' ---------------------------------------------------------------
' Tool Buttons
' ---------------------------------------------------------------
Public Const TOOL_BUTTON_BORDER_WIDTH As Single = 0
Public Const TOOL_BUTTON_FILL_1 As Long = COL_WHITE
Public Const TOOL_BUTTON_FILL_2 As Long = COL_WHITE
Public Const TOOL_BUTTON_SHADOW As Long = msoShadow21
Public Const TOOL_BUTTON_FONT_STYLE As String = "Eras Medium ITC"
Public Const TOOL_BUTTON_FONT_SIZE As Integer = 9
Public Const TOOL_BUTTON_FONT_COLOUR As Long = COL_AMBER
Public Const TOOL_BUTTON_FONT_BOLD As Boolean = False
Public Const TOOL_BUTTON_FONT_X_JUST As Integer = xlHAlignCenter
Public Const TOOL_BUTTON_FONT_Y_JUST As Integer = xlVAlignCenter

' ---------------------------------------------------------------
' Support Button
' ---------------------------------------------------------------
Public Const BTN_SUPPORT_BORDER_WIDTH As Single = 0.5
Public Const BTN_SUPPORT_BORDER_COLOUR As Long = COL_GREEN
Public Const BTN_SUPPORT_FILL_1 As Long = COL_WHITE
Public Const BTN_SUPPORT_FILL_2 As Long = COL_WHITE
Public Const BTN_SUPPORT_SHADOW As Long = 0
Public Const BTN_SUPPORT_FONT_STYLE As String = "Eras Medium ITC"
Public Const BTN_SUPPORT_FONT_SIZE As Integer = 9
Public Const BTN_SUPPORT_FONT_COLOUR As Long = COL_GREEN
Public Const BTN_SUPPORT_FONT_BOLD As Boolean = False
Public Const BTN_SUPPORT_FONT_X_JUST As Integer = xlHAlignCenter
Public Const BTN_SUPPORT_FONT_Y_JUST As Integer = xlVAlignCenter

' ===============================================================
' Main Screen
' ===============================================================
Public Const SCREEN_BORDER_WIDTH As Single = 0
Public Const SCREEN_FILL_1 As Long = COL_OFF_WHITE
Public Const SCREEN_FILL_2 As Long = COL_OFF_WHITE
Public Const SCREEN_SHADOW As Long = 0

Public Const MENUBAR_BORDER_WIDTH As Single = 1
Public Const MENUBAR_FILL_1 As Long = COL_WHITE
Public Const MENUBAR_FILL_2 As Long = COL_WHITE
Public Const MENUBAR_SHADOW As Long = msoShadow21

Public Const BUTTON_UNSET_BORDER_WIDTH As Single = 1
Public Const BUTTON_UNSET_BORDER_COLOUR As Long = COL_BLUE
Public Const BUTTON_UNSET_FILL_1 As Long = COL_WHITE
Public Const BUTTON_UNSET_FILL_2 As Long = COL_WHITE
Public Const BUTTON_UNSET_SHADOW As Long = 0
Public Const BUTTON_UNSET_FONT_STYLE As String = "Eras Medium ITC"
Public Const BUTTON_UNSET_FONT_SIZE As Integer = 12
Public Const BUTTON_UNSET_FONT_COLOUR As Long = COL_BLACK
Public Const BUTTON_UNSET_FONT_X_JUST As Integer = xlHAlignCenter
Public Const BUTTON_UNSET_FONT_Y_JUST As Integer = xlVAlignCenter

Public Const BUTTON_SET_BORDER_WIDTH As Single = 0
Public Const BUTTON_SET_BORDER_COLOUR As Long = COL_BLUE
Public Const BUTTON_SET_FILL_1 As Long = COL_PINK
Public Const BUTTON_SET_FILL_2 As Long = COL_PINK
Public Const BUTTON_SET_SHADOW As Long = 0
Public Const BUTTON_SET_FONT_STYLE As String = "Eras Medium ITC"
Public Const BUTTON_SET_FONT_SIZE As Integer = 12
Public Const BUTTON_SET_FONT_COLOUR As Long = COL_WHITE
Public Const BUTTON_SET_FONT_X_JUST As Integer = xlHAlignCenter
Public Const BUTTON_SET_FONT_Y_JUST As Integer = xlVAlignCenter

Public Const MAIN_FRAME_BORDER_WIDTH As Single = 0
Public Const MAIN_FRAME_FILL_1 As Long = COL_WHITE
Public Const MAIN_FRAME_FILL_2 As Long = COL_WHITE
Public Const MAIN_FRAME_SHADOW As Long = msoShadow21

Public Const BUTTON_FRAME_BORDER_WIDTH As Single = 0
Public Const BUTTON_FRAME_FILL_1 As Long = COL_WHITE
Public Const BUTTON_FRAME_FILL_2 As Long = COL_WHITE
Public Const BUTTON_FRAME_SHADOW As Long = msoShadow21

Public Const HEADER_BORDER_WIDTH As Single = 0
Public Const HEADER_FILL_1 As Long = COL_PINK
Public Const HEADER_FILL_2 As Long = COL_PINK
Public Const HEADER_SHADOW As Long = 0
Public Const HEADER_FONT_STYLE As String = "Eras Medium ITC"
Public Const HEADER_FONT_SIZE As Integer = 12
Public Const HEADER_FONT_COLOUR As Long = COL_WHITE
Public Const HEADER_FONT_BOLD As Boolean = False
Public Const HEADER_FONT_X_JUST As Integer = xlHAlignCenter
Public Const HEADER_FONT_Y_JUST As Integer = xlVAlignCenter

Public Const BTN_MAIN_BORDER_WIDTH As Single = 0
Public Const BTN_MAIN_FILL_1 As Long = COL_BLACK
Public Const BTN_MAIN_FILL_2 As Long = COL_BLACK
Public Const BTN_MAIN_SHADOW As Long = msoShadow21
Public Const BTN_MAIN_FONT_STYLE As String = "Calibri"
Public Const BTN_MAIN_FONT_SIZE As Integer = 32
Public Const BTN_MAIN_FONT_COLOUR As Long = COL_WHITE
Public Const BTN_MAIN_FONT_BOLD As Boolean = True
Public Const BTN_MAIN_FONT_X_JUST As Integer = xlHAlignCenter
Public Const BTN_MAIN_FONT_Y_JUST As Integer = xlVAlignCenter

