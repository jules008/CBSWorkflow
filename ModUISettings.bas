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
Public Const SCREEN_WIDTH As Integer = 1805
Public Const MAIN_FRAME_TOP As Integer = 90
Public Const MAIN_FRAME_LEFT As Integer = 170
Public Const MAIN_FRAME_WIDTH As Integer = 1080
Public Const MAIN_FRAME_HEIGHT As Integer = 360

' ---------------------------------------------------------------
' Main Frame 2
' ---------------------------------------------------------------
Public Const MAIN_FRAME_2_TOP As Integer = 300
Public Const MAIN_FRAME_2_LEFT As Integer = 170
Public Const MAIN_FRAME_2_WIDTH As Integer = 1000
Public Const MAIN_FRAME_2_HEIGHT As Integer = 360

' ---------------------------------------------------------------
' Button Frame
' ---------------------------------------------------------------
Public Const BUTTON_FRAME_TOP As Integer = 5
Public Const BUTTON_FRAME_LEFT As Integer = 240
Public Const BUTTON_FRAME_WIDTH As Integer = 1100
Public Const BUTTON_FRAME_HEIGHT As Integer = 60

' ---------------------------------------------------------------
' Generic Button
' ---------------------------------------------------------------
Public Const GENERIC_BUTTON_WIDTH As Integer = 100
Public Const GENERIC_BUTTON_HEIGHT As Integer = 40

' ---------------------------------------------------------------
' ToDo Button
' ---------------------------------------------------------------
Public Const TODO_BUTTON_WIDTH As Integer = 140
Public Const TODO_BUTTON_LEFT As Integer = 0
Public Const TODO_BUTTON_TOP As Integer = 25
Public Const TODO_ICON_FILE As String = "ToDo.png"
Public Const TODO_ICON_WIDTH As String = 18
Public Const TODO_ICON_TOP As String = 11
Public Const TODO_ICON_LEFT As String = 5
Public Const TODO_BADGE_FILE As String = "ToDo.png"
Public Const TODO_BADGE_WIDTH As String = 25
Public Const TODO_BADGE_HEIGHT As String = 18
Public Const TODO_BADGE_TOP As String = 11
Public Const TODO_BADGE_LEFT As String = 107

' ---------------------------------------------------------------
' Generic Table Settings
' ---------------------------------------------------------------
Public Const GENERIC_TABLE_ROW_HEIGHT As Integer = 20
Public Const GENERIC_TABLE_WIDTH As Integer = 900
Public Const GENERIC_TABLE_LEFT As Integer = 0
Public Const GENERIC_TABLE_TOP As Integer = 25
Public Const GENERIC_TABLE_ROWOFFSET As Integer = 0
Public Const GENERIC_TABLE_COLOFFSET As Integer = 0
Public Const GENERIC_TABLE_HEADING_HEIGHT As Integer = 20
Public Const GENERIC_TABLE_EXPAND_ICON As String = "Expand.png"


' ---------------------------------------------------------------
' Table Lender Icon Settings
' ---------------------------------------------------------------
Public Const TABLE_ICON_TICK As String = "Tick.png"
Public Const TABLE_ICON_EXCLAM As String = "Exclamation.png"
Public Const TABLE_ICON_CROSS As String = "Cross.png"
Public Const TABLE_ICON_OFFSET As Integer = 15
Public Const TABLE_ICON_SIZE As Integer = 13
Public Const TABLE_ICON_TOP As Integer = 3
Public Const TABLE_ICON_LEFT As Integer = 10
Public Const TABLE_ICON_COL As Integer = 2
Public Const TABLE_ICON_NO_ICONS As Integer = 4

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
Public Const BUTTON_TEXT = "For Action:Projects:Active:Closed:CRM:Clients:SPVs:Contacts:Projects:Lenders:Dashboard:Reports:Admin:Users:Email Templates:Workflows:Workflow Types:Exit"
Public Const BUTTON_INDEX = "1:2:2.1:2.2:3:3.1:3.2:3.3:3.4:3.5:4:5:6:6.1:6.2:6.3:6.4:7"
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
Public Const PROJECT_TABLE_COL_WIDTHS As String = "50:50:80:100:100:100:100:50:220:130:100"
Public Const PROJECT_CELL_ROW_HEIGHTS As String = "50:50"
Public Const PROJECT_TABLE_STYLES As String = "GENERIC_TABLE"
Public Const PROJECT_TABLE_TITLES As String = "Expand:Project No:Lender Status:Project Name:Client Name:SPV Name:Case Manager:Step No:Step Name:Progress:Status"
Public Const PROJECT_MAX_LINES As Integer = 150
Public Const PROJECT_BTN_MAIN_1_LEFT As Integer = 790
Public Const PROJECT_BTN_MAIN_1_TOP As Integer = 20
Public Const PROJECT_BTN_MAIN_2_LEFT As Integer = 900
Public Const PROJECT_BTN_MAIN_2_TOP As Integer = 20

' ---------------------------------------------------------------
' Project Sub Table Settings
' ---------------------------------------------------------------
Public Const PROJECT_SUB_TABLE_COL_WIDTHS As String = "100:150:150:100:220:130:100"
Public Const PROJECT_SUB_TABLE_TITLES As String = "Workflow No:Workflow Type:Lender Name:Step No:Step Name:Progress:Status"

' ---------------------------------------------------------------
' Project Table Progress Bar Settings
' ---------------------------------------------------------------
Public Const TABLE_PROGRESS_BADGE_LEFT As Integer = 0
Public Const TABLE_PROGRESS_BADGE_TOP As Integer = 5
Public Const TABLE_PROGRESS_BADGE_MARGIN_TOP As Integer = 4
Public Const TABLE_PROGRESS_FONT_SIZE As Integer = 10
Public Const TABLE_PROGRESS_FONT_COLOUR As Long = COL_WHITE
Public Const TABLE_PROGRESS_BORDER_WEIGHT As Integer = 0
Public Const TABLE_PROGRESS_BADGE_HEIGHT As Integer = 10
Public Const TABLE_PROGRESS_BADGE_FILL_COLOUR As Long = COL_AQUA
Public Const TABLE_PROGRESS_CELL_X_JUST As Integer = xlHAlignRight
Public Const TABLE_PROGRESS_CELL_Y_JUST As Integer = xlVAlignCenter

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
Public Const CRM_CELL_ROW_HEIGHTS As String = "50:50"

' ---------------------------------------------------------------
' CRM Screen Btn 1 Settings
' ---------------------------------------------------------------
Public Const CRM_BTN_MAIN_1_LEFT As Integer = 900
Public Const CRM_BTN_MAIN_1_TOP As Integer = 20

' ---------------------------------------------------------------
' CRM Screen Btn 2 Settings
' ---------------------------------------------------------------
Public Const CRM_BTN_MAIN_2_LEFT As Integer = 790
Public Const CRM_BTN_MAIN_2_TOP As Integer = 20

' ---------------------------------------------------------------
' CRM Screen Btn 3 Settings
' ---------------------------------------------------------------
Public Const CRM_BTN_MAIN_3_LEFT As Integer = 680
Public Const CRM_BTN_MAIN_3_TOP As Integer = 20

' ---------------------------------------------------------------
' CRM Clients Screen
' ---------------------------------------------------------------
Public Const CRM_CLIENT_TABLE_COL_WIDTHS As String = "100:150:150:300:150:150:80"
Public Const CRM_CLIENT_TABLE_TITLES As String = "Client No:Client Name:CBS/HP:Address:Phone No:Url:"
Public Const CRM_CLIENT_MAX_LINES As Integer = 150

' ---------------------------------------------------------------
' CRM SPV Screen
' ---------------------------------------------------------------
Public Const CRM_SPV_TABLE_COL_WIDTHS As String = "100:300:300:300:80"
Public Const CRM_SPV_TABLE_TITLES As String = "SPV No:SPV Name: : :"
Public Const CRM_SPV_MAX_LINES As Integer = 150

' ---------------------------------------------------------------
' CRM Contact Screen
' ---------------------------------------------------------------
Public Const CRM_CONTACT_TABLE_COL_WIDTHS As String = "50:150:100:200:200:100:200:80"
Public Const CRM_CONTACT_TABLE_TITLES As String = "Contact No:Contact Name:Contact Type:Organisation:Position:Phone No:Email Address:"
Public Const CRM_CONTACT_MAX_LINES As Integer = 150

' ---------------------------------------------------------------
' CRM Projects Screen
' ---------------------------------------------------------------
Public Const CRM_PROJECT_TABLE_COL_WIDTHS As String = "100:150:150:150:100:100:250:80"
Public Const CRM_PROJECT_TABLE_TITLES As String = "Project No:Project Name:Client Name:SPV:Case Manager:::"
Public Const CRM_PROJECT_MAX_LINES As Integer = 150

' ---------------------------------------------------------------
' CRM Lenders Screen
' ---------------------------------------------------------------
Public Const CRM_LENDER_TABLE_COL_WIDTHS As String = "100:200:150:150:250:150:80"
Public Const CRM_LENDER_TABLE_TITLES As String = "Lender No:Name:Phone No:Lender Type:Address::"
Public Const CRM_LENDER_MAX_LINES As Integer = 150

'===============================================================
' Dashboard Screen
'===============================================================
Public Const GRAPH_1_TOP As Integer = 45
Public Const GRAPH_1_LEFT As Integer = 20
Public Const GRAPH_1_HEIGHT As Integer = 120
Public Const GRAPH_1_COL_1 As Long = COL_GREEN
Public Const GRAPH_1_COL_2 As Long = COL_PINK
Public Const GRAPH_1_BACK_COL As Long = COL_WHITE

'===============================================================
' Report Screen
'===============================================================
' Frame 1
' ---------------------------------------------------------------

Public Const REP_FRAME_TOP As Integer = 90
Public Const REP_FRAME_LEFT As Integer = 170
Public Const REP_FRAME_WIDTH As Integer = 1000
Public Const REP_FRAME_HEIGHT As Integer = 270

' ---------------------------------------------------------------
' Frame 2
' ---------------------------------------------------------------
Public Const REP_FRAME_2_TOP As Integer = 400
Public Const REP_FRAME_2_LEFT As Integer = 170
Public Const REP_FRAME_2_WIDTH As Integer = 1000
Public Const REP_FRAME_2_HEIGHT As Integer = 270

' ---------------------------------------------------------------
' Buttons
' ---------------------------------------------------------------
Public Const BTN_REP_1_HEIGHT As Integer = 70
Public Const BTN_REP_1_LEFT As Integer = 200
Public Const BTN_REP_1_TOP As Integer = 140
Public Const BTN_REP_1_WIDTH As Integer = 150

Public Const BTN_REP_2_HEIGHT As Integer = 70
Public Const BTN_REP_2_LEFT As Integer = 380
Public Const BTN_REP_2_TOP As Integer = 140
Public Const BTN_REP_2_WIDTH As Integer = 150

Public Const BTN_REP_3_HEIGHT As Integer = 70
Public Const BTN_REP_3_LEFT As Integer = 560
Public Const BTN_REP_3_TOP As Integer = 140
Public Const BTN_REP_3_WIDTH As Integer = 150

Public Const BTN_REP_4_HEIGHT As Integer = 70
Public Const BTN_REP_4_LEFT As Integer = 740
Public Const BTN_REP_4_TOP As Integer = 140
Public Const BTN_REP_4_WIDTH As Integer = 150

Public Const BTN_EXP_5_HEIGHT As Integer = 70
Public Const BTN_EXP_5_LEFT As Integer = 200
Public Const BTN_EXP_5_TOP As Integer = 450
Public Const BTN_EXP_5_WIDTH As Integer = 150

Public Const BTN_EXP_6_HEIGHT As Integer = 70
Public Const BTN_EXP_6_LEFT As Integer = 380
Public Const BTN_EXP_6_TOP As Integer = 450
Public Const BTN_EXP_6_WIDTH As Integer = 150

Public Const BTN_EXP_7_HEIGHT As Integer = 70
Public Const BTN_EXP_7_LEFT As Integer = 560
Public Const BTN_EXP_7_TOP As Integer = 450
Public Const BTN_EXP_7_WIDTH As Integer = 150

Public Const BTN_EXP_8_HEIGHT As Integer = 70
Public Const BTN_EXP_8_LEFT As Integer = 740
Public Const BTN_EXP_8_TOP As Integer = 450
Public Const BTN_EXP_8_WIDTH As Integer = 150

'===============================================================
' Admin Screen
'===============================================================
' Admin Users Screen
' ---------------------------------------------------------------
Public Const ADM_USERS_TABLE_COL_WIDTHS As String = "50:200:200:200:200:230"
Public Const ADM_USERS_TABLE_TITLES As String = "User No:User Name:User Level:Position:Phone No:Supervisor"
Public Const ADM_USERS_MAX_LINES As Integer = 150

' ---------------------------------------------------------------
' Admin Emails Screen
' ---------------------------------------------------------------
Public Const ADM_EMAILS_TABLE_COL_WIDTHS As String = "100:250:250:300:180"
Public Const ADM_EMAILS_TABLE_TITLES As String = "Email No:Template Name:To:Subject:"
Public Const ADM_EMAILS_MAX_LINES As Integer = 150

' ---------------------------------------------------------------
' Admin Workflows Screen
' ---------------------------------------------------------------
Public Const ADM_WFLOWS_TABLE_COL_WIDTHS As String = "100:250:200:300:230"
Public Const ADM_WFLOWS_TABLE_TITLES As String = "Workflow No:Name:Step No:Step Name:"
Public Const ADM_WFLOWS_MAX_LINES As Integer = 150

' ---------------------------------------------------------------
' Admin WorkflowTypes Screen
' ---------------------------------------------------------------
Public Const ADM_WFTYPES_TABLE_COL_WIDTHS As String = "100:250:300:430"
Public Const ADM_WFTYPES_TABLE_TITLES As String = "No:Loan Type:Second Tier:"
Public Const ADM_WFTYPES_MAX_LINES As Integer = 150

' ---------------------------------------------------------------
' Admin Screen Btn 1 Settings
' ---------------------------------------------------------------
Public Const ADM_BTN_MAIN_1_LEFT As Integer = 900
Public Const ADM_BTN_MAIN_1_TOP As Integer = 20

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
Public TODO_BUTTON As ClsUIStyle
Public TODO_BADGE As ClsUIStyle
Public GENERIC_TABLE As ClsUIStyle
Public GREEN_CELL As ClsUIStyle
Public AMBER_CELL As ClsUIStyle
Public RED_CELL As ClsUIStyle
Public GENERIC_TABLE_HEADER As ClsUIStyle
Public SUB_TABLE_HEADER As ClsUIStyle
Public TABLE_PROGRESS_STYLE As ClsUIStyle
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
Public Const GENERIC_BUTTON_FONT_x_JUST As Integer = xlHAlignCenter
Public Const GENERIC_BUTTON_FONT_Y_JUST As Integer = xlVAlignCenter

' ---------------------------------------------------------------
' Todo Button
' ---------------------------------------------------------------
Public Const TODO_BUTTON_BORDER_WIDTH As Single = 0
Public Const TODO_BUTTON_FILL_1 As Long = COL_AQUA
Public Const TODO_BUTTON_FILL_2 As Long = COL_AQUA
Public Const TODO_BUTTON_SHADOW As Long = msoShadow21
Public Const TODO_BUTTON_FONT_STYLE As String = "Eras Medium ITC"
Public Const TODO_BUTTON_FONT_SIZE As Integer = 12
Public Const TODO_BUTTON_FONT_COLOUR As Long = COL_WHITE
Public Const TODO_BUTTON_FONT_BOLD As Boolean = False
Public Const TODO_BUTTON_FONT_x_JUST As Integer = xlHAlignCenter
Public Const TODO_BUTTON_FONT_Y_JUST As Integer = xlVAlignCenter

' ---------------------------------------------------------------
' Todo Badge
' ---------------------------------------------------------------
Public Const TODO_BADGE_BORDER_WIDTH As Single = 1
Public Const TODO_BADGE_BORDER_COLOUR As Long = COL_WHITE
Public Const TODO_BADGE_FILL_1 As Long = COL_AQUA
Public Const TODO_BADGE_FILL_2 As Long = COL_AQUA
Public Const TODO_BADGE_SHADOW As Long = 0
Public Const TODO_BADGE_FONT_STYLE As String = "Eras Medium ITC"
Public Const TODO_BADGE_FONT_SIZE As Integer = 12
Public Const TODO_BADGE_FONT_COLOUR As Long = COL_WHITE
Public Const TODO_BADGE_FONT_BOLD As Boolean = False
Public Const TODO_BADGE_FONT_x_JUST As Integer = xlHAlignRight
Public Const TODO_BADGE_FONT_Y_JUST As Integer = xlVAlignCenter

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
Public Const GENERIC_TABLE_FONT_x_JUST As Integer = xlHAlignCenter
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
Public Const GENERIC_TABLE_HEADER_FONT_x_JUST As Integer = xlHAlignCenter
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
Public Const TRANSPARENT_TEXT_BOX_FONT_x_JUST As Integer = xlHAlignLeft
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
Public Const GREEN_CELL_FONT_x_JUST As Integer = xlHAlignCenter
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
Public Const AMBER_CELL_FONT_x_JUST As Integer = xlHAlignCenter
Public Const AMBER_CELL_FONT_Y_JUST As Integer = xlVAlignCenter

' ---------------------------------------------------------------
' Cell Red
' ---------------------------------------------------------------
Public Const RED_CELL_BORDER_WIDTH As Single = 0
Public Const RED_CELL_FILL_1 As Long = COL_RED
Public Const RED_CELL_FILL_2 As Long = COL_LT_RED
Public Const RED_CELL_SHADOW As Long = 0
Public Const RED_CELL_FONT_STYLE As String = "Eras Medium ITC"
Public Const RED_CELL_FONT_SIZE As Integer = 11
Public Const RED_CELL_FONT_COLOUR As Long = COL_WHITE
Public Const RED_CELL_FONT_BOLD As Boolean = False
Public Const RED_CELL_FONT_x_JUST As Integer = xlHAlignCenter
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
Public Const TOOL_BUTTON_FONT_x_JUST As Integer = xlHAlignCenter
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
Public Const BTN_SUPPORT_FONT_x_JUST As Integer = xlHAlignCenter
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
Public Const BUTTON_UNSET_FONT_x_JUST As Integer = xlHAlignCenter
Public Const BUTTON_UNSET_FONT_Y_JUST As Integer = xlVAlignCenter

Public Const BUTTON_SET_BORDER_WIDTH As Single = 0
Public Const BUTTON_SET_BORDER_COLOUR As Long = COL_BLUE
Public Const BUTTON_SET_FILL_1 As Long = COL_PINK
Public Const BUTTON_SET_FILL_2 As Long = COL_PINK
Public Const BUTTON_SET_SHADOW As Long = 0
Public Const BUTTON_SET_FONT_STYLE As String = "Eras Medium ITC"
Public Const BUTTON_SET_FONT_SIZE As Integer = 12
Public Const BUTTON_SET_FONT_COLOUR As Long = COL_WHITE
Public Const BUTTON_SET_FONT_x_JUST As Integer = xlHAlignCenter
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
Public Const HEADER_FONT_x_JUST As Integer = xlHAlignCenter
Public Const HEADER_FONT_Y_JUST As Integer = xlVAlignCenter

Public Const BTN_MAIN_BORDER_WIDTH As Single = 0
Public Const BTN_MAIN_FILL_1 As Long = COL_BLUE
Public Const BTN_MAIN_FILL_2 As Long = COL_BLUE
Public Const BTN_MAIN_SHADOW As Long = msoShadow21
Public Const BTN_MAIN_FONT_STYLE As String = "Calibri"
Public Const BTN_MAIN_FONT_SIZE As Integer = 10
Public Const BTN_MAIN_FONT_COLOUR As Long = COL_WHITE
Public Const BTN_MAIN_FONT_BOLD As Boolean = False
Public Const BTN_MAIN_FONT_x_JUST As Integer = xlHAlignCenter
Public Const BTN_MAIN_FONT_Y_JUST As Integer = xlVAlignCenter

