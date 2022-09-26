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
Public Const MAIN_FRAME_TOP As Integer = 70
Public Const MAIN_FRAME_LEFT As Integer = 170
Public Const MAIN_FRAME_WIDTH As Integer = 890
Public Const MAIN_FRAME_HEIGHT As Integer = 360
Public Const MAIN_FRAME_2_TOP As Integer = 440
Public Const MAIN_FRAME_2_LEFT As Integer = 170
Public Const MAIN_FRAME_2_WIDTH As Integer = 890
Public Const MAIN_FRAME_2_HEIGHT As Integer = 360

' ---------------------------------------------------------------
' Generic Cell Settings
' ---------------------------------------------------------------
Public Const GENERIC_Cell_HEIGHT As Integer = 21
Public Const GENERIC_Cell_WIDTH As Integer = 550
Public Const GENERIC_Cell_LEFT As Integer = 0
Public Const GENERIC_Cell_TOP As Integer = 21
Public Const GENERIC_Cell_ROWOFFSET As Integer = 21

' ---------------------------------------------------------------
' Menu Bar
' ---------------------------------------------------------------
Public Const Button_HEIGHT As Integer = 31
Public Const MENUBAR_HEIGHT As Integer = 2000
Public Const MENUBAR_WIDTH As Integer = 155
Public Const MENUBAR_TOP As Integer = 0
Public Const MENUBAR_LEFT As Integer = 0
Public Const MENU_TOP As Integer = 150
Public Const MENU_LEFT As Integer = 10
Public Const Button_WIDTH As Integer = 150
Public Const Button_COUNT As Integer = 4
Public Const Button_TEXT = "For Action:Active:Complete:Exit"
'Public Const Button_ICONS = "TEMPLATE - For Action:TEMPLATE - Active:TEMPLATE - Complete:TEMPLATE - Dashboard:TEMPLATE - Reports:TEMPLATE - Admin:TEMPLATE - Exit"
'Public Const Button_BADGES = "TEMPLATE - No Action Items::::::"
'Public Const Button_ICON_TOP As Integer = 6
'Public Const Button_ICON_LEFT As Integer = 5
'Public Const Button_BADGE_TOP As Integer = -6
'Public Const Button_BADGE_LEFT As Integer = 125
Public Const LOGO_TOP As Integer = 13
Public Const LOGO_LEFT As Integer = 10
Public Const LOGO_WIDTH As Integer = 139
Public Const LOGO_HEIGHT As Integer = 90
'
Public Const HEADER_HEIGHT As Integer = 25
Public Const HEADER_ICON_TOP As Integer = 2
Public Const HEADER_ICON_RIGHT As Integer = 10

Public Const BTN_MAIN_1_TOP As Integer = 15
Public Const BTN_MAIN_1_LEFT As Integer = 960
Public Const BTN_MAIN_1_WIDTH As Integer = 100
Public Const BTN_MAIN_1_HEIGHT As Integer = 40

'Public Const BTN_MAIN_2_TOP As Integer = 143
'Public Const BTN_MAIN_2_LEFT As Integer = 1075
'Public Const BTN_MAIN_2_WIDTH As Integer = 80
'Public Const BTN_MAIN_2_HEIGHT As Integer = 20
'
'Public Const BTN_MAIN_3_TOP As Integer = 15
'Public Const BTN_MAIN_3_LEFT As Integer = 280
'Public Const BTN_MAIN_3_WIDTH As Integer = 100
'Public Const BTN_MAIN_3_HEIGHT As Integer = 40
'
'Public Const BTN_MAIN_4_TOP As Integer = 15
'Public Const BTN_MAIN_4_LEFT As Integer = 170
'Public Const BTN_MAIN_4_WIDTH As Integer = 100
'Public Const BTN_MAIN_4_HEIGHT As Integer = 40
'
'Public Const BTN_MAIN_5_TOP As Integer = 113
'Public Const BTN_MAIN_5_LEFT As Integer = 1075
'Public Const BTN_MAIN_5_WIDTH As Integer = 80
'Public Const BTN_MAIN_5_HEIGHT As Integer = 20
'
'Public Const BTN_SUPPORT_TOP As Integer = 500
'Public Const BTN_SUPPORT_LEFT As Integer = 20
'Public Const BTN_SUPPORT_WIDTH As Integer = 110
'Public Const BTN_SUPPORT_HEIGHT As Integer = 20

' ---------------------------------------------------------------
' For Action Screen
' ---------------------------------------------------------------
Public Const FOR_ACTION_Cell_NOCOLS As Integer = 5
Public Const FOR_ACTION_Cell_COL_WIDTHS As String = "100:200:200:290:100"
Public Const FOR_ACTION_Cell_TITLES As String = "Workflow No:Name:Step No:Step Name:Status"
Public Const FOR_ACTION_MAX_LINES As Integer = 40

' ---------------------------------------------------------------
' Active Screen
' ---------------------------------------------------------------
Public Const ACTIVE_Cell_NOCOLS As Integer = 5
Public Const ACTIVE_Cell_COL_WIDTHS As String = "100:200:200:290:100"
Public Const ACTIVE_Cell_TITLES As String = "Workflow No:Name:Step No:Step Name:Status"
Public Const ACTIVE_MAX_LINES As Integer = 150

' ---------------------------------------------------------------
' Complete Screen
' ---------------------------------------------------------------
Public Const COMPLETE_Cell_NOCOLS As Integer = 5
Public Const COMPLETE_Cell_COL_WIDTHS As String = "100:200:200:290:100"
Public Const COMPLETE_Cell_TITLES As String = "Workflow No:Name:Step No:Step Name:Status"
Public Const COMPLETE_MAX_LINES As Integer = 100

' ---------------------------------------------------------------
' Admin Screen
' ---------------------------------------------------------------
'Public Const BTN_ADMIN_1_TOP As Integer = 70
'Public Const BTN_ADMIN_1_LEFT As Integer = 1075
'Public Const BTN_ADMIN_1_WIDTH As Integer = 80
'Public Const BTN_ADMIN_1_HEIGHT As Integer = 20
'
'Public Const BTN_ADMIN_2_TOP As Integer = 100
'Public Const BTN_ADMIN_2_LEFT As Integer = 1075
'Public Const BTN_ADMIN_2_WIDTH As Integer = 80
'Public Const BTN_ADMIN_2_HEIGHT As Integer = 20
'
'Public Const BTN_ADMIN_3_TOP As Integer = 130
'Public Const BTN_ADMIN_3_LEFT As Integer = 1075
'Public Const BTN_ADMIN_3_WIDTH As Integer = 80
'Public Const BTN_ADMIN_3_HEIGHT As Integer = 20
'
'Public Const BTN_ADMIN_4_TOP As Integer = 160
'Public Const BTN_ADMIN_4_LEFT As Integer = 1075
'Public Const BTN_ADMIN_4_WIDTH As Integer = 80
'Public Const BTN_ADMIN_4_HEIGHT As Integer = 20
'
'Public Const BTN_ADMIN_5_TOP As Integer = 190
'Public Const BTN_ADMIN_5_LEFT As Integer = 1075
'Public Const BTN_ADMIN_5_WIDTH As Integer = 80
'Public Const BTN_ADMIN_5_HEIGHT As Integer = 20
'
'Public Const BTN_ADMIN_6_TOP As Integer = 220
'Public Const BTN_ADMIN_6_LEFT As Integer = 1075
'Public Const BTN_ADMIN_6_WIDTH As Integer = 80
'Public Const BTN_ADMIN_6_HEIGHT As Integer = 20
'
'Public Const BTN_ADMIN_7_TOP As Integer = 310
'Public Const BTN_ADMIN_7_LEFT As Integer = 1075
'Public Const BTN_ADMIN_7_WIDTH As Integer = 80
'Public Const BTN_ADMIN_7_HEIGHT As Integer = 20
'
'Public Const BTN_ADMIN_8_TOP As Integer = 280
'Public Const BTN_ADMIN_8_LEFT As Integer = 1075
'Public Const BTN_ADMIN_8_WIDTH As Integer = 80
'Public Const BTN_ADMIN_8_HEIGHT As Integer = 20
'
'Public Const BTN_ADMIN_9_TOP As Integer = 250
'Public Const BTN_ADMIN_9_LEFT As Integer = 1075
'Public Const BTN_ADMIN_9_WIDTH As Integer = 80
'Public Const BTN_ADMIN_9_HEIGHT As Integer = 20
'
'Public Const ADMIN_MEMBER_Cell_NOCOLS As Integer = 10
'Public Const ADMIN_MEMBER_Cell_COL_WIDTHS As String = "120:80:80:60:60:60:100:80:180:70"
'Public Const ADMIN_MEMBER_Cell_TITLES As String = "Name:SSN:Student ID:Watch:Grade:FIN No:Position:DoB:Email:Active"
'Public Const ADMIN_MEMBER_MAX_LINES As Integer = 500
'
'Public Const ADMIN_EMAIL_ADDR_Cell_NOCOLS As Integer = 3
'Public Const ADMIN_EMAIL_ADDR_Cell_COL_WIDTHS As String = "200:200:490"
'Public Const ADMIN_EMAIL_ADDR_Cell_TITLES As String = "Name:Email Address:"
'Public Const ADMIN_EMAIL_ADDR_MAX_LINES As Integer = 50
'
'Public Const ADMIN_EMAIL_Cell_NOCOLS As Integer = 5
'Public Const ADMIN_EMAIL_Cell_COL_WIDTHS As String = "80:150:200:300:160"
'Public Const ADMIN_EMAIL_Cell_TITLES As String = "Email No:Template Name:Mail To:Mail CC:Attachment"
'Public Const ADMIN_EMAIL_MAX_LINES As Integer = 100
'
'Public Const ADMIN_DOC_Cell_NOCOLS As Integer = 5
'Public Const ADMIN_DOC_Cell_COL_WIDTHS As String = "80:200:350:180:80"
'Public Const ADMIN_DOC_Cell_TITLES As String = "Doc No:Title:Description:Filename:"
'Public Const ADMIN_DOC_MAX_LINES As Integer = 100
'
'Public Const ADMIN_DOD_CERT_Cell_NOCOLS As Integer = 4
'Public Const ADMIN_DOD_CERT_Cell_COL_WIDTHS As String = "100:200:200:390"
'Public Const ADMIN_DOD_CERT_Cell_TITLES As String = "DoD Cert No:Name:Group:"
'Public Const ADMIN_DOD_CERT_MAX_LINES As Integer = 100
'
'Public Const ADMIN_WFLOW_Cell_NOCOLS As Integer = 6
'Public Const ADMIN_WFLOW_Cell_COL_WIDTHS As String = "100:300:190:100:100:100" '890
'Public Const ADMIN_WFLOW_Cell_TITLES As String = "Step No:Name:Step Type:Email No:Alt Email No:Document No"
'Public Const ADMIN_WFLOW_MAX_LINES As Integer = 100
'
'Public Const ADMIN_WFTYPES_Cell_NOCOLS As Integer = 3
'Public Const ADMIN_WFTYPES_Cell_COL_WIDTHS As String = "200:400:290" '890
'Public Const ADMIN_WFTYPES_Cell_TITLES As String = "Workflow Type:Description:"
'Public Const ADMIN_WFTYPES_MAX_LINES As Integer = 20
'
'Public Const ADMIN_ROLES_Cell_NOCOLS As Integer = 3
'Public Const ADMIN_ROLES_Cell_COL_WIDTHS As String = "200:200:490"
'Public Const ADMIN_ROLES_Cell_TITLES As String = "Role:Name:"
'Public Const ADMIN_ROLES_MAX_LINES As Integer = 20
'
'Public Const ADMIN_LISTS_Cell_NOCOLS As Integer = 2
'Public Const ADMIN_LISTS_Cell_COL_WIDTHS As String = "200:690"
'Public Const ADMIN_LISTS_Cell_TITLES As String = "List::"
'Public Const ADMIN_LISTS_MAX_LINES As Integer = 20
'
' ===============================================================
' Style Declarations
' ---------------------------------------------------------------
' Main Screen
' ---------------------------------------------------------------
Public SCREEN_STYLE As TypeStyle
Public MENUBAR_STYLE As TypeStyle
Public BUTTON_SET_STYLE As TypeStyle
Public BUTTON_UNSET_STYLE As TypeStyle
Public MAIN_FRAME_STYLE As TypeStyle
Public HEADER_STYLE As TypeStyle
Public BTN_MAIN_STYLE As TypeStyle
Public GENERIC_BUTTON As TypeStyle
'Public TOOL_BUTTON As TypeStyle
'Public BTN_SUPPORT As TypeStyle
Public GENERIC_Cell As TypeStyle
Public GREEN_Cell As TypeStyle
Public AMBER_Cell As TypeStyle
Public RED_Cell As TypeStyle
Public GENERIC_Cell_HEADER As TypeStyle
'Public TRANSPARENT_TEXT_BOX As TypeStyle
'Public VERT_Cell_HEADER As TypeStyle
'Public MATRIX_DEF As TypeStyle
'Public MATRIX_1 As TypeStyle
'Public MATRIX_3 As TypeStyle
'Public MATRIX_4 As TypeStyle
'
' ---------------------------------------------------------------
' New Order Workflow
' ---------------------------------------------------------------
'Public WF_MAINSCREEN_STYLE As TypeStyle

' ===============================================================
' Style Definitions
' ===============================================================
' Generic Styles
' ---------------------------------------------------------------
' Buttons
' ---------------------------------------------------------------
Public Const GENERIC_BUTTON_BORDER_WIDTH As Single = 0
Public Const GENERIC_BUTTON_FILL_1 As Long = COLOUR_3
Public Const GENERIC_BUTTON_FILL_2 As Long = COLOUR_3
Public Const GENERIC_BUTTON_SHADOW As Long = msoShadow21
Public Const GENERIC_BUTTON_FONT_STYLE As String = "Eras Medium ITC"
Public Const GENERIC_BUTTON_FONT_SIZE As Integer = 12
Public Const GENERIC_BUTTON_FONT_COLOUR As Long = COLOUR_5
Public Const GENERIC_BUTTON_FONT_BOLD As Boolean = False
Public Const GENERIC_BUTTON_FONT_X_JUST As Integer = xlHAlignCenter
Public Const GENERIC_BUTTON_FONT_Y_JUST As Integer = xlVAlignCenter

' ---------------------------------------------------------------
' Cells
' ---------------------------------------------------------------
Public Const GENERIC_Cell_BORDER_WIDTH As Single = 0
Public Const GENERIC_Cell_FILL_1 As Long = COLOUR_5
Public Const GENERIC_Cell_FILL_2 As Long = COLOUR_5
Public Const GENERIC_Cell_SHADOW As Long = 0
Public Const GENERIC_Cell_FONT_STYLE As String = "Eras Medium ITC"
Public Const GENERIC_Cell_FONT_SIZE As Integer = 10
Public Const GENERIC_Cell_FONT_COLOUR As Long = COLOUR_4
Public Const GENERIC_Cell_FONT_BOLD As Boolean = False
Public Const GENERIC_Cell_FONT_X_JUST As Integer = xlHAlignCenter
Public Const GENERIC_Cell_FONT_Y_JUST As Integer = xlVAlignBottom

' ---------------------------------------------------------------
' Cell Headers
' ---------------------------------------------------------------
Public Const GENERIC_Cell_HEADER_BORDER_WIDTH As Single = 0
Public Const GENERIC_Cell_HEADER_FILL_1 As Long = COLOUR_3
Public Const GENERIC_Cell_HEADER_FILL_2 As Long = COLOUR_3
Public Const GENERIC_Cell_HEADER_SHADOW As Long = 0
Public Const GENERIC_Cell_HEADER_FONT_STYLE As String = "Calibri"
Public Const GENERIC_Cell_HEADER_FONT_SIZE As Integer = 10
Public Const GENERIC_Cell_HEADER_FONT_COLOUR As Long = COLOUR_5
Public Const GENERIC_Cell_HEADER_FONT_BOLD As Boolean = False
Public Const GENERIC_Cell_HEADER_FONT_X_JUST As Integer = xlHAlignCenter
Public Const GENERIC_Cell_HEADER_FONT_Y_JUST As Integer = xlVAlignBottom

' ---------------------------------------------------------------
' Text Box
' ---------------------------------------------------------------
Public Const TRANSPARENT_TEXT_BOX_BORDER_WIDTH As Single = 0
Public Const TRANSPARENT_TEXT_BOX_FILL_1 As Long = COLOUR_3
Public Const TRANSPARENT_TEXT_BOX_FILL_2 As Long = COLOUR_3
Public Const TRANSPARENT_TEXT_BOX_SHADOW As Long = 0
Public Const TRANSPARENT_TEXT_BOX_FONT_STYLE As String = "Eras Medium ITC"
Public Const TRANSPARENT_TEXT_BOX_FONT_SIZE As Integer = 10
Public Const TRANSPARENT_TEXT_BOX_FONT_COLOUR As Long = COLOUR_2
Public Const TRANSPARENT_TEXT_BOX_FONT_BOLD As Boolean = False
Public Const TRANSPARENT_TEXT_BOX_FONT_X_JUST As Integer = xlHAlignLeft
Public Const TRANSPARENT_TEXT_BOX_FONT_Y_JUST As Integer = xlVAlignCenter

' ===============================================================
' Special Styles
' ===============================================================
' Cell Green
' ---------------------------------------------------------------
Public Const GREEN_Cell_BORDER_WIDTH As Single = 0
Public Const GREEN_Cell_FILL_1 As Long = COLOUR_9
Public Const GREEN_Cell_FILL_2 As Long = COLOUR_9
Public Const GREEN_Cell_SHADOW As Long = 0
Public Const GREEN_Cell_FONT_STYLE As String = "Eras Medium ITC"
Public Const GREEN_Cell_FONT_SIZE As Integer = 11
Public Const GREEN_Cell_FONT_COLOUR As Long = COLOUR_4
Public Const GREEN_Cell_FONT_BOLD As Boolean = False
Public Const GREEN_Cell_FONT_X_JUST As Integer = xlHAlignCenter
Public Const GREEN_Cell_FONT_Y_JUST As Integer = xlVAlignCenter

' ---------------------------------------------------------------
' Cell Amber
' ---------------------------------------------------------------
Public Const AMBER_Cell_BORDER_WIDTH As Single = 0
Public Const AMBER_Cell_FILL_1 As Long = COLOUR_8
Public Const AMBER_Cell_FILL_2 As Long = COLOUR_8
Public Const AMBER_Cell_SHADOW As Long = 0
Public Const AMBER_Cell_FONT_STYLE As String = "Eras Medium ITC"
Public Const AMBER_Cell_FONT_SIZE As Integer = 11
Public Const AMBER_Cell_FONT_COLOUR As Long = COLOUR_4
Public Const AMBER_Cell_FONT_BOLD As Boolean = False
Public Const AMBER_Cell_FONT_X_JUST As Integer = xlHAlignCenter
Public Const AMBER_Cell_FONT_Y_JUST As Integer = xlVAlignCenter

' ---------------------------------------------------------------
' Cell Red
' ---------------------------------------------------------------
Public Const RED_Cell_BORDER_WIDTH As Single = 0
Public Const RED_Cell_FILL_1 As Long = COLOUR_10
Public Const RED_Cell_FILL_2 As Long = COLOUR_10
Public Const RED_Cell_SHADOW As Long = 0
Public Const RED_Cell_FONT_STYLE As String = "Eras Medium ITC"
Public Const RED_Cell_FONT_SIZE As Integer = 11
Public Const RED_Cell_FONT_COLOUR As Long = COLOUR_5
Public Const RED_Cell_FONT_BOLD As Boolean = False
Public Const RED_Cell_FONT_X_JUST As Integer = xlHAlignCenter
Public Const RED_Cell_FONT_Y_JUST As Integer = xlVAlignCenter

' ---------------------------------------------------------------
' Tool Buttons
' ---------------------------------------------------------------
Public Const TOOL_BUTTON_BORDER_WIDTH As Single = 0
Public Const TOOL_BUTTON_FILL_1 As Long = COLOUR_5
Public Const TOOL_BUTTON_FILL_2 As Long = COLOUR_5
Public Const TOOL_BUTTON_SHADOW As Long = msoShadow21
Public Const TOOL_BUTTON_FONT_STYLE As String = "Eras Medium ITC"
Public Const TOOL_BUTTON_FONT_SIZE As Integer = 9
Public Const TOOL_BUTTON_FONT_COLOUR As Long = COLOUR_8
Public Const TOOL_BUTTON_FONT_BOLD As Boolean = False
Public Const TOOL_BUTTON_FONT_X_JUST As Integer = xlHAlignCenter
Public Const TOOL_BUTTON_FONT_Y_JUST As Integer = xlVAlignCenter

' ---------------------------------------------------------------
' Support Button
' ---------------------------------------------------------------
Public Const BTN_SUPPORT_BORDER_WIDTH As Single = 0.5
Public Const BTN_SUPPORT_BORDER_COLOUR As Long = COLOUR_9
Public Const BTN_SUPPORT_FILL_1 As Long = COLOUR_5
Public Const BTN_SUPPORT_FILL_2 As Long = COLOUR_5
Public Const BTN_SUPPORT_SHADOW As Long = 0
Public Const BTN_SUPPORT_FONT_STYLE As String = "Eras Medium ITC"
Public Const BTN_SUPPORT_FONT_SIZE As Integer = 9
Public Const BTN_SUPPORT_FONT_COLOUR As Long = COLOUR_9
Public Const BTN_SUPPORT_FONT_BOLD As Boolean = False
Public Const BTN_SUPPORT_FONT_X_JUST As Integer = xlHAlignCenter
Public Const BTN_SUPPORT_FONT_Y_JUST As Integer = xlVAlignCenter

' ===============================================================
' Main Screen
' ===============================================================
Public Const SCREEN_BORDER_WIDTH As Single = 0
Public Const SCREEN_FILL_1 As Long = COLOUR_7
Public Const SCREEN_FILL_2 As Long = COLOUR_7
Public Const SCREEN_SHADOW As Long = 0

Public Const MENUBAR_BORDER_WIDTH As Single = 1
Public Const MENUBAR_FILL_1 As Long = COLOUR_5
Public Const MENUBAR_FILL_2 As Long = COLOUR_5
Public Const MENUBAR_SHADOW As Long = msoShadow21

Public Const BUTTON_UNSET_BORDER_WIDTH As Single = 1
Public Const BUTTON_UNSET_BORDER_COLOUR As Long = COLOUR_3
Public Const BUTTON_UNSET_FILL_1 As Long = COLOUR_5
Public Const BUTTON_UNSET_FILL_2 As Long = COLOUR_5
Public Const BUTTON_UNSET_SHADOW As Long = 0
Public Const BUTTON_UNSET_FONT_STYLE As String = "Eras Medium ITC"
Public Const BUTTON_UNSET_FONT_SIZE As Integer = 12
Public Const BUTTON_UNSET_FONT_COLOUR As Long = COLOUR_6
Public Const BUTTON_UNSET_FONT_X_JUST As Integer = xlHAlignCenter
Public Const BUTTON_UNSET_FONT_Y_JUST As Integer = xlVAlignCenter

Public Const BUTTON_SET_BORDER_WIDTH As Single = 0
Public Const BUTTON_SET_BORDER_COLOUR As Long = COLOUR_3
Public Const BUTTON_SET_FILL_1 As Long = COLOUR_2
Public Const BUTTON_SET_FILL_2 As Long = COLOUR_2
Public Const BUTTON_SET_SHADOW As Long = 0
Public Const BUTTON_SET_FONT_STYLE As String = "Eras Medium ITC"
Public Const BUTTON_SET_FONT_SIZE As Integer = 12
Public Const BUTTON_SET_FONT_COLOUR As Long = COLOUR_5
Public Const BUTTON_SET_FONT_X_JUST As Integer = xlHAlignCenter
Public Const BUTTON_SET_FONT_Y_JUST As Integer = xlVAlignCenter

Public Const MAIN_FRAME_BORDER_WIDTH As Single = 0
Public Const MAIN_FRAME_FILL_1 As Long = COLOUR_5
Public Const MAIN_FRAME_FILL_2 As Long = COLOUR_5
Public Const MAIN_FRAME_SHADOW As Long = msoShadow21

Public Const HEADER_BORDER_WIDTH As Single = 0
Public Const HEADER_FILL_1 As Long = COLOUR_2
Public Const HEADER_FILL_2 As Long = COLOUR_2
Public Const HEADER_SHADOW As Long = 0
Public Const HEADER_FONT_STYLE As String = "Eras Medium ITC"
Public Const HEADER_FONT_SIZE As Integer = 12
Public Const HEADER_FONT_COLOUR As Long = COLOUR_5
Public Const HEADER_FONT_BOLD As Boolean = False
Public Const HEADER_FONT_X_JUST As Integer = xlHAlignCenter
Public Const HEADER_FONT_Y_JUST As Integer = xlVAlignCenter

Public Const BTN_MAIN_BORDER_WIDTH As Single = 0
Public Const BTN_MAIN_FILL_1 As Long = COLOUR_6
Public Const BTN_MAIN_FILL_2 As Long = COLOUR_6
Public Const BTN_MAIN_SHADOW As Long = msoShadow21
Public Const BTN_MAIN_FONT_STYLE As String = "Calibri"
Public Const BTN_MAIN_FONT_SIZE As Integer = 32
Public Const BTN_MAIN_FONT_COLOUR As Long = COLOUR_5
Public Const BTN_MAIN_FONT_BOLD As Boolean = True
Public Const BTN_MAIN_FONT_X_JUST As Integer = xlHAlignCenter
Public Const BTN_MAIN_FONT_Y_JUST As Integer = xlVAlignCenter

