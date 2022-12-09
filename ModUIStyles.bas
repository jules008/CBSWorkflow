Attribute VB_Name = "ModUIStyles"
'===============================================================
' Module ModUIStyles
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 16 Nov 20
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModUIStyles"

' ===============================================================
' BuildScreenStyles
' Builds the UI styles for use on the screen
' ---------------------------------------------------------------
Public Function BuildScreenStyles() As Boolean
    Const StrPROCEDURE As String = "BuildScreenStyles()"

    Set SCREEN_STYLE = New ClsUIStyle
    Set MENUBAR_STYLE = New ClsUIStyle
    Set BUTTON_UNSET_STYLE = New ClsUIStyle
    Set BUTTON_SET_STYLE = New ClsUIStyle
    Set MAIN_FRAME_STYLE = New ClsUIStyle
    Set BUTTON_FRAME_STYLE = New ClsUIStyle
    Set BTN_MAIN_STYLE = New ClsUIStyle
    Set GENERIC_BUTTON = New ClsUIStyle
    Set TODO_BUTTON = New ClsUIStyle
    Set TODO_BADGE = New ClsUIStyle
    Set HEADER_STYLE = New ClsUIStyle
    Set GENERIC_TABLE = New ClsUIStyle
    Set GREEN_CELL = New ClsUIStyle
    Set AMBER_CELL = New ClsUIStyle
    Set RED_CELL = New ClsUIStyle
    Set GENERIC_TABLE_HEADER = New ClsUIStyle
    Set SUB_TABLE_HEADER = New ClsUIStyle
    Set TABLE_PROGRESS_STYLE = New ClsUIStyle

    With SCREEN_STYLE
        .Name = "SCREEN_STYLE"
        .BorderWidth = SCREEN_BORDER_WIDTH
        .Fill1 = SCREEN_FILL_1
        .Fill2 = SCREEN_FILL_2
        .Shadow = SCREEN_SHADOW
    End With

    With MENUBAR_STYLE
        .Name = "MENUBAR_STYLE"
        .BorderWidth = MENUBAR_BORDER_WIDTH
        .Fill1 = MENUBAR_FILL_1
        .Fill2 = MENUBAR_FILL_2
        .Shadow = MENUBAR_SHADOW
    End With

    With BUTTON_UNSET_STYLE
        .Name = "BUTTON_UNSET_STYLE"
        .BorderWidth = BUTTON_UNSET_BORDER_WIDTH
        .BorderColour = BUTTON_UNSET_BORDER_COLOUR
        .Fill1 = BUTTON_UNSET_FILL_1
        .Fill2 = BUTTON_UNSET_FILL_2
        .Shadow = BUTTON_UNSET_SHADOW
        .FontStyle = BUTTON_UNSET_FONT_STYLE
        .FontSize = BUTTON_UNSET_FONT_SIZE
        .FontColour = BUTTON_UNSET_FONT_COLOUR
        .FontXJust = BUTTON_UNSET_FONT_x_JUST
        .FontVJust = BUTTON_UNSET_FONT_Y_JUST
    End With

    With BUTTON_SET_STYLE
        .Name = "BUTTON_SET_STYLE"
        .BorderWidth = BUTTON_SET_BORDER_WIDTH
        .BorderColour = BUTTON_SET_BORDER_COLOUR
        .Fill1 = BUTTON_SET_FILL_1
        .Fill2 = BUTTON_SET_FILL_2
        .Shadow = BUTTON_SET_SHADOW
        .FontStyle = BUTTON_SET_FONT_STYLE
        .FontSize = BUTTON_SET_FONT_SIZE
        .FontColour = BUTTON_SET_FONT_COLOUR
        .FontXJust = BUTTON_SET_FONT_x_JUST
        .FontVJust = BUTTON_SET_FONT_Y_JUST
    End With
    
    With MAIN_FRAME_STYLE
        .Name = "MAIN_FRAME_STYLE"
        .BorderWidth = MAIN_FRAME_BORDER_WIDTH
        .Fill1 = MAIN_FRAME_FILL_1
        .Fill2 = MAIN_FRAME_FILL_2
        .Shadow = MAIN_FRAME_SHADOW
    End With
    
    With BUTTON_FRAME_STYLE
        .Name = "BUTTON_FRAME_STYLE"
        .BorderWidth = BUTTON_FRAME_BORDER_WIDTH
        .Fill1 = BUTTON_FRAME_FILL_1
        .Fill2 = BUTTON_FRAME_FILL_2
        .Shadow = BUTTON_FRAME_SHADOW
    End With

    With HEADER_STYLE
        .Name = "HEADER_STYLE"
        .BorderWidth = HEADER_BORDER_WIDTH
        .Fill1 = HEADER_FILL_1
        .Fill2 = HEADER_FILL_2
        .Shadow = HEADER_SHADOW
        .FontStyle = HEADER_FONT_STYLE
        .FontSize = HEADER_FONT_SIZE
        .FontBold = HEADER_FONT_BOLD
        .FontColour = HEADER_FONT_COLOUR
        .FontXJust = HEADER_FONT_x_JUST
        .FontVJust = HEADER_FONT_Y_JUST
    End With

    With BTN_MAIN_STYLE
        .Name = "BTN_MAIN_STYLE"
        .BorderWidth = BTN_MAIN_BORDER_WIDTH
        .Fill1 = BTN_MAIN_FILL_1
        .Fill2 = BTN_MAIN_FILL_2
        .Shadow = BTN_MAIN_SHADOW
        .FontStyle = BTN_MAIN_FONT_STYLE
        .FontSize = BTN_MAIN_FONT_SIZE
        .FontBold = BTN_MAIN_FONT_BOLD
        .FontColour = BTN_MAIN_FONT_COLOUR
        .FontXJust = BTN_MAIN_FONT_x_JUST
        .FontVJust = BTN_MAIN_FONT_Y_JUST
    End With

    With GENERIC_BUTTON
        .Name = "GENERIC_BUTTON"
        .BorderWidth = GENERIC_BUTTON_BORDER_WIDTH
        .Fill1 = GENERIC_BUTTON_FILL_1
        .Fill2 = GENERIC_BUTTON_FILL_2
        .Shadow = GENERIC_BUTTON_SHADOW
        .FontStyle = GENERIC_BUTTON_FONT_STYLE
        .FontSize = GENERIC_BUTTON_FONT_SIZE
        .FontBold = GENERIC_BUTTON_FONT_BOLD
        .FontColour = GENERIC_BUTTON_FONT_COLOUR
        .FontXJust = GENERIC_BUTTON_FONT_x_JUST
        .FontVJust = GENERIC_BUTTON_FONT_Y_JUST
    End With

    With TODO_BUTTON
        .Name = "TODO_BUTTON"
        .BorderWidth = TODO_BUTTON_BORDER_WIDTH
        .Fill1 = TODO_BUTTON_FILL_1
        .Fill2 = TODO_BUTTON_FILL_2
        .Shadow = TODO_BUTTON_SHADOW
        .FontStyle = TODO_BUTTON_FONT_STYLE
        .FontSize = TODO_BUTTON_FONT_SIZE
        .FontBold = TODO_BUTTON_FONT_BOLD
        .FontColour = TODO_BUTTON_FONT_COLOUR
        .FontXJust = TODO_BUTTON_FONT_x_JUST
        .FontVJust = TODO_BUTTON_FONT_Y_JUST
    End With

    With TODO_BADGE
        .Name = "TODO_BADGE"
        .BorderWidth = TODO_BADGE_BORDER_WIDTH
        .BorderColour = TODO_BADGE_BORDER_COLOUR
        .Fill1 = TODO_BADGE_FILL_1
        .Fill2 = TODO_BADGE_FILL_2
        .Shadow = TODO_BADGE_SHADOW
        .FontStyle = TODO_BADGE_FONT_STYLE
        .FontSize = TODO_BADGE_FONT_SIZE
        .FontBold = TODO_BADGE_FONT_BOLD
        .FontColour = TODO_BADGE_FONT_COLOUR
        .FontXJust = TODO_BADGE_FONT_x_JUST
        .FontVJust = TODO_BADGE_FONT_Y_JUST
    End With

'    With TOOL_BUTTON
'        .BorderWidth = TOOL_BUTTON_BORDER_WIDTH
'        .Fill1 = TOOL_BUTTON_FILL_1
'        .Fill2 = TOOL_BUTTON_FILL_2
'        .Shadow = TOOL_BUTTON_SHADOW
'        .FontStyle = TOOL_BUTTON_FONT_STYLE
'        .FontSize = TOOL_BUTTON_FONT_SIZE
'        .FontBold = TOOL_BUTTON_FONT_BOLD
'        .FontColour = TOOL_BUTTON_FONT_COLOUR
'        .FontXJust = TOOL_BUTTON_FONT_x_JUST
'        .FontVJust = TOOL_BUTTON_FONT_Y_JUST
'    End With

    With GENERIC_TABLE
        .Name = "GENERIC_TABLE"
        .BorderWidth = GENERIC_TABLE_BORDER_WIDTH
        .Fill1 = GENERIC_TABLE_FILL_1
        .Fill2 = GENERIC_TABLE_FILL_2
        .Shadow = GENERIC_TABLE_SHADOW
        .FontStyle = GENERIC_TABLE_FONT_STYLE
        .FontSize = GENERIC_TABLE_FONT_SIZE
        .FontBold = GENERIC_TABLE_FONT_BOLD
        .FontColour = GENERIC_TABLE_FONT_COLOUR
        .FontXJust = GENERIC_TABLE_FONT_x_JUST
        .FontVJust = GENERIC_TABLE_FONT_Y_JUST
    End With

    With TABLE_PROGRESS_STYLE
        .Name = "TABLE_PROGRESS_STYLE"
        .BorderWidth = GENERIC_TABLE_BORDER_WIDTH
        .Fill1 = GENERIC_TABLE_FILL_1
        .Fill2 = GENERIC_TABLE_FILL_2
        .Shadow = GENERIC_TABLE_SHADOW
        .FontStyle = GENERIC_TABLE_FONT_STYLE
        .FontSize = TABLE_PROGRESS_FONT_SIZE
        .FontBold = GENERIC_TABLE_FONT_BOLD
        .FontColour = GENERIC_TABLE_FONT_COLOUR
        .FontXJust = TABLE_PROGRESS_CELL_X_JUST
        .FontVJust = TABLE_PROGRESS_CELL_Y_JUST
    End With

    With GREEN_CELL
        .Name = "GREEN_CELL"
        .BorderWidth = GREEN_CELL_BORDER_WIDTH
        .Fill1 = GREEN_CELL_FILL_1
        .Fill2 = GREEN_CELL_FILL_2
        .Shadow = GREEN_CELL_SHADOW
        .FontStyle = GREEN_CELL_FONT_STYLE
        .FontSize = GREEN_CELL_FONT_SIZE
        .FontBold = GREEN_CELL_FONT_BOLD
        .FontColour = GREEN_CELL_FONT_COLOUR
        .FontXJust = GREEN_CELL_FONT_x_JUST
        .FontVJust = GREEN_CELL_FONT_Y_JUST
    End With

    With AMBER_CELL
        .Name = "AMBER_CELL"
        .BorderWidth = AMBER_CELL_BORDER_WIDTH
        .Fill1 = AMBER_CELL_FILL_1
        .Fill2 = AMBER_CELL_FILL_2
        .Shadow = AMBER_CELL_SHADOW
        .FontStyle = AMBER_CELL_FONT_STYLE
        .FontSize = AMBER_CELL_FONT_SIZE
        .FontBold = AMBER_CELL_FONT_BOLD
        .FontColour = AMBER_CELL_FONT_COLOUR
        .FontXJust = AMBER_CELL_FONT_x_JUST
        .FontVJust = AMBER_CELL_FONT_Y_JUST
    End With

    With RED_CELL
        .Name = "RED_CELL"
        .BorderWidth = RED_CELL_BORDER_WIDTH
        .Fill1 = RED_CELL_FILL_1
        .Fill2 = RED_CELL_FILL_2
        .Shadow = RED_CELL_SHADOW
        .FontStyle = RED_CELL_FONT_STYLE
        .FontSize = RED_CELL_FONT_SIZE
        .FontBold = RED_CELL_FONT_BOLD
        .FontColour = RED_CELL_FONT_COLOUR
        .FontXJust = RED_CELL_FONT_x_JUST
        .FontVJust = RED_CELL_FONT_Y_JUST
    End With

    With GENERIC_TABLE_HEADER
        .Name = "GENERIC_TABLE_HEADER"
        .BorderWidth = GENERIC_TABLE_HEADER_BORDER_WIDTH
        .Fill1 = GENERIC_TABLE_HEADER_FILL_1
        .Fill2 = GENERIC_TABLE_HEADER_FILL_2
        .Shadow = GENERIC_TABLE_HEADER_SHADOW
        .FontStyle = GENERIC_TABLE_HEADER_FONT_STYLE
        .FontSize = GENERIC_TABLE_HEADER_FONT_SIZE
        .FontBold = GENERIC_TABLE_HEADER_FONT_BOLD
        .FontColour = GENERIC_TABLE_HEADER_FONT_COLOUR
        .FontXJust = GENERIC_TABLE_HEADER_FONT_x_JUST
        .FontVJust = GENERIC_TABLE_HEADER_FONT_Y_JUST
    End With

    With SUB_TABLE_HEADER
        .Name = "GENERIC_TABLE_HEADER"
        .BorderWidth = GENERIC_TABLE_HEADER_BORDER_WIDTH
        .Fill1 = SUB_TABLE_HEADER_FILL_1
        .Fill2 = SUB_TABLE_HEADER_FILL_2
        .Shadow = GENERIC_TABLE_HEADER_SHADOW
        .FontStyle = GENERIC_TABLE_HEADER_FONT_STYLE
        .FontSize = GENERIC_TABLE_HEADER_FONT_SIZE
        .FontBold = GENERIC_TABLE_HEADER_FONT_BOLD
        .FontColour = GENERIC_TABLE_HEADER_FONT_COLOUR
        .FontXJust = GENERIC_TABLE_HEADER_FONT_x_JUST
        .FontVJust = GENERIC_TABLE_HEADER_FONT_Y_JUST
    End With

    BuildScreenStyles = True

Exit Function
    
    
ErrorExit:

    BuildScreenStyles = False
    
Exit Function

ErrorHandler:
'If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
'        Stop
'        Resume
'    Else
'        Resume ErrorExit
'    End If
End Function

' ===============================================================
' DestroyScreenStyles
' Destroys the UI styles for use on the screen
' ---------------------------------------------------------------
Public Function DestroyScreenStyles() As Boolean
    Const StrPROCEDURE As String = "DestroyScreenStyles()"

'    On Error GoTo ErrorHandler
    
    Set SCREEN_STYLE = Nothing
    Set MENUBAR_STYLE = Nothing
    Set BUTTON_UNSET_STYLE = Nothing
    Set BUTTON_SET_STYLE = Nothing
    Set MAIN_FRAME_STYLE = Nothing
    Set BTN_MAIN_STYLE = Nothing
    Set GENERIC_BUTTON = Nothing
'    Set TOOL_BUTTON = Nothing
End Function
