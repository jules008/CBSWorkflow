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
    Set BTN_MAIN_STYLE = New ClsUIStyle
    Set GENERIC_BUTTON = New ClsUIStyle
    Set HEADER_STYLE = New ClsUIStyle
    Set GENERIC_TABLE = New ClsUIStyle
    Set GREEN_CELL = New ClsUIStyle
    Set AMBER_CELL = New ClsUIStyle
    Set RED_CELL = New ClsUIStyle
    Set GENERIC_TABLE_HEADER = New ClsUIStyle

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
        .FontXJust = BUTTON_UNSET_FONT_X_JUST
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
        .FontXJust = BUTTON_SET_FONT_X_JUST
        .FontVJust = BUTTON_SET_FONT_Y_JUST
    End With
    
    With MAIN_FRAME_STYLE
        .Name = "MAIN_FRAME_STYLE"
        .BorderWidth = MAIN_FRAME_BORDER_WIDTH
        .Fill1 = MAIN_FRAME_FILL_1
        .Fill2 = MAIN_FRAME_FILL_2
        .Shadow = MAIN_FRAME_SHADOW
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
        .FontXJust = HEADER_FONT_X_JUST
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
        .FontXJust = BTN_MAIN_FONT_X_JUST
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
        .FontXJust = GENERIC_BUTTON_FONT_X_JUST
        .FontVJust = GENERIC_BUTTON_FONT_Y_JUST
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
'        .FontXJust = TOOL_BUTTON_FONT_X_JUST
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
        .FontXJust = GENERIC_TABLE_FONT_X_JUST
        .FontVJust = GENERIC_TABLE_FONT_Y_JUST
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
        .FontXJust = GREEN_CELL_FONT_X_JUST
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
        .FontXJust = AMBER_CELL_FONT_X_JUST
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
        .FontXJust = RED_CELL_FONT_X_JUST
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
        .FontXJust = GENERIC_TABLE_HEADER_FONT_X_JUST
        .FontVJust = GENERIC_TABLE_HEADER_FONT_Y_JUST
    End With

'    With TRANSPARENT_TEXT_BOX
'        .BorderWidth = TRANSPARENT_TEXT_BOX_BORDER_WIDTH
'        .Fill1 = TRANSPARENT_TEXT_BOX_FILL_1
'        .Fill2 = TRANSPARENT_TEXT_BOX_FILL_2
'        .Shadow = TRANSPARENT_TEXT_BOX_SHADOW
'        .FontStyle = TRANSPARENT_TEXT_BOX_FONT_STYLE
'        .FontSize = TRANSPARENT_TEXT_BOX_FONT_SIZE
'        .FontBold = TRANSPARENT_TEXT_BOX_FONT_BOLD
'        .FontColour = TRANSPARENT_TEXT_BOX_FONT_COLOUR
'        .FontXJust = TRANSPARENT_TEXT_BOX_FONT_X_JUST
'        .FontVJust = TRANSPARENT_TEXT_BOX_FONT_Y_JUST
'    End With
'
'    With TRANSPARENT_TEXT_BOX
'        .BorderWidth = TRANSPARENT_TEXT_BOX_BORDER_WIDTH
'        .Fill1 = TRANSPARENT_TEXT_BOX_FILL_1
'        .Fill2 = TRANSPARENT_TEXT_BOX_FILL_2
'        .Shadow = TRANSPARENT_TEXT_BOX_SHADOW
'        .FontStyle = TRANSPARENT_TEXT_BOX_FONT_STYLE
'        .FontSize = TRANSPARENT_TEXT_BOX_FONT_SIZE
'        .FontBold = TRANSPARENT_TEXT_BOX_FONT_BOLD
'        .FontColour = TRANSPARENT_TEXT_BOX_FONT_COLOUR
'        .FontXJust = TRANSPARENT_TEXT_BOX_FONT_X_JUST
'        .FontVJust = TRANSPARENT_TEXT_BOX_FONT_Y_JUST
'    End With
'
'    With VERT_Cell_HEADER
'        .BorderWidth = VERT_Cell_HEADER_BORDER_WIDTH
'        .Fill1 = VERT_Cell_HEADER_FILL_1
'        .Fill2 = VERT_Cell_HEADER_FILL_2
'        .Shadow = VERT_Cell_HEADER_SHADOW
'        .FontStyle = VERT_Cell_HEADER_FONT_STYLE
'        .FontSize = VERT_Cell_HEADER_FONT_SIZE
'        .FontBold = VERT_Cell_HEADER_FONT_BOLD
'        .FontColour = VERT_Cell_HEADER_FONT_COLOUR
'        .FontXJust = VERT_Cell_HEADER_FONT_X_JUST
'        .FontVJust = VERT_Cell_HEADER_FONT_Y_JUST
'        .TextDir = VERT_Cell_HEADER_TEXT_DIR
'    End With
'
'    With MATRIX_DEF
'        .BorderWidth = MATRIX_DEF_BORDER_WIDTH
'        .Fill1 = MATRIX_DEF_FILL_1
'        .Fill2 = MATRIX_DEF_FILL_2
'        .Shadow = MATRIX_DEF_SHADOW
'        .FontStyle = MATRIX_DEF_FONT_STYLE
'        .FontSize = MATRIX_DEF_FONT_SIZE
'        .FontBold = MATRIX_DEF_FONT_BOLD
'        .FontColour = MATRIX_DEF_FONT_COLOUR
'        .FontXJust = MATRIX_DEF_FONT_X_JUST
'        .FontVJust = MATRIX_DEF_FONT_Y_JUST
'    End With
'
'    With MATRIX_1
'        .BorderWidth = MATRIX_1_BORDER_WIDTH
'        .BorderColour = MATRIX_1_BORDER_COLOUR
'        .Fill1 = MATRIX_1_FILL_1
'        .Fill2 = MATRIX_1_FILL_2
'        .Shadow = MATRIX_1_SHADOW
'        .FontStyle = MATRIX_1_FONT_STYLE
'        .FontSize = MATRIX_1_FONT_SIZE
'        .FontBold = MATRIX_1_FONT_BOLD
'        .FontColour = MATRIX_1_FONT_COLOUR
'        .FontXJust = MATRIX_1_FONT_X_JUST
'        .FontVJust = MATRIX_1_FONT_Y_JUST
'    End With
'
'    With MATRIX_3
'        .BorderWidth = MATRIX_3_BORDER_WIDTH
'        .BorderColour = MATRIX_3_BORDER_COLOUR
'        .Fill1 = MATRIX_3_FILL_1
'        .Fill2 = MATRIX_3_FILL_2
'        .Shadow = MATRIX_3_SHADOW
'        .FontStyle = MATRIX_3_FONT_STYLE
'        .FontSize = MATRIX_3_FONT_SIZE
'        .FontBold = MATRIX_3_FONT_BOLD
'        .FontColour = MATRIX_3_FONT_COLOUR
'        .FontXJust = MATRIX_3_FONT_X_JUST
'        .FontVJust = MATRIX_3_FONT_Y_JUST
'    End With
'
'    With MATRIX_4
'        .BorderWidth = MATRIX_4_BORDER_WIDTH
'        .BorderColour = MATRIX_4_BORDER_COLOUR
'        .Fill1 = MATRIX_4_FILL_1
'        .Fill2 = MATRIX_4_FILL_2
'        .Shadow = MATRIX_4_SHADOW
'        .FontStyle = MATRIX_4_FONT_STYLE
'        .FontSize = MATRIX_4_FONT_SIZE
'        .FontBold = MATRIX_4_FONT_BOLD
'        .FontColour = MATRIX_4_FONT_COLOUR
'        .FontXJust = MATRIX_4_FONT_X_JUST
'        .FontVJust = MATRIX_4_FONT_Y_JUST
'    End With
'
'    With BTN_SUPPORT
'        .BorderWidth = BTN_SUPPORT_BORDER_WIDTH
'        .BorderColour = BTN_SUPPORT_BORDER_COLOUR
'        .Fill1 = BTN_SUPPORT_FILL_1
'        .Fill2 = BTN_SUPPORT_FILL_2
'        .Shadow = BTN_SUPPORT_SHADOW
'        .FontStyle = BTN_SUPPORT_FONT_STYLE
'        .FontSize = BTN_SUPPORT_FONT_SIZE
'        .FontBold = BTN_SUPPORT_FONT_BOLD
'        .FontColour = BTN_SUPPORT_FONT_COLOUR
'        .FontXJust = BTN_SUPPORT_FONT_X_JUST
'        .FontVJust = BTN_SUPPORT_FONT_Y_JUST
'    End With
'
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
