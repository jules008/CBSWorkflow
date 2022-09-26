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

    On Error GoTo ErrorHandler

    With SCREEN_STYLE
        .BorderWidth = SCREEN_BORDER_WIDTH
        .Fill1 = SCREEN_FILL_1
        .Fill2 = SCREEN_FILL_2
        .Shadow = SCREEN_SHADOW
    End With

    With MENUBAR_STYLE
        .BorderWidth = MENUBAR_BORDER_WIDTH
        .Fill1 = MENUBAR_FILL_1
        .Fill2 = MENUBAR_FILL_2
        .Shadow = MENUBAR_SHADOW
    End With

    With BUTTON_UNSET_STYLE
        .BorderWidth = BUTTON_UNSET_BORDER_WIDTH
        .BorderColour = BUTTON_UNSET_BORDER_COLOUR
        .Fill1 = BUTTON_UNSET_FILL_1
        .Fill2 = BUTTON_UNSET_FILL_2
        .Shadow = BUTTON_UNSET_SHADOW
        .FontStyle = BUTTON_UNSET_FONT_STYLE
        .FontSize = BUTTON_UNSET_FONT_SIZE
        .FontColour = BUTTON_UNSET_FONT_COLOUR
        .FontXJust = BUTTON_UNSET_FONT_X_JUST
        .FontYJust = BUTTON_UNSET_FONT_Y_JUST
    End With

    With BUTTON_SET_STYLE
        .BorderWidth = BUTTON_SET_BORDER_WIDTH
        .BorderColour = BUTTON_SET_BORDER_COLOUR
        .Fill1 = BUTTON_SET_FILL_1
        .Fill2 = BUTTON_SET_FILL_2
        .Shadow = BUTTON_SET_SHADOW
        .FontStyle = BUTTON_SET_FONT_STYLE
        .FontSize = BUTTON_SET_FONT_SIZE
        .FontColour = BUTTON_SET_FONT_COLOUR
        .FontXJust = BUTTON_SET_FONT_X_JUST
        .FontYJust = BUTTON_SET_FONT_Y_JUST
    End With
    
    With MAIN_FRAME_STYLE
        .BorderWidth = MAIN_FRAME_BORDER_WIDTH
        .Fill1 = MAIN_FRAME_FILL_1
        .Fill2 = MAIN_FRAME_FILL_2
        .Shadow = MAIN_FRAME_SHADOW
    End With

    With HEADER_STYLE
        .BorderWidth = HEADER_BORDER_WIDTH
        .Fill1 = HEADER_FILL_1
        .Fill2 = HEADER_FILL_2
        .Shadow = HEADER_SHADOW
        .FontStyle = HEADER_FONT_STYLE
        .FontSize = HEADER_FONT_SIZE
        .FontBold = HEADER_FONT_BOLD
        .FontColour = HEADER_FONT_COLOUR
        .FontXJust = HEADER_FONT_X_JUST
        .FontYJust = HEADER_FONT_Y_JUST
    End With

    With BTN_MAIN_STYLE
        .BorderWidth = BTN_MAIN_BORDER_WIDTH
        .Fill1 = BTN_MAIN_FILL_1
        .Fill2 = BTN_MAIN_FILL_2
        .Shadow = BTN_MAIN_SHADOW
        .FontStyle = BTN_MAIN_FONT_STYLE
        .FontSize = BTN_MAIN_FONT_SIZE
        .FontBold = BTN_MAIN_FONT_BOLD
        .FontColour = BTN_MAIN_FONT_COLOUR
        .FontXJust = BTN_MAIN_FONT_X_JUST
        .FontYJust = BTN_MAIN_FONT_Y_JUST
    End With

    With GENERIC_BUTTON
        .BorderWidth = GENERIC_BUTTON_BORDER_WIDTH
        .Fill1 = GENERIC_BUTTON_FILL_1
        .Fill2 = GENERIC_BUTTON_FILL_2
        .Shadow = GENERIC_BUTTON_SHADOW
        .FontStyle = GENERIC_BUTTON_FONT_STYLE
        .FontSize = GENERIC_BUTTON_FONT_SIZE
        .FontBold = GENERIC_BUTTON_FONT_BOLD
        .FontColour = GENERIC_BUTTON_FONT_COLOUR
        .FontXJust = GENERIC_BUTTON_FONT_X_JUST
        .FontYJust = GENERIC_BUTTON_FONT_Y_JUST
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
'        .FontYJust = TOOL_BUTTON_FONT_Y_JUST
'    End With

    With GENERIC_Cell
        .BorderWidth = GENERIC_Cell_BORDER_WIDTH
        .Fill1 = GENERIC_Cell_FILL_1
        .Fill2 = GENERIC_Cell_FILL_2
        .Shadow = GENERIC_Cell_SHADOW
        .FontStyle = GENERIC_Cell_FONT_STYLE
        .FontSize = GENERIC_Cell_FONT_SIZE
        .FontBold = GENERIC_Cell_FONT_BOLD
        .FontColour = GENERIC_Cell_FONT_COLOUR
        .FontXJust = GENERIC_Cell_FONT_X_JUST
        .FontYJust = GENERIC_Cell_FONT_Y_JUST
    End With

    With GREEN_Cell
        .BorderWidth = GREEN_Cell_BORDER_WIDTH
        .Fill1 = GREEN_Cell_FILL_1
        .Fill2 = GREEN_Cell_FILL_2
        .Shadow = GREEN_Cell_SHADOW
        .FontStyle = GREEN_Cell_FONT_STYLE
        .FontSize = GREEN_Cell_FONT_SIZE
        .FontBold = GREEN_Cell_FONT_BOLD
        .FontColour = GREEN_Cell_FONT_COLOUR
        .FontXJust = GREEN_Cell_FONT_X_JUST
        .FontYJust = GREEN_Cell_FONT_Y_JUST
    End With

    With AMBER_Cell
        .BorderWidth = AMBER_Cell_BORDER_WIDTH
        .Fill1 = AMBER_Cell_FILL_1
        .Fill2 = AMBER_Cell_FILL_2
        .Shadow = AMBER_Cell_SHADOW
        .FontStyle = AMBER_Cell_FONT_STYLE
        .FontSize = AMBER_Cell_FONT_SIZE
        .FontBold = AMBER_Cell_FONT_BOLD
        .FontColour = AMBER_Cell_FONT_COLOUR
        .FontXJust = AMBER_Cell_FONT_X_JUST
        .FontYJust = AMBER_Cell_FONT_Y_JUST
    End With

    With RED_Cell
        .BorderWidth = RED_Cell_BORDER_WIDTH
        .Fill1 = RED_Cell_FILL_1
        .Fill2 = RED_Cell_FILL_2
        .Shadow = RED_Cell_SHADOW
        .FontStyle = RED_Cell_FONT_STYLE
        .FontSize = RED_Cell_FONT_SIZE
        .FontBold = RED_Cell_FONT_BOLD
        .FontColour = RED_Cell_FONT_COLOUR
        .FontXJust = RED_Cell_FONT_X_JUST
        .FontYJust = RED_Cell_FONT_Y_JUST
    End With

    With GENERIC_Cell_HEADER
        .BorderWidth = GENERIC_Cell_HEADER_BORDER_WIDTH
        .Fill1 = GENERIC_Cell_HEADER_FILL_1
        .Fill2 = GENERIC_Cell_HEADER_FILL_2
        .Shadow = GENERIC_Cell_HEADER_SHADOW
        .FontStyle = GENERIC_Cell_HEADER_FONT_STYLE
        .FontSize = GENERIC_Cell_HEADER_FONT_SIZE
        .FontBold = GENERIC_Cell_HEADER_FONT_BOLD
        .FontColour = GENERIC_Cell_HEADER_FONT_COLOUR
        .FontXJust = GENERIC_Cell_HEADER_FONT_X_JUST
        .FontYJust = GENERIC_Cell_HEADER_FONT_Y_JUST
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
'        .FontYJust = TRANSPARENT_TEXT_BOX_FONT_Y_JUST
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
'        .FontYJust = TRANSPARENT_TEXT_BOX_FONT_Y_JUST
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
'        .FontYJust = VERT_Cell_HEADER_FONT_Y_JUST
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
'        .FontYJust = MATRIX_DEF_FONT_Y_JUST
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
'        .FontYJust = MATRIX_1_FONT_Y_JUST
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
'        .FontYJust = MATRIX_3_FONT_Y_JUST
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
'        .FontYJust = MATRIX_4_FONT_Y_JUST
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
'        .FontYJust = BTN_SUPPORT_FONT_Y_JUST
'    End With
'
    BuildScreenStyles = True

Exit Function
    
    
ErrorExit:

    BuildScreenStyles = False
    
Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


