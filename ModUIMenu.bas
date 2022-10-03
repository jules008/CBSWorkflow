Attribute VB_Name = "ModUIMenu"
'===============================================================
' Module ModUIMenu
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 25 Jun 20
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModUIMenu"

' ===============================================================
' BuildMenu
' Builds the menu using shapes
' ---------------------------------------------------------------
Public Function BuildMenu() As Boolean
    
    Const StrPROCEDURE As String = "BuildMenu()"

    On Error GoTo ErrorHandler
        
    Set MainScreen = New ClsUIScreen
    Set MenuBar = New ClsUIFrame

    If Not BuildBackDrop Then Err.Raise HANDLED_ERROR
    If Not BuildMenuBar Then Err.Raise HANDLED_ERROR
    
    BuildMenu = True
       
Exit Function

ErrorExit:

    BuildMenu = False
    
Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildBackDrop
' Builds the background image
' ---------------------------------------------------------------
Private Function BuildBackDrop() As Boolean
    Const StrPROCEDURE As String = "BuildBackDrop()"

    On Error GoTo ErrorHandler

    'Main Screen
    With MainScreen
        .Style = SCREEN_STYLE
        .Name = "Main Screen"
        .Top = 0
        .Left = 0
        .Height = SCREEN_HEIGHT
        .Width = SCREEN_WIDTH
    End With
    
    BuildBackDrop = True

Exit Function

ErrorExit:

    BuildBackDrop = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


' ===============================================================
' BuildMenuBar
' Builds menu on menu bar
' ---------------------------------------------------------------
Private Function BuildMenuBar() As Boolean
    Dim ButtonText() As String
    Dim ButtonIcon() As String
    Dim ButtonBadge() As String
    Dim ButtonIndex() As String
    Dim Button As ClsUIButton
    Dim i As Integer
    
    Const StrPROCEDURE As String = "BuildMenuBar()"

    On Error GoTo ErrorHandler
    
    Set Logo = New ClsUIDashObj
'    Set BtnSupport = New ClsUIButton
    
    MainScreen.Frames.AddItem MenuBar, "MenuBar"
   
   'Menubar
    With MenuBar
        .Top = MENUBAR_TOP
        .Left = MENUBAR_LEFT
        .Height = MENUBAR_HEIGHT
        .Width = MENUBAR_WIDTH
        .Style = MENUBAR_STYLE
        .Header.Visible = False
        .EnableHeader = False
        .ZOrder = 0
    End With

    'Logo
    MenuBar.DashObjs.AddItem Logo
    
    With Logo
        .EnumObjType = ObjImage
        .ShpDashObj = ShtMain.Shapes("TEMPLATE - Logo").Duplicate
        .Name = "Logo"
        .Visible = True
        .Top = LOGO_TOP
        .Left = LOGO_LEFT
        .Width = LOGO_WIDTH
        .Height = LOGO_HEIGHT
    End With

    'menu
    With MenuBar.Menu
        .Top = MENU_TOP
        .Left = MENU_LEFT
    End With

    'Menu Items
    ButtonText() = Split(BUTTON_TEXT, ":")
    ButtonIndex() = Split(BUTTON_INDEX, ":")
'    ButtonIcon() = Split(Button_ICONS, ":")
'    ButtonBadge() = Split(Button_BADGES, ":")

    For i = 0 To BUTTON_COUNT - 1

        Set Button = New ClsUIButton
    
        With Button
            .SelectStyle = BUTTON_SET_STYLE
            .UnSelectStyle = BUTTON_UNSET_STYLE
            .Height = BUTTON_HEIGHT
            .Width = BUTTON_WIDTH
            .Text = ButtonText(i)
            .ButtonIndex = ButtonIndex(i)
            .Name = "Menu Btn - " & .ButtonIndex
            MenuBar.Menu.AddButton Button

        End With
    Next
    
    Set Button = Nothing

    BuildMenuBar = True

Exit Function

ErrorExit:

    Set Button = Nothing
    
    BuildMenuBar = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' ButtonClickEvent
' Handles Button Click Events
' ---------------------------------------------------------------
Public Function ButtonClickEvent(ButtonIndex As String) As Boolean
    
    Const StrPROCEDURE As String = "ButtonClickEvent()"

    On Error GoTo ErrorHandler
    
    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART
    
    MenuBar.Menu.ButtonClick ButtonIndex
        
    ButtonClickEvent = True

Exit Function

ErrorExit:

    ButtonClickEvent = False
    If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

Exit Function

ErrorHandler:
    
    If Err.Number >= 2000 And Err.Number <= 2500 Then
        CustomErrorHandler Err.Number
        Resume Next
    End If
    
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' ResetScreen
' resets to start up screen for transitions between pages
' ---------------------------------------------------------------
Public Function ResetScreen() As Boolean
    Dim Frame As ClsUIFrame
    
    Const StrPROCEDURE As String = "ResetScreen()"

    On Error Resume Next
    
    ShtMain.Unprotect PROTECT_KEY
        
    For Each Frame In MainScreen.Frames
        If Frame.Name <> "MenuBar" Then
            MainScreen.Frames.RemoveItem Frame.Name
            Frame.Terminate
            Set Frame = Nothing
        End If
    Next
    
    If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY
        
    ResetScreen = True
        
Exit Function

ErrorExit:

    ResetScreen = False
    If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

Exit Function

ErrorHandler:
    
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function



