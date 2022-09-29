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
    ButtonText() = Split(Button_TEXT, ":")
'    ButtonIcon() = Split(Button_ICONS, ":")
'    ButtonBadge() = Split(Button_BADGES, ":")

    For i = 0 To Button_COUNT - 1

        Set Button = New ClsUIButton
    
        With Button
            .SelectStyle = BUTTON_SET_STYLE
            .UnSelectStyle = BUTTON_UNSET_STYLE
            .Height = Button_HEIGHT
            .Width = Button_WIDTH
            .Text = ButtonText(i)
            .Name = "Button - " & .Text
            .OnAction = "'ModUIMenu.ProcessBtnPress(" & i + 1 & ")'"
'            .Icon = ShtMain.Shapes(ButtonIcon(i)).Duplicate
'            If ButtonBadge(i) <> "" Then .Badge = ShtMain.Shapes(ButtonBadge(i)).Duplicate

            MenuBar.Menu.AddButton Button

            .Top = MENU_TOP + (i * .Height) - i
            .Left = .Left
            .Selected = False

'            With .Icon
'                .Visible = True
'                .Name = "Icon - " & Button.Text
'                .Left = Button.Left + Button_ICON_LEFT
'                .Top = Button.Top + Button_ICON_TOP
'            End With
            
'            If ButtonBadge(i) <> "" Then
'                With .Badge
'                    .Visible = True
'                    .Name = "Icon - " & Button.Text
'                    .Left = Button.Left + Button_BADGE_LEFT
'                    .Top = Button.Top + Button_BADGE_TOP
'                End With
'                .BadgeText = "0"
'           End If
        End With
    Next
    
'    With BtnSupport
'        .UnSelectStyle = BTN_SUPPORT
'        .Selected = False
'        .Height = BTN_SUPPORT_HEIGHT
'        .Width = BTN_SUPPORT_WIDTH
'        .Top = BTN_SUPPORT_TOP
'        .Left = BTN_SUPPORT_LEFT
'        .Text = "Send Support Message"
'        .Name = "Button - " & .Text
'        .OnAction = "'ModUIMenu.ProcessBtnPress(" & i + 1 & ")'"
'
'        MenuBar.Menu.AddItem BtnSupport
'
'    End With
    
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
' ProcessBtnPress
' Receives all button presses and processes
' ---------------------------------------------------------------
Public Function ProcessBtnPress(Optional ButtonNo As EnumBtnNo) As Boolean
    Dim Response As Integer
    Dim RngMenu
    Const StrPROCEDURE As String = "ProcessBtnPress()"

    On Error GoTo ErrorHandler
    
    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART
    
Restart:
    Application.StatusBar = ""
            
    If ButtonNo = 0 Then
        If Not ShtMain.[Button] Is Nothing Then
            ButtonNo = ShtMain.[Button]
        Else
            ButtonNo = enBtnActive
        End If
    Else
        If ButtonNo < 4 And ButtonNo = ShtMain.[Button] Then Exit Function
    End If
    
    Select Case ButtonNo

        Case enBtnForAction
            
            ShtMain.Unprotect PROTECT_KEY
            ShtMain.[Button] = 1

            If Not ResetScreen Then Err.Raise HANDLED_ERROR
            If Not ModUIForAction.BuildScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Unprotect PROTECT_KEY

            MenuBar.Menu.ButtonClick "Button - For Action"

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY
            
        Case enBtnActive

            ShtMain.Unprotect PROTECT_KEY
            ShtMain.[Button] = 2

            If Not ResetScreen Then Err.Raise HANDLED_ERROR
            If Not ModUIActive.BuildScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Unprotect PROTECT_KEY

            MenuBar.Menu.ButtonClick "Button - Active"

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enBtnComplete

            ShtMain.Unprotect PROTECT_KEY
            ShtMain.[Button] = 3

            If Not ResetScreen Then Err.Raise HANDLED_ERROR
            If Not ModUIComplete.BuildScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Unprotect PROTECT_KEY

            MenuBar.Menu.ButtonClick "Button - Complete"

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enBtnExit

            Response = MsgBox("Are you sure you want to exit?", vbExclamation + vbYesNo + vbDefaultButton2, APP_NAME)

            If Response = 6 Then

                If Workbooks.Count = 1 Then
                    With Application
                        .DisplayAlerts = False
                        .Quit
                        .DisplayAlerts = True
                    End With
                Else
                    ThisWorkbook.Close savechanges:=False
                End If
            End If
            
    End Select
        
        
GracefulExit:
    
    ModLibrary.PerfSettingsOff

    ProcessBtnPress = True

Exit Function

ErrorExit:

    Application.DisplayAlerts = True

    ProcessBtnPress = False

Exit Function

ErrorHandler:
    
    If Err.Number >= 2000 And Err.Number <= 2500 Then
        If CustomErrorHandler(Err.Number) = SYSTEM_RESTART Then
            Resume Restart
        Else
            Resume GracefulExit
        End If
    End If
    
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
       Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' ResetScreen
' Functions for graceful close down of system
' ---------------------------------------------------------------
Public Function ResetScreen() As Boolean
    Dim Frame As ClsUIFrame
    
    Const StrPROCEDURE As String = "ResetScreen()"

'    On Error Resume Next
    
    ShtMain.Unprotect PROTECT_KEY
        
    For Each Frame In MainScreen.Frames
        If Frame.Name <> "MenuBar" Then
            MainScreen.Frames.RemoveItem Frame.Name
            Frame.Terminate
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



