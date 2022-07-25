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
' BuildStylesMenu
' Builds the UI styles for use on the menu and back drop
' ---------------------------------------------------------------
'Public Function BuildStylesMenu() As Boolean
'    Const StrPROCEDURE As String = "BuildStylesMenu()"
'
'    On Error GoTo ErrorHandler
'
'    With SCREEN_STYLE
'        .BorderWidth = SCREEN_BORDER_WIDTH
'        .Fill1 = SCREEN_FILL_1
'        .Fill2 = SCREEN_FILL_2
'        .Shadow = SCREEN_SHADOW
'    End With
'
'    With MENUBAR_STYLE
'        .BorderWidth = MENUBAR_BORDER_WIDTH
'        .Fill1 = MENUBAR_FILL_1
'        .Fill2 = MENUBAR_FILL_2
'        .Shadow = MENUBAR_SHADOW
'    End With
'
'    With MENUITEM_UNSET_STYLE
'        .BorderWidth = MENUITEM_UNSET_BORDER_WIDTH
'        .BorderColour = MENUITEM_UNSET_BORDER_COLOUR
'        .Fill1 = MENUITEM_UNSET_FILL_1
'        .Fill2 = MENUITEM_UNSET_FILL_2
'        .Shadow = MENUITEM_UNSET_SHADOW
'        .FontStyle = MENUITEM_UNSET_FONT_STYLE
'        .FontSize = MENUITEM_UNSET_FONT_SIZE
'        .FontColour = MENUITEM_UNSET_FONT_COLOUR
'        .FontXJust = MENUITEM_UNSET_FONT_X_JUST
'        .FontYJust = MENUITEM_UNSET_FONT_Y_JUST
'    End With
'
'    With MENUITEM_SET_STYLE
'        .BorderWidth = MENUITEM_SET_BORDER_WIDTH
'        .BorderColour = MENUITEM_SET_BORDER_COLOUR
'        .Fill1 = MENUITEM_SET_FILL_1
'        .Fill2 = MENUITEM_SET_FILL_2
'        .Shadow = MENUITEM_SET_SHADOW
'        .FontStyle = MENUITEM_SET_FONT_STYLE
'        .FontSize = MENUITEM_SET_FONT_SIZE
'        .FontColour = MENUITEM_SET_FONT_COLOUR
'        .FontXJust = MENUITEM_SET_FONT_X_JUST
'        .FontYJust = MENUITEM_SET_FONT_Y_JUST
'    End With
'
'    BuildStylesMenu = True
'
'Exit Function
'
'ErrorExit:
'
'    BuildStylesMenu = False
'
'Exit Function
'
'ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
'        Stop
'        Resume
'    Else
'        Resume ErrorExit
'    End If
'End Function

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
    Dim MenuItemText() As String
    Dim MenuItemIcon() As String
    Dim MenuItemBadge() As String
    Dim i As Integer
    
    Const StrPROCEDURE As String = "BuildMenuBar()"

    On Error GoTo ErrorHandler
    
'    Set Logo = New ClsUIDashObj
'    Set BtnSupport = New ClsUIMenuItem
    
    MainScreen.Frames.AddItem MenuBar
   
   'Menubar
    With MenuBar
        .Top = MENUBAR_TOP
        .Left = MENUBAR_LEFT
        .Height = MENUBAR_HEIGHT
        .Width = MENUBAR_WIDTH
        .Name = "MenuBar"
        .Style = MENUBAR_STYLE
        .Header.Visible = False
        .EnableHeader = False
    End With

    'Logo
'    MenuBar.DashObs.AddItem Logo
    
'    With Logo
'        .EnumObjType = ObjImage
'        .ShpDashObj = ShtMain.Shapes("TEMPLATE - Logo").Duplicate
'        .Name = "Logo"
'        .Visible = True
'        .Top = LOGO_TOP
'        .Left = LOGO_LEFT
'        .Width = LOGO_WIDTH
'        .Height = LOGO_HEIGHT
'    End With

    'menu
    With MenuBar.Menu
        .Top = MENU_TOP
        .Left = MENU_LEFT
    End With

    'Menu Items
    MenuItemText() = Split(MENUITEM_TEXT, ":")
'    MenuItemIcon() = Split(MENUITEM_ICONS, ":")
'    MenuItemBadge() = Split(MENUITEM_BADGES, ":")

    For i = 0 To MENUITEM_COUNT - 1

        Set MenuItem = New ClsUIMenuItem
    
        With MenuItem
            .SelectStyle = MENUITEM_SET_STYLE
            .UnSelectStyle = MENUITEM_UNSET_STYLE
            .Height = MENUITEM_HEIGHT
            .Width = MENUITEM_WIDTH
            .Text = MenuItemText(i)
            .Name = "MenuItem - " & .Text
            .OnAction = "'ModUIMenu.ProcessBtnPress(" & i + 1 & ")'"
'            .Icon = ShtMain.Shapes(MenuItemIcon(i)).Duplicate
'            If MenuItemBadge(i) <> "" Then .Badge = ShtMain.Shapes(MenuItemBadge(i)).Duplicate

            MenuBar.Menu.AddItem MenuItem

            .Top = MENU_TOP + (i * .Height) - i
            .Left = .Left
            .Selected = False

'            With .Icon
'                .Visible = True
'                .Name = "Icon - " & MenuItem.Text
'                .Left = MenuItem.Left + MENUITEM_ICON_LEFT
'                .Top = MenuItem.Top + MENUITEM_ICON_TOP
'            End With
            
'            If MenuItemBadge(i) <> "" Then
'                With .Badge
'                    .Visible = True
'                    .Name = "Icon - " & MenuItem.Text
'                    .Left = MenuItem.Left + MENUITEM_BADGE_LEFT
'                    .Top = MenuItem.Top + MENUITEM_BADGE_TOP
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
'        .Name = "MenuItem - " & .Text
'        .OnAction = "'ModUIMenu.ProcessBtnPress(" & i + 1 & ")'"
'
'        MenuBar.Menu.AddItem BtnSupport
'
'    End With
    
    Set MenuItem = Nothing

    BuildMenuBar = True

Exit Function

ErrorExit:

    Set MenuItem = Nothing
    
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
    
    Const StrPROCEDURE As String = "ProcessBtnPress()"

    On Error GoTo ErrorHandler
    
    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART
    
Restart:
    Application.StatusBar = ""
            
    If Not ModUIActive.BuildScreen Then Err.Raise HANDLED_ERROR
    
'    If ButtonNo = 0 Then
'        ButtonNo = [menuitemno]
'    Else
'        If ButtonNo < 7 And ButtonNo = [menuitemno] Then Exit Function
'    End If
    
'    Select Case ButtonNo
'
'        Case enBtnForAction
            
'            ShtMain.Unprotect PROTECT_KEY
'            [menuitemno] = 1
'
'            If Not ResetScreen Then Err.Raise HANDLED_ERROR
'            If Not ModUIForAction.BuildScreen Then Err.Raise HANDLED_ERROR
'
'            ShtMain.Unprotect PROTECT_KEY
'
'            With MenuBar
'                .Menu(1).Selected = True
'                .Menu(2).Selected = False
'                .Menu(3).Selected = False
'                .Menu(4).Selected = False
'                .Menu(5).Selected = False
'                .Menu(6).Selected = False
'                .Menu(7).Selected = False
'            End With
'
'            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY
            
'        Case enBtnActive
'
'            ShtMain.Unprotect PROTECT_KEY
'            [menuitemno] = 2
'
'            If Not ResetScreen Then Err.Raise HANDLED_ERROR
'            If Not ModUIActive.BuildScreen Then Err.Raise HANDLED_ERROR
'
'            ShtMain.Unprotect PROTECT_KEY
'
'            With MenuBar
'                .Menu(1).Selected = False
'                .Menu(2).Selected = True
'                .Menu(3).Selected = False
'                .Menu(4).Selected = False
'                .Menu(5).Selected = False
'                .Menu(6).Selected = False
'                .Menu(7).Selected = False
'            End With
'
'            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY
'
'        Case enBtnComplete

'            ShtMain.Unprotect PROTECT_KEY
'            [menuitemno] = 3
'
'            If Not ResetScreen Then Err.Raise HANDLED_ERROR
'            If Not ModUIComplete.BuildScreen Then Err.Raise HANDLED_ERROR
'
''                If CurrentUser.AccessLvl < ManagerLvl_4 Then Err.Raise ACCESS_DENIED
'
'            ShtMain.Unprotect PROTECT_KEY
'
'            With MenuBar
'                .Menu(1).Selected = False
'                .Menu(2).Selected = False
'                .Menu(3).Selected = True
'                .Menu(4).Selected = False
'                .Menu(5).Selected = False
'                .Menu(6).Selected = False
'                .Menu(7).Selected = False
'            End With
'
'            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

'        Case enBtnDashboard

'            ShtMain.Unprotect PROTECT_KEY
'            [menuitemno] = 4
'
'            If Not ResetScreen Then Err.Raise HANDLED_ERROR
'            If Not ModReports.ExpQryCertsMatrix(False) Then Err.Raise HANDLED_ERROR
'            If Not ModUIDashboard.BuildScreen Then Err.Raise HANDLED_ERROR
'
'            ShtMain.Unprotect PROTECT_KEY
'
'            With MenuBar
'                .Menu(1).Selected = False
'                .Menu(2).Selected = False
'                .Menu(3).Selected = False
'                .Menu(4).Selected = True
'                .Menu(5).Selected = False
'                .Menu(6).Selected = False
'                .Menu(7).Selected = False
'            End With
'
'            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY
'
'            ModLibrary.PerfSettingsOff
'
'        Case enBtnRepPage
        
'            ShtMain.Unprotect PROTECT_KEY
'            [menuitemno] = 5
'
'            If Not ResetScreen Then Err.Raise HANDLED_ERROR
'            If Not ModUIReports.BuildScreen Then Err.Raise HANDLED_ERROR
'
'            ShtMain.Unprotect PROTECT_KEY
'
'            With MenuBar
'                .Menu(1).Selected = False
'                .Menu(2).Selected = False
'                .Menu(3).Selected = False
'                .Menu(4).Selected = False
'                .Menu(5).Selected = True
'                .Menu(6).Selected = False
'                .Menu(7).Selected = False
'           End With
'
'            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY
'
'        Case enBtnAdmin

'            If CurrentUser.UserLvl = enBasic Then Err.Raise ACCESS_DENIED
'
'            ShtMain.Unprotect PROTECT_KEY
'            [menuitemno] = 6
'
'            If Not ResetScreen Then Err.Raise HANDLED_ERROR
'            If Not ModUIAdmin.BuildScreen Then Err.Raise HANDLED_ERROR
'
'            ShtMain.Unprotect PROTECT_KEY
'
'            With MenuBar
'                .Menu(1).Selected = False
'                .Menu(2).Selected = False
'                .Menu(3).Selected = False
'                .Menu(4).Selected = False
'                .Menu(5).Selected = False
'                .Menu(6).Selected = True
'                .Menu(7).Selected = False
'            End With
'
'            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY
'
'        Case enBtnExit

'            Response = MsgBox("Are you sure you want to exit?", vbExclamation + vbYesNo + vbDefaultButton2, APP_NAME)
'
'            If Response = 6 Then
'
'                If Workbooks.Count = 1 Then
'                    With Application
'                        .DisplayAlerts = False
'                        .Quit
'                        .DisplayAlerts = True
'                    End With
'                Else
'                    ThisWorkbook.Close savechanges:=False
'                End If
'
'            End If
            
'        Case enBtnSupport
'            If Not FrmSupportForm.ShowForm Then Err.Raise HANDLED_ERROR
'    End Select
        
        
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
    Dim UILineitem As ClsUILineitem
'    Dim DashObj As ClsUIDashObj
    Dim MenuItem As ClsUIMenuItem
    
    Const StrPROCEDURE As String = "ResetScreen()"

    On Error Resume Next
    
    ShtMain.Unprotect PROTECT_KEY
        
    For Each MenuItem In MainScreen.Menu
        MainScreen.Menu.RemoveItem MenuItem.Name
        MenuItem.ShpMenuItem.Delete
        Set MenuItem = Nothing
    Next
    
    For Each Frame In MainScreen.Frames
        If Frame.Name <> "MenuBar" Then
'            For Each DashObj In Frame.DashObs
'                Frame.DashObs.RemoveItem DashObj.Name
'                DashObj.ShpDashObj.Delete
'                Set DashObj = Nothing
'            Next
            
            For Each UILineitem In Frame.Lineitems
                Frame.Lineitems.RemoveItem UILineitem.Name
                UILineitem.ShpLineitem.Delete
                Set UILineitem = Nothing
            Next
            
            For Each UILineitem In Frame.LineitemHeaders
                Frame.LineitemHeaders.RemoveItem UILineitem.Name
                UILineitem.ShpLineitem.Delete
                Set UILineitem = Nothing
            Next
            
            For Each MenuItem In Frame.Menu
                Frame.Menu.RemoveItem MenuItem.Name
                MenuItem.ShpMenuItem.Delete
                Set MenuItem = Nothing
            Next
            
            Frame.Header.ShpHeader.Delete
            Frame.Header.Icon.Delete
            Set Frame.Header = Nothing
            
            MainScreen.Frames.RemoveItem Frame.Name
            Frame.ShpFrame.Delete
            Set Frame = Nothing
            
        End If
    Next
    
    For Each MenuItem In MainScreen.Menu
        MainScreen.Menu.RemoveItem MenuItem.Name
        Set MenuItem = Nothing
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



