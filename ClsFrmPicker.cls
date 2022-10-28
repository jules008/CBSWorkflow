VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsFrmPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsFrmPicker
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 07 Oct 22
'===============================================================
' Methods
'---------------------------------------------------------------
' ClearForm - Clears form ready to start over
' SearchTextChanged - Event on search text change.  lists all matching data in list box
' ResultsListSelect - Event on results list being clicked.  moves selected item to search box
'===============================================================
'---------------------------------------------------------------
Option Explicit

Private pData As Recordset
Private WithEvents pFormPicker As FrmPicker
Attribute pFormPicker.VB_VarHelpID = -1
Private pSelectedItem As String
Private pCreateNew As Boolean
Private pSelectionMade As Boolean
Private pTitle As String
Private pInstructions As String

'---------------------------------------------------------------
Public Property Let Data(ByVal vNewValue As Recordset)
    Set pData = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get SelectedItem() As String
    SelectedItem = pSelectedItem
End Property

Public Property Let SelectedItem(ByVal vNewValue As String)
    pSelectedItem = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get CreateNew() As Boolean
    CreateNew = pCreateNew
End Property

Public Property Let Title(ByVal vNewValue As String)
    pFormPicker.TtlTop = vNewValue
End Property

Public Property Let Instructions(ByVal vNewValue As String)
    pFormPicker.LblInstructions = vNewValue
End Property

Public Property Let Show(ByVal vNewValue As Boolean)
    If vNewValue Then
        pFormPicker.Show
    End If
End Property

Public Property Let Hide(ByVal vNewValue As Boolean)
    If vNewValue Then
        pFormPicker.Hide
    End If
End Property

' ===============================================================
' Method ClearForm
' Clears form ready to start over
'---------------------------------------------------------------
Public Sub ClearForm()
    With pFormPicker
        .LstResults.Clear
        .TxtSearch = ""
    End With
End Sub

' ===============================================================
' Event pFormPicker_ItemSelected
' Event on select button being pressed.  Selects item in txtsearch and passes it on
'---------------------------------------------------------------
Private Sub pFormPicker_ItemSelected()
    If Not pSelectionMade Then
        MsgBox "Please Select an Item", vbInformation + vbOKOnly
    Else
        pSelectedItem = pFormPicker.TxtSearch
        pFormPicker.Hide
    End If
End Sub

' ===============================================================
' Event pFormPicker_ResultsListSelect
' Event on results list being clicked.  moves selected item to search box
'---------------------------------------------------------------
Private Sub pFormPicker_ResultsListSelect()
    With pFormPicker
        If .LstResults.ListIndex <> -1 Then
            .TxtSearch = .LstResults.List(.LstResults.ListIndex)
            pSelectionMade = True
        End If
    End With
End Sub

' ===============================================================
' Event pformpicker_CreateNew
' Event on new button being pressed
'---------------------------------------------------------------
Private Sub pformpicker_CreateNew()
    pCreateNew = True
End Sub

' ===============================================================
' Method SearchTextChanged
' Event on search text change.  lists all matching data in list box
'---------------------------------------------------------------
Private Sub pFormPicker_SearchTextChanged()
    Dim SearchTxt As String
    Dim FieldName As String
    
    SearchTxt = pFormPicker.TxtSearch
    
    If SearchTxt = "" Then
        Do While Not pData.EOF
            With pFormPicker.LstResults
                If Not IsNull(pData.Fields(0)) Then .AddItem pData.Fields(0)
            End With
            pData.MoveNext
        Loop
    Else
        FieldName = pData.Fields(0).Name
        If pData.RecordCount > 0 Then
        pData.MoveFirst
        pFormPicker.LstResults.Clear
        Do While Not pData.EOF
            If InStr(UCase(pData.Fields(0)), UCase(SearchTxt)) Then
                With pFormPicker.LstResults
                    If Not IsNull(pData.Fields(0)) Then .AddItem pData.Fields(0)
                End With
            End If
            pData.MoveNext
        Loop
        End If
    End If
End Sub

' ===============================================================
Private Sub Class_Initialize()
    Set pFormPicker = FrmPicker
End Sub

Private Sub Class_Terminate()
    Set pFormPicker = Nothing
End Sub
' ===============================================================
