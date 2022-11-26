VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUIStyles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUIStyles
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 23 Sep 22
'===============================================================
' Methods
'---------------------------------------------------------------
' Find - Finds Item from index
' Add - Add item to Collection
' Remove - Removes item from collection
' Count - Returns number of items in collection
' GetCollection - Gets all items in collection
' SaveCollection - Saves all items in collection to database
' RemoveCollection - Removes all items in collection
' DeleteCollection - Deletes all items in collection
' Destroy - Removes item from collection and deletes it
'===============================================================
'---------------------------------------------------------------
Option Explicit

Private pStyles As Collection


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pStyles.[_NewEnum]
End Function

' ===============================================================
' Method Find
' Finds Item from index
' ---------------------------------------------------------------
Public Function Find(Name As String) As ClsUIStyle
    On Error Resume Next
    If Name <> "" Then
        Set Find = pStyles.item(Name)
    End If
End Function

' ===============================================================
' Method Add
' Add item to Collection
' ---------------------------------------------------------------
Public Sub Add(Style As ClsUIStyle)
    pStyles.Add Style, Key:=CStr(Style.Name)
End Sub

' ===============================================================
' Method Remove
' Removes item from collection
' ---------------------------------------------------------------
Public Sub Remove(Name As Variant)
    pStyles.Remove Name
End Sub

' ===============================================================
' Method Count
' Returns number of items in collection
' ---------------------------------------------------------------
Public Function Count() As Integer
    Count = pStyles.Count
End Function

' ===============================================================
' ===============================================================
' Method RemoveCollection
' Deletes all items in collection
' ---------------------------------------------------------------
Public Sub RemoveCollection()
    Dim Style As ClsUIStyle
    For Each Style In pStyles
        pStyles.Remove CStr(Style.Name)
    Next
End Sub

' ===============================================================

Private Sub Class_Initialize()
    Set pStyles = New Collection

End Sub
' ===============================================================
Private Sub Class_Terminate()
    Set pStyles = Nothing
End Sub
' ===============================================================
