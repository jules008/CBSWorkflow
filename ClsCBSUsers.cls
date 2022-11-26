VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsCBSUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsCBSUsers
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 10 Oct 22
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

Private pCBSUsers As Collection


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pCBSUsers.[_NewEnum]
End Function

' ===============================================================
' Method Find
' Finds Item from index
' ---------------------------------------------------------------
Public Function Find(CBSUserNo As Variant) As ClsCBSUser
Attribute Find.VB_UserMemId = 0
    On Error Resume Next
    Set Find = pCBSUsers.item(CBSUserNo)
End Function

' ===============================================================
' Method Add
' Add item to Collection
' ---------------------------------------------------------------
Public Sub Add(CBSUser As ClsCBSUser)
    pCBSUsers.Add CBSUser, Key:=CStr(CBSUser.CBSUserNo)
End Sub

' ===============================================================
' Method Remove
' Removes item from collection
' ---------------------------------------------------------------
Public Sub Remove(CBSUserNo As Variant)
    pCBSUsers.Remove CBSUserNo
End Sub

' ===============================================================
' Method Count
' Returns number of items in collection
' ---------------------------------------------------------------
Public Function Count() As Integer
    Count = pCBSUsers.Count
End Function

' ===============================================================
' Method GetCollection
' Gets all items in collection
' ---------------------------------------------------------------
Public Sub GetCollection()
    Dim RstCBSUser As Recordset
    Dim CBSUser As ClsCBSUser
    Dim i As Integer

    Set RstCBSUser = ModDatabase.SQLQuery("TblCBSUser")
    With RstCBSUser
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            For i = 1 To .RecordCount
                Set CBSUser = New ClsCBSUser
                CBSUser.DBGet !CBSUserNo
                pCBSUsers.Add CBSUser, Key:=CStr(CBSUser.CBSUserNo)
                .MoveNext
            Next
        End If
    End With
End Sub

' ===============================================================
' Method SaveCollection
' Saves all items in collection to database
' ---------------------------------------------------------------
Public Sub SaveCollection()
    Dim CBSUser As ClsCBSUser
    For Each CBSUser In pCBSUsers
        CBSUser.DBSave
    Next
End Sub

' ===============================================================
' Method RemoveCollection
' Deletes all items in collection
' ---------------------------------------------------------------
Public Sub RemoveCollection()
    Dim CBSUser As ClsCBSUser
    For Each CBSUser In pCBSUsers
        pCBSUsers.Remove CStr(CBSUser.CBSUserNo)
    Next
End Sub

' ===============================================================
' Method DeleteCollection
' Deletes all items in collection
' ---------------------------------------------------------------
Public Sub DeleteCollection()
    Dim CBSUser As ClsCBSUser
    For Each CBSUser In pCBSUsers
        pCBSUsers.Remove CStr(CBSUser.CBSUserNo)
        CBSUser.DBDelete
    Next
End Sub

' ===============================================================
' Method Destroy
' Removes item from collection and destroys it
' ---------------------------------------------------------------
Public Sub Destroy(CBSUserNo As Variant)
    Dim CBSUser As ClsCBSUser

    Set CBSUser = pCBSUsers(CBSUserNo)
    pCBSUsers.Remove CBSUserNo
    CBSUser.DBDelete True

    Set CBSUser = Nothing
End Sub
' ===============================================================

Private Sub Class_Initialize()
    Set pCBSUsers = New Collection

End Sub
' ===============================================================
Private Sub Class_Terminate()
    Set pCBSUsers = Nothing
End Sub
' ===============================================================
