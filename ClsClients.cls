VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsClients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsClients
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 13 Sep 22
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

Private pClients As Collection


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pClients.[_NewEnum]
End Function

' ===============================================================
' Method Find
' Finds Item from index
' ---------------------------------------------------------------
Public Function Find(ClientNo As Variant) As ClsClient
Attribute Find.VB_UserMemId = 0
'    On Error Resume Next
    Set Find = pClients.Item(ClientNo)
End Function

' ===============================================================
' Method Add
' Add item to Collection
' ---------------------------------------------------------------
Public Sub Add(Client As ClsClient)
    pClients.Add Client, Key:=CStr(Client.ClientNo)
End Sub

' ===============================================================
' Method Remove
' Removes item from collection
' ---------------------------------------------------------------
Public Sub Remove(ClientNo As Variant)
    pClients.Remove ClientNo
End Sub

' ===============================================================
' Method Count
' Returns number of items in collection
' ---------------------------------------------------------------
Public Function Count() As Integer
    Count = pClients.Count
End Function

' ===============================================================
' Method GetCollection
' Gets all items in collection
' ---------------------------------------------------------------
Public Sub GetCollection()
    Dim RstClient As Recordset
    Dim Client As ClsClient
    Dim i As Integer

    Set RstClient = ModDatabase.SQLQuery("TblClient")
    With RstClient
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            For i = 1 To .RecordCount
                Set Client = New ClsClient
                Client.DBGet !ClientNo
                pClients.Add Client, Key:=CStr(Client.ClientNo)
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
    Dim Client As ClsClient
    For Each Client In pClients
        Client.DBSave
    Next
End Sub

' ===============================================================
' Method RemoveCollection
' Deletes all items in collection
' ---------------------------------------------------------------
Public Sub RemoveCollection()
    Dim Client As ClsClient
    For Each Client In pClients
        pClients.Remove CStr(Client.ClientNo)
    Next
End Sub

' ===============================================================
' Method DeleteCollection
' Deletes all items in collection
' ---------------------------------------------------------------
Public Sub DeleteCollection()
    Dim Client As ClsClient
    For Each Client In pClients
        pClients.Remove CStr(Client.ClientNo)
        Client.DBDelete
    Next
End Sub

' ===============================================================
' Method Destroy
' Removes item from collection and destroys it
' ---------------------------------------------------------------
Public Sub Destroy(ClientNo As Variant)
    Dim Client As ClsClient

    Set Client = pClients(ClientNo)
    pClients.Remove ClientNo
    Client.DBDelete True

    Set Client = Nothing
End Sub
' ===============================================================

Private Sub Class_Initialize()
    Set pClients = New Collection

End Sub
' ===============================================================
Private Sub Class_Terminate()
    Set pClients = Nothing
End Sub
' ===============================================================
