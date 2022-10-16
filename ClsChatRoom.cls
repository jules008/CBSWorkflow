VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsChatRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsChatRoom
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 16 Oct 22
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

Private pChatRoom As Collection
Private WithEvents pChatRoomForm As FrmChatRoom
Attribute pChatRoomForm.VB_VarHelpID = -1

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pChatRoom.[_NewEnum]
End Function

' ===============================================================
' Method Add
' Add item to Collection
' ---------------------------------------------------------------
Public Sub Add(ByVal ChatEntry As ClsChatEntry)
    
    pChatRoom.Add ChatEntry
    
    With ChatEntry
        DB.Execute "INSERT INTO TblChatRoom (ProjectNo, Username, EntryDateTime, Message) " _
                    & "VALUES (" & .ProjectNo & ", '" & .UserName & "', #" & .EntryDateTime & "#, '" & .Message & "')"
                    
    End With
End Sub

' ===============================================================
' Method Count
' Returns number of items in collection
' ---------------------------------------------------------------
Public Function Count() As Integer
    Count = pChatRoom.Count
End Function

' ===============================================================
' Method GetCollection
' Gets all items in collection
' ---------------------------------------------------------------
Private Sub GetCollection(ProjectNo As Integer)
    Dim RstChatRoom As Recordset
    Dim ChatEntry As ClsChatEntry
    Dim i As Integer

    Set RstChatRoom = ModDatabase.SQLQuery("SELECT * FROM TblChatRoom WHERE ProjectNo = " & ProjectNo)
    With RstChatRoom
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            For i = 1 To .RecordCount
                Set ChatEntry = New ClsChatEntry
                
                ChatEntry.ProjectNo = !ProjectNo
                ChatEntry.EntryDateTime = !EntryDateTime
                ChatEntry.UserName = !UserName
                ChatEntry.Message = !Message
                
                pChatRoom.Add ChatEntry
                .MoveNext
            Next
        End If
    End With
End Sub

' ===============================================================
' Method DisplayForm
' displays client form
' ---------------------------------------------------------------
Public Sub DisplayForm(ProjectNo As Integer)
    Dim ChatEntry As ClsChatEntry
    
    GetCollection ProjectNo
    
    With pChatRoomForm
        .ClearForm
        Sin 0
        For Each ChatEntry In pChatRoom
            .xTxtAllMessages = .xTxtAllMessages & ChatEntry.EntryDateTime _
                               & " " & ChatEntry.UserName _
                               & " said """ & ChatEntry.Message & """" _
                               & vbCrLf
            Debug.Print .xTxtAllMessages
        Next
        .ProjectNo = ProjectNo
        .Show
    End With
End Sub

' ===============================================================
' Method RemoveCollection
' Deletes all items in collection
' ---------------------------------------------------------------
Public Sub RemoveCollection()
    Dim ChatEntry As ClsChatEntry
    For Each ChatEntry In pChatRoom
        pChatRoom.Remove ChatEntry
    Next
End Sub

' ===============================================================
' Method DeleteCollection
' Deletes all items in collection
' ---------------------------------------------------------------
Public Sub DeleteCollection()
    Dim ChatEntry As ClsChatEntry
    For Each ChatEntry In pChatRoom
        pChatRoom.Remove ChatEntry
    Next
End Sub

' ===============================================================
' Event pChatRoomForm_SendMessage
' new message has been created
' ---------------------------------------------------------------
Private Sub pChatRoomForm_SendMessage()
    Dim ChatEntry As ClsChatEntry
    
    Set ChatEntry = New ClsChatEntry
    
    With ChatEntry
        .ProjectNo = pChatRoomForm.ProjectNo
        .UserName = Application.UserName
        .Message = pChatRoomForm.xTxtNewMessage
        .EntryDateTime = Now
    End With
    
    With pChatRoomForm
        .xTxtAllMessages = .xTxtAllMessages & ChatEntry.EntryDateTime _
                               & " " & ChatEntry.UserName _
                               & " said """ & ChatEntry.Message & """" _
                               & vbCrLf
    End With

    Add ChatEntry
    
    Set ChatEntry = Nothing
End Sub
' ===============================================================

Private Sub Class_Initialize()
    Set pChatRoom = New Collection
    Set pChatRoomForm = New FrmChatRoom
End Sub
' ===============================================================
Private Sub Class_Terminate()
    Set pChatRoom = Nothing
    Set pChatRoomForm = Nothing
End Sub
' ===============================================================

