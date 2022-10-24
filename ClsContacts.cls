VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsContacts
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

Private pContacts As Collection
Private pPrimaryContact As ClsContact

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pContacts.[_NewEnum]
End Function

'---------------------------------------------------------------
Public Property Get PrimaryContact() As ClsContact
    Set PrimaryContact = pPrimaryContact
End Property

Public Property Let PrimaryContact(ByRef vNewValue As ClsContact)
    Set pPrimaryContact = vNewValue
End Property

' ===============================================================
' Method Find
' Finds Item from index
' ---------------------------------------------------------------
Public Function Find(ContactNo As Variant) As ClsContact
Attribute Find.VB_UserMemId = 0
    On Error Resume Next
    Set Find = pContacts.Item(ContactNo)
End Function

' ===============================================================
' Method Add
' Add item to Collection
' ---------------------------------------------------------------
Public Sub Add(Contact As ClsContact)
    Dim LocContact As ClsContact
    
    Set LocContact = Find(CStr(Contact.ContactNo))
        
    If Contact.PrimaryContact Then Set pPrimaryContact = Contact
    
    If LocContact Is Nothing Then
    pContacts.Add Contact, Key:=CStr(Contact.ContactNo)
    End If
End Sub

' ===============================================================
' Method Remove
' Removes item from collection
' ---------------------------------------------------------------
Public Sub Remove(ContactNo As Variant)
    pContacts.Remove ContactNo
End Sub

' ===============================================================
' Method Count
' Returns number of items in collection
' ---------------------------------------------------------------
Public Function Count() As Integer
    Count = pContacts.Count
End Function

' ===============================================================
' Method GetCollection
' Gets all items in collection
' ---------------------------------------------------------------
Public Sub GetCollection(ContactType As String, OrgName As String)
    Dim RstContact As Recordset
    Dim Contact As ClsContact
    Dim i As Integer

    RemoveCollection
    
    Set RstContact = ModDatabase.SQLQuery("SELECT ContactNo FROM TblContact WHERE ContactType = '" & ContactType & "'" _
                                            & "AND Organisation = '" & OrgName & "'")
    With RstContact
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            For i = 1 To .RecordCount
                Set Contact = New ClsContact
                Contact.DBGet !ContactNo
                pContacts.Add Contact, Key:=CStr(Contact.ContactNo)
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
    Dim Contact As ClsContact
    For Each Contact In pContacts
        Contact.DBSave
    Next
End Sub

' ===============================================================
' Method RemoveCollection
' Deletes all items in collection
' ---------------------------------------------------------------
Public Sub RemoveCollection()
    Dim Contact As ClsContact
    For Each Contact In pContacts
        pContacts.Remove CStr(Contact.ContactNo)
    Next
End Sub

' ===============================================================
' Method DeleteCollection
' Deletes all items in collection
' ---------------------------------------------------------------
Public Sub DeleteCollection()
    Dim Contact As ClsContact
    For Each Contact In pContacts
        pContacts.Remove CStr(Contact.ContactNo)
        Contact.DBDelete
    Next
End Sub

' ===============================================================
' Method Destroy
' Removes item from collection and destroys it
' ---------------------------------------------------------------
Public Sub Destroy(ContactNo As Variant)
    Dim Contact As ClsContact

    Set Contact = pContacts(ContactNo)
    pContacts.Remove ContactNo
    Contact.DBDelete True

    Set Contact = Nothing
End Sub
' ===============================================================

Private Sub Class_Initialize()
    Set pContacts = New Collection
    Set pPrimaryContact = New ClsContact
End Sub
' ===============================================================
Private Sub Class_Terminate()
    Set pContacts = Nothing
    Set pPrimaryContact = Nothing
End Sub
' ===============================================================
