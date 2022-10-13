VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsLenders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsLenders
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 13 Oct 22
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

Private pLenders As Collection


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pLenders.[_NewEnum]
End Function

' ===============================================================
' Method Find
' Finds Item from index
' ---------------------------------------------------------------
Public Function Find(LenderNo As Variant) As ClsLender
Attribute Find.VB_UserMemId = 0
    On Error Resume Next
    Set Find = pLenders.Item(LenderNo)
End Function

' ===============================================================
' Method Add
' Add item to Collection
' ---------------------------------------------------------------
Public Sub Add(Lender As ClsLender)
    pLenders.Add Lender, Key:=CStr(Lender.LenderNo)
End Sub

' ===============================================================
' Method Remove
' Removes item from collection
' ---------------------------------------------------------------
Public Sub Remove(LenderNo As Variant)
    pLenders.Remove LenderNo
End Sub

' ===============================================================
' Method Count
' Returns number of items in collection
' ---------------------------------------------------------------
Public Function Count() As Integer
    Count = pLenders.Count
End Function

' ===============================================================
' Method GetCollection
' Gets all items in collection
' ---------------------------------------------------------------
Public Sub GetCollection()
    Dim RstLender As Recordset
    Dim Lender As ClsLender
    Dim i As Integer

    Set RstLender = ModDatabase.SQLQuery("TblLender")
    With RstLender
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            For i = 1 To .RecordCount
                Set Lender = New ClsLender
                Lender.DBGet !LenderNo
                pLenders.Add Lender, Key:=CStr(Lender.LenderNo)
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
    Dim Lender As ClsLender
    For Each Lender In pLenders
        Lender.DBSave
    Next
End Sub

' ===============================================================
' Method RemoveCollection
' Deletes all items in collection
' ---------------------------------------------------------------
Public Sub RemoveCollection()
    Dim Lender As ClsLender
    For Each Lender In pLenders
        pLenders.Remove CStr(Lender.LenderNo)
    Next
End Sub

' ===============================================================
' Method DeleteCollection
' Deletes all items in collection
' ---------------------------------------------------------------
Public Sub DeleteCollection()
    Dim Lender As ClsLender
    For Each Lender In pLenders
        pLenders.Remove CStr(Lender.LenderNo)
        Lender.DBDelete
    Next
End Sub

' ===============================================================
' Method Destroy
' Removes item from collection and destroys it
' ---------------------------------------------------------------
Public Sub Destroy(LenderNo As Variant)
    Dim Lender As ClsLender

    Set Lender = pLenders(LenderNo)
    pLenders.Remove LenderNo
    Lender.DBDelete True

    Set Lender = Nothing
End Sub
' ===============================================================

Private Sub Class_Initialize()
    Set pLenders = New Collection

End Sub
' ===============================================================
Private Sub Class_Terminate()
    Set pLenders = Nothing
End Sub
' ===============================================================
