VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsSPVs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsSPVs
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 12 Sep 22
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

Private pSPVs As Collection
Private pParent As ClsClient

'---------------------------------------------------------------
Public Property Get Parent() As ClsClient
    Set Parent = pParent
End Property

Public Property Set Parent(obj As ClsClient)
    Set pParent = obj
End Property
'---------------------------------------------------------------

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pSPVs.[_NewEnum]
End Function

' ===============================================================
' Method Find
' Finds Item from index
' ---------------------------------------------------------------
Public Function Find(SPVNo As Variant) As ClsSPV
Attribute Find.VB_UserMemId = 0
'    On Error Resume Next
    Set Find = pSPVs.Item(SPVNo)
End Function

' ===============================================================
' Method Add
' Add item to Collection
' ---------------------------------------------------------------
Public Sub Add(Spv As ClsSPV)
    Spv.Initialize Me.Parent
    pSPVs.Add Spv, Key:=CStr(Spv.SPVNo)
End Sub

' ===============================================================
' Method Remove
' Removes item from collection
' ---------------------------------------------------------------
Public Sub Remove(SPVNo As Variant)
    pSPVs(SPVNo).Parent = Nothing
    pSPVs.Remove SPVNo
End Sub

' ===============================================================
' Method Count
' Returns number of items in collection
' ---------------------------------------------------------------
Public Function Count() As Integer
    Count = pSPVs.Count
End Function

' ===============================================================
' Method GetCollection
' Gets all items in collection
' ---------------------------------------------------------------
Public Sub GetCollection()
    Dim RstSPV As Recordset
    Dim Spv As ClsSPV
    Dim i As Integer

    Set RstSPV = ModDatabase.SQLQuery("TblSPV")
    With RstSPV
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            For i = 1 To .RecordCount
                Set Spv = New ClsSPV
                Spv.DBGet !SPVNo
                pSPVs.Add Spv, Key:=CStr(Spv.SPVNo)
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
    Dim Spv As ClsSPV
    For Each Spv In pSPVs
        Spv.DBSave
    Next
End Sub

' ===============================================================
' Method RemoveCollection
' Deletes all items in collection
' ---------------------------------------------------------------
Public Sub RemoveCollection()
    Dim Spv As ClsSPV
    For Each Spv In pSPVs
        pSPVs.Remove CStr(Spv.SPVNo)
    Next
End Sub

' ===============================================================
' Method DeleteCollection
' Deletes all items in collection
' ---------------------------------------------------------------
Public Sub DeleteCollection()
    Dim Spv As ClsSPV
    For Each Spv In pSPVs
        pSPVs.Remove CStr(Spv.SPVNo)
        Spv.DBDelete
    Next
End Sub

' ===============================================================
' Method Destroy
' Removes item from collection and destroys it
' ---------------------------------------------------------------
Public Sub Destroy(SPVNo As Variant)
    Dim Spv As ClsSPV

    Set Spv = pSPVs(SPVNo)
    Spv.Parent = Nothing
    Spv.Terminate
    pSPVs.Remove SPVNo
    Spv.DBDelete True

    Set Spv = Nothing
End Sub
' ===============================================================
Public Sub Initialize(obj As ClsClient)
    Set Me.Parent = obj
End Sub
'---------------------------------------------------------------

Private Sub Class_Initialize()
    Set pSPVs = New Collection

End Sub
' ===============================================================
Public Sub Terminate()

    Set Me.Parent = Nothing

End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pSPVs = Nothing
End Sub
' ===============================================================