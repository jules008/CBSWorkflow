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
    On Error Resume Next
    Set Find = pSPVs.item(CStr(SPVNo))
    
End Function

' ===============================================================
' Method Add
' Add item to Collection
' ---------------------------------------------------------------
Public Sub Add(ByRef SPV As ClsSPV)
    Dim LocSPV As ClsSPV
    
    Set LocSPV = Find(SPV.SPVNo)
    
    If LocSPV Is Nothing Then
        SPV.Initialize Me.Parent
        pSPVs.Add SPV, Key:=CStr(SPV.SPVNo)
    End If
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
' Gets all items in collection for parent ClientNo
' ---------------------------------------------------------------
Public Sub GetCollection()
    Dim RstSPV As Recordset
    Dim SPV As ClsSPV
    Dim ClientNo As Integer
    Dim i As Integer

    ClientNo = pParent.ClientNo
    
    Set RstSPV = ModDatabase.SQLQuery("SELECT * FROM TblSPV WHERE ClientNo = " & ClientNo)
    With RstSPV
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            For i = 1 To .RecordCount
                Set SPV = New ClsSPV
                SPV.DBGet !SPVNo
                pSPVs.Add SPV, Key:=CStr(SPV.SPVNo)
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
    Dim SPV As ClsSPV
    For Each SPV In pSPVs
        SPV.DBSave
    Next
End Sub

' ===============================================================
' Method RemoveCollection
' Deletes all items in collection
' ---------------------------------------------------------------
Public Sub RemoveCollection()
    Dim SPV As ClsSPV
    For Each SPV In pSPVs
        pSPVs.Remove CStr(SPV.SPVNo)
    Next
End Sub

' ===============================================================
' Method DeleteCollection
' Deletes all items in collection
' ---------------------------------------------------------------
Public Sub DeleteCollection()
    Dim SPV As ClsSPV
    For Each SPV In pSPVs
        pSPVs.Remove CStr(SPV.SPVNo)
        SPV.DBDelete
    Next
End Sub

' ===============================================================
' Method Destroy
' Removes item from collection and destroys it
' ---------------------------------------------------------------
Public Sub Destroy(SPVNo As Variant)
    Dim SPV As ClsSPV

    Set SPV = pSPVs(SPVNo)
    SPV.Parent = Nothing
    SPV.Terminate
    pSPVs.Remove SPVNo
    SPV.DBDelete True

    Set SPV = Nothing
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
