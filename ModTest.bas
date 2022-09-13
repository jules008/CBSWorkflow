Attribute VB_Name = "ModTest"
Option Explicit


Public Sub TestClass()
    Dim Clients As ClsClients
    Dim Client As ClsClient
    Dim Spv As ClsSPV
    Dim Project As ClsProject
    Dim Contact As ClsContact
    Dim i
    
    Set Clients = New ClsClients
    Set Spv = New ClsSPV
    Set Project = New ClsProject
    
Debug.Assert Not Clients Is Nothing
    
    For i = 1 To 5
        Set Client = New ClsClient
        With Client
            .Name = "Frogspawn Ltd - " & i
            .ClientNo = i
            .PhoneNo = "14234242"
            .Url = "www.hockey.com"
            .DBSave
        End With
        Clients.Add Client
        
        Set Spv = New ClsSPV
        With Spv
            .Name = "Frogspawn Ltd - " & i
            .SPVNo = i
            .DBSave
        End With
        Clients(1).SPVs.Add Spv

Debug.Assert Clients.Count = i
Debug.Assert Clients(1).SPVs.Count = i
Next
    
    Set Contact = New ClsContact
    
    With Contact
        .ContactName = "Dom"
    End With
    
    With Project
        .ExitFee = True
        .ProjectNo = 1
    End With

    Client.Contacts.Add Contact
    Clients("1").SPVs("1").Contacts.Add Contact
    Clients("1").SPVs("1").Projects.Add Project
    Set Client = Nothing
    
    Set Client = Clients("2")
            
Debug.Assert Client.ClientNo = 2

    With Client
        Debug.Print .Name, .ClientNo, .PhoneNo, .Url
        
        .Name = "New Name"
        .DBSave
        Debug.Print .Name, .ClientNo, .PhoneNo, .Url
    End With
    
    Clients.RemoveCollection
    
Debug.Assert Clients.Count = 0

    Set Clients = Nothing
    
Debug.Assert Clients Is Nothing

    Set Clients = New ClsClients
    
    Clients.GetCollection
    
    Stop
Debug.Assert Clients.Count = 5

    Clients.Remove 4

Debug.Assert Clients.Count = 4

    Clients.RemoveCollection
    Clients.GetCollection
    
Debug.Assert Clients.Count = 5

    Clients.Destroy 4
    
Debug.Assert Clients.Count = 4

    Clients.RemoveCollection
    Clients.GetCollection
    
Debug.Assert Clients.Count = 4
    
    For Each Client In Clients
        Debug.Print Client.ClientNo
    Next
    
    Clients.DeleteCollection
    
Debug.Assert Clients.Count = 0

    
    DB.Execute "DELETE * FROM TblClient"
    
    Set Client = Nothing
    Set Clients = Nothing
    Set Spv = Nothing
    Set Contact = Nothing
    Set Project = Nothing
End Sub


