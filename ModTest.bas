Attribute VB_Name = "ModTest"
Option Explicit


Public Sub TestClass()
    Dim Clients As ClsClients
    Dim Client As ClsClient
    Dim i
    
    Set Clients = New ClsClients
    
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
        
Debug.Assert Clients.Count = i
    
    Next

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

    Stop
    
    DB.Execute "DELETE * FROM TblClient"
    
    Set Client = Nothing
    Set Clients = Nothing
    
End Sub


