Attribute VB_Name = "ModTests"
Option Compare Database
Option Explicit

Public Sub TestGenerateUUID4()

    Dim uuid As String
    uuid = ModUUIDGenerator.GenerateUUIDv4

    Debug.Print uuid

End Sub

Public Sub TestGenerateUUID7()

    Dim uuid As String
    uuid = ModUUIDGenerator.GenerateUUIDv7

    Debug.Print uuid

End Sub

Public Sub TestTimeStamps()
    
    Dim ts As String
    ts = ModTimeStampGenerator.GetUtcIsoTimestamp
    Debug.Print ts

End Sub


Public Sub TestWriteEvent()

    Dim evtStore As ClsEventStore
    Set evtStore = New ClsEventStore
    
    Dim ws As DAO.Workspace
    Set ws = DBEngine.CreateWorkspace("MyWorkspace", "Admin", "")
    Call evtStore.Init(ws, CurrentProject.Path & "\EventSourcingBackend.accdb")
    
    Dim subject As String
    subject = ModUUIDGenerator.GenerateUUIDv7
    
    Dim evt As ClsDomainEvent
    Set evt = New ClsDomainEvent
    Call evt.Init("user.registered", "machine-x", subject)
    
    Dim evt2 As ClsDomainEvent
    Set evt2 = New ClsDomainEvent
    Call evt2.Init("user.changed", "machine-x", subject)
    evt2.Version = evt.Version + 1
    
    Dim events As Collection: Set events = New Collection
    Call events.Add(evt)
    Call events.Add(evt2)
    
    evtStore.AddEvents events
   
    'On Error Resume Next
    evtStore.CloseEventStore
    Dim db As DAO.Database
    ws.Close
    Set evtStore = Nothing
    Set ws = Nothing

End Sub

