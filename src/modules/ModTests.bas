Attribute VB_Name = "ModTests"
Option Compare Database
Option Explicit

Public Sub TestAll()

    TestGenerateUUID4
    TestGenerateUUID7
    TestTimeStamps
    TestWriteEvent
    TestReadEventById
    TestReadEventsBySubject
    TestReadAllEventsAfter
    TestReadAllEventsAfter
    TestEventPayloadAddAttr
    TestEventPayloadGetAttributes
    TestEventPayloadAttribNotFound
    TestEventPayloadReadXml
    TestCoopAcceptMembership
    TestErrorDoubleCoopAcceptMembership
    TestCoopChangeName
    TestReplayCoop

End Sub

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

    On Error GoTo ErrTestWriteEvent

    Dim evtStore As ClsEventStore
    Set evtStore = New ClsEventStore
    
    Dim ws As DAO.Workspace
    Set ws = DBEngine.CreateWorkspace("MyWorkspace", "Admin", "")
    Call evtStore.Init(ws, CurrentProject.Path & "\EventSourcingBackend.accdb")
    
    ws.Databases(0).Execute "DELETE FROM domain_events"
    
    Dim Subject As String
    Subject = ModUUIDGenerator.GenerateUUIDv7
    
    Dim Evt As ClsDomainEvent
    Set Evt = New ClsDomainEvent
    Call Evt.Init("user", "user.registered", "machine-x", Subject)
    Evt.AddEventSource "DlsMember"
    
    Dim evt2 As ClsDomainEvent
    Set evt2 = New ClsDomainEvent
    Call evt2.Init("user", "user.changed", "machine-x", Subject)
    evt2.AddEventSource "DlsMember"
    evt2.Version = Evt.Version + 1
    
    Dim Events As Collection: Set Events = New Collection
    Call Events.Add(Evt)
    Call Events.Add(evt2)
    
    evtStore.AddEvents Events
    
    Dim Assert As ClsAsserts: Set Assert = New ClsAsserts
    Assert.CountEquals "TestWriteEvent", 2, evtStore.ReadAllEventsAfter(0).Count
    
    Exit Sub
    
ErrTestWriteEvent:

    Dim errNo As Long: errNo = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    Dim errSrc As String: errSrc = Err.Source

    On Error Resume Next
    evtStore.CloseEventStore
    ws.Close
    
    Set evtStore = Nothing
    Set ws = Nothing
    
    Err.Raise errNo, errSrc, errMsg

End Sub

Public Sub TestReadEventById()

    On Error GoTo ErrTestReadEventById

    Dim evtStore As ClsEventStore
    Set evtStore = New ClsEventStore
    
    Dim ws As DAO.Workspace
    Set ws = DBEngine.CreateWorkspace("MyWorkspace", "Admin", "")
    Call evtStore.Init(ws, CurrentProject.Path & "\EventSourcingBackend.accdb")
    
    ws.Databases(0).Execute "DELETE FROM domain_events"
    
    Dim Subject As String
    Subject = ModUUIDGenerator.GenerateUUIDv7
    
    Dim Evt As ClsDomainEvent
    Set Evt = New ClsDomainEvent
    Call Evt.Init("user.registered", "machine-x", Subject)
    Evt.AddEventSource "DlsManager"
    
    Dim evtId As String
    evtId = Evt.EventId
    
    Dim evt2 As ClsDomainEvent
    Set evt2 = New ClsDomainEvent
    Call evt2.Init("user.changed", "machine-x", Subject)
    evt2.AddEventSource "DlsManager"
    evt2.Version = Evt.Version + 1
    
    Dim Events As Collection: Set Events = New Collection
    Call Events.Add(Evt)
    Call Events.Add(evt2)
    
    evtStore.AddEvents Events
    
    Dim EvtFromDb As ClsDomainEvent
    Set EvtFromDb = evtStore.ReadEventById(evtId)
    
    Exit Sub
    
ErrTestReadEventById:
    Dim errNo As Long: errNo = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    Dim errSrc As String: errSrc = Err.Source
    
    On Error Resume Next
    evtStore.CloseEventStore
    ws.Close
    
    Set evtStore = Nothing
    Set ws = Nothing
    
    Err.Raise errNo, errSrc, errMsg

End Sub

Public Sub TestReadEventsBySubject()

    On Error GoTo ErrTestReadEventsBySubject

    Dim Asserts As ClsAsserts
    Set Asserts = New ClsAsserts

    Dim evtStore As ClsEventStore
    Set evtStore = New ClsEventStore
    
    Dim ws As DAO.Workspace
    Set ws = DBEngine.CreateWorkspace("MyWorkspace", "Admin", "")
    Call evtStore.Init(ws, CurrentProject.Path & "\EventSourcingBackend.accdb")
    
    ws.Databases(0).Execute "DELETE FROM domain_events"
    
    Dim Subject As String
    Subject = ModUUIDGenerator.GenerateUUIDv7
    
    Dim Evt As ClsDomainEvent
    Set Evt = New ClsDomainEvent
    Call Evt.Init("user", "user.registered", Subject)
    Evt.AddEventSource "DlsManager"
    
    Dim evtId As String
    evtId = Evt.EventId
    
    Dim evt2 As ClsDomainEvent
    Set evt2 = New ClsDomainEvent
    Call evt2.Init("user", "user.changed", Subject)
    evt2.AddEventSource "DlsManager"
    evt2.Version = Evt.Version + 1
    
    Dim Events As Collection: Set Events = New Collection
    Call Events.Add(Evt)
    Call Events.Add(evt2)
    
    evtStore.AddEvents Events
    
    Dim EvtFromDb As ClsDomainEvent
    Dim EventsFromDb As Collection
    Set EventsFromDb = evtStore.ReadEventsBySubject(Subject)
    
    Asserts.CountEquals "ModTests.TestReadEventsBySubject", 2, EventsFromDb.Count
    
    evtStore.CloseEventStore
    ws.Close
    Set evtStore = Nothing
    Set ws = Nothing
    Exit Sub
    
ErrTestReadEventsBySubject:
    Dim errNo As Long: errNo = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    Dim errSrc As String: errSrc = Err.Source

    On Error Resume Next
    evtStore.CloseEventStore
    ws.Close
    
    Set evtStore = Nothing
    Set ws = Nothing
    
    Err.Raise errNo, errSrc, errMsg
    
End Sub

Public Sub TestReadAllEventsAfter()

    On Error GoTo ErrTestReadAllEventsAfter

    Dim Asserts As ClsAsserts
    Set Asserts = New ClsAsserts

    Dim evtStore As ClsEventStore
    Set evtStore = New ClsEventStore
    
    Dim ws As DAO.Workspace
    Set ws = DBEngine.CreateWorkspace("MyWorkspace", "Admin", "")
    Call evtStore.Init(ws, CurrentProject.Path & "\EventSourcingBackend.accdb")
   
    ws.Databases(0).Execute "DELETE FROM domain_events"

    Dim Subject As String
    Subject = ModUUIDGenerator.GenerateUUIDv7
    
    Dim Evt As ClsDomainEvent
    Set Evt = New ClsDomainEvent
    Call Evt.Init("user", "user.registered", Subject)
    Evt.AddEventSource "DlsManager"
    
    Dim evtId As String
    evtId = Evt.EventId
    
    Dim evt2 As ClsDomainEvent
    Set evt2 = New ClsDomainEvent
    Call evt2.Init("user", "user.changed", Subject)
    evt2.AddEventSource "DlsManager"
    evt2.Version = Evt.Version + 1
    
    Dim Events As Collection: Set Events = New Collection
    Call Events.Add(Evt)
    Call Events.Add(evt2)
    
    evtStore.AddEvents Events
    
    Dim EventList As Collection
    Set EventList = evtStore.ReadAllEventsAfter(1)
    Asserts.CountEquals "ModTests.TestReadAllEventsAfter", 2, EventList.Count
    
    evtStore.CloseEventStore
    ws.Close
    
    Set evtStore = Nothing
    Set ws = Nothing
    Exit Sub
    
ErrTestReadAllEventsAfter:
    Dim errNo As Long: errNo = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    Dim errSrc As String: errSrc = Err.Source

    On Error Resume Next
    evtStore.CloseEventStore
    ws.Close
    
    Set evtStore = Nothing
    Set ws = Nothing
    
    Err.Raise errNo, errSrc, errMsg
    
End Sub

Public Sub TestEventPayloadAddAttr()

    Dim Payload As ClsEventPayload
    Set Payload = New ClsEventPayload
    
    Payload.AddAttribute "UserName", "Matthias"
    Payload.AddAttribute "UserName", "Sandra"
    Payload.AddAttribute "Role", "Admin"
    
    Debug.Print Payload.AsXML.xml

End Sub

Public Sub TestEventPayloadGetAttributes()

    Dim Payload As ClsEventPayload
    Set Payload = New ClsEventPayload
    
    Payload.AddAttribute "UserName", "Matthias"
    Payload.AddAttribute "UserName", "Sandra"
    Payload.AddAttribute "Role", "Admin"
    
    Dim Attributes As Collection
    Set Attributes = Payload.GetAttributes

    Dim Assert As ClsAsserts
    Set Assert = New ClsAsserts
    Assert.CountEquals "TestEventPayloadGetAttributes", 2, Attributes.Count
    
    Dim i As Long
    For i = 1 To Attributes.Count
        Debug.Print Attributes(i), Payload.GetAttribute(Attributes(i))
    Next i

End Sub

Public Sub TestEventPayloadAttribNotFound()

    Dim Payload As ClsEventPayload
    Set Payload = New ClsEventPayload
    
    Payload.AddAttribute "UserName", "Matthias"
    Payload.AddAttribute "UserName", "Sandra"
    Payload.AddAttribute "Role", "Admin"
    
    Dim Attributes As Collection
    Set Attributes = Payload.GetAttributes

    Dim Assert As ClsAsserts
    Set Assert = New ClsAsserts
    
    On Error GoTo ErrorFound
    Payload.GetAttribute "roles"
    
    On Error GoTo 0
    Assert.ErrExpected "TestEventPayloadAttribNotFound"
    
ErrorFound:
    
End Sub

Public Sub TestEventPayloadReadXml()

    Dim Payload As ClsEventPayload
    Set Payload = New ClsEventPayload
    
    Payload.AddAttribute "UserName", "Matthias"
    Payload.AddAttribute "UserName", "Sandra"
    Payload.AddAttribute "Role", "Admin"
    
    Dim xml As String
    xml = Payload.AsXML.xml
    
    Dim Payload2 As ClsEventPayload
    Set Payload2 = New ClsEventPayload
    Payload2.ReadXml xml
    
    Dim Attributes As Collection
    Set Attributes = Payload2.GetAttributes
    
    Dim i As Long
    For i = 1 To Attributes.Count
        Debug.Print Attributes(i), Payload2.GetAttribute(Attributes(i))
    Next i

End Sub


Public Sub TestCoopAcceptMembership()

    On Error GoTo ErrTestCoopAcceptMembership
        
    Dim Asserts As ClsAsserts
    Set Asserts = New ClsAsserts

    Dim evtStore As ClsEventStore
    Set evtStore = New ClsEventStore
    
    Dim ws As DAO.Workspace
    Set ws = DBEngine.CreateWorkspace("MyWorkspace", "Admin", "")
    Call evtStore.Init(ws, CurrentProject.Path & "\EventSourcingBackend.accdb")
   
    ws.Databases(0).Execute "DELETE FROM domain_events"
    
    Dim Member As ClsCoopMember
    Set Member = New ClsCoopMember
    
    Dim PostalAddr As ClsPostalAddress
    Set PostalAddr = New ClsPostalAddress
    With PostalAddr
        .City = "Schorndorf"
        .CountryCode = "DE"
        .Street = "Schurwaldstr."
        .StreetNumber = 105
        .ZipCode = "73614"
    End With
    Member.AcceptMembership "Hotzel", "Matthias", "Herr", _
        PostalAddr, _
        4, DateValue("2026-05-01")
        
    Dim Events As Collection
    Set Events = Member.IAggregateRoot_ReleaseEvents
    Asserts.CountEquals "TestCoopAcceptMembership", 1, Events.Count
    
    Dim Evt As ClsDomainEvent
    Set Evt = Events(1)
    
    Dim XmlString As String: XmlString = Evt.Data
    
    Debug.Print XmlString
    
    Exit Sub

    
ErrTestCoopAcceptMembership:
    
    Dim errNo As Long: errNo = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    Dim errSrc As String: errSrc = Err.Source
    
    On Error Resume Next
    
    evtStore.CloseEventStore
    ws.Close
    
    Err.Raise errNo, errSrc, errSrc & " - " & errMsg
    
End Sub

Public Sub TestErrorDoubleCoopAcceptMembership()
       
    Dim Asserts As ClsAsserts
    Set Asserts = New ClsAsserts

    
    Dim Member As ClsCoopMember
    Set Member = New ClsCoopMember
    
    Dim PostalAddr As ClsPostalAddress
    Set PostalAddr = New ClsPostalAddress
    With PostalAddr
        .City = "Schorndorf"
        .CountryCode = "DE"
        .Street = "Schurwaldstr."
        .ZipCode = "73614"
        .StreetNumber = 105
    End With
    
    Member.AcceptMembership _
        "Hotzel", "Matthias", "Herr", _
        PostalAddr, _
        4, DateValue("2026-05-01")
    
    On Error GoTo ErrFound
        
    Member.AcceptMembership _
        "Hotzel", "Matthias", "Herr", _
        PostalAddr, _
        4, DateValue("2026-05-01")
    
    On Error GoTo 0
    
    Asserts.ErrExpected "TestErrorDoubleCoopAcceptMembership"
    Exit Sub

    
ErrFound:
    
    
End Sub

Public Sub TestCoopChangeName()
    
    Dim Asserts As ClsAsserts: Set Asserts = New ClsAsserts
    
    Dim Member As ClsCoopMember
    Set Member = New ClsCoopMember
    
    Dim PostalAddr As ClsPostalAddress
    Set PostalAddr = New ClsPostalAddress
    With PostalAddr
        .City = "Schorndorf"
        .Street = "Schurwaldstr."
        .StreetNumber = 105
        .ZipCode = "73614"
    End With
    
    Member.AcceptMembership "Hotzel", "Matthias", "Herr", _
        PostalAddr, _
        4, DateSerial(2026, 3, 1)
    
    Asserts.Equals "TestCoopChangeName", "Hotzel", Member.LastName
    
    Member.ChangeName "Lenker"
    Asserts.Equals "TestCoopChangeName", "Lenker", Member.LastName
    
    Dim ws As DAO.Workspace
    Set ws = DBEngine.CreateWorkspace("MyWorkspace", "Admin", "")
    
    Dim evtStore As ClsEventStore
    Set evtStore = New ClsEventStore
    evtStore.Init ws, CurrentProject.Path & "\EventSourcingBackend.accdb"
    
    Dim MemberId As String: MemberId = Member.MemberId
    
    evtStore.AddEvents Member.IAggregateRoot_ReleaseEvents

End Sub

Public Sub TestReplayCoop()
    
    Dim Asserts As ClsAsserts: Set Asserts = New ClsAsserts
    
    Dim Member As ClsCoopMember
    Set Member = New ClsCoopMember
    
    Dim PostalAddr As ClsPostalAddress
    Set PostalAddr = New ClsPostalAddress
    With PostalAddr
        .City = "Schorndorf"
        .Street = "Schurwaldstr."
        .StreetNumber = 105
        .ZipCode = "73614"
    End With
    
    Member.AcceptMembership "Hotzel", "Matthias", "Herr", _
        PostalAddr, _
        4, DateSerial(2026, 3, 1)
    
    Asserts.Equals "TestCoopChangeName", "Hotzel", Member.LastName
    
    Member.ChangeName "Lenker"
    Asserts.Equals "TestCoopChangeName", "Lenker", Member.LastName
    
    Dim ws As DAO.Workspace
    Set ws = DBEngine.CreateWorkspace("MyWorkspace", "Admin", "")
    
    Dim evtStore As ClsEventStore
    Set evtStore = New ClsEventStore
    evtStore.Init ws, CurrentProject.Path & "\EventSourcingBackend.accdb"
    
    ws.Databases(0).Execute "DELETE FROM domain_events"
    
    Dim MemberId As String: MemberId = Member.MemberId
    
    evtStore.AddEvents Member.IAggregateRoot_ReleaseEvents
    
    Dim Member2 As ClsCoopMember: Set Member2 = New ClsCoopMember
    evtStore.Replay Member2, evtStore.ReadEventsBySubject(MemberId)

    Asserts.Equals "TestReplayCoop", "Lenker", Member2.LastName
    
End Sub
