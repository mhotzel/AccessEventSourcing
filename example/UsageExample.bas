Attribute VB_Name = "UsageExample"
Option Compare Database
Option Explicit

' =============================================================================
' UsageExample
' Demonstrates how to implement a concrete aggregate and use the EventStore.
'
' This file is NOT meant to be imported into a production database as-is.
' Copy the patterns you need into your own modules.
' =============================================================================

' -----------------------------------------------------------------------------
' Example: run from the Immediate Window or a button click to test the setup
' -----------------------------------------------------------------------------
Public Sub RunExample()
    ' Initialise the store (creates table if it doesn't exist)
    Dim store As New EventStore
    store.Init CurrentDb()

    ' --- Register a new user ---
    Dim user As New UserAggregate
    Dim userId As String
    userId = GenerateUUID()   ' from EventStoreHelper

    user.Register userId, "Alice", "alice@example.com"

    ' Write the recorded events; subject is new, so pass EXPECTED_VERSION_NEW
    store.AddEvents user.ReleaseEvents(), EventStore.EXPECTED_VERSION_NEW

    Debug.Print "User registered, version=" & user.GetVersion()

    ' --- Rebuild the aggregate from the event store ---
    Dim events As Collection
    Set events = store.ReadEventsBySubject(userId)

    Dim rebuilt As New UserAggregate
    rebuilt.Rebuild events

    Debug.Print "Rebuilt user name: " & rebuilt.UserName
    Debug.Print "Rebuilt user email: " & rebuilt.Email

    ' --- Update the user (optimistic locking) ---
    rebuilt.ChangeEmail "alice@newdomain.com"
    store.AddEvents rebuilt.ReleaseEvents(), rebuilt.GetVersion() - 1

    Debug.Print "Email updated, version=" & rebuilt.GetVersion()

    ' --- Read events for a projection ---
    Dim batch As Collection
    Dim checkpoint As Long
    checkpoint = 0

    Set batch = store.ReadAllEventsAfter(checkpoint, 500)

    Dim item As Collection
    For Each item In batch
        Dim pos As Long
        Dim evt As DomainEvent
        pos = item(1)
        Set evt = item(2)
        Debug.Print "position=" & pos & "  type=" & evt.EventType & _
                    "  subject=" & evt.Subject & "  version=" & evt.Version
        checkpoint = pos
    Next item
End Sub

' =============================================================================
' UserAggregate
' A minimal example aggregate. In a real project this would be a separate
' class module (UserAggregate.cls). It is shown here inline so the example
' compiles as a single standard module for illustration purposes only.
'
' In production: create UserAggregate.cls and move this code there.
' =============================================================================

' NOTE: In Access VBA a standard module cannot define class state.
' The code block below is pseudo-code illustrating what UserAggregate.cls
' would look like. Copy it into a new class module named "UserAggregate".

'   VERSION 1.0 CLASS
'   ...
'   Attribute VB_Name = "UserAggregate"
'   Option Compare Database
'   Option Explicit
'
'   ' --- Aggregate state ---
'   Private m_Root    As AggregateRoot
'   Private m_UserId  As String
'   Private m_Name    As String
'   Private m_Email   As String
'
'   Private Sub Class_Initialize()
'       Set m_Root = New AggregateRoot
'   End Sub
'
'   ' --- Read-only state accessors ---
'   Public Property Get UserId()  As String : UserId  = m_UserId  : End Property
'   Public Property Get UserName() As String : UserName = m_Name   : End Property
'   Public Property Get Email()   As String : Email   = m_Email   : End Property
'
'   ' --- Domain methods ---
'
'   ' Registers a new user (acts as a named constructor).
'   Public Sub Register(ByVal userId As String, ByVal name As String, ByVal email As String)
'       Dim keys(2) As String : keys(0) = "userId" : keys(1) = "name" : keys(2) = "email"
'       Dim vals(2) As String : vals(0) = userId   : vals(1) = name   : vals(2) = email
'       Dim evt As New DomainEvent
'       evt.Init "user.registered", userId, XmlBuildData(keys, vals)
'       RecordThat evt
'   End Sub
'
'   ' Changes the user's email address.
'   Public Sub ChangeEmail(ByVal newEmail As String)
'       Dim keys(0) As String : keys(0) = "email"
'       Dim vals(0) As String : vals(0) = newEmail
'       Dim evt As New DomainEvent
'       evt.Init "user.emailChanged", m_UserId, XmlBuildData(keys, vals)
'       RecordThat evt
'   End Sub
'
'   ' --- Rebuild from event stream ---
'   Public Sub Rebuild(ByVal events As Collection)
'       Dim evt As DomainEvent
'       For Each evt In events
'           m_Root.SetVersion evt.Version
'           Apply evt
'       Next evt
'   End Sub
'
'   ' --- Internal: record, apply, queue ---
'   Private Sub RecordThat(ByVal evt As DomainEvent)
'       m_Root.PrepareEvent evt   ' stamps the next version number onto the event
'       Apply evt                 ' update in-memory state
'       m_Root.QueueEvent evt     ' queue for writing
'   End Sub
'
'   ' Routes an event to the matching state-update handler.
'   Private Sub Apply(ByVal evt As DomainEvent)
'       Select Case evt.EventType
'           Case "user.registered"
'               m_UserId = XmlGetValue(evt.DataXml, "userId")
'               m_Name   = XmlGetValue(evt.DataXml, "name")
'               m_Email  = XmlGetValue(evt.DataXml, "email")
'           Case "user.emailChanged"
'               m_Email  = XmlGetValue(evt.DataXml, "email")
'       End Select
'   End Sub
'
'   ' --- Delegation to AggregateRoot ---
'   Public Function ReleaseEvents() As Collection
'       Set ReleaseEvents = m_Root.ReleaseEvents()
'   End Function
'
'   Public Function GetVersion() As Long
'       GetVersion = m_Root.GetVersion()
'   End Function
