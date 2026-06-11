VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AggregateRoot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =============================================================================
' AggregateRoot
' Manages the version counter and the list of recorded (unpersisted) events
' for a domain aggregate.
'
' Because VBA does not support class inheritance, concrete aggregates hold an
' AggregateRoot instance as a private field and delegate to it.
' Each concrete aggregate class is responsible for implementing its own Apply()
' dispatcher and calling RecordThat() to record new events.
'
' Typical pattern in a concrete aggregate class (e.g. UserAggregate.cls):
'
'   Private m_Root As AggregateRoot
'
'   Private Sub Class_Initialize()
'       Set m_Root = New AggregateRoot
'   End Sub
'
'   ' Domain "constructor" – creates a new user
'   Public Sub Register(ByVal userName As String, ByVal email As String)
'       Dim keys(1) As String : keys(0) = "userName" : keys(1) = "email"
'       Dim vals(1) As String : vals(0) = userName   : vals(1) = email
'
'       Dim evt As New DomainEvent
'       evt.Init "user.registered", GenerateUUID(), XmlBuildData(keys, vals)
'
'       RecordThat evt
'   End Sub
'
'   ' Rebuilds state from a Collection of DomainEvent objects
'   Public Sub Rebuild(ByVal events As Collection)
'       Dim evt As DomainEvent
'       For Each evt In events
'           m_Root.SetVersion evt.Version
'           Apply evt
'       Next evt
'   End Sub
'
'   ' Routes events to the matching internal handler
'   Private Sub Apply(ByVal evt As DomainEvent)
'       Select Case evt.EventType
'           Case "user.registered"
'               m_UserName = XmlGetValue(evt.DataXml, "userName")
'               m_Email    = XmlGetValue(evt.DataXml, "email")
'       End Select
'   End Sub
'
'   ' Records an event: assigns version, applies it, then queues it for storage
'   Private Sub RecordThat(ByVal evt As DomainEvent)
'       m_Root.PrepareEvent evt   ' increments version and stamps it on the event
'       Apply evt
'       m_Root.QueueEvent evt
'   End Sub
'
'   Public Function ReleaseEvents() As Collection
'       Set ReleaseEvents = m_Root.ReleaseEvents()
'   End Function
'
'   Public Function GetVersion() As Long
'       GetVersion = m_Root.GetVersion()
'   End Function
' =============================================================================

' Current highest version of the aggregate (-1 means no events yet)
Private m_Version As Long

' Events recorded in this session but not yet written to the EventStore
Private m_RecordedEvents As Collection

Private Sub Class_Initialize()
    m_Version = -1
    Set m_RecordedEvents = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_RecordedEvents = Nothing
End Sub

' =============================================================================
' Called by the concrete aggregate's RecordThat method before Apply.
' Increments the version counter and stamps it onto the event.
' =============================================================================
Public Sub PrepareEvent(ByVal evt As DomainEvent)
    m_Version = m_Version + 1
    evt.Version = m_Version
End Sub

' =============================================================================
' Adds a prepared event to the internal queue.
' Call this after Apply() has updated the aggregate state.
' =============================================================================
Public Sub QueueEvent(ByVal evt As DomainEvent)
    m_RecordedEvents.Add evt
End Sub

' =============================================================================
' Called during Rebuild to synchronise the version counter with the event
' stream read from the EventStore (without recording or applying).
' =============================================================================
Public Sub SetVersion(ByVal version As Long)
    m_Version = version
End Sub

' =============================================================================
' Returns all recorded (unpersisted) events and clears the internal queue.
' Pass the result directly to EventStore.AddEvents().
' =============================================================================
Public Function ReleaseEvents() As Collection
    Dim released As Collection
    Set released = m_RecordedEvents
    Set m_RecordedEvents = New Collection   ' reset queue
    Set ReleaseEvents = released
End Function

' =============================================================================
' Returns the current version of the aggregate.
' The EventStore uses this for the optimistic locking check.
' =============================================================================
Public Function GetVersion() As Long
    GetVersion = m_Version
End Function
