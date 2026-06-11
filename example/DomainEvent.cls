VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DomainEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =============================================================================
' DomainEvent
' Represents a single immutable domain event, loosely modelled after the
' CloudEvents v1.0 standard (https://cloudevents.io/).
'
' Metadata fields are stored as individual DB columns.
' The payload (data) is stored as an XML string in the data_xml column.
' Use the helper functions in EventStoreHelper to build and read the payload.
'
' Usage – creating a new event:
'   Dim evt As New DomainEvent
'   Dim keys(1) As String : keys(0) = "userName" : keys(1) = "email"
'   Dim vals(1) As String : vals(0) = "Alice"    : vals(1) = "alice@example.com"
'   evt.Init "user.registered", "subject-uuid", XmlBuildData(keys, vals)
'
' Usage – reading payload:
'   Dim name As String
'   name = XmlGetValue(evt.DataXml, "userName")
' =============================================================================

' --- Private fields ---
Private m_EventId   As String   ' UUID of this event
Private m_EventType As String   ' e.g. "user.registered"
Private m_Subject   As String   ' aggregate/stream ID (UUID)
Private m_Version   As Long     ' position within the subject stream (for optimistic locking)
Private m_EventTime As String   ' UTC timestamp, ISO 8601, e.g. "2026-05-30T14:23:05.123Z"
Private m_DataXml   As String   ' XML payload, built with XmlBuildData()
Private m_AuthId    As String   ' optional: ID of the user/process that triggered the event

' =============================================================================
' Initialises a new event. Call this immediately after "Dim evt As New DomainEvent".
'
' Parameters:
'   eventType  - required; event name in dot-notation, e.g. "user.registered"
'   subject    - required; UUID of the aggregate this event belongs to
'   dataXml    - optional; XML payload, use XmlBuildData() from EventStoreHelper
'   authId     - optional; identifier of the triggering user or process
' =============================================================================
Public Sub Init(ByVal eventType As String, ByVal subject As String, _
                Optional ByVal dataXml As String = "", _
                Optional ByVal authId As String = "")
    m_EventId   = GenerateUUID()
    m_EventType = eventType
    m_Subject   = subject
    m_EventTime = GetUtcTimestamp()
    m_DataXml   = dataXml
    m_AuthId    = authId
    m_Version   = 0   ' will be set by AggregateRoot.RecordThat before queuing
End Sub

' =============================================================================
' Initialises the event from raw DB column values (used by EventStore when
' reading events back from the database).
' =============================================================================
Public Sub InitFromDb(ByVal eventId As String, ByVal eventType As String, _
                      ByVal subject As String, ByVal version As Long, _
                      ByVal eventTime As String, ByVal dataXml As String, _
                      ByVal authId As String)
    m_EventId   = eventId
    m_EventType = eventType
    m_Subject   = subject
    m_Version   = version
    m_EventTime = eventTime
    m_DataXml   = dataXml
    m_AuthId    = authId
End Sub

' --- Property accessors (read-only after Init) ---

Public Property Get EventId() As String
    EventId = m_EventId
End Property

Public Property Get EventType() As String
    EventType = m_EventType
End Property

Public Property Get Subject() As String
    Subject = m_Subject
End Property

Public Property Get Version() As Long
    Version = m_Version
End Property

Public Property Get EventTime() As String
    EventTime = m_EventTime
End Property

Public Property Get DataXml() As String
    DataXml = m_DataXml
End Property

Public Property Get AuthId() As String
    AuthId = m_AuthId
End Property

' --- Package-level setter for Version (called by AggregateRoot only) ---
' VBA has no "friend" visibility; document by convention that only
' AggregateRoot should call this.
Public Property Let Version(ByVal val As Long)
    m_Version = val
End Property
