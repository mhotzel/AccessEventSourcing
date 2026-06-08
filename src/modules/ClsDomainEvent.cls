VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsDomainEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =============================================================================

' This Class implements a Domain Event, which can be stored in the database
' and be serialized as a xml document.
'
' =============================================================================

' --- Private Constants ---

Private Const ERRNO_INVALID_XML As Long = 50001
Private Const ERRMSG_INVALID_XML As String = "Das XML-Dokument ist nicht valide"

' --- Private fields ---
Private m_EventId           As String               ' UUID of this event
Private m_EventType         As String               ' e.g. "user.registered"
Private m_EventSource       As String               ' The Source, which has generated the Event. For Instance die Machine, combined with the user
Private m_Subject           As String               ' aggregate/stream ID (UUID)
Private m_Version           As Long                 ' position within the subject stream (for optimistic locking)
Private m_EventTime         As String               ' UTC timestamp, ISO 8601, e.g. "2026-05-30T14:23:05.123Z"
Private m_Data              As String               ' Payload, should by XML
Private m_AuthId            As String               ' optional: ID of the user/process that triggered the event


' =============================================================================
' Initialises a new event. Call this immediately after "Dim evt As New DomainEvent".
'
' Parameters:
'   pEventType   - required; event name in dot-notation, e.g. "user.registered"
'   pEventSource - required; source of the Event to identify the machine, process, maybe combined with the user, which has the event produced
'   pSubject     - required; UUID of the aggregate this event belongs to
'   pAuthId      - optional; identifier of the triggering user or process
' =============================================================================
Public Sub Init(ByVal pEventType As String, ByVal pEventSource As String, _
                ByVal pSubject As String, Optional ByVal pData As String = "", _
                Optional ByVal pAuthId As String = "")
    m_EventId = ModUUIDGenerator.GenerateUUIDv7()
    m_EventType = pEventType
    m_EventSource = pEventSource
    m_Subject = pSubject
    m_EventTime = ModTimeStampGenerator.GetUtcIsoTimestamp()
    m_Data = pData
    m_AuthId = pAuthId
    m_Version = 0     ' will be set by AggregateRoot.RecordThat before queuing
End Sub

' =============================================================================
' Initialises the event from raw DB column values (used by EventStore when
' reading events back from the database).
' =============================================================================
Public Sub InitFromDb(ByVal pEventId As String, ByVal pEventType As String, _
                      ByVal pEventSource As String, ByVal pSubject As String, _
                      ByVal pVersion As Long, ByVal pEventTime As String, _
                      ByVal pData As String, ByVal pAuthId As String)
    m_EventId = pEventId
    m_EventType = pEventType
    m_EventSource = pEventSource
    m_Subject = pSubject
    m_Data = pData
    m_Version = pVersion
    m_EventTime = pEventTime
    m_AuthId = pAuthId
End Sub

' --- Property accessors (read-only after Init) ---

Public Property Get EventId() As String
    EventId = m_EventId
End Property

Public Property Get EventType() As String
    EventType = m_EventType
End Property

Public Property Get EventSource() As String
    EventSource = m_EventSource
End Property

Public Property Get subject() As String
    subject = m_Subject
End Property

Public Property Get Version() As Long
    Version = m_Version
End Property

Public Property Get EventTime() As String
    EventTime = m_EventTime
End Property

Public Property Get Data() As String
    Data = m_Data
End Property

Public Property Get DataContentType() As String
    If Len(m_Data) > 0 Then
        DataContentType = "application/xml"
    Else
        DataContentType = ""
    End If
End Property

Public Property Get AuthId() As String
    AuthId = m_AuthId
End Property

' =============================================================================
' --- Package-level setter for Version (called by AggregateRoot only) ---
' VBA has no "friend" visibility; document by convention that only
' AggregateRoot should call this.
' =============================================================================
Public Property Let Version(ByVal val As Long)
    m_Version = val
End Property

