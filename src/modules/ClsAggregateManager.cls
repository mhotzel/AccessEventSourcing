VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsAggregateManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=================================================================
' Aggregate Root, which can be used in business entities to write
' an read events
'=================================================================

' ----------------- Event sourcing related variables and consts --
Private Const ERRNO_STREAM_EXISTS = 50104
Private Const ERRNO_WRONG_IDENT_TYPE = 50105
Private Const ERRNO_NOT_INITIALIZED = 50106

Private m_EventBuffer As Collection
Private m_Version As Long
Private isInitialized As Boolean

Private m_AggregateRoot As IAggregateRoot

Private m_UserName As String
Private m_ComputerName As String
Private m_EventSource As String

'=================================================================
' Initializes the AggregateRoot, "Constructor"
'=================================================================
Public Sub Init(ByRef pAggregate As IAggregateRoot, ByVal EventSource As String)

    isInitialized = True
    Set m_AggregateRoot = pAggregate
    m_EventSource = EventSource
    m_UserName = Trim$(LCase$(Environ$("USERNAME")))
    m_ComputerName = Trim$(LCase$(Environ$("COMPUTERNAME")))

End Sub

'=================================================================
' Returns true, if there are no Events in the Event buffer and
' no events where applied
'=================================================================
Public Property Get IsEmpty() As Boolean

    If m_Version > 0 Or EventBuffer.Count > 0 Then
        IsEmpty = False
    Else
        IsEmpty = True
    End If

End Property


'=================================================================
' Rebuilds the Aggregate from e collection of events
'=================================================================
Public Sub Rebuild(ByRef pEvents As Collection)

    Dim Evt As ClsDomainEvent
    For Each Evt In pEvents
        m_AggregateRoot.Apply Evt
    Next Evt
    
End Sub

'=================================================================
' Returns all buffered events and deletes them from the buffer
'=================================================================
Public Function ReleaseEvents() As Collection

    If isInitialized = False Then
        Err.Raise ERRNO_NOT_INITIALIZED, "ClsAggregateManager.ReleaseEvents", "Das Objekt wurde nicht mittels 'Init' initialisiert"
    End If

    Dim EventsToRelease As Collection
    Set EventsToRelease = New Collection
    Dim Evt As ClsDomainEvent
    
    Dim i As Long
    For i = 1 To EventBuffer.Count
        Set Evt = EventBuffer(i)
        m_Version = m_Version + 1
        Evt.Version = m_Version
        EventsToRelease.Add Evt
    Next i

    Set m_EventBuffer = Nothing
    Set ReleaseEvents = EventsToRelease

End Function

'=================================================================
' Builds the Auth ID from based on the computer name and user name
'=================================================================
Private Function BuildAuthId() As String

    BuildAuthId = m_ComputerName & "." & m_UserName

End Function
'=================================================================
' Applies the Event to this AggregateRoot und writes it to the eventbuffer
'=================================================================
Public Sub RecordThat(ByRef pEvent As ClsDomainEvent)

    If isInitialized = False Then
        Err.Raise ERRNO_NOT_INITIALIZED, "ClsAggregateManager.ReleaseEvents", "Das Objekt wurde nicht mittels 'Init' initialisiert"
    End If
    
    pEvent.AddEventSource m_EventSource
    pEvent.AddAuthId BuildAuthId
    m_AggregateRoot.Apply pEvent
    EventBuffer.Add pEvent
End Sub

Private Property Get EventBuffer() As Collection
    If m_EventBuffer Is Nothing Then
        Set m_EventBuffer = New Collection
    End If
    Set EventBuffer = m_EventBuffer
End Property
