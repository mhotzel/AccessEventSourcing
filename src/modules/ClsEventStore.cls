VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsEventStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =============================================================================
' Implements an EventStore according to best practices for Event Sourcing
' =============================================================================

' - Private Constants -----------------
Private Const ERRNO_NOT_INITIALIZED As Long = 50101
Private Const ERRMSG_NOT_INITIALIZED As String = "Das EventStore-Objekt wurde nicht mit 'Init' initialisiert"

' - Private variables -----------------
Private isInitialized As Boolean
Private m_DbConn As DAO.Database
Private m_Workspace As DAO.Workspace
Private m_Listeners As Collection
Private m_QuerySqlMaxVersion As QueryDef
Private m_QuerySqlEventsById As QueryDef
Private m_QuerySqlEventsBySubject As QueryDef

' =============================================================================
' Initializes the store with the given DAO Workspace, which also has to be reference and ensures the
' domain_events table and its indexes exist.
' Parameters:
'   pWorkspace:     a DAO.Workspace for transaction handling, connection handling and so on
'   pDbPath         a String with the Path to the database file
' =============================================================================
Public Sub Init(ByRef pWorkspace As DAO.Workspace, pDbPath)
    
    isInitialized = True
    
    Set m_Workspace = pWorkspace
    Set m_DbConn = m_Workspace.OpenDatabase(pDbPath)
    Call EnsureTable(m_DbConn)
    Call InitQueries(m_DbConn)

End Sub


' =============================================================================
' Writes a Collection of DomainEvent objects to the store atomically.
' Runs an optimistic locking check before inserting.
'
' Parameters:
'   events          - Collection of DomainEvent; must not be empty
' =============================================================================
Public Sub AddEvents(ByVal events As Collection)

    If Not isInitialized Then
        Err.Raise ERRNO_NOT_INITIALIZED, "ClsEventStore.AddEvents", ERRMSG_NOT_INITIALIZED
    End If

    If events.Count = 0 Then Exit Sub

    Dim firstEvt As ClsDomainEvent
    Set firstEvt = events(1)

    ' --- Write all events inside a single DAO transaction ---
    m_Workspace.BeginTrans
    
    On Error GoTo ErrAddEvents

    Dim rs As DAO.Recordset
    Set rs = m_DbConn.OpenRecordset("domain_events", dbOpenDynaset)

    Dim evt As ClsDomainEvent
    For Each evt In events
        Call InsertEvent(rs, evt)
        Call NotifyListeners(evt)
    Next evt

    rs.Close
    Set rs = Nothing

    m_Workspace.CommitTrans

    Exit Sub

ErrAddEvents:
    Dim errNum  As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    m_Workspace.Rollback

    Err.Raise errNum, "ClsEventStore.AddEvents", errDesc
End Sub


' =============================================================================
' Reads a single event by its UUID. Returns Nothing if not found.
' Parameters:
'   pEventId:    a DAO.Workspace for transaction handling, connection handling and so on
' =============================================================================
Public Function ReadEventById(ByVal pEventId As String) As ClsDomainEvent

    On Error GoTo ErrReadEventById

    Dim rs As DAO.Recordset
    m_QuerySqlEventsById.Parameters("pEventId").Value = pEventId
    Set rs = m_QuerySqlEventsById.OpenRecordset(dbOpenSnapshot, dbReadOnly)

    If rs.EOF Then
        Set ReadEventById = Nothing
    Else
        Set ReadEventById = RecordsetToEvent(rs)
    End If

    rs.Close
    Set rs = Nothing
    Exit Function
    
ErrReadEventById:
    Dim errNo As Integer: errNo = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    Dim errSrc As String: errSrc = "ClsEventStore.ReadEventById"
    On Error Resume Next
    
    rs.Close
    Set rs = Nothing
    Err.Raise errNo, errSrc, errMsg
    
End Function

' =============================================================================
' Reads all events for a given subject, ordered by version (ascending).
' Returns a Collection of ClsDomainEvent objects.
' Parameters:
'   subject:    a DAO.Workspace for transaction handling, connection handling and so on
' =============================================================================
Public Function ReadEventsBySubject(ByVal pSubject As String) As Collection

    On Error GoTo ErrReadEventsBySubject

    Dim result As New Collection
    Dim rs As DAO.Recordset
    
    m_QuerySqlEventsBySubject.Parameters("pSubject").Value = pSubject
    Set rs = m_QuerySqlEventsBySubject.OpenRecordset(dbOpenSnapshot, dbReadOnly)

    Do While Not rs.EOF
        result.Add RecordsetToEvent(rs)
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Set ReadEventsBySubject = result
    Exit Function
    
ErrReadEventsBySubject:
    Dim errNo As Integer: errNo = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    Dim errSrc As String: errSrc = "ClsEventStore.ReadEventsBySubject"
    On Error Resume Next
    
    rs.Close
    Set rs = Nothing
    Err.Raise errNo, errSrc, errMsg
    
End Function

' =============================================================================
' Reads up to [limit] events with position > lastPosition, ordered by position.
' Returns a Collection of items, each being a two-element Collection:
'   item(1) = position As Long
'   item(2) = event    As DomainEvent
'
' Usage (projection loop):
'   Dim batch As Collection
'   Set batch = store.ReadAllEventsAfter(checkpoint, 500)
'   Dim item As Collection
'   For Each item In batch
'       Dim pos As Long      : pos = item(1)
'       Dim evt As DomainEvent : Set evt = item(2)
'       ' process evt ...
'       checkpoint = pos
'   Next item
' =============================================================================
Public Function ReadAllEventsAfter(ByVal lastPosition As Long, _
                                   Optional ByVal limit As Long = 500) As Collection
    
    On Error GoTo ErrReadAllEventsAfter
    Dim result As New Collection

    Dim SqlEventsAfterPos
    SqlEventsAfterPos = "SELECT TOP " & limit & " * FROM domain_events " & _
                        "WHERE position > " & lastPosition & " " & _
                        "ORDER BY domain_events.position ASC ;"

    Dim rs As DAO.Recordset
    Set rs = m_DbConn.OpenRecordset(SqlEventsAfterPos, dbOpenSnapshot, dbReadOnly)
    Do While Not rs.EOF
        Dim item As New Collection
        item.Add CLng(rs!position)          ' item(1) = position
        item.Add RecordsetToEvent(rs)       ' item(2) = DomainEvent
        result.Add item
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Set ReadAllEventsAfter = result
    Exit Function
    
ErrReadAllEventsAfter:
    Dim errNo As Integer: errNo = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    Dim errSrc As String: errSrc = "ClsEventStore.ReadAllEventsAfter"
    On Error Resume Next
    
    rs.Close
    Set rs = Nothing
    Err.Raise errNo, errSrc, errMsg
    
End Function


' =============================================================================
' Closes the internal Workspace and the database connections
' =============================================================================
Public Sub CloseEventStore()

    If Not isInitialized Then
        Err.Raise ERRNO_NOT_INITIALIZED, "ClsEventStore.AddEvents", ERRMSG_NOT_INITIALIZED
    End If

    On Error Resume Next
   
    If Not m_Workspace Is Nothing Then
        m_Workspace.Rollback
    End If
    
    If Not m_DbConn Is Nothing Then
        m_DbConn.Close
    End If
    
    Set m_Listeners = Nothing
    Set m_DbConn = Nothing
End Sub

' =============================================================================
' Registers a listener that is called within the DAO transaction.
' Use this for read-model updates that must be atomic with the event write.
' =============================================================================
Public Sub AddListener(ByVal listener As IfaceEventListener)

    If Not isInitialized Then
        Err.Raise ERRNO_NOT_INITIALIZED, "ClsEventStore.AddEvents", ERRMSG_NOT_INITIALIZED
    End If

    If Not ListenerExists(m_Listeners, listener) Then
        m_Listeners.Add listener
    End If
End Sub

' =============================================================================
' Returns True if the listener object is already in the collection
' (reference equality).
' =============================================================================
Private Function ListenerExists(ByVal col As Collection, ByVal listener As IfaceEventListener) As Boolean
    Dim item As IfaceEventListener
    For Each item In col
        If item Is listener Then
            ListenerExists = True
            Exit Function
        End If
    Next item
    ListenerExists = False
End Function

' Calls Handle() on every Event listener.
Private Sub NotifyListeners(ByVal evt As ClsDomainEvent)
    Dim listener As IfaceEventListener
    For Each listener In m_Listeners
        listener.Handle evt
    Next listener
End Sub

' =============================================================================
' Inserts a single event row using a DAO Recordset (avoids SQL injection risk
' from payload content that might contain single quotes or special characters).
' Parameters:
'   rs: an opened Recordset of type DAO.Recordset, where the event has to be written.
'   evt: the event of type ClsDomainEvent, which has to be written
' =============================================================================
Private Sub InsertEvent(ByRef rs As DAO.Recordset, ByVal evt As ClsDomainEvent)
    
    On Error GoTo ErrInsertEvent
       
    rs.AddNew
    rs("event_id") = evt.EventId
    rs("event_type") = evt.EventType
    rs("subject") = evt.subject
    rs("version") = evt.Version
    rs("event_time") = evt.EventTime
    rs("event_source") = evt.EventSource
    rs("data") = evt.Data
    rs("content_type") = evt.DataContentType
    rs("auth_id") = evt.AuthId
    rs.Update
    Exit Sub
    
ErrInsertEvent:
    Dim errNo As Integer: errNo = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    Dim errSrc As String: errSrc = "ClsEventStore.InsertEvent"
    
    If errNo = 3022 Then
        errMsg = "Die Eindeutigkeit der Events ist nicht gewährleistet"
    End If
    On Error Resume Next
    Err.Raise errNo, errSrc, errMsg
    
End Sub


' =============================================================================
' Returns the highest version currently stored for the given subject, or -1
' if no events exist for that subject yet.
' =============================================================================
Private Function GetMaxVersion(ByVal subject As String) As Long

    On Error GoTo ErrGetMaxVersion
    
    m_QuerySqlMaxVersion.Parameters("pSubject").Value = subject
    Dim rs As DAO.Recordset
    Set rs = m_QuerySqlMaxVersion.OpenRecordset(dbOpenSnapshot, dbForwardOnly)
    
    If rs.EOF Or IsNull(rs("maxVer")) Then
        GetMaxVersion = -1
    Else
        GetMaxVersion = CLng(rs("maxVer"))
    End If
    
    rs.Close
    Set rs = Nothing
    Exit Function
    
    
ErrGetMaxVersion:
    Dim errNo As Integer: errNo = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    Dim errSrc As String: errSrc = "ClsEventStore.GetMaxVersion"
    
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
    End If
    Set rs = Nothing
    
End Function

' =============================================================================
' Initializes the prepared statements
' Parameters: pDb - required; the Database
'
' =============================================================================
Private Sub InitQueries(ByRef pDb As DAO.Database)
    Dim SqlMaxVersion As String
    SqlMaxVersion = "PARAMETERS " & _
                    "pSubject String; " & _
                    "SELECT MAX(version) AS maxVer FROM domain_events WHERE subject = pSubject"
    
    Set m_QuerySqlMaxVersion = pDb.CreateQueryDef("", SqlMaxVersion)
    
    Dim SqlEventsById As String
    SqlEventsById = "PARAMETERS " & _
                    "pEventId String; " & _
                    "SELECT * FROM domain_events WHERE event_id = pEventId"
                    
    Set m_QuerySqlEventsById = pDb.CreateQueryDef("", SqlEventsById)
    
    Dim SqlEventsBySubject As String
    SqlEventsBySubject = "PARAMETERS " & _
                         "pSubject String; " & _
                         "SELECT * FROM domain_events WHERE subject = pSubject"
    
    Set m_QuerySqlEventsBySubject = pDb.CreateQueryDef("", SqlEventsBySubject)
    
End Sub

' =============================================================================
' Reads a recordset from the eventstore table and builds a DomainEvent from it
' =============================================================================
Private Function RecordsetToEvent(ByRef rs As DAO.Recordset) As ClsDomainEvent

    On Error GoTo ErrRecordsetToEvent

    Dim DomainEvent As ClsDomainEvent
    Set DomainEvent = New ClsDomainEvent
    
    Dim pEventId As String
    Dim pEventType As String
    Dim pEventSource As String
    Dim pSubject As String
    Dim pVersion As Long
    Dim pEventTime As String
    Dim pEventData As String
    Dim pAuthId As String
    
    pEventId = rs("event_id").Value
    pEventType = rs("event_type").Value
    pEventSource = rs("event_source").Value
    pSubject = rs("subject").Value
    pVersion = rs("version").Value
    pEventTime = rs("event_time").Value
    pEventData = rs("data").Value
    pAuthId = rs("auth_id").Value
    
    Call DomainEvent.InitFromDb(pEventId, pEventType, pEventSource, pSubject, pVersion, pEventTime, pEventData, pAuthId)
    Set RecordsetToEvent = DomainEvent
    Exit Function
    
ErrRecordsetToEvent:
    Dim errNo As Integer: errNo = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    Dim errSrc As String: errSrc = "ClsEventStore.RecordsetToEvent"
    
    On Error Resume Next
    Set DomainEvent = Nothing
    Err.Raise errNo, errSrc, errMsg

End Function



' =============================================================================
' Table and index setup
' Creates the domain_events table and its indexes if they do not exist yet.
' Called once from Init().
' =============================================================================
Private Sub EnsureTable(ByRef pDbConn As DAO.Database)

    ' Check if the table already exists
    Dim tbl As DAO.TableDef
    For Each tbl In pDbConn.TableDefs
        If tbl.Name = "domain_events" Then
            Exit Sub   ' table exists, nothing to do
        End If
    Next tbl

    ' --- Create table ---
    Dim newTable As DAO.TableDef
    Set newTable = pDbConn.CreateTableDef("domain_events")

    ' position: auto-increment primary key for global ordering
    Dim fldPosition As DAO.Field
    Set fldPosition = newTable.CreateField("position", dbLong)
    fldPosition.Attributes = dbAutoIncrField
    newTable.Fields.Append fldPosition

    ' event_id: UUID string, must be unique
    Dim fldId As DAO.Field
    Set fldId = newTable.CreateField("event_id", dbText, 36)
    fldId.Required = True
    newTable.Fields.Append fldId

    ' event_type: e.g. "user.registered"
    Dim fldType As DAO.Field
    Set fldType = newTable.CreateField("event_type", dbText, 255)
    fldType.Required = True
    newTable.Fields.Append fldType
    
    ' event_source: e.g. "machine_x"
    Dim fldSource As DAO.Field
    Set fldSource = newTable.CreateField("event_source", dbText, 255)
    fldSource.Required = True
    newTable.Fields.Append fldSource

    ' subject: aggregate/stream UUID
    Dim fldSubject As DAO.Field
    Set fldSubject = newTable.CreateField("subject", dbText, 36)
    fldSubject.Required = True
    newTable.Fields.Append fldSubject

    ' version: per-subject version number
    Dim fldVersion As DAO.Field
    Set fldVersion = newTable.CreateField("version", dbLong)
    fldVersion.Required = True
    newTable.Fields.Append fldVersion

    ' event_time: UTC ISO 8601 timestamp string
    Dim fldTime As DAO.Field
    Set fldTime = newTable.CreateField("event_time", dbText, 33)
    fldTime.Required = True
    newTable.Fields.Append fldTime

    ' data: XML payload (Memo = LongText, no size limit)
    Dim fldData As DAO.Field
    Set fldData = newTable.CreateField("data", dbMemo)
    fldData.AllowZeroLength = True
    newTable.Fields.Append fldData
    
    ' content_type: "application/xml" or ""
    Dim fldContentType As DAO.Field
    Set fldContentType = newTable.CreateField("content_type", dbText, 255)
    fldContentType.Required = False
    fldContentType.AllowZeroLength = True
    newTable.Fields.Append fldContentType

    ' auth_id: optional triggering user/process ID
    Dim fldAuth As DAO.Field
    Set fldAuth = newTable.CreateField("auth_id", dbText, 36)
    fldAuth.Required = False
    fldAuth.AllowZeroLength = True
    newTable.Fields.Append fldAuth

    pDbConn.TableDefs.Append newTable

    ' --- Primary key on position ---
    Dim idxPk As DAO.Index
    Set idxPk = newTable.CreateIndex("PK_domain_events")
    idxPk.Primary = True
    idxPk.Unique = True
    Dim fldPk As DAO.Field
    Set fldPk = idxPk.CreateField("position")
    idxPk.Fields.Append fldPk
    newTable.Indexes.Append idxPk

    ' --- Unique index on event_id ---
    Dim idxId As DAO.Index
    Set idxId = newTable.CreateIndex("UQ_event_id")
    idxId.Unique = True
    Dim fldIdIdx As DAO.Field
    Set fldIdIdx = idxId.CreateField("event_id")
    idxId.Fields.Append fldIdIdx
    newTable.Indexes.Append idxId

    ' --- Unique composite index on (subject, version) ? optimistic locking ---
    Dim idxSv As DAO.Index
    Set idxSv = newTable.CreateIndex("UQ_subject_version")
    idxSv.Unique = True
    Dim fldSvSubject As DAO.Field
    Dim fldSvVersion As DAO.Field
    Set fldSvSubject = idxSv.CreateField("subject")
    Set fldSvVersion = idxSv.CreateField("version")
    idxSv.Fields.Append fldSvSubject
    idxSv.Fields.Append fldSvVersion
    newTable.Indexes.Append idxSv

    ' --- Index on event_type for type-based queries ---
    Dim idxType As DAO.Index
    Set idxType = newTable.CreateIndex("IDX_event_type")
    Dim fldTypeIdx As DAO.Field
    Set fldTypeIdx = idxType.CreateField("event_type")
    idxType.Fields.Append fldTypeIdx
    newTable.Indexes.Append idxType

    pDbConn.TableDefs.Refresh
    
End Sub

Private Sub Class_Initialize()
    isInitialized = False
    Set m_Listeners = New Collection
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    
    Set m_Listeners = Nothing
    Set m_DbConn = Nothing
    
    m_QuerySqlMaxVersion.Close
    Set m_QuerySqlMaxVersion = Nothing
    
    m_QuerySqlEventsBySubject.Close
    Set m_QuerySqlEventsBySubject = Nothing
    
    m_QuerySqlEventsById.Close
    
End Sub

