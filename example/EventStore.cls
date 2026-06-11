VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =============================================================================
' EventStore
' DAO-based event store for Microsoft Access.
'
' Table layout (created automatically on first use):
'
'   domain_events
'   ┌──────────────┬──────────────┬─────────────────────────────────────────┐
'   │ Column       │ Type         │ Notes                                   │
'   ├──────────────┼──────────────┼─────────────────────────────────────────┤
'   │ position     │ AutoNumber   │ PK; global ordering for projections     │
'   │ event_id     │ Text(36)     │ UUID; unique                            │
'   │ event_type   │ Text(255)    │ e.g. "user.registered"                  │
'   │ subject      │ Text(36)     │ aggregate UUID (stream ID)              │
'   │ version      │ Long         │ per-subject version; (subject,version)  │
'   │              │              │ is unique → optimistic locking          │
'   │ event_time   │ Text(50)     │ UTC ISO 8601 timestamp                  │
'   │ data_xml     │ Memo         │ XML payload, built with XmlBuildData()  │
'   │ auth_id      │ Text(36)     │ optional triggering user/process ID     │
'   └──────────────┴──────────────┴─────────────────────────────────────────┘
'
' Optimistic locking:
'   Before inserting, the store checks that the current max version for the
'   subject in the DB equals (first event's version - 1).
'   A unique index on (subject, version) provides an additional DB-level guard.
'
' Listeners:
'   - Transactional listeners run inside the DAO transaction (same workspace).
'   - Post-commit listeners run after a successful commit.
'   Register via AddTransactionalListener / AddPostCommitListener.
'
' Usage:
'   Dim store As New EventStore
'   store.Init CurrentDb()
'
'   ' -- write --
'   store.AddEvents userAggregate.ReleaseEvents()
'
'   ' -- read stream --
'   Dim events As Collection
'   Set events = store.ReadEventsBySubject(subjectId)
'   userAggregate.Rebuild events
'
'   ' -- read for projections --
'   Dim batch As Collection
'   Set batch = store.ReadAllEventsAfter(lastPosition, 500)
' =============================================================================

' Sentinel values mirroring EventStoreInterface constants in PHP
Public Const EXPECTED_VERSION_NEW     As Long = -1   ' subject must not exist yet
Public Const EXPECTED_VERSION_NOCHECK As Long = -999 ' skip optimistic locking check

' Registered listeners
Private m_TransactListeners   As Collection
Private m_PostCommitListeners As Collection

' DAO database reference
Private m_Db As DAO.Database

' =============================================================================
' Class lifecycle
' =============================================================================
Private Sub Class_Initialize()
    Set m_TransactListeners   = New Collection
    Set m_PostCommitListeners = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_TransactListeners   = Nothing
    Set m_PostCommitListeners = Nothing
    Set m_Db                  = Nothing
End Sub

' =============================================================================
' Initialises the store with a DAO database reference and ensures the
' domain_events table and its indexes exist.
'
' Typically called with CurrentDb():
'   store.Init CurrentDb()
' =============================================================================
Public Sub Init(ByVal db As DAO.Database)
    Set m_Db = db
    EnsureTable
End Sub

' =============================================================================
' Writes a Collection of DomainEvent objects to the store atomically.
' Runs an optimistic locking check before inserting.
'
' Parameters:
'   events          - Collection of DomainEvent; must not be empty
'   expectedVersion - EXPECTED_VERSION_NEW    : subject must not exist yet
'                     EXPECTED_VERSION_NOCHECK: skip the version check
'                     any other Long value    : max version in DB must equal this value
'
' Raises:
'   Error 50001  Concurrency conflict (version mismatch)
'   Error 50002  Duplicate event_id (should not happen with UUID generation)
' =============================================================================
Public Sub AddEvents(ByVal events As Collection, _
                     Optional ByVal expectedVersion As Long = EXPECTED_VERSION_NOCHECK)
    If events.Count = 0 Then Exit Sub

    Dim firstEvt As DomainEvent
    Set firstEvt = events(1)

    ' --- Optimistic locking check (application level) ---
    If expectedVersion <> EXPECTED_VERSION_NOCHECK Then
        Dim currentMax As Long
        currentMax = GetMaxVersion(firstEvt.Subject)

        If expectedVersion = EXPECTED_VERSION_NEW Then
            ' The subject must not exist yet
            If currentMax >= 0 Then
                Err.Raise 50001, "EventStore.AddEvents", _
                    "Concurrency conflict: subject '" & firstEvt.Subject & _
                    "' already exists (currentVersion=" & currentMax & ")."
            End If
        Else
            ' The current DB version must match the caller's expectation
            If currentMax <> expectedVersion Then
                Err.Raise 50001, "EventStore.AddEvents", _
                    "Concurrency conflict: expected version " & expectedVersion & _
                    " for subject '" & firstEvt.Subject & _
                    "' but found " & currentMax & "."
            End If
        End If
    End If

    ' --- Write all events inside a single DAO transaction ---
    Dim ws As DAO.Workspace
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans

    On Error GoTo RollbackAndRaise

    Dim evt As DomainEvent
    For Each evt In events
        InsertEvent evt
        NotifyTransactListeners evt
    Next evt

    ws.CommitTrans

    ' --- Notify post-commit listeners after successful commit ---
    For Each evt In events
        NotifyPostCommitListeners evt
    Next evt

    Set ws = Nothing
    Exit Sub

RollbackAndRaise:
    Dim errNum  As Long
    Dim errDesc As String
    errNum  = Err.Number
    errDesc = Err.Description

    ws.Rollback
    Set ws = Nothing

    ' Re-raise the original error
    Err.Raise errNum, "EventStore.AddEvents", errDesc
End Sub

' =============================================================================
' Reads a single event by its UUID. Returns Nothing if not found.
' =============================================================================
Public Function ReadEventById(ByVal eventId As String) As DomainEvent
    Dim sql As String
    sql = "SELECT * FROM domain_events WHERE event_id = '" & EscapeStr(eventId) & "'"

    Dim rs As DAO.Recordset
    Set rs = m_Db.OpenRecordset(sql, dbOpenSnapshot)

    If rs.EOF Then
        Set ReadEventById = Nothing
    Else
        Set ReadEventById = RecordsetToEvent(rs)
    End If

    rs.Close
    Set rs = Nothing
End Function

' =============================================================================
' Reads all events for a given subject, ordered by version (ascending).
' Returns a Collection of DomainEvent objects.
' =============================================================================
Public Function ReadEventsBySubject(ByVal subject As String) As Collection
    Dim result As New Collection

    Dim sql As String
    sql = "SELECT * FROM domain_events " & _
          "WHERE subject = '" & EscapeStr(subject) & "' " & _
          "ORDER BY version ASC"

    Dim rs As DAO.Recordset
    Set rs = m_Db.OpenRecordset(sql, dbOpenSnapshot)

    Do While Not rs.EOF
        result.Add RecordsetToEvent(rs)
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Set ReadEventsBySubject = result
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
    Dim result As New Collection

    Dim sql As String
    sql = "SELECT TOP " & limit & " * FROM domain_events " & _
          "WHERE position > " & lastPosition & " " & _
          "ORDER BY position ASC"

    Dim rs As DAO.Recordset
    Set rs = m_Db.OpenRecordset(sql, dbOpenSnapshot)

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
End Function

' =============================================================================
' Registers a listener that is called within the DAO transaction.
' Use this for read-model updates that must be atomic with the event write.
' =============================================================================
Public Sub AddTransactionalListener(ByVal listener As IEventListener)
    If Not ListenerExists(m_TransactListeners, listener) Then
        m_TransactListeners.Add listener
    End If
End Sub

' =============================================================================
' Registers a listener that is called after a successful commit.
' Use this for side-effects like sending notifications.
' =============================================================================
Public Sub AddPostCommitListener(ByVal listener As IEventListener)
    If Not ListenerExists(m_PostCommitListeners, listener) Then
        m_PostCommitListeners.Add listener
    End If
End Sub

' =============================================================================
' Private helpers
' =============================================================================

' Inserts a single event row using a DAO Recordset (avoids SQL injection risk
' from payload content that might contain single quotes or special characters).
Private Sub InsertEvent(ByVal evt As DomainEvent)
    Dim rs As DAO.Recordset
    Set rs = m_Db.OpenRecordset("domain_events", dbOpenTable)

    rs.AddNew
    rs!event_id   = evt.EventId
    rs!event_type = evt.EventType
    rs!subject    = evt.Subject
    rs!version    = evt.Version
    rs!event_time = evt.EventTime
    rs!data_xml   = evt.DataXml
    rs!auth_id    = evt.AuthId
    rs.Update

    rs.Close
    Set rs = Nothing
End Sub

' Returns the highest version currently stored for the given subject, or -1
' if no events exist for that subject yet.
Private Function GetMaxVersion(ByVal subject As String) As Long
    Dim sql As String
    sql = "SELECT MAX(version) AS maxVer FROM domain_events " & _
          "WHERE subject = '" & EscapeStr(subject) & "'"

    Dim rs As DAO.Recordset
    Set rs = m_Db.OpenRecordset(sql, dbOpenSnapshot)

    If rs.EOF Or IsNull(rs!maxVer) Then
        GetMaxVersion = -1
    Else
        GetMaxVersion = CLng(rs!maxVer)
    End If

    rs.Close
    Set rs = Nothing
End Function

' Reconstitutes a DomainEvent from the current row of a Recordset.
Private Function RecordsetToEvent(ByVal rs As DAO.Recordset) As DomainEvent
    Dim evt As New DomainEvent
    evt.InitFromDb _
        CStr(rs!event_id), _
        CStr(rs!event_type), _
        CStr(rs!subject), _
        CLng(rs!version), _
        CStr(rs!event_time), _
        Nz(rs!data_xml, ""), _
        Nz(rs!auth_id, "")
    Set RecordsetToEvent = evt
End Function

' Calls Handle() on every transactional listener.
Private Sub NotifyTransactListeners(ByVal evt As DomainEvent)
    Dim listener As IEventListener
    For Each listener In m_TransactListeners
        listener.Handle evt
    Next listener
End Sub

' Calls Handle() on every post-commit listener.
Private Sub NotifyPostCommitListeners(ByVal evt As DomainEvent)
    Dim listener As IEventListener
    For Each listener In m_PostCommitListeners
        listener.Handle evt
    Next listener
End Sub

' Returns True if the listener object is already in the collection
' (reference equality).
Private Function ListenerExists(ByVal col As Collection, ByVal listener As IEventListener) As Boolean
    Dim item As IEventListener
    For Each item In col
        If item Is listener Then
            ListenerExists = True
            Exit Function
        End If
    Next item
    ListenerExists = False
End Function

' Escapes a string for safe embedding in SQL (doubles single quotes).
' Note: use this only for metadata fields (IDs, types). Payload goes via
' the Recordset approach in InsertEvent, which is injection-safe.
Private Function EscapeStr(ByVal s As String) As String
    EscapeStr = Replace(s, "'", "''")
End Function

' =============================================================================
' Table and index setup
' Creates the domain_events table and its indexes if they do not exist yet.
' Called once from Init().
' =============================================================================
Private Sub EnsureTable()
    ' Check if the table already exists
    Dim tbl As DAO.TableDef
    For Each tbl In m_Db.TableDefs
        If tbl.Name = "domain_events" Then
            Exit Sub   ' table exists, nothing to do
        End If
    Next tbl

    ' --- Create table ---
    Dim newTable As DAO.TableDef
    Set newTable = m_Db.CreateTableDef("domain_events")

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
    Set fldTime = newTable.CreateField("event_time", dbText, 50)
    fldTime.Required = True
    newTable.Fields.Append fldTime

    ' data_xml: XML payload (Memo = LongText, no size limit)
    Dim fldData As DAO.Field
    Set fldData = newTable.CreateField("data_xml", dbMemo)
    newTable.Fields.Append fldData

    ' auth_id: optional triggering user/process ID
    Dim fldAuth As DAO.Field
    Set fldAuth = newTable.CreateField("auth_id", dbText, 36)
    fldAuth.Required = False
    newTable.Fields.Append fldAuth

    m_Db.TableDefs.Append newTable

    ' --- Primary key on position ---
    Dim idxPk As DAO.Index
    Set idxPk = newTable.CreateIndex("PK_domain_events")
    idxPk.Primary = True
    idxPk.Unique  = True
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

    ' --- Unique composite index on (subject, version) → optimistic locking ---
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

    m_Db.TableDefs.Refresh
End Sub
