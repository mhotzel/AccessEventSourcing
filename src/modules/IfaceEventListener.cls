VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IfaceEventListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =============================================================================
' IEventListener
' Interface class for event listeners.
'
' Usage:
'   Create a new class module, add "Implements IEventListener" at the top,
'   then implement the Handle method:
'
'   ' MyProjectionListener.cls
'   Implements IEventListener
'
'   Private Sub IEventListener_Handle(ByVal evt As DomainEvent)
'       Select Case evt.EventType
'           Case "user.registered"
'               ' update read model ...
'       End Select
'   End Sub
' =============================================================================

' Implement this method in every listener class.
' Called by the EventStore for each event, either within the transaction
' (transactional listener) or after a successful commit (post-commit listener).
Public Sub Handle(ByVal evt As ClsDomainEvent)
End Sub
