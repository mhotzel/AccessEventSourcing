VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAggregateRoot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=================================================================
' Dispatches the event to set the state of the Aggregate root
' based on the events. Has to be implemented by concrete classes
'=================================================================
Public Sub Apply(ByRef pEvent As ClsDomainEvent)

End Sub

'=================================================================
' Returns all buffered events and deletes them from the buffer
'=================================================================
Public Function ReleaseEvents() As Collection
    
End Function

