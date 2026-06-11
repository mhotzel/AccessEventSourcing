VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsAsserts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'==================================================
'Support for Test cases
'==================================================

Private Const ERRNO_TESTS = 56001

Public Sub CountEquals(ByVal Src As String, ByVal expected As Long, ByVal actual As Long)
    
    If expected <> actual Then
        Err.Raise ERRNO_TESTS, Src, Src & " - Expected: " & expected & " - got: " & actual
    End If
    
End Sub

Public Sub ErrExpected(ByVal Src As String)
    
    Err.Raise ERRNO_TESTS, Src, Src & " - Error was expected, but not found"
    
End Sub


Public Sub Equals(ByVal Src As String, ByVal expected As Variant, ByVal actual As Variant)

    If expected <> actual Then
        Err.Raise ERRNO_TESTS, Src, Src & " - Expected: " & expected & " - got: " & actual
    End If

End Sub
