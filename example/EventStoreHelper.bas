Attribute VB_Name = "EventStoreHelper"
Option Compare Database
Option Explicit

' =============================================================================
' EventStoreHelper
' Standard module with utility functions used across the event sourcing modules:
'   - UUID generation
'   - UTC timestamp creation
'   - XML payload builder and reader
' =============================================================================

' Win32 API for reading the current UTC system time
#If VBA7 Then
    Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
#Else
    Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
#End If

Private Type SYSTEMTIME
    wYear         As Integer
    wMonth        As Integer
    wDayOfWeek    As Integer
    wDay          As Integer
    wHour         As Integer
    wMinute       As Integer
    wSecond       As Integer
    wMilliseconds As Integer
End Type

' -----------------------------------------------------------------------------
' Generates a new UUID (GUID) string, e.g. "3f2504e0-4f89-11d3-9a0c-0305e82c3301"
' Requires: Windows Scripting runtime (available by default in Access)
' -----------------------------------------------------------------------------
Public Function GenerateUUID() As String
    Dim typeLib As Object
    Set typeLib = CreateObject("Scriptlet.TypeLib")
    ' Raw GUID has the form {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}
    ' Strip braces and trailing null character
    Dim raw As String
    raw = typeLib.Guid
    GenerateUUID = LCase(Mid(raw, 2, 36))
    Set typeLib = Nothing
End Function

' -----------------------------------------------------------------------------
' Returns the current UTC timestamp as an ISO 8601 string with milliseconds,
' e.g. "2026-05-30T14:23:05.123Z"
' -----------------------------------------------------------------------------
Public Function GetUtcTimestamp() As String
    Dim st As SYSTEMTIME
    GetSystemTime st

    GetUtcTimestamp = Format(st.wYear, "0000") & "-" & _
                      Format(st.wMonth, "00") & "-" & _
                      Format(st.wDay, "00") & "T" & _
                      Format(st.wHour, "00") & ":" & _
                      Format(st.wMinute, "00") & ":" & _
                      Format(st.wSecond, "00") & "." & _
                      Format(st.wMilliseconds, "000") & "Z"
End Function

' -----------------------------------------------------------------------------
' Builds an XML string from a flat key/value array.
'
' Usage:
'   Dim keys(1)   As String : keys(0)   = "userName" : keys(1)   = "email"
'   Dim values(1) As String : values(0) = "Alice"    : values(1) = "alice@example.com"
'   Dim xml As String
'   xml = XmlBuildData(keys, values)
'
' Result:
'   <data><field name="userName">Alice</field><field name="email">alice@example.com</field></data>
'
' Note: Use XmlGetValue / XmlGetAllValues to read back individual fields.
' -----------------------------------------------------------------------------
Public Function XmlBuildData(ByRef keys() As String, ByRef values() As String) As String
    Dim doc As Object
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False

    ' Create root <data> element
    Dim root As Object
    Set root = doc.createElement("data")
    doc.appendChild root

    Dim i As Integer
    For i = 0 To UBound(keys)
        Dim fieldNode As Object
        Set fieldNode = doc.createElement("field")
        ' Store the key as an attribute for easy lookup
        fieldNode.setAttribute "name", keys(i)
        fieldNode.Text = values(i)
        root.appendChild fieldNode
    Next i

    XmlBuildData = doc.XML
    Set doc = Nothing
End Function

' -----------------------------------------------------------------------------
' Reads a single field value from an XML data string produced by XmlBuildData.
'
' Usage:
'   Dim name As String
'   name = XmlGetValue(evt.DataXml, "userName")  ' -> "Alice"
'
' Returns an empty string if the field is not found.
' -----------------------------------------------------------------------------
Public Function XmlGetValue(ByVal xml As String, ByVal fieldName As String) As String
    If Len(Trim(xml)) = 0 Then
        XmlGetValue = ""
        Exit Function
    End If

    Dim doc As Object
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    doc.loadXML xml

    Dim node As Object
    Set node = doc.selectSingleNode("//field[@name='" & fieldName & "']")

    If node Is Nothing Then
        XmlGetValue = ""
    Else
        XmlGetValue = node.Text
    End If

    Set doc = Nothing
End Function

' -----------------------------------------------------------------------------
' Returns all field names contained in an XML data string as a Collection.
' Useful for iterating over an unknown payload structure.
' -----------------------------------------------------------------------------
Public Function XmlGetFieldNames(ByVal xml As String) As Collection
    Dim result As New Collection

    If Len(Trim(xml)) = 0 Then
        Set XmlGetFieldNames = result
        Exit Function
    End If

    Dim doc As Object
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    doc.loadXML xml

    Dim nodes As Object
    Set nodes = doc.selectNodes("//field")

    Dim node As Object
    For Each node In nodes
        result.Add node.getAttribute("name")
    Next node

    Set XmlGetFieldNames = result
    Set doc = Nothing
End Function

' -----------------------------------------------------------------------------
' Validates whether a string is well-formed XML.
' Returns True if valid, False otherwise.
' -----------------------------------------------------------------------------
Public Function IsValidXml(ByVal xml As String) As Boolean
    If Len(Trim(xml)) = 0 Then
        IsValidXml = False
        Exit Function
    End If

    Dim doc As Object
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    IsValidXml = doc.loadXML(xml)
    Set doc = Nothing
End Function
