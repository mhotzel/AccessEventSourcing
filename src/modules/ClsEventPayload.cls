VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsEventPayload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =============================================================================
' simple implementation, that is able to export data as XML-output.
' in the moment, there is no support for nested tags and it is intended to be
' used to output key-value-pairs in Domain Events.
' =============================================================================

Private Const ERRNO_NO_ATTR_FOUND = 50102

Private m_Keys As Collection
Private m_Values As Collection

' =============================================================================
' Adds an attribute to the payload, which will be represented as a xml-tag
' =============================================================================
Public Sub AddAttribute(ByVal AttrName As String, ByVal AttrVal As String)
    
    Dim sKey As String
    sKey = LCase$(AttrName)
    
    Dim i As Long
    For i = 1 To m_Keys.Count
        If m_Keys(i) = sKey Then
            m_Keys.Remove i
            m_Values.Remove i
            Exit For
        End If
    Next i
    
    m_Keys.Add sKey
    m_Values.Add AttrVal
    
End Sub

' =============================================================================
' Returns a Collection of all Attributes in the Payload
' =============================================================================
Public Function GetAttributes() As Collection

    Dim AllAttributes As Collection
    Set AllAttributes = New Collection
    
    Dim i As Long
    For i = 1 To m_Keys.Count
        AllAttributes.Add m_Keys(i)
    Next i
    
    Set GetAttributes = AllAttributes

End Function

' =============================================================================
' Returns the Value of an attribute in the payload. Throws an Error
' with the number 56002, if no attribute was found
' Parameters:
'   AttrName:   The name of an attribute
' =============================================================================
Public Function GetAttribute(ByVal AttrName As String) As String

    Dim i As Long
    Dim result As String
    For i = 1 To m_Keys.Count
        If m_Keys(i) = AttrName Then
            result = m_Values(i)
            GetAttribute = result
            Exit Function
        End If
    Next i
    
    Err.Raise ERRNO_NO_ATTR_FOUND, "ClsEventPayload.GetAttribute", "Kein Attribut mit dem Namen '" & AttrName & "' gefunden"

End Function

' =============================================================================
' Exports the Payload as an MSXML2.DOMDocument60
' =============================================================================
Public Function AsXML() As MSXML2.DOMDocument60

    Dim Doc As MSXML2.DOMDocument60
    Set Doc = New MSXML2.DOMDocument60
    Doc.async = False
    
    Doc.appendChild Doc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")
    
    Dim RootElem As MSXML2.IXMLDOMElement
    Set RootElem = Doc.createElement("payload")
    Doc.appendChild RootElem
    
    Dim i As Long
    Dim DataElem As MSXML2.IXMLDOMElement
    For i = 1 To m_Keys.Count
        Set DataElem = Doc.createElement(m_Keys(i))
        DataElem.Text = m_Values(i)
        RootElem.appendChild DataElem
    Next i

    Set AsXML = Doc

End Function

' =============================================================================
' Imports an XML-String
' Parameters:
'   XmlString:  An XML-Document as as String, for instance from the database
' =============================================================================
Public Sub ReadXml(ByVal XmlString As String)
    
    Set m_Keys = Nothing
    Set m_Values = Nothing
    
    Set m_Keys = New Collection
    Set m_Values = New Collection
    
    Dim Doc As MSXML2.DOMDocument60
    Set Doc = New MSXML2.DOMDocument60
    Doc.loadXML XmlString

    Dim RootElem As MSXML2.IXMLDOMElement
    Set RootElem = Doc.documentElement
    
    Dim DataElem As MSXML2.IXMLDOMElement
    For Each DataElem In RootElem.childNodes
        m_Keys.Add DataElem.nodeName
        m_Values.Add DataElem.Text
    Next DataElem
    
End Sub


Private Sub Class_Initialize()
    Set m_Keys = New Collection
    Set m_Values = New Collection
End Sub
