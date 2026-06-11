VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsPostalAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=================================================================
' Implements a postal address of a organization or an individual
' person
'=================================================================

Private m_Street As String
Private m_StreetNumber As String
Private m_ZipCode As String
Private m_City As String
Private m_CountryCode As String

Public Property Get Street() As String
    Street = m_Street
End Property

Public Property Get StreetNumber() As String
    StreetNumber = m_StreetNumber
End Property

Public Property Get ZipCode() As String
    ZipCode = m_ZipCode
End Property

Public Property Get City() As String
    City = m_City
End Property

Public Property Get CountryCode() As String
    CountryCode = m_CountryCode
End Property

Public Property Let Street(ByVal pStreet As String)
    m_Street = pStreet
End Property

Public Property Let StreetNumber(ByVal pStreetNum As String)
    m_StreetNumber = pStreetNum
End Property

Public Property Let ZipCode(ByVal pZipCode As String)
    m_ZipCode = pZipCode
End Property

Public Property Let City(ByVal pCity As String)
    m_City = pCity
End Property

Public Property Let CountryCode(ByVal pCountryCode As String)
    m_CountryCode = pCountryCode
End Property

'=================================================================
' clones the postal address as e real copy
'=================================================================
Public Function Clone() As ClsPostalAddress

    Dim addr As ClsPostalAddress
    Set addr = New ClsPostalAddress
    
    With addr
        .City = m_City
        .CountryCode = m_CountryCode
        .Street = m_Street
        .StreetNumber = m_StreetNumber
        .ZipCode = m_ZipCode
    End With
    
    Set Clone = addr

End Function
