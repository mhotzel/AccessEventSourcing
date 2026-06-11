VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsCoopMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=================================================================
' Implements a individual (natural) Member for the cooperative
'=================================================================

Implements IAggregateRoot

' ----------------- Event sourcing related variables and consts --
Private Const AGG_TYPE_MEMBER = "Member"
Private Const EVT_TYPE_MEMBER_ACCEPTED = "Member.Accepted"
Private Const EVT_TYPE_NAME_CHANGED = "Member.Name.Changed"
Private Const EVT_TYPE_PAYMENT_RECEIVED = "Member.Payment.Received"

Private Const ERRNO_STREAM_EXISTS = 50104

Private m_Version As Long
Private m_UserName As String
Private m_ComputerName As String

Private m_AggregateManager As ClsAggregateManager

' ----------------- Busines related variables and consts ---------
Private m_MemberId As String
Private m_LastName As String
Private m_FirstName As String
Private m_Title As String
Private m_PostalAddress As ClsPostalAddress
Private m_RequestedShares As Long
Private m_PayendShares As Long
Private m_MemberSince As Variant
Private m_MembershipAccepted As Variant

Public Property Get MemberId() As String
    MemberId = m_MemberId
End Property

Public Property Get MemberSince() As Date
    MemberSince = m_MemberSince
End Property

Public Property Get LastName() As String
    LastName = m_LastName
End Property

Public Property Get FirstName() As String
    FirstName = m_FirstName
End Property

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Get PostalAddress() As ClsPostalAddress
    Set PostalAddress = m_PostalAddress
End Property

Public Property Get RequestedShares() As Long
    RequestedShares = m_RequestedShares
End Property

Public Property Get PayendShares() As Long
    PayendShares = m_PayendShares
End Property

'=================================================================
' Accepts the Application for the membership of a legal party or individual person
'   Parameters:
'       pLastName:  The last name of a natural person or the name of the organization
'=================================================================
Public Sub AcceptMembership( _
            ByVal pLastName As String, ByVal pFirstName As String, ByVal pTitle As String, _
            ByRef pPostalAddress As ClsPostalAddress, _
            ByVal pNumberOfShares As Integer, ByVal pDateAccepted As Date)

    If AggregateManager.IsEmpty = False Then
        Err.Raise ERRNO_STREAM_EXISTS, "ClsMember.AcceptMembership", "Das Mitglied wurde bereits registriert"
    End If
    
    Dim Payload As ClsEventPayload
    Set Payload = New ClsEventPayload
    
    With Payload
        .AddAttribute "lastname", pLastName
        .AddAttribute "firstname", pFirstName
        .AddAttribute "title", pTitle
        
        .AddAttribute "city", pPostalAddress.City
        .AddAttribute "countrycode", pPostalAddress.CountryCode
        .AddAttribute "street", pPostalAddress.Street
        .AddAttribute "streetnumber", pPostalAddress.StreetNumber
        .AddAttribute "zipcode", pPostalAddress.ZipCode
        
        .AddAttribute "shares", pNumberOfShares
        .AddAttribute "accepted", Format$(pDateAccepted, "yyyy-MM-dd")
        
    End With
    
    Dim De As ClsDomainEvent: Set De = New ClsDomainEvent
    De.Init EVT_TYPE_MEMBER_ACCEPTED, ModUUIDGenerator.GenerateUUIDv7, Payload.AsXML.xml
    AggregateManager.RecordThat De

End Sub


'=================================================================
' Changes the name of the member
'=================================================================
Public Sub ChangeName(ByVal pLastName As String, _
                         Optional ByVal pValidFrom As Date = 0)
    
    Dim Payload As ClsEventPayload
    Set Payload = New ClsEventPayload
    
    If pValidFrom = 0 Then
        pValidFrom = Now
    End If
    
    Payload.AddAttribute "lastname", pLastName
    Payload.AddAttribute "validfrom", Format$(pValidFrom, "yyyy-MM-dd")
    
    Dim De As ClsDomainEvent: Set De = New ClsDomainEvent
    De.Init EVT_TYPE_NAME_CHANGED, Me.MemberId, Payload.AsXML.xml
    AggregateManager.RecordThat De

End Sub

'=================================================================
' Changes the first name of the member
'=================================================================
Public Sub ChangeFirstName(ByVal pFirstName As String, _
                         Optional ByVal pEffectiveDate As Date = 0)

End Sub

'=================================================================
' Changes the address of the member
'=================================================================
Public Sub Relocate(ByVal pStreet As String, ByVal pStreetNo, _
                         ByVal pZipCode As String, ByVal pCity, _
                         Optional ByVal pEffectiveDate As Date = 0)

End Sub


'=================================================================
' Confirms that the member payed the shares. The membership is
' active, when the member payed the shares,
'=================================================================
Public Sub ConfirmPayment(ByVal pSharesPayed As Long, ByVal pPaymentRecieved As Date)
    
    Dim Payload As ClsEventPayload
    Set Payload = New ClsEventPayload
    Payload.AddAttribute "paymentreceived", Format(pPaymentRecieved, "yyyy-MM-dd")
    Payload.AddAttribute "sharespayed", pSharesPayed
    
    Dim De As ClsDomainEvent: Set De = New ClsDomainEvent
    De.Init EVT_TYPE_PAYMENT_RECEIVED, Me.MemberId, Payload.AsXML.xml
    AggregateManager.RecordThat De
End Sub

'=================================================================
' Adds or changes the phone number
'=================================================================
Public Sub ChangePhoneNumber(ByVal pNewPhoneNo As String)

End Sub

'=================================================================
' Adds or changes the email address
'=================================================================
Public Sub ChangeEmailAddress(ByVal pNewEmail As String)

End Sub

'=================================================================
' Applies the Event to this AggregateRoot. Here you have to
' implement the logic to update the state of this aggregate root
'=================================================================
Public Sub IAggregateRoot_Apply(ByRef pEvent As ClsDomainEvent)

    Select Case pEvent.EventType
        Case EVT_TYPE_MEMBER_ACCEPTED
            ApplyMemberAccepted pEvent
        Case EVT_TYPE_NAME_CHANGED
            ApplyMemberNameChanged pEvent
        Case EVT_TYPE_PAYMENT_RECEIVED
            ApplyPaymentReceived pEvent
    End Select

End Sub

'=================================================================
' Applies the Acceptance of a new member
'=================================================================
Private Sub ApplyMemberAccepted(ByRef pEvent As ClsDomainEvent)
    
    Dim Payload As ClsEventPayload
    Set Payload = New ClsEventPayload
    Payload.ReadXml pEvent.Data
    
    m_LastName = Payload.GetAttribute("lastname")
    m_FirstName = Payload.GetAttribute("firstname")
    m_Title = Payload.GetAttribute("title")
    m_MemberId = pEvent.Subject
    
    Set m_PostalAddress = New ClsPostalAddress
    With m_PostalAddress
        .City = Payload.GetAttribute("city")
        .CountryCode = Payload.GetAttribute("countrycode")
        .Street = Payload.GetAttribute("street")
        .StreetNumber = Payload.GetAttribute("streetnumber")
        .ZipCode = Payload.GetAttribute("zipcode")
    End With
    
    m_RequestedShares = CLng(Payload.GetAttribute("shares"))
    m_MembershipAccepted = DateValue(Payload.GetAttribute("accepted"))
    
End Sub

Public Function IAggregateRoot_ReleaseEvents() As Collection
    Set IAggregateRoot_ReleaseEvents = m_AggregateManager.ReleaseEvents
End Function

'=================================================================
' Applies the Change of a member name
'=================================================================
Public Sub ApplyMemberNameChanged(ByRef pEvent As ClsDomainEvent)

    Dim Payload As ClsEventPayload
    Set Payload = New ClsEventPayload
    Payload.ReadXml pEvent.Data
    
    m_LastName = Payload.GetAttribute("lastname")
    
End Sub

'=================================================================
' Applies the Confirmation, that the member payed the shares
'=================================================================
Public Sub ApplyPaymentReceived(ByRef pEvent As ClsDomainEvent)

    Dim Payload As ClsEventPayload
    Set Payload = New ClsEventPayload
    Payload.ReadXml pEvent.Data
    
    m_PayendShares = m_PayendShares + CLng(Payload.GetAttribute("sharespayed"))
    If PayendShares >= RequestedShares Then
        m_MemberSince = Payload.GetAttribute("paymentreceived")
    End If
    
End Sub

Private Sub Class_Initialize()

    m_UserName = Environ("USERNAME")
    m_ComputerName = Environ("COMPUTERNAME")
    m_MemberSince = Null
    m_MembershipAccepted = Null

End Sub

Private Property Get AggregateManager() As ClsAggregateManager
    If m_AggregateManager Is Nothing Then
        Set m_AggregateManager = New ClsAggregateManager
    End If
    Set AggregateManager = m_AggregateManager
    AggregateManager.Init Me, "DlsManager", "Member"
End Property
