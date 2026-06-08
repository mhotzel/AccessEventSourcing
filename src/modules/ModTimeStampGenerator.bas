Attribute VB_Name = "ModTimeStampGenerator"
Option Compare Database
Option Explicit

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

' Nur noch zwei API-Funktionen nötig!
#If VBA7 Then
    Private Declare PtrSafe Sub GetSystemTimePreciseAsFileTime Lib "kernel32" (ByRef lpSystemTimeAsFileTime As Currency)
    Private Declare PtrSafe Function FileTimeToSystemTime Lib "kernel32" (ByRef lpFileTime As Currency, ByRef lpSystemTime As SYSTEMTIME) As Long
#Else
    Private Declare Sub GetSystemTimePreciseAsFileTime Lib "kernel32" (ByRef lpSystemTimeAsFileTime As Currency)
    Private Declare Function FileTimeToSystemTime Lib "kernel32" (ByRef lpFileTime As Currency, ByRef lpSystemTime As SYSTEMTIME) As Long
#End If

Public Function GetUtcIsoTimestamp() As String
    Dim utcTime As Currency
    Dim st As SYSTEMTIME
    
    ' 1. Zeit in UTC abrufen (mikrosekundengenau)
    GetSystemTimePreciseAsFileTime utcTime
    
    ' 2. Den rohen UTC-Wert DIREKT in Jahr, Monat, Tag, etc. übersetzen (ohne Umweg über LocalTime)
    FileTimeToSystemTime utcTime, st
    
    ' 3. Mikrosekunden berechnen (.u)
    Dim decTime As Variant
    Dim raw100ns As Variant
    Dim secFraction As Variant
    Dim microSecs As Long
    
    decTime = CDec(utcTime)
    raw100ns = decTime * 10000
    
    ' Alles über 1 Sekunde (10.000.000 Intervalle) abschneiden
    secFraction = raw100ns - Fix(raw100ns / 10000000) * 10000000
    
    ' Durch 10 teilen ergibt die 6-stelligen Mikrosekunden
    microSecs = Fix(secFraction / 10)
    
    ' 4. Den finalen String zusammensetzen
    ' Das "Z" am Ende ist der weltweite ISO-Standard für UTC.
    ' (Falls Ihr empfangendes System zwingend "+00:00" statt "Z" erwartet,
    ' tauschen Sie das "Z" hier einfach gegen "+00:00" aus).
    GetUtcIsoTimestamp = Format$(st.wYear, "0000") & "-" & _
                         Format$(st.wMonth, "00") & "-" & _
                         Format$(st.wDay, "00") & "T" & _
                         Format$(st.wHour, "00") & ":" & _
                         Format$(st.wMinute, "00") & ":" & _
                         Format$(st.wSecond, "00") & "." & _
                         Format$(microSecs, "000000") & ":+00:00"
                         
End Function


