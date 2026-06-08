Attribute VB_Name = "ModUUIDGenerator"
#If VBA7 Then
    Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (pguid As Any) As Long
    
    Private Declare PtrSafe Sub GetSystemTimeAsFileTime Lib "kernel32" (ByRef lpSystemTimeAsFileTime As Currency)
    Private Declare PtrSafe Sub GetSystemTimePreciseAsFileTime Lib "kernel32" (ByRef lpSystemTimeAsFileTime As Currency)
    
    Private Declare PtrSafe Function BCryptGenRandom Lib "bcrypt.dll" ( _
        ByVal hAlgorithm As LongPtr, _
        ByRef pbBuffer As Any, _
        ByVal cbBuffer As Long, _
        ByVal dwFlags As Long) As Long
    
#Else
    Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As Any) As Long
    
    Private Declare Sub GetSystemTimeAsFileTime Lib "kernel32" (ByRef lpSystemTimeAsFileTime As Currency)
    Private Declare Sub GetSystemTimePreciseAsFileTime Lib "kernel32" (ByRef lpSystemTimeAsFileTime As Currency)
        
    Private Declare Function BCryptGenRandom Lib "bcrypt.dll" ( _
        ByVal hAlgorithm As Long, _
        ByRef pbBuffer As Any, _
        ByVal cbBuffer As Long, _
        ByVal dwFlags As Long) As Long
        
#End If

' Flag, damit Windows seinen eigenen bevorzugten Zufallsgenerator nutzt
Private Const BCRYPT_USE_SYSTEM_FAVORED_RNG As Long = 2

'Erzeugt eine UUID Version 4
Public Function GenerateUUIDv4() As String
    Dim id(15) As Byte
    Dim result As String
    
    If CoCreateGuid(id(0)) = 0 Then
        ' Beachtung der Little-Endian Struktur von Windows-GUIDs bei der Textwandlung:
        ' Block 1 (Bytes 3, 2, 1, 0)
        result = result & Right$("0" & Hex$(id(3)), 2) & Right$("0" & Hex$(id(2)), 2) & _
                          Right$("0" & Hex$(id(1)), 2) & Right$("0" & Hex$(id(0)), 2) & "-"
        
        ' Block 2 (Bytes 5, 4)
        result = result & Right$("0" & Hex$(id(5)), 2) & Right$("0" & Hex$(id(4)), 2) & "-"
        
        ' Block 3 (Bytes 7, 6) -> Hier wird die "4" sichtbar!
        result = result & Right$("0" & Hex$(id(7)), 2) & Right$("0" & Hex$(id(6)), 2) & "-"
        
        ' Block 4 (Bytes 8, 9) -> Hier wird 8, 9, A oder B sichtbar!
        result = result & Right$("0" & Hex$(id(8)), 2) & Right$("0" & Hex$(id(9)), 2) & "-"
        
        ' Block 5 (Bytes 10 bis 15)
        Dim i As Long
        For i = 10 To 15
            result = result & Right$("0" & Hex$(id(i)), 2)
        Next i
        
        GenerateUUIDv4 = LCase$(result)
    End If
End Function

'Erzeugt eine UUID Version 7, also chronologisch monoton aufsteigend
Public Function GenerateUUIDv7() As String
    Dim fileTime As Currency
    Dim tempMs As Currency
    Dim remainder As Currency
    Dim b(0 To 15) As Byte
    Dim i As Long
    
    ' 1. Hochpräzise Zeit holen (Genauigkeit < 1 Mikrosekunde)
    GetSystemTimePreciseAsFileTime fileTime
    
    ' 2. Sub-Millisekunden-Anteil extrahieren
    ' Hinter dem Komma von fileTime stehen die 100-ns Ticks (0 bis 9999)
    Dim subMsTicks As Long
    subMsTicks = CLng((fileTime - Int(fileTime)) * 10000@)
    
    ' Diese 0-9999 Ticks skalieren wir exakt auf die verfügbaren 12 Bits (0 bis 4095)
    Dim subMs12Bit As Long
    subMs12Bit = Int(subMsTicks * 4096 / 10000)
    
    ' 3. Umrechnung des Ganzzahl-Anteils in Unix-Zeit (Millisekunden)
    tempMs = Int(fileTime - 11644473600000@)
    
    ' 4. Die 48-Bit-Zeit in Bytes 0-5 schreiben
    For i = 5 To 0 Step -1
        remainder = tempMs - Int(tempMs / 256@) * 256@
        b(i) = CByte(remainder)
        tempMs = Int(tempMs / 256@)
    Next i
    
    ' 5. Restliche 10 Bytes mit Krypto-Zufall füllen
    If BCryptGenRandom(0, b(6), 10, BCRYPT_USE_SYSTEM_FAVORED_RNG) <> 0 Then
        GenerateUUIDv7 = "ERROR: Zufallsgenerator fehlgeschlagen"
        Exit Function
    End If
    
    ' 6. Sub-Millisekunden-Bits und Version/Variante einweben
    ' Byte 6: Version 7 (0x70) + die oberen 4 Bits des skalierten Zeit-Anteils
    b(6) = &H70 Or ((subMs12Bit \ 256) And &HF)
    
    ' Byte 7: Die unteren 8 Bits des skalierten Zeit-Anteils
    b(7) = subMs12Bit And &HFF
    
    ' Byte 8: Variante 1 (0x80)
    b(8) = (b(8) And &H3F) Or &H80
    
    ' 7. In Hex-String wandeln
    Dim uuidStr As String
    For i = 0 To 15
        uuidStr = uuidStr & LCase$(Right$("0" & Hex$(b(i)), 2))
        If i = 3 Or i = 5 Or i = 7 Or i = 9 Then
            uuidStr = uuidStr & "-"
        End If
    Next i
    
    GenerateUUIDv7 = uuidStr
End Function

'Erzeugt einen 16-Byte-Wert aus einer UUID, um diese
'platzsparend in einem MS-Access-Feld vom Typ "Replikations-ID", das ist der Datentyp "Guid" zu speichern.
'
'Achtung: Das Speichern sollte dann aber über DAO-RecordSets erfolgen!
'
'Bei Nutzung von QueryDefs: Der Datentyp nach "PARAMETERS " lautet "Guid", also z.B.
' "PARAMETERS [prmID] Guid"
'
Public Function UuidStringToAccessBinary(ByVal strUuid As String) As Byte()
    Dim b(0 To 15) As Byte
    Dim cleanUuid As String
    Dim i As Integer, pos As Integer
    
    ' 1. Alle Bindestriche und mögliche Leerzeichen entfernen
    cleanUuid = Replace(strUuid, "-", "")
    cleanUuid = Trim$(cleanUuid)
    
    ' Sicherheitspruefung: Eine UUID muss ohne Bindestriche exakt 32 Zeichen lang sein
    If Len(cleanUuid) <> 32 Then
        Err.Raise 5, "UuidV7StringToAccessBinary", "Ungültige UUID-Länge!"
    End If
    
    ' 2. Der magische Byte-Swap für die Access/Windows Little-Endian-Logik
    
    ' Block 1: Bytes 0 bis 3 (Datentyp Long -> komplett umdrehen)
    b(3) = CByte("&H" & Mid$(cleanUuid, 1, 2))
    b(2) = CByte("&H" & Mid$(cleanUuid, 3, 2))
    b(1) = CByte("&H" & Mid$(cleanUuid, 5, 2))
    b(0) = CByte("&H" & Mid$(cleanUuid, 7, 2))
    
    ' Block 2: Bytes 4 bis 5 (Datentyp Integer -> umdrehen)
    b(5) = CByte("&H" & Mid$(cleanUuid, 9, 2))
    b(4) = CByte("&H" & Mid$(cleanUuid, 11, 2))
    
    ' Block 3: Bytes 6 bis 7 (Datentyp Integer -> umdrehen)
    b(7) = CByte("&H" & Mid$(cleanUuid, 13, 2))
    b(6) = CByte("&H" & Mid$(cleanUuid, 15, 2))
    
    ' Block 4 & 5: Bytes 8 bis 15 (Datentyp Byte-Array -> normale Reihenfolge!)
    pos = 17
    For i = 8 To 15
        b(i) = CByte("&H" & Mid$(cleanUuid, pos, 2))
        pos = pos + 2
    Next i
    
    ' 3. Das fertige 16-Byte-Array zurückgeben
    UuidStringToAccessBinary = b
End Function

'Mit dieser Funktion kann man den 16-Byte-Wert wieder zurueck in die String-Repräsentation zumwandeln
Public Function AccessBinaryToUuidString(ByRef b() As Byte) As String
    Dim hexParts(0 To 15) As String
    Dim i As Integer
    
    ' Sicherheitsprüfung: Ist es wirklich ein 16-Byte-Array?
    If UBound(b) < 15 Then
        Err.Raise 5, "AccessBinaryToUuidString", "Das Array muss 16 Bytes (0 bis 15) enthalten!"
    End If
    
    ' 1. Alle Bytes in zweistellige Hexadezimal-Werte umwandeln
    For i = 0 To 15
        ' Right$ und "0" sorgen dafür, dass z. B. aus einem "A" ein "0A" wird
        hexParts(i) = Right$("0" & Hex$(b(i)), 2)
    Next i
    
    ' 2. Den String zusammenbauen und den Byte-Swap RÜCKGÄNGIG machen.
    ' Block 1 bis 3 werden wieder rückwärts gelesen, Block 4 und 5 normal.
    ' LCase$ wandelt die großen Hex-Buchstaben in die Standard-Kleinschreibung von UUIDs um.
    
    AccessBinaryToUuidString = LCase$( _
        hexParts(3) & hexParts(2) & hexParts(1) & hexParts(0) & "-" & _
        hexParts(5) & hexParts(4) & "-" & _
        hexParts(7) & hexParts(6) & "-" & _
        hexParts(8) & hexParts(9) & "-" & _
        hexParts(10) & hexParts(11) & hexParts(12) & hexParts(13) & hexParts(14) & hexParts(15) _
    )
End Function

