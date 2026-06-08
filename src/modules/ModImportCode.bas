Attribute VB_Name = "ModImportCode"
Option Compare Database
Option Explicit

Private Const ModulesDir = "modules"
Private Const FormsDir = "forms"
Private Const ReportsDir = "reports"
Private Const TablesDir = "tables"
Private Const QueriesDir = "queries"

' Importiert die VBA-Quellen aus einem Unterverzeichnis "src",
' um eine Datenbank wieder aus den Sourcen herzustellen.
Public Sub StartImport()

    Dim GuidForScriptingRuntime As String
    GuidForScriptingRuntime = "{420B2830-E71F-11D0-893D-00A0C9054228}"
    
    Dim GuidForVbeExtensibility As String
    GuidForVbeExtensibility = "{0002E157-0000-0000-C000-000000000046}"
    
    ' Verweise dynamisch hinzufügen
    AddReferenceByGuid GuidForScriptingRuntime, 1, 0, "Microsoft Scripting Runtime 1.0"
    AddReferenceByGuid GuidForVbeExtensibility, 5, 3, "Microsoft Visual Basic for Applications Extensibility 5.3"
    
    Call ImportAll

End Sub

Private Sub ImportAll()

    ' Late Binding: Als Object deklariert, um Kompilierfehler vor dem Laden der Verweise zu verhindern
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim ExportDir As String: ExportDir = CurrentProject.Path & "\src"
    
    If Not fso.FolderExists(ExportDir) Then
        MsgBox "Der Quellordner 'src' wurde nicht gefunden!", vbCritical, "Fehler"
        Exit Sub
    End If
    
    ' 1. Die Module und Klassenmodule
    Dim SubFolder As String: SubFolder = ExportDir & "\" & ModulesDir
    If fso.FolderExists(SubFolder) Then
        ImportModules SubFolder
    End If
    
    ' 2. Dann die Tabellendefinitionen
    SubFolder = ExportDir & "\" & TablesDir
    If fso.FolderExists(SubFolder) Then
        ImportTables SubFolder
    End If
    
    ' 3. Danach erst die Abfragen
    SubFolder = ExportDir & "\" & QueriesDir
    If fso.FolderExists(SubFolder) Then
        ImportQueries SubFolder
    End If
    
    ' 4. Jetzt nun die Forms
    SubFolder = ExportDir & "\" & FormsDir
    If fso.FolderExists(SubFolder) Then
        ImportForms SubFolder
    End If
    
    ' 5. Am Ende noch die Reports
    SubFolder = ExportDir & "\" & ReportsDir
    If fso.FolderExists(SubFolder) Then
        ImportReports SubFolder
    End If
    
    Set fso = Nothing
    
    MsgBox "Import beendet. Die Datenbank wird nun zum Speichern geschlossen.", vbOKOnly, "Datenbank schließen"
    Application.CloseCurrentDatabase

End Sub

' Fügt die Verweise ins Projekt ein
Public Sub AddReferenceByGuid(GuidStr As String, Major As Long, Minor As Long, RefName As String)
    On Error Resume Next
    
    Application.References.AddFromGuid GuidStr, Major, Minor
    
    Select Case Err.Number
        Case 0
            Debug.Print "Verweis erfolgreich hinzugefügt: " & RefName
        Case 32813
            Debug.Print "Verweis existiert bereits: " & RefName
        Case Else
            Debug.Print "Fehler beim Hinzufügen von '" & RefName & "': " & Err.Description
    End Select
    
    On Error GoTo 0
End Sub

Private Sub ImportModules(pImportDir As String)
    
    Dim VbComp As Object
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim ModName As String
    Dim File As Object
    
    For Each File In fso.GetFolder(pImportDir).Files
        If Right$(File.Name, 4) = ".cls" Or Right$(File.Name, 4) = ".bas" Then
            ModName = Left(File.Name, InStrRev(File.Name, ".") - 1)
            
            ' Annahme: Dieses Modul selbst heißt "ModImportCode" und darf nicht überschrieben werden
            If ModName <> "ModImportCode" Then
                On Error Resume Next
                Set VbComp = Application.VBE.ActiveVBProject.VBComponents(ModName)
                
                On Error GoTo 0
                If Not VbComp Is Nothing Then
                    Application.VBE.ActiveVBProject.VBComponents.Remove VbComp
                End If
                
                Application.VBE.ActiveVBProject.VBComponents.Import File.Path
            End If
        End If
    Next File

End Sub

Private Sub ImportTables(pImportDir As String)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim TableName As String
    Dim File As Object
    
    For Each File In fso.GetFolder(pImportDir).Files
        If Right$(File.Name, 4) = ".xml" Then
            TableName = Left$(File.Name, InStrRev(File.Name, ".") - 1)
            
            ' Existierende Tabelle löschen (Achtung: Löscht Daten, falls vorhanden!)
            On Error Resume Next
            DoCmd.DeleteObject acTable, TableName
            On Error GoTo 0
            
            Application.ImportXML File.Path, acStructureOnly
        End If
    Next File

End Sub

Private Sub ImportQueries(pImportDir As String)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim File As Object
    For Each File In fso.GetFolder(pImportDir).Files
        If Right$(File.Name, 4) = ".sql" Then
            Application.LoadFromText acQuery, Left(File.Name, InStrRev(File.Name, ".") - 1), File.Path
        End If
    Next File

End Sub

Private Sub ImportForms(pImportDir As String)
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim FormName As String
    Dim File As Object
    For Each File In fso.GetFolder(pImportDir).Files
        If Right$(File.Name, 4) = ".acf" Then
            FormName = Left(File.Name, InStrRev(File.Name, ".") - 1)
            
            Application.LoadFromText acForm, FormName, File.Path
        End If
    Next File

End Sub

Private Sub ImportReports(pImportDir As String)
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim ReportName As String
    Dim File As Object
    For Each File In fso.GetFolder(pImportDir).Files
        If Right$(File.Name, 4) = ".rpt" Then
            ReportName = Left(File.Name, InStrRev(File.Name, ".") - 1)
            
            ' KORREKTUR: Hier muss acReport stehen!
            Application.LoadFromText acReport, ReportName, File.Path
        End If
    Next File

End Sub
