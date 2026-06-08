Attribute VB_Name = "ModExportCode"
Option Compare Database
Option Explicit

Private Const ModulesDir = "modules"
Private Const FormsDir = "forms"
Private Const ReportsDir = "reports"
Private Const TablesDir = "tables"
Private Const QueriesDir = "queries"


' Exportiert die VBA-Quellen in ein Unterverzeichnis "src",
' um diese zum Beispiel in ein Versionskontrollsystem einzuchecken.
' Es werden automatisch die Unterverzeichnisse für Forms, Reports,
' Module und Klassenmodule sowie Tabellen- und Abfragedefinitionen angelegt.
'
Public Sub ExportAll()

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim ExportDir As String: ExportDir = CurrentProject.Path & "\src"
    Dim SubFolderObj As Object
    
    If Not fso.FolderExists(ExportDir) Then
        fso.CreateFolder ExportDir
    End If
    
    Dim SubFolder As String: SubFolder = ExportDir & "\" & ModulesDir
    If Not fso.FolderExists(SubFolder) Then
        Set SubFolderObj = fso.CreateFolder(SubFolder)
    End If
    
    ExportModules SubFolder
    
    SubFolder = ExportDir & "\" & FormsDir
    If Not fso.FolderExists(SubFolder) Then
        fso.CreateFolder SubFolder
    End If
    
    ExportForms SubFolder
    
    SubFolder = ExportDir & "\" & ReportsDir
    If Not fso.FolderExists(SubFolder) Then
        fso.CreateFolder SubFolder
    End If
    
    ExportReports SubFolder
    
    SubFolder = ExportDir & "\" & TablesDir
    If Not fso.FolderExists(SubFolder) Then
        fso.CreateFolder SubFolder
    End If
    
    ExportTables SubFolder
    
    SubFolder = ExportDir & "\" & QueriesDir
    If Not fso.FolderExists(SubFolder) Then
        fso.CreateFolder SubFolder
    End If
    
    ExportQueries SubFolder
    
    Set fso = Nothing

End Sub

Private Sub ExportForms(pExportDir As String)

    Dim FormDef As Object
    For Each FormDef In CurrentProject.AllForms
        Application.SaveAsText acForm, FormDef.Name, pExportDir & "\" & FormDef.Name & ".acf"
    Next FormDef
    
End Sub

Private Sub ExportReports(pExportDir As String)

    Dim ReportDef As Object
    For Each ReportDef In CurrentProject.AllReports
        Application.SaveAsText acReport, ReportDef.Name, pExportDir & "\" & ReportDef.Name & ".rpt"
    Next ReportDef
    
End Sub

Private Sub ExportModules(pExportDir As String)

    Dim VbComp As VBIDE.VBComponent
    For Each VbComp In Application.VBE.ActiveVBProject.VBComponents
        Select Case VbComp.Type
            Case vbext_ct_StdModule
                VbComp.Export pExportDir & "\" & VbComp.Name & ".bas"
            Case vbext_ct_ClassModule
                VbComp.Export pExportDir & "\" & VbComp.Name & ".cls"
        End Select
    Next VbComp
    
End Sub

Private Sub ExportTables(pExportDir As String)

    Dim tbl As DAO.TableDef
    For Each tbl In CurrentDb.TableDefs
        'Keine Systemtabellen
        If Not (tbl.Name Like "MSys*" Or tbl.Name Like "~*") Then
            Application.ExportXML _
                ObjectType:=acTable, _
                DataSource:=tbl.Name, _
                SchemaTarget:=pExportDir & "\" & tbl.Name & ".xml", _
                OtherFlags:=acStructureOnly
            End If
    Next tbl

End Sub

Private Sub ExportQueries(pExportDir As String)

    Dim Query As DAO.QueryDef
    For Each Query In CurrentDb.QueryDefs
        Application.SaveAsText acQuery, Query.Name, pExportDir & "\" & Query.Name & ".sql"
    Next Query
    
End Sub


