Attribute VB_Name = "createdb"
Sub CreateDBDirectory()
Dim strFolder As String
Dim RetVal As Long
Dim i As Integer
Dim sqlstatus As Boolean
Dim rngHeader As Range

    
    ' This procedure creates a blank database,its tables and all directory folders to be used by the DermaDB spreadsheet.
    ' Will be called on the first initialization of the spreadsheet or if a new database is to be created.
    
    ' Select Folder path
    strFolder = GetFolder
    If strFolder = "" Then Exit Sub
    
    'Create DLL Folder and place dlls in folders
    On Error Resume Next
    MkDir Path:=strFolder & "\dll"
    MkDir Path:=strFolder & "\dll\x32"
    MkDir Path:=strFolder & "\dll\x64"
    MkDir Path:=strFolder & "\Patient"
    On Error GoTo 0
    
    FileCopy Application.ActiveWorkbook.Path & "\dll\x32\sqlite3.dll", _
                 strFolder & "\dll\x32\sqlite3.dll"
    FileCopy Application.ActiveWorkbook.Path & "\dll\x32\SQLite3_StdCall.dll", _
                 strFolder & "\dll\x32\SQLite3_StdCall.dll"
    FileCopy Application.ActiveWorkbook.Path & "\dll\x64\sqlite3.dll", _
                strFolder & "\dll\x64\sqlite3.dll"
                 
    ' Create New database
    ConnectDB.initDLL strDBpath:=strFolder
    RetVal = lib_Sqlite3.SQLite3Open(strFolder & "\DermaDB.db3", myDbHandle)
    
    ' Save Folder path and DBnames
    DBStore.Range("DBPath") = strFolder
    DBStore.Range("DBName") = "DermaDB.db3"
    
    ' Create Tables
    For i = 2 To DBTables.UsedRange.Columns.Count
        Set rngHeader = DBTables.Cells(1, i)
        Set rngHeader = DBTables.Range(rngHeader, rngHeader.End(xlDown))
        
        sqlstatus = sql_query.createtable(rngHeader)
        
        If Not sqlstatus Then
            MsgBox Prompt:="Creation of Table (" & rngHeader(1, 1) & ") failed.", Buttons:=vbOKOnly
            ConnectDB.closeDB myDbHandle:=myDbHandle
            Exit Sub
        End If
    Next

    'Create Table Indices
    For i = 2 To DBIndex.UsedRange.Columns.Count
    Set rngHeader = DBIndex.Cells(1, i)
        Set rngHeader = DBIndex.Range(rngHeader, rngHeader.End(xlDown))
        
        sqlstatus = sql_query.createindex(rngHeader)
        
        If Not sqlstatus Then
            MsgBox Prompt:="Creation of Index(" & rngHeader(1, 1) & ") on table (" & _
                            rngHeader(2, 1) & ") failed.", Buttons:=vbOKOnly
            ConnectDB.closeDB myDbHandle:=myDbHandle
            Exit Sub
        End If
    Next
    
    
    'Add Administrator Username and Password
    sqlstatus = sql_query.insert(db_table:="UserProfiles", values:="NULL,'ADMIN','ADMIN','ADMIN','ADMIN', 'ADMIN1'")
    
    If Not sqlstatus Then
        MsgBox Prompt:="Insert of Admin Level User failed.", Buttons:=vbOKOnly
        ConnectDB.closeDB myDbHandle:=myDbHandle
        Exit Sub
    End If
    
    ConnectDB.closeDB myDbHandle:=myDbHandle
    
End Sub

Function GetFolder()
Dim dirFolder As FileDialog
Dim sItem As String

    ' This procedure creates a blank database,its tables and all directory folders to be used by the DermaDB spreadsheet.
    ' Will be called on the first initialization of the spreadsheet or if a new database is to be created.
    
    ' Select Folder path
    Set dirFolder = Application.FileDialog(msoFileDialogFolderPicker)
    With dirFolder
        .Title = "Select A Folder"
        .AllowMultiSelect = False
        .InitialFileName = "C:\"
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
GetFolder = sItem
Set dirFolder = Nothing
End Function


