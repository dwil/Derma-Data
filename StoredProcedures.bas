Attribute VB_Name = "StoredProcedures"
Option Explicit
#If Win64 Then
    ' These handles are used for accessing the database.
    ' They need to be public because they will be used by multiple subs.
    Public myDbHandle As LongPtr, myStmtHandle As LongPtr
#Else
    Public myDbHandle As Long, myStmtHandle As Long
#End If

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
        
        ConnectDB.initDLL strDBpath:=strFolder ' Initialize the DLLs to run statements
        sqlstatus = sp_createtable(rngHeader)
        
        If Not sqlstatus Then
            MsgBox prompt:="Creation of Table (" & rngHeader(1, 1) & ") failed.", Buttons:=vbOKOnly
            ConnectDB.closeDB myDbHandle:=myDbHandle
            Exit Sub
        End If
    Next

    'Create Table Indices
    For i = 2 To DBIndex.UsedRange.Columns.Count
    Set rngHeader = DBIndex.Cells(1, i)
        Set rngHeader = DBIndex.Range(rngHeader, rngHeader.End(xlDown))
        
        ConnectDB.initDLL strDBpath:=strFolder ' Initialize the DLLs to run statements
        sqlstatus = sp_createindex(rngHeader)
        
        If Not sqlstatus Then
            MsgBox prompt:="Creation of Index(" & rngHeader(1, 1) & ") on table (" & _
                            rngHeader(2, 1) & ") failed.", Buttons:=vbOKOnly
            ConnectDB.closeDB myDbHandle:=myDbHandle
            Exit Sub
        End If
    Next
    
    
    'Add Administrator Username and Password
    ConnectDB.initDLL strDBpath:=strFolder ' Initialize the DLLs to run statements
    sqlstatus = sp_insert(db_table:="UserProfiles", values:="NULL,'ADMIN','ADMIN','ADMIN','ADMIN', 'ADMIN1'")
    
    If Not sqlstatus Then
        MsgBox prompt:="Insert of Admin Level User failed.", Buttons:=vbOKOnly
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

Function sp_createtable(rngHeaders As Range) As Boolean
Dim i As Integer
Dim strQuery As String
Dim RetVal As Long
    
    ' Compile Create Table statement
    strQuery = "CREATE TABLE " & rngHeaders(1, 1) & " ("

    For i = 2 To rngHeaders.Rows.Count
        strQuery = strQuery & rngHeaders(i, 1) & ", "
    Next
    
    strQuery = Left(strQuery, Len(strQuery) - 2) & ")"
    
    'Get statement handle for create table
    RetVal = lib_Sqlite3.SQLite3PrepareV2(myDbHandle, strQuery, myStmtHandle)
    
    'Run statement
    RetVal = lib_Sqlite3.SQLite3Step(myStmtHandle)
    
    'Check if statement Failed
    If RetVal <> lib_Sqlite3.SQLITE_DONE Then
        sp_createtable = False
    Else
        sp_createtable = True
    End If
    
    ' Finalize Statement
    RetVal = lib_Sqlite3.SQLite3Finalize(myStmtHandle)
End Function

Function sp_createindex(rngHeaders As Range) As Boolean
Dim i As Integer
Dim strQuery As String
Dim RetVal As Long
    
    ' Compile Create Table statement
    strQuery = "CREATE INDEX " & rngHeaders(1, 1) & " ON " & rngHeaders(2, 1) & " ("

    For i = 3 To rngHeaders.Rows.Count
        strQuery = strQuery & rngHeaders(i, 1) & ", "
    Next
    
    strQuery = Left(strQuery, Len(strQuery) - 2) & ")"
    
    'Get statement handle for create table
    RetVal = lib_Sqlite3.SQLite3PrepareV2(myDbHandle, strQuery, myStmtHandle)
    
    'Run statement
    RetVal = lib_Sqlite3.SQLite3Step(myStmtHandle)
    
    'Check if statement Failed
    If RetVal <> lib_Sqlite3.SQLITE_DONE Then
        sp_createindex = False
    Else
        sp_createindex = True
    End If
    
    ' Finalize Statement
    RetVal = lib_Sqlite3.SQLite3Finalize(myStmtHandle)
End Function

Function sp_insert(db_table As String, values As String) As Boolean
Dim strQuery As String
Dim RetVal As Long

    strQuery = "INSERT INTO " & db_table & " VALUES (" & values & ")"
        
        'Get statement handle for create table
    RetVal = lib_Sqlite3.SQLite3PrepareV2(myDbHandle, strQuery, myStmtHandle)
    
    'Run statement
    RetVal = lib_Sqlite3.SQLite3Step(myStmtHandle)
    
    'Check if statement Failed
    If RetVal <> lib_Sqlite3.SQLITE_DONE Then
        sp_insert = False
    Else
        sp_insert = True
    End If
    
    RetVal = lib_Sqlite3.SQLite3Finalize(myStmtHandle)
End Function
