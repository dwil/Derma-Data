Attribute VB_Name = "sql_query"
Option Explicit
#If Win64 Then
    ' These handles are used for accessing the database.
    ' They need to be public because they will be used by multiple subs.
    Public myDbHandle As LongPtr, myStmtHandle As LongPtr
#Else
    Public myDbHandle As Long, myStmtHandle As Long
#End If

Public Sub QueryFail()
    MsgBox Prompt:="Database operations have failed. Please check that your database connection is correct.", _
           Buttons:=vbokayonly, Title:="Database operations failed."
End Sub


#If Win64 Then
Sub PrintColumns(ByVal stmtHandle As LongPtr, shtOutput As Worksheet, nrow As Integer)
#Else
Sub PrintColumns(ByVal stmtHandle As Long, shtOutput As Worksheet, nrow As Integer)
#End If
' Needed to return query results. Need to look at this.
    Dim colCount As Long
    Dim colName As String
    Dim colType As Long
    Dim colTypeName As String
    Dim colValue As Variant
    
    Dim i As Long
    
    colCount = SQLite3ColumnCount(stmtHandle)
    Debug.Print "Column count: " & colCount
    For i = 0 To colCount - 1
        colName = SQLite3ColumnName(stmtHandle, i)
        colType = SQLite3ColumnType(stmtHandle, i)
        shtOutput.Cells(nrow, i + 1) = ColumnValue(stmtHandle, i, colType)
    Next
End Sub

#If Win64 Then
Function ColumnValue(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long, ByVal SQLiteType As Long) As Variant
#Else
Function ColumnValue(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long, ByVal SQLiteType As Long) As Variant
#End If
' Needed to return query results. Need to look at this.
    Select Case SQLiteType
        Case SQLITE_INTEGER:
            ColumnValue = SQLite3ColumnInt32(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_FLOAT:
            ColumnValue = SQLite3ColumnDouble(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_TEXT:
            ColumnValue = SQLite3ColumnText(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_BLOB:
            ColumnValue = SQLite3ColumnText(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_NULL:
            ColumnValue = Null
    End Select
End Function

Function createtable(rngHeaders As Range) As Boolean
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
        createtable = False
    Else
        createtable = True
    End If
    
    ' Finalize Statement
    RetVal = lib_Sqlite3.SQLite3Finalize(myStmtHandle)
End Function

Function createindex(rngHeaders As Range) As Boolean
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
        createindex = False
    Else
        createindex = True
    End If
    
    ' Finalize Statement
    RetVal = lib_Sqlite3.SQLite3Finalize(myStmtHandle)
End Function

Function insert(db_table As String, values As String) As Boolean
Dim strQuery As String
Dim RetVal As Long

    strQuery = "INSERT INTO " & db_table & " VALUES (" & values & ")"
        
    'Get statement handle for create table
    RetVal = lib_Sqlite3.SQLite3PrepareV2(myDbHandle, strQuery, myStmtHandle)
    
    'Run statement
    RetVal = lib_Sqlite3.SQLite3Step(myStmtHandle)
    
    'Check if statement Failed
    If RetVal <> lib_Sqlite3.SQLITE_DONE Then
        insert = False
    Else
        insert = True
    End If
    
    RetVal = lib_Sqlite3.SQLite3Finalize(myStmtHandle)
End Function

Function qry_select(db_table As String, strRetSheet As String) As Boolean
Dim strQuery As String
Dim shtReturn As Worksheet
Dim RetVal As Long
Dim i As Integer

    ' Delete return sheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets(strRetSheet).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Create return sheet
    Set shtReturn = Sheets.Add(After:=Worksheets(Worksheets.Count))
    shtReturn.Name = strRetSheet
    
    ' Create Select Query

    strQuery = "SELECT * FROM [" & db_table & "]"
    
    RetVal = lib_Sqlite3.SQLite3PrepareV2(myDbHandle, strQuery, myStmtHandle)
    
    'Run statement
    RetVal = lib_Sqlite3.SQLite3Step(myStmtHandle)
    
    If RetVal = SQLITE_MISUSE Then
        qry_select = False
        RetVal = lib_Sqlite3.SQLite3Finalize(myStmtHandle)
        Exit Function
    End If
    
    i = 1
    
    Do While RetVal = SQLITE_ROW
        PrintColumns myStmtHandle, shtReturn, i
        RetVal = lib_Sqlite3.SQLite3Step(myStmtHandle)
        i = i + 1
    Loop
    
    qry_select = True
    RetVal = lib_Sqlite3.SQLite3Finalize(myStmtHandle)

End Function

