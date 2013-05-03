Attribute VB_Name = "ConnectDB"
Option Explicit

Sub DBConnect()
Dim strDBpath As String
Dim strDBName As String
Dim a As String
    
    ' Retrieve database path and name.
    strDBpath = DBStore.Range("DBPath")
    strDBName = DBStore.Range("DBName")
    
    ' If path and name are not present then ask user to select a database.
    If strDBpath = "" Or strDBName = "" Then
        MsgBox Prompt:="There is no link to a database. Please select a database to link to.", _
               Buttons:=vbOKOnly
        strDBpath = ConnectDB.selectDBpath
        
        If strDBpath = "FALSE" Then Exit Sub
        
        strDBName = Right(strDBpath, Len(strDBpath) - InStrRev(strDBpath, "\"))
        strDBpath = Left(strDBpath, InStrRev(strDBpath, "\") - 1)
        
        DBStore.Range("DBPath") = strDBpath
        DBStore.Range("DBName") = strDBName
    End If
    
    
    ConnectDB.initDLL strDBpath:=strDBpath
    ConnectDB.initDB strDBpath:=strDBpath & "\" & strDBName
End Sub

Public Sub initDLL(strDBpath As String)
    Dim InitReturn As Long
    
    #If Win64 Then
        '64-bit .dll in subdirectory
        InitReturn = lib_Sqlite3.SQLite3Initialize(strDBpath & "\dll\x64")
    #Else
        InitReturn = lib_Sqlite3.SQLite3Initialize(strDBpath & "\dll\x32")
    #End If
    
    If InitReturn <> SQLITE_INIT_OK Then
        Exit Sub
    End If
    
End Sub

Public Sub initDB(strDBpath As String)
    Dim RetVal As Long
    RetVal = lib_Sqlite3.SQLite3Open(strDBpath, myDbHandle)
End Sub

Public Sub closeDB(myDbHandle)
    Dim RetVal As Long
    RetVal = lib_Sqlite3.SQLite3Close(myDbHandle)
End Sub


Function selectDBpath()
    selectDBpath = Application.GetOpenFilename("Select Database (*.db3), *.db3", , "Select a database file.")
End Function
