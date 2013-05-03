Attribute VB_Name = "sql_sp"
Option Explicit
Function userprofile_get() As Boolean
Dim sqlstatus As Boolean
    'Retrieve User Profiles
    ConnectDB.DBConnect
    sqlstatus = sql_query.qry_select(db_table:="UserProfiles", strRetSheet:="UserProfiles")
    
    userprofile_get = sqlstatus
    
    If Not sqlstatus Then
        sql_query.QueryFail
        Exit Function
    End If
    ConnectDB.closeDB myDbHandle:=myDbHandle
End Function

Function userprofile_insert(ByVal userinfo) As Boolean
' Userinfo is a vector of 5 strings as follows:
    ' 1: First name
    ' 2: Last name
    ' 3: PermissionType
    ' 4: Username
    ' 5: Password
Dim sqlstatus As Boolean
Dim strValues As String
Dim i As Integer

' check to sure vector is of correct length
    If UBound(userinfo) <> 5 Then
         userprofile_insert = False
        sql_query.QueryFail
        Exit Function
    End If
    
' Concatinate string
    strValues = "NULL,"
    
    For i = 1 To 5
        strValues = strValues & "'" & userinfo(i) & "',"
    Next
    
    strValues = Left(strValues, Len(strValues) - 1)

' Run Insert Query
    ConnectDB.DBConnect
    
    sqlstatus = sql_query.insert(db_table:="UserProfiles", values:=strValues)
    
    userprofile_insert = sqlstatus

    If Not sqlstatus Then
        sql_query.QueryFail
        Exit Function
    End If
    
    ConnectDB.closeDB myDbHandle:=myDbHandle
End Function
