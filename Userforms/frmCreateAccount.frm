VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCreateAccount 
   Caption         =   "Create User Account"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6150
   OleObjectBlob   =   "frmCreateAccount.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCreateAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
Dim shtProfile As Worksheet

    On Error Resume Next
        Set shtProfile = Sheets("UserProfiles")

        Application.DisplayAlerts = False
        shtProfile.Delete
        Application.DisplayAlerts = False
    On Error GoTo 0
    
    Unload Me
End Sub

Private Sub btnCreateAccount_Click()
Dim userinfo(1 To 5) As String
Dim i As Integer
Dim shtProfile As Worksheet
Dim spstatus As Boolean, cPwd As Boolean



    On Error Resume Next
    Set shtProfile = Sheets("UserProfiles")
    On Error GoTo 0
    
    'Check if UserProfiles have been extracted from sheet
    If shtProfile Is Nothing Then
        sql_sp.userprofile_get
        Set shtProfile = Sheets("UserProfiles")
    End If

    userinfo(1) = Me.bxFirstname
    userinfo(2) = Me.bxLastName
    userinfo(3) = Me.cbxEmployee
    userinfo(4) = Me.bxUsername
    userinfo(5) = Me.bxPassword
    
    ' Checking constraints on passwords. Passwords much be atleast six letters long and contain at least one number
    cPwd = checkpwd(Me.bxPassword.Text)
    
    If Not cPwd Then
        PwdError
        Me.bxPassword = ""
        Me.bxPasswordConfirm = ""
        Exit Sub
    End If
    
    ' Check if password if correct
    If Me.bxPassword <> Me.bxPasswordConfirm Then
        MsgBox Prompt:="The password you have entered does not match. Please reenter your password.", _
               Buttons:=vbOKOnly, Title:="Password Mismatch."
        Me.bxPassword = ""
        Me.bxPasswordConfirm = ""
        Exit Sub
    End If
    
    ' Check if username exists
    For i = 1 To shtProfile.UsedRange.Rows.Count
        If Me.bxUsername = shtProfile.Cells(i, 5) Then
            MsgBox Prompt:="The username you are attempting to create already exists. Please try another username.", _
               Buttons:=vbOKOnly, Title:="Duplicate Username."
               Me.bxUsername = ""
               Me.bxPassword = ""
               Me.bxPasswordConfirm = ""
            Exit Sub
        End If
    Next
    
    ' Insert User account information into database
    spstatus = sql_sp.userprofile_insert(userinfo:=userinfo)
    
    On Error Resume Next
        Application.DisplayAlerts = False
        shtProfile.Delete
        Application.DisplayAlerts = False
    On Error GoTo 0
    
    Unload Me
    
End Sub

Private Sub UserForm_Initialize()
    'Centres the user form
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
    
    ' Add options to to the employee type drop box
    With Me.cbxEmployee
        .AddItem "Employee"
        .AddItem "Admin"
        .Value = "Employee"
    End With

End Sub

Private Function checkpwd(strPwd As String) As Boolean
Dim i As Integer
Dim strCurrent As String
Dim cLen As Boolean, cNumber As Boolean, cLetter As Boolean
Dim test As Variant
    cLen = Len(strPwd) >= 6

    For i = 1 To Len(strPwd)
        strCurrent = Mid(strPwd, i, 1)
        Select Case True
            Case IsNumeric(strCurrent)
                cNumber = True
            Case Asc(strCurrent) >= 65 And Asc(strCurrent) <= 90 ' English characters
                cLetter = True
            Case Asc(strCurrent) >= 97 And Asc(strCurrent) <= 122 ' English characters
                cLetter = True
            Case AscW(strCurrent) >= 3585 And AscW(strCurrent) <= 3642 ' Thai characters
                cLetter = True
            Case AscW(strCurrent) >= 3647 And AscW(strCurrent) <= 3675 ' Thai characters
                cLetter = True
        End Select
    Next
    
    checkpwd = cLen And cNumber And cLetter

End Function

Private Sub PwdError()
    MsgBox Prompt:="The password is invalid. It must be at least six characters and contain atleast one letter" & _
                   " and one number. Please try again.", _
            Buttons:=vbOKOnly, Title:="Invalid Password."
End Sub
