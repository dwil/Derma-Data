VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "Enter Username and Password"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7305
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
Dim shtProfile As Worksheet

    On Error Resume Next
        Set shtProfile = Sheets("UserProfiles")
 
        CurrentUser.Range("UserID") = ""
        CurrentUser.Range("Username") = ""
        CurrentUser.Range("UserValidation") = False
    
        Application.DisplayAlerts = False
        shtProfile.Delete
        Application.DisplayAlerts = False
    On Error GoTo 0
    
    Unload Me
End Sub


Private Sub btnCreateAccount_Click()
    frmCreateAccount.Show
End Sub

Private Sub btnOkay_Click()
Dim shtProfile As Worksheet
Dim Match As Boolean
Dim i As Integer

    On Error Resume Next
    Set shtProfile = Sheets("UserProfiles")
    On Error GoTo 0
    
    'Check if UserProfiles have been extracted from sheet
    If shtProfile Is Nothing Then
        sql_sp.userprofile_get
        Set shtProfile = Sheets("UserProfiles")
    End If
    
    Match = False
    
    ' Search Profiles to find matching User name and password combination.
    For i = 1 To shtProfile.UsedRange.Rows.Count
        If Me.bxUsername = shtProfile.Cells(i, 5) And Me.bxPassword = shtProfile.Cells(i, 6) Then
            Match = True
            Exit For
        End If
    Next
    
    If Match Then
        CurrentUser.Range("UserID") = shtProfile.Cells(i, 1)
        CurrentUser.Range("Username") = shtProfile.Cells(i, 5)
        CurrentUser.Range("LastActivity") = Now
        Unload Me
    Else
        Me.bxUsername.Text = ""
        Me.bxPassword.Text = ""
        CurrentUser.Range("UserID") = ""
        CurrentUser.Range("Username") = ""
        MsgBox Prompt:="The username and/or password you have entered is incorrect. Usernames and passwords " & _
                       "are case sensitive. Please try again.", _
               Buttons:=vbOKOnly, Title:="Incorrect username/password."
    End If
    
    On Error Resume Next
        Application.DisplayAlerts = False
        shtProfile.Delete
        Application.DisplayAlerts = False
    On Error GoTo 0
    
    CurrentUser.Range("UserValidation") = Match
        
End Sub

Private Sub UserForm_Initialize()
Dim spstatus As Boolean
    'Centers the userform on the screen
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
        
    spstatus = sql_sp.userprofile_get
    
End Sub

