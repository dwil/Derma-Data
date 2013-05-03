VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCreateAccount 
   Caption         =   "Create User Account"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6090
   OleObjectBlob   =   "frmCreateAccount.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCreateAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnCreateAccount_Click()
Dim userinfo(1 To 5) As String
Dim spstatus As Boolean

    userinfo(1) = Me.bxFirstname
    userinfo(2) = Me.bxLastName
    userinfo(3) = "Employee"
    userinfo(4) = Me.bxUsername
    userinfo(5) = Me.bxPassword

    spstatus = sql_sp.userprofile_insert(userinfo:=userinfo)
    
End Sub

Private Sub UserForm_Initialize()
    'Centres the user form
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2

End Sub
