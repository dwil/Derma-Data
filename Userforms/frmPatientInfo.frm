VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPatientInfo 
   Caption         =   "Patient Information"
   ClientHeight    =   13575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11640
   OleObjectBlob   =   "frmPatientInfo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPatientInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    'Centres the user form
    Me.StartUpPosition = 0
    Me.Height = 475.5
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
    
    ' Initialize Comboboxes with values
    With cbxPatientTitle
        .AddItem "Mr."
        .AddItem "Mrs."
        .AddItem "Ms."
    End With
    
    With cbxPatientIDType
        .AddItem "Passport"
        .AddItem "Government ID"
        .AddItem "Driver's License"
        .AddItem "Student ID"
    End With
    
    With lbxFileMenu
        .AddItem "Create New User Account"
        .AddItem "Sign In as Another User"
        .AddItem "Sign Out"
    End With
    
    With lbxPatientMenu
        .AddItem "Save As New Patient"
        .AddItem "Save Current Patient"
        .AddItem "Clear Current Patient"
        .AddItem "Search Patient"
    End With
    
    With lbxAppointmentMenu
        .AddItem "Create New Appointment"
        .AddItem "Save Current Appointment"
    End With
    
End Sub

Private Sub ResetMenuBar()
' Reset Label colours when mouse move off of label
    lblFile.BackStyle = fmBackStyleTransparent
    lblFile.BorderStyle = fmBorderStyleNone
    lbxFileMenu.Visible = False
    lblPatient.BackStyle = fmBackStyleTransparent
    lblPatient.BorderStyle = fmBorderStyleNone
    lbxPatientMenu.Visible = False
    lblAppointment.BackStyle = fmBackStyleTransparent
    lblAppointment.BorderStyle = fmBorderStyleNone
    lbxAppointmentMenu.Visible = False
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetMenuBar
End Sub
Private Sub fmPatientInfo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetMenuBar
End Sub
Private Sub fmIdentification_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetMenuBar
End Sub

'#################################################################################################################'
' Sub Selecting and Showing the Menu Bar
'#################################################################################################################'

Private Sub lblFile_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblFile.BackColor = RGB(255, 218, 91)
    lblFile.BackStyle = fmBackStyleOpaque
    lblFile.BorderStyle = fmBorderStyleSingle
    lblFile.BorderColor = RGB(255, 194, 33)
End Sub
Private Sub lblFile_Click()
' This sub opens the file menu bar when the file label is clicked.
' Position the File Menu list box
    lbxFileMenu.Left = lblFile.Left
    lbxFileMenu.Top = lblFile.Top + lblFile.Height
    lbxFileMenu.Visible = True
End Sub
Private Sub lblPatient_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblPatient.BackColor = RGB(255, 218, 91)
    lblPatient.BackStyle = fmBackStyleOpaque
    lblPatient.BorderStyle = fmBorderStyleSingle
    lblPatient.BorderColor = RGB(255, 194, 33)
End Sub
Private Sub lblPatient_Click()
' This sub opens the file menu bar when the file label is clicked.
' Position the File Menu list box
    lbxPatientMenu.Left = lblPatient.Left
    lbxPatientMenu.Top = lblPatient.Top + lblPatient.Height
    lbxPatientMenu.Visible = True
End Sub
Private Sub lblAppointment_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAppointment.BackColor = RGB(255, 218, 91)
    lblAppointment.BackStyle = fmBackStyleOpaque
    lblAppointment.BorderStyle = fmBorderStyleSingle
    lblAppointment.BorderColor = RGB(255, 194, 33)
End Sub
Private Sub lblAppointment_Click()
' This sub opens the file menu bar when the file label is clicked.
' Position the File Menu list box
    lbxAppointmentMenu.Left = lblAppointment.Left
    lbxAppointmentMenu.Top = lblAppointment.Top + lblAppointment.Height
    lbxAppointmentMenu.Visible = True
End Sub

'#################################################################################################################'
' Subs for Selecting items in the menu you bar.
'#################################################################################################################'

Private Sub lbxFileMenu_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim intOption As Integer
' Actions for File menu Options
    intOption = lbxFileMenu.ListIndex
Select Case intOption
    Case 0
        frmCreateAccount.Show
    Case 1
        frmLogin.Show
    Case 2
End Select

End Sub

