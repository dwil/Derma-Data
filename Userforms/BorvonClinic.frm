VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DrBorvonClinic 
   Caption         =   "UserForm1"
   ClientHeight    =   13695
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   13635
   OleObjectBlob   =   "BorvonClinic.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DrBorvonClinic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ExaminationCheckBox_Click()

End Sub

Private Sub ComboBox4_Change()

End Sub

Private Sub CommandButtonCancel_Click()

End Sub

Private Sub CommandButtonCancel1_Click()
Unload Me
End Sub

Private Sub CommandButtonCancel2_Click()
Unload Me
End Sub

Private Sub CommandButtonClearForm1_Click()
Dim cCont As Control

    For Each cCont In Me.MultiPage1.Pages(0).Controls
        If TypeName(cCont) = "MultiPage" Or TypeName(cCont) = "Label" Or TypeName(cCont) = "Frame" Or TypeName(cCont) = "CommandButton" Then
        Else
            cCont.Value = ""
        End If
    Next cCont
    
    Call UserForm_Initialize
End Sub

Private Sub CommandButtonClearForm2_Click()
Call UserForm_Initialize
End Sub

Private Sub CommandButtonSave1_Click()



End Sub

Private Sub DateComboBoxD_Change()

End Sub

Private Sub ExaminationCheckBox1_Click()

End Sub

Private Sub Label28_Click()

End Sub

Private Sub Label29_Click()

End Sub

Private Sub IDTypeComboBox_Change()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label43_Click()

End Sub

Private Sub Label48_Click()

End Sub

Private Sub Label60_Click()

End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub NOOptionButton_Click()

End Sub

Private Sub TextBox15_Change()

End Sub

Private Sub TitleComboBoxA_Change()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
 With TitleComboBoxA
    .AddItem "Mr."
    .AddItem "Mrs."
    .AddItem "Ms"
    .AddItem "Jr.Mr."
    .AddItem "Jr.Miss"
    
End With

With ComboBoxDateA
    .AddItem "1"
    .AddItem "2"
    .AddItem "3"
    .AddItem "4"
    .AddItem "5"
    .AddItem "6"
    .AddItem "7"
    .AddItem "8"
    .AddItem "9"
    .AddItem "10"
    .AddItem "11"
    .AddItem "12"
    .AddItem "13"
    .AddItem "14"
    .AddItem "15"
    .AddItem "16"
    .AddItem "17"
    .AddItem "18"
    .AddItem "19"
    .AddItem "20"
    .AddItem "21"
    .AddItem "22"
    .AddItem "23"
    .AddItem "24"
    .AddItem "25"
    .AddItem "26"
    .AddItem "27"
    .AddItem "28"
    .AddItem "29"
    .AddItem "30"
    .AddItem "31"
    
End With

With ComboBoxMonthA
    .AddItem "January"
    .AddItem "February"
    .AddItem "March"
    .AddItem "April"
    .AddItem "May"
    .AddItem "June"
    .AddItem "July"
    .AddItem "August"
    .AddItem "September"
    .AddItem "October"
    .AddItem "November"
    .AddItem "December"
    
End With
    
    
 With ComboBoxDateB
    .AddItem "1"
    .AddItem "2"
    .AddItem "3"
    .AddItem "4"
    .AddItem "5"
    .AddItem "6"
    .AddItem "7"
    .AddItem "8"
    .AddItem "9"
    .AddItem "10"
    .AddItem "11"
    .AddItem "12"
    .AddItem "13"
    .AddItem "14"
    .AddItem "15"
    .AddItem "16"
    .AddItem "17"
    .AddItem "18"
    .AddItem "19"
    .AddItem "20"
    .AddItem "21"
    .AddItem "22"
    .AddItem "23"
    .AddItem "24"
    .AddItem "25"
    .AddItem "26"
    .AddItem "27"
    .AddItem "28"
    .AddItem "29"
    .AddItem "30"
    .AddItem "31"
    
End With

With ComboBoxMonthB
    .AddItem "January"
    .AddItem "February"
    .AddItem "March"
    .AddItem "April"
    .AddItem "May"
    .AddItem "June"
    .AddItem "July"
    .AddItem "August"
    .AddItem "September"
    .AddItem "October"
    .AddItem "November"
    .AddItem "December"
    
End With

With ComboBoxDateC
    .AddItem "1"
    .AddItem "2"
    .AddItem "3"
    .AddItem "4"
    .AddItem "5"
    .AddItem "6"
    .AddItem "7"
    .AddItem "8"
    .AddItem "9"
    .AddItem "10"
    .AddItem "11"
    .AddItem "12"
    .AddItem "13"
    .AddItem "14"
    .AddItem "15"
    .AddItem "16"
    .AddItem "17"
    .AddItem "18"
    .AddItem "19"
    .AddItem "20"
    .AddItem "21"
    .AddItem "22"
    .AddItem "23"
    .AddItem "24"
    .AddItem "25"
    .AddItem "26"
    .AddItem "27"
    .AddItem "28"
    .AddItem "29"
    .AddItem "30"
    .AddItem "31"
    
End With

With ComboBoxMonthC
    .AddItem "January"
    .AddItem "February"
    .AddItem "March"
    .AddItem "April"
    .AddItem "May"
    .AddItem "June"
    .AddItem "July"
    .AddItem "August"
    .AddItem "September"
    .AddItem "October"
    .AddItem "November"
    .AddItem "December"
    
End With
    
'Existing Pation Option
YESOptionButton2.Value = True

'Charges checkbox option
ExaminationCheckBox1.Value = False


With DateComboBoxD
    .AddItem "1"
    .AddItem "2"
    .AddItem "3"
    .AddItem "4"
    .AddItem "5"
    .AddItem "6"
    .AddItem "7"
    .AddItem "8"
    .AddItem "9"
    .AddItem "10"
    .AddItem "11"
    .AddItem "12"
    .AddItem "13"
    .AddItem "14"
    .AddItem "15"
    .AddItem "16"
    .AddItem "17"
    .AddItem "18"
    .AddItem "19"
    .AddItem "20"
    .AddItem "21"
    .AddItem "22"
    .AddItem "23"
    .AddItem "24"
    .AddItem "25"
    .AddItem "26"
    .AddItem "27"
    .AddItem "28"
    .AddItem "29"
    .AddItem "30"
    .AddItem "31"
    
End With

With MonthComboBoxD
    .AddItem "January"
    .AddItem "February"
    .AddItem "March"
    .AddItem "April"
    .AddItem "May"
    .AddItem "June"
    .AddItem "July"
    .AddItem "August"
    .AddItem "September"
    .AddItem "October"
    .AddItem "November"
    .AddItem "December"
    
End With

 With IDTypeComboBox
    .AddItem "Passport"
    .AddItem "Thai ID"
    .AddItem "Driving License"
    .AddItem "Student ID"
   
    
End With

End Sub

Private Sub YESOptionButton_Click()

End Sub



