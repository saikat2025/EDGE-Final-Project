VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   7245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14610
   OleObjectBlob   =   "UserForm2 13.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()

End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub SpinButton1_Change()

End Sub

Private Sub SpinButton1_SpinDown()
If TextBox4.Value > 0 Then
TextBox4 = TextBox4 - 1
End If
End Sub

Private Sub SpinButton1_SpinUp()
TextBox4 = TextBox4 + 1
End Sub

Private Sub UserForm_Click()

End Sub
