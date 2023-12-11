VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InputData 
   Caption         =   "Input Data"
   ClientHeight    =   2840
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5080
   OleObjectBlob   =   "InputData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InputData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton2_Click()

InputData.Hide

End Sub

Private Sub Submit_Click()


InputData.Hide
Simulations.Show

End Sub
