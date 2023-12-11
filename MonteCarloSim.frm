VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MonteCarloSim 
   Caption         =   "Monte Carlo Simulation"
   ClientHeight    =   3700
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5080
   OleObjectBlob   =   "MonteCarloSim.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MonteCarloSim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

MonteCarloSim.Hide
SelectData.Show

End Sub

Private Sub CommandButton2_Click()

MonteCarloSim.Hide
InputData.Show

End Sub

Private Sub CommandButton3_Click()

MonteCarloSim.Hide

End Sub
