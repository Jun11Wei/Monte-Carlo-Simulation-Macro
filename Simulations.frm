VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Simulations 
   Caption         =   "Input Simulations"
   ClientHeight    =   2380
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "Simulations.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Simulations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton2_Click()

Simulations.Hide

End Sub

Private Sub CommandButton3_Click()

Simulations.Hide


    Dim Simulation As Long
    Dim results() As Double
    ReDim results(1 To Sims.Value)
    
    For Simulation = 1 To Sims.Value
        results(Simulation) = Application.WorksheetFunction.Norm_Inv(Rnd(), InputData.Mean.Value, InputData.Std.Value)
    Next Simulation
    
Dim avgValue As Double
Dim stdDev As Double
Dim minValue As Double
Dim maxValue As Double
    

    avgValue = Application.WorksheetFunction.Average(results)
    stdDev = Application.WorksheetFunction.StDev(results)
    minValue = Application.WorksheetFunction.Min(results)
    maxValue = Application.WorksheetFunction.Max(results)


Range(Selection.Address) = "Descriptive Statistics"
Range(Selection.Address).Offset(1, 0) = "Mean"
Range(Selection.Address).Offset(2, 0) = "Standard Deviation"
Range(Selection.Address).Offset(3, 0) = "Min"
Range(Selection.Address).Offset(4, 0) = "Max"

Range(Selection.Address).Offset(1, 1) = avgValue
Range(Selection.Address).Offset(2, 1) = stdDev
Range(Selection.Address).Offset(3, 1) = minValue
Range(Selection.Address).Offset(4, 1) = maxValue

Dim confidenceIntervals(1 To 3) As Double
    confidenceIntervals(1) = Application.WorksheetFunction.Percentile_Inc(results, 0.05) ' 90% CI
    confidenceIntervals(2) = Application.WorksheetFunction.Percentile_Inc(results, 0.025) ' 95% CI
    confidenceIntervals(3) = Application.WorksheetFunction.Percentile_Inc(results, 0.005) ' 99% CI

Range(Selection.Address).Offset(0, 3) = "Confidence Intervals"
Range(Selection.Address).Offset(1, 3) = "90% CI"
Range(Selection.Address).Offset(1, 4) = confidenceIntervals(1)
Range(Selection.Address).Offset(2, 3) = "95% CI"
Range(Selection.Address).Offset(2, 4) = confidenceIntervals(2)
Range(Selection.Address).Offset(3, 3) = "99% CI"
Range(Selection.Address).Offset(3, 4) = confidenceIntervals(3)



Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "Data"
    
    ' Output results to a column in the new worksheet
    ws.Range("A1").Value = "Simulation Results"
    For Simulation = 1 To Sims.Value
        ws.Cells(Simulation + 1, 1).Value = results(Simulation)
    Next Simulation


End Sub
