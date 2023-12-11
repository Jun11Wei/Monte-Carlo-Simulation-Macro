VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Simulations2 
   Caption         =   "Input Simulations"
   ClientHeight    =   2380
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "Simulations2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Simulations2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton2_Click()

Simulations2.Hide

End Sub

Private Sub CommandButton3_Click()

Simulations2.Hide
    
    Dim MeanValue As Double
    Dim StandardDev As Double
    Dim Simulation2 As Long
    Dim results2() As Double
    ReDim results2(1 To Sims.Value)
        
        MeanValue = Application.WorksheetFunction.Average(Range(SelectData.Ref.Value))
        StandardDev = Application.WorksheetFunction.StDev(Range(SelectData.Ref.Value))

For Simulation2 = 1 To Sims.Value
            results2(Simulation2) = Application.WorksheetFunction.Norm_Inv(Rnd(), MeanValue, StandardDev)
        Next Simulation2
        
    Dim avgValue2 As Double
    Dim stdDev2 As Double
    Dim minValue2 As Double
    Dim maxValue2 As Double
        
    
        avgValue2 = Application.WorksheetFunction.Average(results2)
        stdDev2 = Application.WorksheetFunction.StDev(results2)
        minValue2 = Application.WorksheetFunction.Min(results2)
        maxValue2 = Application.WorksheetFunction.Max(results2)
    
    
    Range(Selection.Address) = "Descriptive Statistics"
    Range(Selection.Address).Offset(1, 0) = "Mean"
    Range(Selection.Address).Offset(2, 0) = "Standard Deviation"
    Range(Selection.Address).Offset(3, 0) = "Min"
    Range(Selection.Address).Offset(4, 0) = "Max"
    
    Range(Selection.Address).Offset(1, 1) = avgValue2
    Range(Selection.Address).Offset(2, 1) = stdDev2
    Range(Selection.Address).Offset(3, 1) = minValue2
    Range(Selection.Address).Offset(4, 1) = maxValue2
    
    Dim confidenceIntervals(1 To 3) As Double
    confidenceIntervals(1) = Application.WorksheetFunction.Percentile_Inc(results2, 0.05) ' 90% CI
    confidenceIntervals(2) = Application.WorksheetFunction.Percentile_Inc(results2, 0.025) ' 95% CI
    confidenceIntervals(3) = Application.WorksheetFunction.Percentile_Inc(results2, 0.005) ' 99% CI

    Range(Selection.Address).Offset(0, 3) = "Confidence Intervals"
    Range(Selection.Address).Offset(1, 3) = "90% CI"
    Range(Selection.Address).Offset(1, 4) = confidenceIntervals(1)
    Range(Selection.Address).Offset(2, 3) = "95% CI"
    Range(Selection.Address).Offset(2, 4) = confidenceIntervals(2)
    Range(Selection.Address).Offset(3, 3) = "99% CI"
    Range(Selection.Address).Offset(3, 4) = confidenceIntervals(3)
    
    
    Dim ws2 As Worksheet
        Set ws2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws2.Name = "Data"
        
        ws2.Range("A1").Value = "Simulation Results"
        For Simulation2 = 1 To Sims.Value
            ws2.Cells(Simulation2 + 1, 1).Value = results2(Simulation2)
        Next Simulation2

End Sub
