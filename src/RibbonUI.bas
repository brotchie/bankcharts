Attribute VB_Name = "RibbonUI"
Option Explicit

Private ribbon As IRibbonUI

Sub OnLoad(ui As IRibbonUI)
    Debug.Print "Loaded"
    Set ribbon = ui
End Sub

Sub InsertFootballChart(control As IRibbonControl)
    Dim ui As New FootballChartUI
    ui.Show False
End Sub

