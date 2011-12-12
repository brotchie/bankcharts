VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FootballChartGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum Col
    Value = 1
    Running = 2
    Fixed = 3
    Bottom = 4
    Top = 5
    UpOffset = 6
    DownOffset = 7
    UpFill = 8
    DownFill = 9
    PositiveUpFill = 10
    PositiveDownFill = 11
    NegativeUpFill = 12
    NegativeDownFill = 13
    lastcolumn = 13
End Enum


Public Sub GenerateChart(source As Range, labels As Range, destination As Range)
    Debug.Assert source.Count = labels.Count
    
    Dim ii As Integer
    Dim formula As String
    
    Application.ScreenUpdating = False
    ' Link Labels.
    LinkCells labels, destination.Offset(1, 0).Resize(labels.Count, 1)
    
    ' Link Data.
    destination.Offset(0, Value).Value = "Value"
    LinkCells source, destination.Offset(1, Value).Resize(labels.Count, 1)
    
    ' Running.
    destination.Offset(0, Running).Value = "Running"
    For ii = 1 To labels.Count
        If ii = 1 Then
            formula = "RC[-1]"
        ElseIf ii = labels.Count Then
            formula = "0"
        Else
            formula = "RC[-1] + R[-1]C"
        End If
        destination.Offset(ii, Running).FormulaR1C1Local = "=" & formula
    Next ii
    
    ' Fixed.
    destination.Offset(0, Fixed).Value = "Fixed"
    For ii = 1 To labels.Count
        If ii = 1 Or ii = labels.Count Then
            formula = "RC[-2]"
        Else
            formula = "#N/A"
        End If
        destination.Offset(ii, Fixed).FormulaR1C1Local = "=" & formula
    Next ii
    
    ' Bottom
    destination.Offset(0, Bottom).Value = "Bottom"
    For ii = 1 To labels.Count
        If ii = 1 Then
            formula = "0"
        Else
            formula = "min(RC[-2], R[-1]C[-2])"
        End If
        destination.Offset(ii, Bottom).FormulaR1C1Local = "=" & formula
    Next ii
    
    ' Top
    destination.Offset(0, Top).Value = "Top"
    destination.Offset(1, Top).Resize(1, 1).FormulaR1C1Local = "=max(RC[-3], 0)"
    destination.Offset(2, Top).Resize(labels.Count - 1, 1).FormulaR1C1Local = "=max(RC[-3], R[-1]C[-3])"
    
    ' Up offset.
    destination.Offset(0, UpOffset).Value = "Up Offset"
    destination.Offset(1, UpOffset).Resize(labels.Count, 1).FormulaR1C1Local = "=max(RC[-2], 0)"
    
    ' Down offset.
    destination.Offset(0, DownOffset).Value = "Down Offset"
    destination.Offset(1, DownOffset).Resize(labels.Count, 1).FormulaR1C1Local = "=min(RC[-2], 0)"
    
    ' Up fill.
    destination.Offset(0, UpFill).Value = "Up Fill"
    destination.Offset(1, UpFill).Resize(1, 1).FormulaR1C1Local = "=#N/A"
    destination.Offset(2, UpFill).Resize(labels.Count - 2, 1).FormulaR1C1Local = "=max(RC[-3] - RC[-2], 0)"
    destination.Offset(labels.Count, UpFill).Resize(1, 1).FormulaR1C1Local = "=#N/A"
    
    ' Down fill.
    destination.Offset(0, DownFill).Value = "Down Fill"
    destination.Offset(1, DownFill).Resize(1, 1).FormulaR1C1Local = "=#N/A"
    destination.Offset(2, DownFill).Resize(labels.Count - 2, 1).FormulaR1C1Local = "=min(RC[-5] - RC[-2], 0)"
    destination.Offset(labels.Count, DownFill).Resize(1, 1).FormulaR1C1Local = "=#N/A"
    
    ' Positive up fill.
    destination.Offset(0, PositiveUpFill).Value = "Positive Up Fill"
    destination.Offset(1, PositiveUpFill).Resize(1, 1).FormulaR1C1Local = "=#N/A"
    destination.Offset(2, PositiveUpFill).Resize(labels.Count - 2, 1).FormulaR1C1Local = "=IF(RC[-9]>=0, RC[-2], #N/A)"
    destination.Offset(labels.Count, PositiveUpFill).Resize(1, 1).FormulaR1C1Local = "=#N/A"
    
    ' Positive down fill.
    destination.Offset(0, PositiveDownFill).Value = "Positive Down Fill"
    destination.Offset(1, PositiveDownFill).Resize(1, 1).FormulaR1C1Local = "=#N/A"
    destination.Offset(2, PositiveDownFill).Resize(labels.Count - 2, 1).FormulaR1C1Local = "=IF(RC[-10]>=0, RC[-2], #N/A)"
    destination.Offset(labels.Count, PositiveDownFill).Resize(1, 1).FormulaR1C1Local = "=#N/A"
    
    ' Negative up fill.
    destination.Offset(0, NegativeUpFill).Value = "Negative Up Fill"
    destination.Offset(1, NegativeUpFill).Resize(1, 1).FormulaR1C1Local = "=#N/A"
    destination.Offset(2, NegativeUpFill).Resize(labels.Count - 2, 1).FormulaR1C1Local = "=IF(RC[-11]<0, RC[-4], #N/A)"
    destination.Offset(labels.Count, NegativeUpFill).Resize(1, 1).FormulaR1C1Local = "=#N/A"
    
    ' Negative down fill.
    destination.Offset(0, NegativeDownFill).Value = "Negative Down Fill"
    destination.Offset(1, NegativeDownFill).Resize(1, 1).FormulaR1C1Local = "=#N/A"
    destination.Offset(2, NegativeDownFill).Resize(labels.Count - 2, 1).FormulaR1C1Local = "=IF(RC[-12]<0, RC[-4], #N/A)"
    destination.Offset(labels.Count, NegativeDownFill).Resize(1, 1).FormulaR1C1Local = "=#N/A"
    
    ' Format table.
    destination.Resize(1, lastcolumn + 1).Font.Bold = True
    destination.Resize(1, lastcolumn + 1).Borders(xlEdgeBottom).Weight = 2
    destination.Offset(1, 0).Resize(labels.Count, lastcolumn + 1).NumberFormat = "#,##0;[Red](#,##0)"
    destination.Resize(1, labels.Count + 1).Font.Bold = True
    
    CreateChart labels, destination
    
    Application.ScreenUpdating = True
End Sub

Private Sub CreateChart(labels As Range, destination As Range)
    Dim chartobj As chartObject
    Set chartobj = destination.Parent.ChartObjects().Add(destination.Left, destination.Offset(labels.Count + 2, 0).Top, 500, 250)
    
    Dim chart As chart
    Set chart = chartobj.chart

    chart.ChartArea.Format.Line.Visible = msoFalse
    
    
    ' Stacked column chart.
    chart.ChartType = xlColumnStacked
    
    Dim sc As SeriesCollection
    Set sc = chart.SeriesCollection()
    
    
    ' Offsets.
    With sc.NewSeries
        .Values = destination.Offset(1, UpOffset).Resize(labels.Count, 1)
        .XValues = labels
        .Format.Fill.Visible = msoFalse
    End With
    With sc.NewSeries
        .Values = destination.Offset(1, DownOffset).Resize(labels.Count, 1)
        .XValues = labels
        .Format.Fill.Visible = msoFalse
    End With
    
    ' Positive.
    With sc.NewSeries
        .Values = destination.Offset(1, PositiveUpFill).Resize(labels.Count, 1)
        .XValues = labels
        .Format.Fill.ForeColor.RGB = RGB(0, 255, 0)
    End With
    With sc.NewSeries
        .Values = destination.Offset(1, PositiveDownFill).Resize(labels.Count, 1)
        .XValues = labels
        .Format.Fill.ForeColor.RGB = RGB(0, 255, 0)
    End With
    
    ' Negative.
    With sc.NewSeries
        .Values = destination.Offset(1, NegativeUpFill).Resize(labels.Count, 1)
        .XValues = labels
        .Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
    End With
    With sc.NewSeries
        .Values = destination.Offset(1, NegativeDownFill).Resize(labels.Count, 1)
        .XValues = labels
        .Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
    End With
    
    ' Fixed.
    With sc.NewSeries
        .Values = destination.Offset(1, Fixed).Resize(labels.Count, 1)
        .XValues = labels
    End With
    
    With chart.Axes(xlValue)
        .HasMajorGridlines = False
        .TickLabels.NumberFormat = "#,##0;[Red](#,##0)"
    End With
    
    ' Hide legend.
    chart.Legend.Clear
End Sub

Private Sub LinkCells(source As Range, destination As Range)
    Debug.Assert source.Rows.Count = destination.Rows.Count
    Debug.Assert source.Columns.Count = destination.Columns.Count
    
    Dim ii As Integer
    For ii = 1 To source.Cells.Count
        destination.Cells(ii).formula = "=" & source.Cells(ii).Address
    Next ii
End Sub

