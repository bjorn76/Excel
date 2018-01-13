Attribute VB_Name = "ActiveChart"
Option Explicit

Private Sub SetChartFont(FontSize As Integer)
    Dim Lent As LegendEntry
    
    ActiveChart.ChartTitle.Format.TextFrame2.TextRange.Font.Size = FontSize
    
    For Each Lent In ActiveChart.Legend.LegendEntries
      Lent.Format.TextFrame2.TextRange.Font.Size = FontSize
    Next Lent
    
    ActiveChart.Axes(xlCategory).AxisTitle.Format.TextFrame2.TextRange.Font.Size = FontSize
    ActiveChart.Axes(xlValue).AxisTitle.Format.TextFrame2.TextRange.Font.Size = FontSize
    
    'To do: size of axis values
    'ActiveChart.Axes(xlCategory).Text.Font = FontSize
    'ActiveChart.Axes(xlCategory).Value.Format.TextFrame2.TextRange.Font.Size = FontSize
    
  
End Sub

Sub Font10()
   SetChartFont (10)
End Sub

Sub Font16()
   SetChartFont (16)
End Sub

Sub Font24()
   SetChartFont (24)
End Sub

Sub ChartLables()
  ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "NOx (g/kWh)"
  ActiveChart.Axes(xlCategory).AxisTitle.Text = "ANR"
  
  ActiveChart.ChartTitle.Text = "ETC" & Chr(13) & "Aged DW3200 systems"
  
  MsgBox ActiveChart.Name
  
  
  'ActiveChart.Type = xlXYScatterLines
  'xlXYScatterLines

End Sub


Sub ChartScale()
    With ActiveChart.Axes(xlValue)
      .MaximumScale = 3
      .MinimumScale = 0
      .MajorUnit = 0.2
      .MinorUnit = 0.1
    End With
    
    
    ActiveChart.Axes(xlCategory).MaximumScale = 1.2
    ActiveChart.Axes(xlCategory).MinimumScale = 0.8
    
End Sub

