Attribute VB_Name = "PrivateMod"
Option Explicit

Sub ChartLables(objChart As Excel.Chart, title, xlabel, ylabel As String)
  
  objChart.ChartTitle.Text = title
  objChart.Axes(xlValue, xlPrimary).AxisTitle.Text = ylabel
  objChart.Axes(xlCategory).AxisTitle.Text = xlabel
  
  
End Sub


Sub ChartScaleUswing(objChart As Excel.Chart)
    With objChart.Axes(xlValue)
      .MaximumScale = 100 '%
      .MinimumScale = 0
      .MajorUnit = 10
      .MinorUnit = 1
    End With
    
    
    objChart.Axes(xlCategory).MinimumScale = 200
    objChart.Axes(xlCategory).MaximumScale = 500 ' grader C
    
    
End Sub










Sub PrintChartDataRef()
  Dim ws As Worksheet
  Dim ct As Chart
  Dim srs As Series
  Dim i, j As Integer
  
  
  i = 1
  Set ws = ActiveWorkbook.Worksheets.Add
  
    
  For Each ct In ActiveWorkbook.Charts
    If i = 1 Then
      ws.Name = Left(ct.Name, 2) & "_(" & Mid(ws.Name, 6, 3) & ")"
    Else
      ws.Name = Left(ct.Name, 2) & "_" & ws.Name
    End If
    
    For Each srs In ct.FullSeriesCollection
      'ws.Cells(i, 1).Value = "'" & ct.FullSeriesCollection(1).Name
      'ws.Cells(i, 2).Value = "'" & ct.FullSeriesCollection(1).Formula
      ws.Cells(i, 1).Value = "'" & ct.Name
      ws.Cells(i, 2).Value = "'" & srs.Name
      ws.Cells(i, 3).Value = "'" & srs.Formula
      i = i + 1
      
     Next srs
     
  Next ct
  ws.Name = "ChartRef_" & ws.Name
  
  
    
    
End Sub


