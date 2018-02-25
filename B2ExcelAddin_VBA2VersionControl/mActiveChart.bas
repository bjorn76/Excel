Attribute VB_Name = "mActiveChart"
Option Explicit

'vbAbort =  3
'vbCancel = 2
'vbIgnore = 5
'vbNo =     7
'vbOK =     1
'vbRetry =  4
'vbYes =    6

'vbOKOnly             0   Display OK button only.
'vbOKCancel           1   Display OK and Cancel buttons.
'vbAbortRetryIgnore   2   Display Abort, Retry, and Ignore buttons.
'vbYesNoCancel        3   Display Yes, No, and Cancel buttons.
'vbYesNo              4   Display Yes and No buttons.
'vbRetryCancel        5   Display Retry and Cancel buttons.
'vbCritical          16  Display Critical Message icon.
'vbQuestion          32  Display Warning Query icon.
'vbExclamation       48  Display Warning Message icon.
'vbInformation       64  Display Information Message icon.


Sub SelectChartFont()
    Dim chrt As Chart
    Dim sFS As String
    Dim iFontSize As Integer
    
    
       
    Set chrt = Application.ActiveWorkbook.ActiveChart

   'User confirmation
    If MsgBox("Modify Chart Font Size?" & Chr(13) & Chr(10) & chrt.Name, vbOKCancel, "?") <> 1 Then
      Exit Sub
    End If
    
    sFS = InputBox("Enter size", "Font Size", 16)
    
    iFontSize = CInt(sFS)
    
    SetChartFont (iFontSize)
    
    

End Sub



Private Sub SetChartFont(FontSize As Integer)
    Dim Lent As LegendEntry
    Dim chrt As Chart
    
    Set chrt = Application.ActiveChart
    
    
    
'    chrt.ChartTitle.Format.TextFrame2.TextRange.Font.Size = FontSize
'
'    For Each Lent In chrt.Legend.LegendEntries
'      Lent.Format.TextFrame2.TextRange.Font.Size = FontSize
'    Next Lent
'
'    chrt.Axes(xlCategory).AxisTitle.Format.TextFrame2.TextRange.Font.Size = FontSize
'    chrt.Axes(xlValue).AxisTitle.Format.TextFrame2.TextRange.Font.Size = FontSize
    
    'chrt.ChartArea.Font.Size = FontSize
    
    With chrt.ChartArea.Font
      .Size = FontSize
      .Bold = True
    
    End With
    
    'To do: size of axis values
    'chrt.Axes(xlCategory).Text.Font = FontSize
    'ActiveChart.Axes(xlCategory).Value.Format.TextFrame2.TextRange.Font.Size = FontSize
'    With chrt.Axes(xlCategory).Format.TextFrame2.TextRange.Font
'      .BaselineOffset = 0
'      .Size = FontSize
'    End With
    
  
End Sub

Public Sub LineW()
    Dim chrt As Chart
    Dim i As Integer
    Dim srs As Series
       
    Set chrt = Application.ActiveWorkbook.ActiveChart
    
   'User confirmation
    If MsgBox("Modify Line weight?" & Chr(13) & Chr(10) & chrt.Name, vbOKCancel, "?") <> 1 Then
      Exit Sub
    End If
    
    
    chrt.Activate 'must activate to access series properties
    'For Each srs In ActiveChart.FullSeriesCollection
    For Each srs In Application.ActiveChart.SeriesCollection
        i = i + 1

        'Common look
        srs.Format.Line.Weight = 0.5
        srs.Format.Line.DashStyle = xlSolid
        'xlNone, xlSolid, xlDash, xlDot, xlDashDot, xlDashDotDot,
        'srs.MarkerSize = 10
        srs.Format.Fill.Visible = msoFalse
        srs.HasDataLabels = False  'new 8 jan
      Next
        
End Sub

Public Sub SetLine()
    Dim chrt As Chart
    Dim i As Integer
    Dim srs As Series
       
    Set chrt = Application.ActiveWorkbook.ActiveChart
    
   'User confirmation
    If MsgBox("Modify Line?" & Chr(13) & Chr(10) & chrt.Name, vbOKCancel, "?") <> 1 Then
      Exit Sub
    End If
    
    ChartEdit.ChartSeriesLine chrt
    
    
    
        
End Sub
Public Sub SetLineColor()
    Dim chrt As Chart
    Dim i As Integer
    Dim srs As Series
       
    Set chrt = Application.ActiveWorkbook.ActiveChart
    
   'User confirmation
    If MsgBox("Modify Line Color?" & Chr(13) & Chr(10) & chrt.Name, vbOKCancel, "?") <> 1 Then
      Exit Sub
    End If
    
    ChartEdit.ChartSeriesColor chrt
    
    
    
        
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
  Application.ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "NOx (g/kWh)"
  Application.ActiveChart.Axes(xlCategory).AxisTitle.Text = "ANR"
  
  Application.ActiveChart.ChartTitle.Text = "ETC" & Chr(13) & "Aged DW3200 systems"
  
  MsgBox Application.ActiveChart.Name
  

  
  
  'ActiveChart.Type = xlXYScatterLines
  'xlXYScatterLines

End Sub


Sub ChartScale()
    With Application.ActiveChart.Axes(xlValue)
      .MaximumScale = 3
      .MinimumScale = 0
      .MajorUnit = 0.2
      .MinorUnit = 0.1
    End With
    
    
    Application.ActiveChart.Axes(xlCategory).MaximumScale = 1.2
    Application.ActiveChart.Axes(xlCategory).MinimumScale = 0.8
    
End Sub


