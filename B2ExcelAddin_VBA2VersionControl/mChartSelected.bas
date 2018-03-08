Attribute VB_Name = "mChartSelected"
Option Explicit
'Functions in this module shall include:
' # User confirmation
' # Act on selected och active chart(s)



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




Sub ChartExpPNG()
  Dim obj As Object
    

  If Not ActiveChart Is Nothing Then
  
    'User confirmation
    If MsgBox("Export Chart?" & Chr(13) & Chr(10) & ActiveChart.Name, vbOKCancel, "?") <> 1 Then
      Exit Sub
    End If
    
    ChartExportPNG ActiveChart
  Else
    'User confirmation
    If MsgBox("Export Charts?" & Chr(13) & Chr(10) & Selection.Count, vbOKCancel, "?") <> 1 Then
      Exit Sub
    End If
    
    For Each obj In Selection
      If TypeName(obj) = "ChartObject" Then
        ChartExportPNG obj.Chart
      End If
    Next
  End If
End Sub
  
  




Sub SetCol()

    Dim chrt As Chart
    Dim i As Integer
    Dim srs As Series
       
    Set chrt = Application.ActiveWorkbook.ActiveChart
    
   'User confirmation
    If MsgBox("Modify Chart?" & Chr(13) & Chr(10) & chrt.Name, vbOKCancel, "?") <> 1 Then
      Exit Sub
    End If
    
    mChart.ChartSeriesClustCol chrt

End Sub


'Sub FormatCharts()
'  Dim obj As Object
'
'  If Not ActiveChart Is Nothing Then
'    FormatOneChart ActiveChart
'  Else
'    For Each obj In Selection
'      If TypeName(obj) = "ChartObject" Then
'        FormatOneChart obj.Chart
'      End If
'    Next
'  End If
'End Sub
'
'Sub FormatOneChart(cht As Chart)
'  ' do all your formatting here, based on cht not on ActiveChart
'End Sub




Sub SelectChartFont()
    Dim chrt As Chart
    Dim sFS As String
    Dim iFontSize As Integer
        
    Dim obj As Object
    
    If Not ActiveChart Is Nothing Then
      'FormatOneChart ActiveChart
      Set chrt = Application.ActiveWorkbook.ActiveChart
      
      'User confirmation
    If MsgBox("Modify Chart Font Size?" & Chr(13) & Chr(10) & chrt.Name, vbOKCancel, "?") <> 1 Then
      Exit Sub
    End If
    
    sFS = InputBox("Enter size", "Font Size", 16)
    
    iFontSize = CInt(sFS)
    
    SetChartFont (iFontSize)
      
      
      
      
      
    Else
      'For Each obj In Selection
           
      
      
      If MsgBox("Modify Chart Font Size?" & Chr(13) & Chr(10) & Selection.Count & " graphs are selected", vbOKCancel, "?") <> 1 Then
        Exit Sub
      End If
      
      
      
      For Each obj In Selection
        If TypeName(obj) = "ChartObject" Then
          'FormatOneChart obj.Chart
          MsgBox obj.Name
      
          
          
        End If
      Next
    End If
 
    
       
    


    
    

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
    Dim obj As Object
    
       
    Set chrt = Application.ActiveWorkbook.ActiveChart
    
   'User confirmation

    
    
  If Not ActiveChart Is Nothing Then
    If MsgBox("Modify Line weight?" & Chr(13) & Chr(10) & chrt.Name, vbOKCancel, "?") <> 1 Then
      Exit Sub
    End If
    mChart.LineW05 Application.ActiveWorkbook.ActiveChart
  
  Else
    For Each obj In Selection
      If TypeName(obj) = "ChartObject" Then
        'FormatOneChart obj.Chart
        'Do do.....
        
        
      End If
    Next
  End If
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
    
    mChart.ChartSeriesLine chrt
    
    
    
        
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
    
    mChart.ChartSeriesColor chrt
    
    
    
        
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


Sub SetY1Scale()
    'Y-Scale
    With Application.ActiveChart.Axes(xlValue)
      .MaximumScale = 350
      .MinimumScale = 0
      .MajorUnit = 50
      .MinorUnit = 10
    End With
    
    
'    Application.ActiveChart.Axes(xlCategory).MaximumScale = 1.2
'    Application.ActiveChart.Axes(xlCategory).MinimumScale = 0.8
    
End Sub


