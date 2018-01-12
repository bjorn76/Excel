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



Sub ExportPNG()
  Dim sName As String
  Dim sPath As String
  Dim sFullName As String
  Dim ct As Chart
  
  
  sPath = ActiveWorkbook.Path & "\ExportPNG\"
  
      
  If Len(Dir(sPath, vbDirectory)) = 0 Then
     MkDir sPath
  End If
  
  For Each ct In ActiveWorkbook.Charts
    sName = ct.Name
    sFullName = sPath & sName & ".png"
    ct.Export (sFullName)
  Next ct
  
End Sub




Sub ChartSeriesFixedLook(objChart As Excel.Chart)
    'SchemeColor:
    '1 = vit
    '2 = grön
    '3 = röd
    '4 = blå,
    '5 = gul,
    '6= violett,
    '7 = turkos
        

    Dim srs As Series
    Dim i As Integer
    Dim mycolors(32) As Long
          
       
    'Different look
    'Mörk till ljusblå
    mycolors(1) = RGB(2, 152, 202) 'Turkose
    mycolors(2) = RGB(57, 90, 131) 'Lilablå
    mycolors(5) = RGB(35, 47, 69) ' Mörk-mörk blå
    
    'Gröna toner
    mycolors(3) = RGB(169, 174, 110) 'Beigegrön
    mycolors(4) = RGB(136, 164, 90) ' SommarGrön
    mycolors(6) = RGB(43, 153, 93) 'blågrön
    
    'Orange till röd
    mycolors(7) = RGB(247, 151, 139) 'Apelsin
    mycolors(8) = RGB(222, 105, 61) 'Lotsbåt
    mycolors(9) = RGB(207, 56, 71) 'Blod
    
 
    
    
    i = 0
    
    objChart.Activate
     
    For Each srs In ActiveChart.FullSeriesCollection
        i = i + 1
        
        'Common look
        srs.Format.line.Weight = 1.25
        srs.Format.line.DashStyle = xlSolid
        'xlNone, xlSolid, xlDash, xlDot, xlDashDot, xlDashDotDot,
        srs.MarkerSize = 10
        srs.Format.Fill.Visible = msoFalse
        srs.HasDataLabels = False  'new 8 jan
       
  
  
       
  
    
     'Template:
     'ActiveChart.SeriesCollection(i).Format.Fill.ForeColor.SchemeColor = 2
     'ActiveChart.SeriesCollection(i).Format.Line.ForeColor.SchemeColor = 2
     'ActiveChart.SeriesCollection(i).Format.Line.ForeColor.RGB = RGB(127, 127, 127)
     'End of template
     
     'ActiveSheet.ChartObjects(1).Activate
     'ActiveChart.SeriesCollection(i).DataLabels.ShowValue = False '/new 2018-01-08
     
    
    Select Case i
       ' Can be filled:                  xlMarkerStyleCircle xlMarkerStyleDiamond xlMarkerStyleTriangle xlMarkerStyleSquare
       ' Markes not possible to fill:    xlMarkerStyleX xlMarkerStylePlus xlMarkerStyleStar
       
       
       Case 1
       With ActiveChart.SeriesCollection(i)
         'B2.Datalab
         .MarkerStyle = xlMarkerStyleCircle
         .Format.line.ForeColor.RGB = mycolors(i)
         .Format.Fill.ForeColor.RGB = mycolors(i)
         '.Format.Fill.Visible = msoFalse
       End With
       
       Case 2
       With ActiveChart.SeriesCollection(i)
         .MarkerStyle = xlMarkerStyleDiamond
         .Format.line.ForeColor.RGB = mycolors(i)
         .Format.Fill.ForeColor.RGB = mycolors(i)
       End With
       
 
  
       Case 3
       With ActiveChart.SeriesCollection(i)
         .MarkerStyle = xlMarkerStyleStar
         .Format.line.ForeColor.RGB = mycolors(i)
         .Format.Fill.ForeColor.RGB = mycolors(i)
       End With
       
       Case 4
       With ActiveChart.SeriesCollection(i)
         .MarkerStyle = xlMarkerStyleX
         .Format.line.ForeColor.RGB = mycolors(i)
         .Format.Fill.ForeColor.RGB = mycolors(i)
       End With
       
       Case 5
       With ActiveChart.SeriesCollection(i)
         .MarkerStyle = xlMarkerStyleTriangle
         .Format.line.ForeColor.RGB = mycolors(i)
         .Format.Fill.ForeColor.RGB = mycolors(i)

       End With
  
       
       
       Case 6
       ActiveChart.SeriesCollection(i).MarkerStyle = xlMarkerStyleCircle
       ActiveChart.SeriesCollection(i).MarkerSize = 6
       With ActiveChart.SeriesCollection(i).Format.line
         '.DashStyle = xlDashDot
         .ForeColor.RGB = mycolors(i)
       End With
       
       With ActiveChart.SeriesCollection(i).Format.Fill
         .Visible = msoTrue
         .ForeColor.RGB = mycolors(i)
         .ForeColor.ObjectThemeColor = msoThemeColorAccent1
         .ForeColor.TintAndShade = 0
         '.ForeColor.Brightness = 0
         .Solid
       End With

    End Select
    
  Next srs

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


