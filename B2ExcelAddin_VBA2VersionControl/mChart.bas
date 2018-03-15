Attribute VB_Name = "mChart"
Option Explicit
'All fuctions in this module takes an chart as the first argument
'No user confirmation



Sub ChartExportPNG(ByRef objChart As Excel.Chart)
'Export all charts as png pict in active workbook. Saved in a subfolder
  Dim sName As String
  Dim sPath As String
  Dim sFullName As String
   
 
  sPath = objChart.Application.ActiveWorkbook.Path & "\ExportPNG\"
      
  If Len(Dir(sPath, vbDirectory)) = 0 Then
     MkDir sPath
  End If
  
  
    sName = objChart.Name
    sFullName = sPath & sName & ".png"

    
    If (objChart.ChartArea.Height < 420) And (TypeName(objChart.Parent) = "ChartObject") Then
      'only charts embeded in a sheet can and needs to be resized prior to export
      ' it' resized to create adjust resolotion on exported graph picture
      'objChart.ChartArea.Height = 210 * 2
      'objChart.ChartArea.Width = 297 * 2
      objChart.ChartArea.Height = objChart.ChartArea.Height * 2
      objChart.ChartArea.Width = objChart.ChartArea.Width * 2
      
    End If


    'objChart.Select

    
    Application.ScreenUpdating = False
    If TypeName(objChart.Parent) = "ChartObject" Then ActiveWindow.Zoom = 400 'temporary maximaze (400%) zoom to increase resolution on exported picture
    objChart.Export (sFullName)
    Application.ScreenUpdating = True
    
    If TypeName(objChart.Parent) = "ChartObject" Then
      'reset zoom after export
      ActiveWindow.Zoom = 50
      ActiveWindow.ScrollRow = 1
      ActiveWindow.ScrollColumn = 1
    End If

End Sub

Sub ChartSeriesClustCol(ByRef objChart As Excel.Chart)
    
    Dim srs As Series
    Dim i As Integer
    Dim mycolors(32) As Long
  
    If objChart.ChartType <> xlColumnClustered Then Exit Sub
            
    mycolors(1) = RGB(219, 106, 41) 'Orange Kellys Color nbr 4
    mycolors(2) = RGB(147, 173, 60) 'Green Kellys Color nbr 18
    mycolors(3) = RGB(2, 152, 202) 'Turquoise
    mycolors(7) = RGB(255, 150, 150) 'Old pink
    mycolors(5) = RGB(172, 255, 128) 'MintGreen
    mycolors(6) = RGB(70, 240, 240) ' Cyan
    mycolors(4) = RGB(240, 30, 230) 'Magneta
    mycolors(8) = RGB(127, 128, 129) 'Grey Kellys nbr 8
    mycolors(9) = RGB(98, 166, 71) 'Green Kellys  nbr 9
    mycolors(10) = RGB(72, 56, 150) 'Purple Kellys nbr 13
    mycolors(11) = RGB(209, 45, 39) 'Red Kellys nbr 20
    mycolors(12) = RGB(235, 205, 62) ' Mustard Kellys nbr 2
    
    i = 0
        
    For Each srs In objChart.FullSeriesCollection
      i = i + 1
      With srs.Format
        .Line.ForeColor.RGB = mycolors(i)
        .Fill.ForeColor.RGB = mycolors(i)
        '.Fill.BackColor.RGB = mycolors(i)
        .Fill.BackColor.RGB = RGB(255, 255, 255)
        '.Fill.Solid
        .Fill.TwoColorGradient msoGradientHorizontal, 1
      End With
    Next srs
  
End Sub


Public Sub ChartFontSize(ByRef objChart As Excel.Chart, fs As Integer)
    With objChart.ChartArea.Font
      .Size = fs
      .Bold = True
    
    End With
End Sub


Public Sub LineW05(ByRef objChart As Excel.Chart)
  If Not ((objChart.ChartType = xlXYScatterLines) Or (objChart.ChartType = xlXYScatterLinesNoMarkers)) Then Exit Sub
  
    
    Dim i As Integer
    Dim srs As Series
       
   
    'objChart.Activate
    
    'For Each srs In ActiveChart.FullSeriesCollection
    For Each srs In objChart.FullSeriesCollection
    'For Each srs In objChart.SeriesCollection
        i = i + 1

        'Common look
        'srs.Format.Line.Weight = 0.5
        srs.Format.Line.Weight = 0.5
        srs.Format.Line.DashStyle = xlSolid
        'xlNone, xlSolid, xlDash, xlDot, xlDashDot, xlDashDotDot,
        'srs.MarkerSize = 10
        srs.Format.Fill.Visible = msoFalse
        srs.HasDataLabels = False  'new 8 jan
        
        'srs.MarkerStyle = xlMarkerStyleNone
        'srs.MarkerBackgroundColorIndex = xlColorIndexNone
        'srs.MarkerForegroundColorIndex = xlColorIndexNone
        
      Next
        'ActiveWindow.Zoom = 130
End Sub



Sub ChartSeriesMarker(ByRef objChart As Excel.Chart)
  If objChart.ChartType <> xlXYScatterLinesNoMarkers Then Exit Sub
 
    objChart.ChartType = xlXYScatterLines
    
End Sub

Sub ChartSeriesNoMarker(ByRef objChart As Excel.Chart)
  If objChart.ChartType <> xlXYScatterLines Then Exit Sub

    
    
    objChart.ChartType = xlXYScatterLinesNoMarkers
    
    
    
'    Dim srs As Series
'    Dim i As Integer
'
'    i = 0
'
'
'    'objChart.Activate
'    For Each srs In objChart.FullSeriesCollection
'        i = i + 1
'
'        srs.Format.Fill.Visible = msoFalse
'        'srs.HasDataLabels = False  'new 8 jan
'        srs.MarkerStyle = xlMarkerStyleNone
'        srs.MarkerBackgroundColorIndex = xlColorIndexNone
'        srs.MarkerForegroundColorIndex = xlColorIndexNone
'
'      Next
    

End Sub


Sub ChartSeriesLineAndMarker(ByRef objChart As Excel.Chart)

  If objChart.ChartType <> xlXYScatterLines Then Exit Sub

    Dim srs As Series
    Dim i As Integer
    
    i = 0

    
    For Each srs In objChart.SeriesCollection
      i = i + 1
      'Common look
      'srs.Format.Line.Weight = 1.25
      srs.Format.Line.Weight = 2
      
      srs.Format.Line.DashStyle = xlSolid
      'xlNone, xlSolid, xlDash, xlDot, xlDashDot, xlDashDotDot
      srs.MarkerSize = 10
      srs.Format.Fill.Visible = msoFalse
      srs.HasDataLabels = False  'new 8 jan
      
      Select Case i
      ' Can be filled:                  xlMarkerStyleCircle xlMarkerStyleDiamond xlMarkerStyleTriangle xlMarkerStyleSquare
      ' Markes not possible to fill:    xlMarkerStyleX xlMarkerStylePlus xlMarkerStyleStar
       
      Case 1, 7
         srs.MarkerStyle = xlMarkerStyleCircle
         srs.MarkerBackgroundColorIndex = xlColorIndexNone
         'srs.Format.Fill.Visible = msoFalse
      Case 2, 8
         srs.MarkerStyle = xlMarkerStyleDiamond
         srs.MarkerBackgroundColorIndex = xlColorIndexNone
      Case 3, 9
           srs.MarkerStyle = xlMarkerStyleStar
           srs.MarkerBackgroundColorIndex = xlColorIndexNone
      Case 4, 10
         srs.MarkerStyle = xlMarkerStyleX
         srs.MarkerBackgroundColorIndex = xlColorIndexNone
      Case 5, 11
         srs.MarkerStyle = xlMarkerStyleTriangle
         srs.MarkerBackgroundColorIndex = xlColorIndexNone
         
      Case 6, 12
         srs.MarkerStyle = xlMarkerStyleCircle
         srs.MarkerSize = 6
      End Select
      
    Next

End Sub


Sub ChartSeriesLineType(ByRef objChart As Excel.Chart)
  If Not ((objChart.ChartType = xlXYScatterLinesNoMarkers) Or (objChart.ChartType = xlXYScatter) Or (objChart.ChartType = xlXYScatterLines)) Then Exit Sub
        
    Dim srs As Series
    Dim i As Integer, j As Integer
    
    i = 0
    j = 0
        
    For Each srs In objChart.FullSeriesCollection
      i = i + 1
      srs.Format.Line.ForeColor.RGB = MyColor(i)
      srs.Format.Line.Weight = 1
      j = i Mod 3
                
      Select Case j
      'xlSolid, xlDash, xlDot, xlDashDot, xlDashDotDot
      Case 1
        'srs.Format.Line
        srs.Format.Line.DashStyle = xlSolid
      Case 2
        srs.Format.Line.DashStyle = xlDashDot
         
         
      Case 0
        srs.Format.Line.DashStyle = xlDashDotDot
         
             
      Case Else
        MsgBox "Logic Error in ChartSeriesLineType sub!"
         
      End Select
      
    Next
    
    
    

End Sub



Sub ChartSeriesLine(ByRef objChart As Excel.Chart)
  If objChart.ChartType <> xlXYScatterLines Then Exit Sub
    Dim srs As Series
    Dim i As Integer
    
    i = 0

    For Each srs In objChart.SeriesCollection
      i = i + 1

      'Common look
      srs.Format.Line.Weight = 1.25
      srs.Format.Line.DashStyle = xlSolid
      'xlNone, xlSolid, xlDash, xlDot, xlDashDot, xlDashDotDot
    Next

End Sub


Private Function MyColor(i As Integer) As Long

    Dim mycolors(32) As Long
          
    mycolors(1) = RGB(219, 106, 41) 'Orange Kellys Color nbr 4
    mycolors(2) = RGB(147, 173, 60) 'Green Kellys Color nbr 18
    mycolors(3) = RGB(2, 152, 202) 'Turquoise
    mycolors(7) = RGB(255, 150, 150) 'Old pink
    mycolors(5) = RGB(172, 255, 128) 'MintGreen
    mycolors(6) = RGB(70, 240, 240) ' Cyan
    mycolors(4) = RGB(240, 30, 230) 'Magneta
    mycolors(8) = RGB(127, 128, 129) 'Grey Kellys nbr 8
    mycolors(9) = RGB(98, 166, 71) 'Green Kellys  nbr 9
    mycolors(10) = RGB(72, 56, 150) 'Purple Kellys nbr 13
    mycolors(11) = RGB(209, 45, 39) 'Red Kellys nbr 20
    mycolors(12) = RGB(235, 205, 62) ' Mustard Kellys nbr 2

MyColor = mycolors(i)


End Function





Sub ChartSeriesColorLine(ByRef objChart As Excel.Chart)
  If objChart.ChartType <> xlXYScatterLines Then Exit Sub
        
    Dim srs As Series
    Dim i As Integer
    
    i = 0
    'objChart.Activate 'Only needed when FullSerisCollectoion is used. Not seriescollection?
    
    For Each srs In objChart.FullSeriesCollection
      i = i + 1
       
       With srs.Format
         .Line.ForeColor.RGB = MyColor(i)
         '.Line.BackColor.RGB = MyColor(i) 'Filled marker color
         '.Fill.ForeColor.RGB = MyColor(i)
         '.Fill.BackColor.RGB = MyColor(i)
         
        End With
        'srs.MarkerForegroundColor = MyColor(i)
        srs.MarkerForegroundColorIndex = xlColorIndexNone
        srs.MarkerBackgroundColorIndex = xlColorIndexNone
           
    Next srs
    
    

End Sub


Sub ChartSeriesColorMarkers(ByRef objChart As Excel.Chart)
  If objChart.ChartType <> xlXYScatterLines Then Exit Sub
    
    Dim srs As Series
    Dim i As Integer
    
    i = 0
    'objChart.Activate 'Only needed when FullSerisCollectoion is used. Not seriescollection?
    
    For Each srs In objChart.FullSeriesCollection
      i = i + 1
       
       With srs.Format
         .Line.ForeColor.RGB = MyColor(i)
         .Line.BackColor.RGB = MyColor(i)
         '.Fill.ForeColor.RGB = 0 'MyColor(i)
         '.Fill.BackColor.RGB = 0 'MyColor(i)
         
        End With
        srs.MarkerForegroundColor = MyColor(i)
        'srs.MarkerForegroundColorIndex = xlColorIndexAutomatic
        
        Select Case i
          Case 6, 12
            srs.MarkerBackgroundColor = MyColor(i)
          Case Else
            srs.MarkerBackgroundColorIndex = xlColorIndexNone
            'srs.MarkerBackgroundColor = MyColor(i)
        End Select
           
    Next srs
    
    

End Sub

    

