Attribute VB_Name = "ChartEdit"
Option Explicit
'All fuctions in this module takes an chart as argument and tweak it


Sub ChartSeriesLine(ByRef objChart As Excel.Chart)

    Dim srs As Series
    Dim i As Integer
    
    i = 0

    
    For Each srs In objChart.SeriesCollection
      i = i + 1

      'Common look
      srs.Format.Line.Weight = 1.25
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





Sub ChartSeriesColor(ByRef objChart As Excel.Chart)
    Dim srs As Series
    Dim i As Integer
    Dim mycolors(32) As Long
          
    mycolors(1) = RGB(219, 106, 41) 'Orange Kellys Color nbr 4
    mycolors(2) = RGB(147, 173, 60) 'Green Kellys Color nbr 18
    mycolors(3) = RGB(2, 152, 202) 'Turquoise
    mycolors(4) = RGB(255, 150, 150) 'Old pink
    mycolors(5) = RGB(172, 255, 128) 'MintGreen
    mycolors(6) = RGB(70, 240, 240) ' Cyan
    mycolors(7) = RGB(240, 30, 230) 'Magneta
    mycolors(8) = RGB(127, 128, 129) 'Grey Kellys nbr 8
    mycolors(9) = RGB(98, 166, 71) 'Green Kellys  nbr 9
    mycolors(10) = RGB(72, 56, 150) 'Purple Kellys nbr 13
    mycolors(11) = RGB(209, 45, 39) 'Red Kellys nbr 20
    mycolors(12) = RGB(235, 205, 62) ' Mustard Kellys nbr 2
    
    i = 0
    'objChart.Activate 'Only needed when FullSerisCollectoion is used. Not seriescollection?
    
    For Each srs In objChart.FullSeriesCollection
      i = i + 1
       
       With srs.Format
         .Line.ForeColor.RGB = mycolors(i)
         .Fill.ForeColor.RGB = mycolors(i)
         .Fill.BackColor.RGB = mycolors(i)
         
        End With
        srs.MarkerForegroundColor = mycolors(i)
        'srs.MarkerForegroundColorIndex = xlColorIndexAutomatic
        
        Select Case i
        
        Case 6, 12
          srs.MarkerBackgroundColor = mycolors(i)
        Case Else
          srs.MarkerBackgroundColorIndex = xlColorIndexNone
        End Select
           
    Next srs

End Sub
