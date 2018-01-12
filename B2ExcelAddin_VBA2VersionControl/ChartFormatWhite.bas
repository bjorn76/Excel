Attribute VB_Name = "ChartFormatWhite"
Option Compare Text
Option Explicit

'Version    Date        Developer   Comments
'1.00       08/11/10    TCW         Add revision list
'1.01       27/06/11    TCW         Adapt for Excel 2010: Set legend border and background; set gridline style depending on Excel version
'1.02       27/06/11    TCW         Specify colours by RGB value rather than location in pallet to make independant of Excel copy
'1.03       03/07/11    TCW         Reorganise code. Add automatic subscripting and addition of degree symbol. Add formating of second y-axis
'1.04       15/07/11    TCW         Improve to work for Area, Bar, Column and Line plots, as well as Scatter plots. Works for 1 and 2 value axes (hopefully).
'1.05       27/07/11    TCW         Add call centre x-axis


Sub FormatChartWhite()
'Formats active chart by calling FormatChart2

    Dim Chrt As Object
    
    'Check active sheet is a chart; if not ActiveChart will fail
    If ActiveSheet.Type = xlWorksheet Then
           Set Chrt = ActiveChart
           If Chrt Is Nothing Then
              Set Chrt = Worksheets(ActiveSheet.Name).ChartObjects(1)
           End If
           Worksheets(ActiveSheet.Name).ChartObjects(1).Placement = xlFreeFloating
'            MsgBox "The ""FormatChart"" macro can only be run when a Chart is active.", vbOKOnly + vbCritical, "FormatChart"
'            Exit Sub
    End If
       
    Set Chrt = ActiveChart
    If Chrt Is Nothing Then
       MsgBox ("No chart selected")
    Else
        With Chrt
            'Hide chart so runs faster. Then format it
    '        .Visible = xlSheetHidden
            FormatChart2 Chrt
       
            'Make chart visible & active again
     '       .Visible = xlSheetVisible
    '        .Activate
        End With
    End If
End Sub


Sub FormatChart2(Chrt As Chart)
'Formats Chrt
    
    'Constants to control aparence of chart
    Const AxisFontSize As Byte = 18 'Font size for axis lables
    Const TickFontSize As Byte = 16 'Font size for axis numbers and legend
    Const LineThick As Integer = xlThin 'Thickness of lines. Options: xlHairline, xlThin, xlMedium, xlThick.


    Dim ChrtType As Integer
    Dim HasCategoryTicks As Boolean
    Dim HasMarkers As Boolean
    Dim IsAreaOrLine As Boolean
    Dim IsBar As Boolean
    Dim xCont As Boolean
    Dim xl2010 As Boolean
    
    Const Factor = 28.346
        
    'Check version of Excel
    xl2010 = Is2010
    
    
    'Detect chart type. Set flags controlling formatting accordingly
    ChrtType = Chrt.ChartType
    If ChrtType = -4111 Then ChrtType = Chrt.SeriesCollection(1).ChartType 'If chart has two y-axis base formatting on chart type of first series
    Select Case ChrtType
        Case xlArea, xlAreaStacked, xlAreaStacked100
            'Area plot
            HasCategoryTicks = True
            HasMarkers = False
            IsAreaOrLine = True
            IsBar = False
            xCont = False
        Case xlBarClustered, xlBarStacked, xlBarStacked100
            'Bar plot
            HasCategoryTicks = False
            HasMarkers = False
            IsAreaOrLine = False
            IsBar = True
            xCont = False
        Case xlColumnClustered, xlColumnStacked, xlColumnStacked100
            'Column plot
            HasCategoryTicks = False
            HasMarkers = False
            IsAreaOrLine = False
            IsBar = False
            xCont = False
        Case xlLine, xlLineMarkers, xlLineMarkersStacked, xlLineMarkersStacked100, xlLineStacked, xlLineStacked100
            'Line plot
            HasCategoryTicks = True
            HasMarkers = True
            IsAreaOrLine = True
            IsBar = False
            xCont = False
        Case xlXYScatter, xlXYScatterLines, xlXYScatterLinesNoMarkers, xlXYScatterSmooth, xlXYScatterSmoothNoMarkers
            'Scatter plot
            HasCategoryTicks = True
            HasMarkers = True
            IsAreaOrLine = False
            IsBar = False
            xCont = True
    End Select
            
    'Format area around chart. Have no fill and no border so works better when paste chart into another application
    With Chrt.ChartArea
        .Border.LineStyle = xlNone
        .Interior.ColorIndex = xlNone
        .Width = 30 * Factor
        .Height = 20 * Factor
    End With
    FormatFont Chrt.ChartArea, TickFontSize
    
    'Format plot border and set location and width
    Chrt.SizeWithWindow = False
            
    FormatLine Chrt.PlotArea, LineThick
    With Chrt.PlotArea
        .Left = 41
        .Top = 55
        .Width = 661
        .Height = Chrt.ChartArea.Height - .Top - 25
        If IsBar And Chrt.Axes.Count > 2 Then 'Bar chart with two value axes (x-axes)
            'Adjust PlotArea to make way for extra y-axis
            .Top = 28
            .Height = 0.883 * Chrt.ChartArea.Height
        End If
    End With
        
    'Format chart background / fill
    'Use different code fo Excel 2000 and 2010. 2000 code will work in 2010, but 2010 offers chance to specify colours by RGB value so apperance is independant of colour pallet (copy of Excel)
    If xl2010 Then 'Excel 2010
        With Chrt.PlotArea.Format.Fill
            .Visible = msoTrue
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .ForeColor.RGB = RGB(255, 255, 255) 'Pale blue
            .BackColor.RGB = RGB(255, 255, 255) 'White
        End With
    Else 'Excel 2000
        With Chrt.PlotArea.Fill
            .TwoColorGradient Style:=msoGradientHorizontal, Variant:=1
            .Visible = True
            .ForeColor.SchemeColor = 37 'Pale blue (hopefully)
            .BackColor.SchemeColor = 2 'White
        End With
    End If
    
    'Format gridlines
    With Chrt.Axes(xlValue).MajorGridlines.Border
        .Color = 0 'Black
        .Weight = xlHairline
        If xl2010 Then 'Excel 2010 syntax
            Chrt.SetElement msoElementPrimaryValueGridLinesMajor
            With .Parent.Format.line
                .Visible = msoTrue
                .DashStyle = msoLineDash
            End With
        Else 'Excel 2000 syntax
            .LineStyle = xlDot
        End If
    End With
   
    'Format axes
    With Chrt
        FormatAxis .Axes(xlValue), AxisFontSize, True, False, True, LineThick, TickFontSize 'xlValue axis
        FormatAxis .Axes(xlCategory), AxisFontSize, HasCategoryTicks, IsAreaOrLine, xCont, LineThick, TickFontSize 'xlCategory axis
        If IsBar Then 'If bar chart then y-axis is category axis
            .Axes(xlCategory).AxisTitle.Left = 7
        Else 'y-axis is value axis
            .Axes(xlValue).AxisTitle.Left = 7
        End If
        If .Axes.Count > 2 Then 'Have second y-axis
            FormatAxis .Axes(xlValue, xlSecondary), AxisFontSize, True, False, True, LineThick, TickFontSize 'Secondary xlValue axis
            .PlotArea.Width = 647
            If Not IsBar Then .Axes(xlValue, xlSecondary).AxisTitle.Left = 694 'Position second value axis if it is a y-axis
        End If
    End With
    If xl2010 Then 'Centre y-axis labels vertically and x-axis labels horizonatally. This doesn't work in Excel 2000
        GraphDrawing.yAxisCentre2 Chrt
        GraphDrawing.xAxisCentre2 Chrt
    End If
      
    'Format legend
    Chrt.HasLegend = True
    FormatFont Chrt.Legend, TickFontSize
    FormatLine Chrt.Legend, LineThick
    With Chrt.Legend.Interior 'Set legend fill
        .Color = RGB(255, 255, 255) 'White
        .PatternColor = RGB(255, 255, 255) 'White
        .Pattern = xlSolid
    End With
    With Chrt.Legend 'Set legend location
      .Shadow = False
    End With
    
    'Add title
    Chrt.HasTitle = True
    FormatFont Chrt.ChartTitle, AxisFontSize
'    AutoSuperScript Chrt.ChartTitle
    AutoSubScript Chrt.ChartTitle
    
    'If x-axis continuos, set limt on time axis to an approriate value
    If xCont And Not xl2010 Then SetTimeAxis Chrt 'SetTimeAxis fails in Excel 2010
    
    'Format lines/borders and markers
    If HasMarkers Then
        'Thicken lines and enlarge points
        GraphDrawing.GrowPoints2 Chrt, True
        GraphDrawing.GrowLines2 Chrt
    Else
        'Thicken all borders
        FormatAllLines Chrt, LineThick
    End If
End Sub


Sub AutoDegreeAdd(obj As Axis)
'Adds a degree symbol to an axis label if the last character is "C" and not already proceeded by a degree symbol
    
    Dim Label As String
    
    With obj
        Label = Trim(.AxisTitle.Text)
        If Right(Label, 1) = "C" And Right(Label, 2) <> "°C" Then
            .AxisTitle.Text = Left(Label, Len(Label) - 1) & "°C"
        End If
    End With
End Sub


Sub AutoSubScript(obj As Variant)
'Searches text in obj for things that should be subscripted and subscripts them

    Dim i As Integer
    Dim n As Integer
    Dim SubList As Variant
    
    'List of items which have the last character subscripted
    SubList = Array("C2", "C3", "C7", "H2", "H4", "H6", "H8", "N2", "NH3", "NO2", "NOX", "O2", "SO3", "SOX")
    
    With obj
        'Loop through SubList looking for things to subscript
        For i = 0 To UBound(SubList)
            n = 0
            Do 'Loop until all instances are found
                n = n + 1
                n = InStr(n, .Text, SubList(i), vbTextCompare)
                If n <> 0 Then 'Item located
                    .Characters(Start:=(n + Len(SubList(i)) - 1), length:=1).Font.Subscript = True
                End If
            Loop Until n = 0
        Next i
    End With
End Sub


Sub AutoSuperScript(obj As Variant)
'Searches text in obj for things that should be superscripted and superscripts them
   
    Dim i As Integer
    Dim n As Integer
    Dim SupList As Variant
    
    'List of items which have the last character subscripted
    SupList = Array("2", "3", "-1", "-2", "-3")
    
    With obj
        'Loop through SubList looking for things to subscript
        For i = 0 To UBound(SupList)
            n = 0
            Do 'Loop until all instances are found
                n = n + 1
                n = InStr(n, .Text, SupList(i), vbTextCompare)
                If n <> 0 Then 'Item located
                    .Characters(Start:=n, length:=Len(SupList(i))).Font.Superscript = True
                End If
            Loop Until n = 0
        Next i
    End With
End Sub


Private Sub FormatAllLines(Chrt As Chart, Thickness As Integer)
'Set thickeness of all the borders/outlines in the SeriesCollection

    Dim line As Series
    
    For Each line In Chrt.SeriesCollection
        FormatLine line, Thickness
    Next line
End Sub


Private Sub FormatAxis(obj As Axis, AxisFontSize As Byte, HasTicks As Boolean, IsAreaLineCat As Boolean, IsCont As Boolean, Thickness As Integer, TickFontSize As Byte)
'Formats chart axis

    Dim form As String
    Dim i As Byte
    Dim n As Byte
    Dim TickUnitStr As String

    With obj
        'If axis does not have label, add one
        If Not .HasTitle Then .HasTitle = True
        
        'Format axis label text and set line thickness
        FormatFont .AxisTitle, AxisFontSize
        FormatLine obj, Thickness
        
        'Set axis ticks (if required)
        If HasTicks Then
            .MajorTickMark = xlCross
            .MinorTickMark = xlInside
            .TickLabelPosition = xlNextToAxis
            If IsAreaLineCat Then .MinorTickMark = xlTickMarkNone 'No intermediate (minor) ticks on the category axis of Area or Line plot
        Else
            .MajorTickMark = xlTickMarkNone
            .MinorTickMark = xlTickMarkNone
        End If
        
        'Set axis scales
        If IsCont Then 'Continuous scale on axis
'            .MinimumScale = 0
'            .MaximumScaleIsAuto = True
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = xlAutomatic
            .ReversePlotOrder = False
            .DisplayUnit = xlNone
            On Error Resume Next 'Ignore errors, as sometimes next line fails for no apparent reason
                .ScaleType = xlLinear
            On Error GoTo 0
        Else 'Discrete scale, non-continuous scale
            .CrossesAt = 1
            .TickLabelSpacing = 1
            .TickMarkSpacing = 1
            .AxisBetweenCategories = Not IsAreaLineCat 'If bar or column category axis this true - ticks between categories (columns). If area or line category axis this false - ticks below markers
            .ReversePlotOrder = False
        End If
        
        'Set font of numbers by tick labels
        FormatFont .TickLabels, TickFontSize
        
        'Set format of tick labels appropriately (continuous axis only)
        If IsCont Then
            TickUnitStr = CStr(.MajorUnit)
            n = InStr(1, TickUnitStr, ".") 'Locate decimal point
            If n = 0 Then
                form = "General"
            Else
                n = Len(TickUnitStr) - n 'Number of decimal places required
                form = "0."
                For i = 1 To n
                    form = form & "0"
                Next i
            End If
            .TickLabels.NumberFormat = form
        End If
        
        'Add degree sysmbol(s) as required
        AutoDegreeAdd obj
        
        'Superscript and subscript axis title as appropriate
        AutoSuperScript obj.AxisTitle
        AutoSubScript obj.AxisTitle
    End With
End Sub


Private Sub FormatFont(obj As Variant, Size As Byte)
'Formates fonts

    obj.AutoScaleFont = True
    With obj.Font
        .Name = "Arial"
        .FontStyle = "Regular"
        .Size = Size
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .Color = 0 'Black
        .Background = xlAutomatic
    End With
End Sub


Private Sub FormatLine(obj As Variant, Thickness As Integer)
'Formats lines

    With obj.Border
        .Color = 0 'Black
        .Weight = Thickness
        .LineStyle = xlContinuous
    End With
End Sub


Private Sub SetTime(obj As Axis, DataEnd As Single, TestEnd As Integer, AxisEnd As Integer)
'Tests whether last time is equal to TestTime (within tolarance) set end time on axis to AxisEnd

    Const Tolerance As Byte = 10
    
    If DataEnd >= (TestEnd - Tolerance) And DataEnd <= (TestEnd + Tolerance) Then
        obj.MaximumScale = AxisEnd
    End If
End Sub


Private Sub SetTimeAxis(Chrt As Chart)
'Attempts to identify which test cycle is being run and sets limit on time axis appropriately

    Dim Finaltime As Single
    Dim xLabels As Variant

    With Chrt
        'Identify last value in x-axis data
        xLabels = .Axes(xlCategory).CategoryNames
        Finaltime = xLabels(UBound(xLabels))

        'Compare end time with those for various test cycles. Set x-axis limit accordingly.
        SetTime .Axes(xlCategory), Finaltime, 1180, 1200 'NEDC
        SetTime .Axes(xlCategory), Finaltime, 1199, 1400 'HD-FTP
        SetTime .Axes(xlCategory), Finaltime, 1239, 1400 'NRTC
        SetTime .Axes(xlCategory), Finaltime, 1800, 2000 'WHTC
    End With
End Sub

Function Is2010() As Boolean
'Returns true if Excel 2007 or later is being used

    Is2010 = Application.Version >= 12 ' Excel 2007 is version 12, so Is2010 true for Excel 2007, Excel 2010 and any subsequent versions

End Function
