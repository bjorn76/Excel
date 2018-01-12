Attribute VB_Name = "GraphDrawing"
Option Explicit
Option Compare Text

'Version    Date        Developer   Comments
'1.00       08/11/10    TCW         Add revision list
'1.01       23/06/11    TCW         Test works with Excel 2010
'1.02       23/06/11    TCW         GrowPoints: Gives 8 point rather than 7 point sysmbols. Add code to set outline thickness with Excel 2010 - this used by new macro GrowPoints_ThickenOutline
'1.03       24/06/11    TCW         Set colour in GrowPoints using RGB value instead of ColorIndex so independant of pallet setup
'1.04       27/06/11    TCW         SwopDataInSeries: Fix bug which caused no replacement to be made if the number of swops was less than expected
'1.05       02/07/11    TCW         Add: SeriesPaste
'1.06       07/07/11    TCW         Add: SeriesSort
'1.07       10/07/11    TCW         Add: yAxisCentre
'1.08       27/07/11    TCW         Add: xAxisCentre
'1.09       21/04/12    TCW         Add: FormatTxtBoxes
'1.10       09/07/12    TCW         Correction to ReplaceInSeries

Sub AddSeriesToCharts()
'Adds new lines to a series of charts
    
    Dim Chrt As String
    Dim ChrtPrefix() As Variant
    Dim ChrtSuffix() As Variant
    Dim col As String
    Dim DataColumns() As Variant
    Dim DataRange As String
    Dim DataSht() As Variant
    Dim Graph As Integer
    Dim lines As SeriesCollection
    Dim NewLine As Series
    Dim RunNum As Integer
    Dim RunNumStr As Variant
    Dim SeriesName As String
    Dim simulation As Integer
    Dim WkBk As Workbook
    
    '-----------------------------------------------------------------------------
    'Set up charts, data series, etc to plot
    'To customise macro, you should only need to change this section

    'Simulations
    ChrtPrefix = Array("EA_", "EB_", "EC_", "FA_", "FB_", "FC_") 'Prefixes of chart names
    DataSht = Array("E_Arun", "E_Brun", "E_Crun", "F_Arun", "F_Brun", "F_Crun") 'Prefixes of data sheet names - correponds with ChrtPrefix

    'Graphs for each simulation
    ChrtSuffix = Array("CO", "THC", "NOx", "NO2", "NO2NOx") 'Suffix of chart names
    DataColumns = Array("CS", "CT", "CV", "CX", "DF") 'Columns for data - corresponds with ChrtSuffix
    '-----------------------------------------------------------------------------
    
    'Get run number from user
    RunNumStr = InputBox("Number of simulation:", "Add Series to Charts", 0)
    If RunNumStr = "" Then 'Exit if cancel pressed
        Exit Sub
    End If
    RunNum = CInt(RunNumStr)
    
    'Get run series name from user
    SeriesName = InputBox("Name of series:", "Add Series to Charts", "R" & CStr(RunNum))
    If SeriesName = "" Then 'Exit if cancel pressed
        Exit Sub
    End If
    
    'Get current workbook
    Set WkBk = ActiveWorkbook
    
    With WkBk
        For simulation = 0 To UBound(DataSht) 'Loop through simulations
            For Graph = 0 To UBound(ChrtSuffix) 'Loop through all graphs for given simulation
                
                'Form chart name
                Chrt = ChrtPrefix(simulation) & ChrtSuffix(Graph)
                
                'Add new series, change name, make line thicker
                Set lines = .Charts(Chrt).SeriesCollection
                col = DataColumns(Graph)
                DataRange = "A:A," & col & ":" & col
                lines.Add Source:=.Sheets(DataSht(simulation) & CStr(RunNum)).Range(DataRange) _
                    , Rowcol:=xlColumns, SeriesLabels:=True, CategoryLabels:=True, Replace:=False
                Set NewLine = lines(lines.Count) 'Get object for new line
                ChangeSeriesName SeriesName, NewLine 'Change series name
                NewLine.Border.Weight = xlMedium 'Thicken line
            Next Graph
        Next simulation
    End With
End Sub


Private Sub CentreAxisH(Ax As Axis)
'If axis Ax is horizonal, then centre it horizontally with the PlotArea. Use "inside PlotArea", ie without axis ticks and labels

    Dim AxTitle As AxisTitle
    
    If Ax.HasTitle Then
        Set AxTitle = Ax.AxisTitle
        With Ax.Parent.PlotArea
            If AxTitle.Orientation = xlHorizontal Then
                AxTitle.Left = (.InsideWidth - AxTitle.Width) / 2 + .InsideLeft
            End If
        End With
    End If
End Sub


Private Sub CentreAxisV(Ax As Axis)
'If axis Ax is vertical, then centre it vertically with the PlotArea. Use "inside PlotArea", ie without axis ticks and labels

    Dim AxTitle As AxisTitle
    
    If Ax.HasTitle Then
        Set AxTitle = Ax.AxisTitle
        With Ax.Parent.PlotArea
            If AxTitle.Orientation <> xlHorizontal Then
                AxTitle.Top = (.InsideHeight - AxTitle.Height) / 2 + .InsideTop
            End If
        End With
    End If
End Sub


'Sub FormatTxtBoxes()
''Formats all the text boxes on the active chart. Attempts to do sub- and super-scripting of characters that need to be
'
'    Dim Chrt As Chart
'    Dim ChrtShape As Shape
'
'    'Check active sheet is a chart; if not ActiveChart will fail
'    If ActiveSheet.Type = xlWorksheet Then
'        MsgBox "The ""FormatTxtBoxes"" macro can only be run when a Chart is active.", vbOKOnly + vbCritical, "FormatTxtBoxes"
'        Exit Sub
'    End If
'
'    'Get object for active chart
'    Set Chrt = ActiveChart
'
'    'Loop through all shapes on chart. If shape is a text box, format it
'    For Each ChrtShape In Chrt.Shapes
'        With ChrtShape
'            If .Type = msoTextBox Then
'                With .TextFrame2.TextRange.Font
'                    .Name = "Arial"
'                    .Size = 18
'                End With
'                With .Fill 'Text box background
'                    .ForeColor.RGB = RGB(255, 255, 255) 'Set fill colour to white
'                    .Solid
'                End With
'                With .line 'Text box outline
'                    .Visible = msoTrue
'                    .Weight = 1 'Outline thickness
'                    .ForeColor.RGB = 0 'Black outline
'                End With
'
'                'Make an attemp to superscript and subscript things that need to be
'                With .TextFrame2
'                    ChartFormat.AutoSuperScript.TextRange 'Important for superscript to be before subscript
'                    ChartFormat.AutoSubScript.TextRange
'                End With
'            End If
'        End With
'    Next ChrtShape
'End Sub
'
'
'Private Sub ChangeSeriesName(Name As String, line As Series)
''Changes name of series Line to Name
'
'    Dim entry As String
'
'    With line
'        entry = .Formula 'Get series entry for line
'        'Change name in series and write it back
'        .Formula = Left(entry, InStr(entry, "(")) & """" & Name & """," & Right(entry, Len(entry) - InStr(entry, ","))
'    End With
'End Sub


Sub ReplaceInSeries()
'Does find and replace in all of the series in the active chart
    Const AppName As String = "Replace In Series"
    
    Dim Chrt As Chart
    Dim colSwop As Boolean
    Dim line As Series
    Dim nSwop As Integer
    Dim OldTxt As String
    Dim NewTxt As String
    
    Set Chrt = ActiveChart
    nSwop = 2
    
    'Get string to find and replace from user
    OldTxt = InputBox("Text to be replaced:", AppName)
    If OldTxt = "" Then 'Exit if cancel pressed
        Exit Sub
    End If
    NewTxt = InputBox("Replacement text:", AppName, OldTxt)
    If NewTxt = "" Then 'Exit if cancel pressed
        Exit Sub
    End If
    
    'Guess if user wants to replace column entries
    If Len(NewTxt) <= 2 And Len(OldTxt) <= 2 Then 'Probably want to swop column
        colSwop = True
    Else
        colSwop = False
    End If
    
    'Loop through all series on chart
    For Each line In Chrt.SeriesCollection
        SwopDataInSeries colSwop, line, NewTxt, OldTxt, nSwop, False
    Next line
End Sub


Sub GrowLines()
'Goes through all series on plot, if series has lines make medium thickness
    GrowLines2 ActiveChart
End Sub


Sub GrowLines2(Chrt As Chart)
'Goes through all series on plot, if series has lines make medium thickness
    
'Excel 2010 note:
'Below runs in Excel 2010, despite not being correct syntax for Excel 2010
'"Series.Border.LineStyle" becomes "Series.Format.Line.Style"  in Excel2010
'"Series.Border.Weight"    becomes "Series.Format.Line.Weight" in Excel2010. In 2000, the value is a code for one of four thicknesses. In 2010, the vaue is the thickness in points. xlMedium is equivalent to 2 point thickness.
'NB If use Excel 2000 syntax changes only the line between the points, while 2010 syntax changes the line around the points as well (a bug?!)
    
    Dim line As Series
    
    For Each line In Chrt.SeriesCollection
        With line.Border
            If .LineStyle <> xlLineStyleNone Then 'Series has a line
                .Weight = xlMedium 'Set thickness to medium
            End If
        End With
    Next line
End Sub


Sub GrowPoints()
'Goes through all series on plot, if series has points make 8 point and give black boarder
    GrowPoints2 ActiveChart, True
End Sub


Sub GrowPoints_ThickenOutline()
'Goes through all series on plot, if series has points make 8 point and give black boarder
'In Excel 2010 also sets line thickness to 1point. In Excel 2000, is the same as GrowPoints
    GrowPoints2 ActiveChart, True
End Sub


Sub GrowPoints2(Chrt As Chart, SetOutlineThickness As Boolean)
'Goes through all series on plot, if series has points make 8 point and give black boarder. If SetOutlineThickness=true then set point outline to 1 point
    
'Excel 2010 note:
'Objects the same in 2000 and 2010: "Series.MarkerStyle", "Series.MarkerSize", "Series.MarkerForegroundColorIndex", "Series.MarkerForegroundColor", etc
'In Excel 2000 line around point is 0.75 points thick. In 2010, the point outline can have any thickness

'KNOWN BUG: If SetOutlineThickness = trueLine, thicknesses and dash types end up as only those available in Excel 2000. This is a consequence of the limitations of VBA in Excel 2010; with the 2010 VBA command lines between points and point outlines are both set the same command.

    Dim line As Series
    Dim lineDash As Integer
    Dim lineVisible As Integer
    Dim lineWeight As Single
    Dim xl2010 As Boolean
    
    'Check Excel version
    xl2010 = Is2010
    
    For Each line In Chrt.SeriesCollection
        With line
            If .MarkerStyle <> xlMarkerStyleNone Then 'Series has symbols
                .MarkerForegroundColor = 0 'Set background black (Specify colour as RGB value so independant of pallet)
                .MarkerSize = 8 'Set size to 8 points
                
                If xl2010 And SetOutlineThickness Then 'In Excel 2010 can change thickness of outline around point. Set this to 1 point.
                    'When change point outline (MarkerLine) it also changes the line between the points. Note values for line so can change back.
                    lineDash = .Border.LineStyle 'Use 2000 syntax so can write back using 2000 syntax
                    lineWeight = .Border.Weight  'Use 2000 syntax. PROBLEM: In 2010 specify Weight in points, but can erite this back without changing points
                    lineVisible = .Format.line.Visible
                    
                    'Change settings for point outline
                    With .Format.line
                        .DashStyle = msoLineSolid
                        .Visible = msoTrue
                        .Weight = 1
                    End With
                        
                    'Put line between points to how it was. Need to use Excel 2000 syntax to change line without changing markers
                    With .Border
                        If lineVisible = msoFalse Then 'No line between points
                            .LineStyle = xlNone
                        Else
                            .LineStyle = lineDash
                            .Weight = lineWeight
                        End If
                    End With
                End If
            End If
        End With
    Next line
End Sub


Sub ReplicateGraph_ChangeColOnly()
'Makes a series of charts with the same format as the active chart, but with different series
'Changes the column for the y data, but leaves the worksheets with the data unchanged

    Dim ChrtSuffix() As Variant
    Dim DataCol1() As Variant
    Dim DataCol2() As Variant
    Dim Graph As Integer
    Dim LastChrt As Chart
    Dim line As Series
    Dim NewChrt As Chart
    Dim OldDataCol1 As String
    Dim OldDataCol2 As String
    Dim RefChrt As Chart
    Dim yLabels As Variant
    
    '-----------------------------------------------------------------------------
    'Set up charts, etc to plot
    'To customise macro, you should only need to change this section

    'Graphs for each simulation
    ChrtSuffix = Array("THC", "NOx", "NO2", "GasT") 'Suffix of chart names
    yLabels = Array("Cumulative THC Emissions / g", "Cumulative NOX Emissions / g", "Cumulative NO2 Emissions / g", "Gas Temperature / °C")
    
    OldDataCol1 = "AQ" 'Column to be replace in series
    DataCol1 = Array("AR", "AT", "AV", "B")  'Columns for data - corresponds with ChrtSuffix
    
    OldDataCol2 = "CS" 'Column to be replace in series
    DataCol2 = Array("CT", "CV", "CX", "BD") 'Columns for data - corresponds with ChrtSuffix
    '-----------------------------------------------------------------------------
    
    'Get objects of reference chart and workbook
    Set RefChrt = ActiveChart
    
    Set LastChrt = RefChrt
    For Graph = 0 To UBound(ChrtSuffix) 'Loop through all graphs
                
        'Create new chart
        RefChrt.Copy After:=LastChrt
        Set NewChrt = ActiveChart
        NewChrt.Name = RefChrt.Name & ChrtSuffix(Graph)
            
        With NewChrt
            'Change y axis label
            .Axes(xlValue).AxisTitle.Characters.Text = yLabels(Graph)
                
            'Loop through all lines on graph
            For Each line In .SeriesCollection
                'Swop column name in series
                SwopDataInSeries True, line, CStr(DataCol1(Graph)), OldDataCol1, 2, True
                SwopDataInSeries True, line, CStr(DataCol2(Graph)), OldDataCol2, 2, True
            Next line
                
            'Update name of last chart
            Set LastChrt = NewChrt
        End With
    Next Graph
End Sub


Sub ReplicateGraph_incSheetChange()
'Makes a series of charts with the same format as the active chart, but with different series

    Dim a As Integer
    Dim b As Integer
    Dim ChrtName As String
    Dim ChrtPrefix() As Variant
    Dim ChrtSuffix() As Variant
    Dim DataCol1() As Variant
    Dim DataCol2() As Variant
    Dim DataSht() As Variant
    Dim Graph As Integer
    Dim LastChrt As Chart
    Dim line As Series
    Dim NewChrt As Chart
    Dim OldDataCol1 As String
    Dim OldDataCol2 As String
    Dim OldDataShtName As String
    Dim OldSeries As String
    Dim RefChrt As Chart
    Dim simulation As Integer
    Dim yLabels As Variant
    
    '-----------------------------------------------------------------------------
    'Set up charts, data series, etc to plot
    'To customise macro, you should only need to change this section

    'Simulations
    ChrtPrefix = Array("EA_", "EB_", "EC_", "FA_", "FB_", "FC_") 'Prefixes of chart names
    DataSht = Array("E_Arun1", "E_Brun1", "E_Crun1", "F_Arun1", "F_Brun1", "F_Crun1") 'Data sheet names - correponds with ChrtPrefix

    'Graphs for each simulation
    ChrtSuffix = Array("CO", "THC", "NOx", "NO2", "NO2NOx", "GasT") 'Suffix of chart names
    yLabels = Array("Cumulative CO Emissions / g", "Cumulative THC Emissions / g", "Cumulative NOX Emissions / g", "Cumulative NO2 Emissions / g", "NO2/NOX Ratio", "Gas Temperature / °C")
    
    OldDataCol1 = "AQ" 'Column to be replace in series
    DataCol1 = Array("AQ", "AR", "AT", "AV", "DF", "B") 'Columns for data - corresponds with ChrtSuffix
    
    OldDataCol2 = "CS" 'Column to be replace in series
    DataCol2 = Array("CS", "CT", "CV", "CX", "DF", "BD") 'Columns for data - corresponds with ChrtSuffix
    '-----------------------------------------------------------------------------
    
    'Get objects of reference chart and workbook
    Set RefChrt = ActiveChart
    
    'Get name of data sheet to be replaced
    OldSeries = RefChrt.SeriesCollection(1).Formula
    a = InStr(OldSeries, ",")
    b = InStr(OldSeries, "!")
    OldDataShtName = Mid(OldSeries, a + 1, b - a - 1)
    
    Set LastChrt = RefChrt
    For simulation = 0 To UBound(DataSht) 'Loop through simulations
        For Graph = 0 To UBound(ChrtSuffix) 'Loop through all graphs for given simulation
            ChrtName = ChrtPrefix(simulation) & ChrtSuffix(Graph)
                
            'Create new chart
            RefChrt.Copy After:=LastChrt
            Set NewChrt = ActiveChart
            NewChrt.Name = ChrtName
            
            With NewChrt
                'Change y axis label
                .Axes(xlValue).AxisTitle.Characters.Text = yLabels(Graph)
                
                'Loop through all lines on graph
                For Each line In .SeriesCollection
                    'Swop data sheet name in series
                    SwopDataInSeries False, line, CStr(DataSht(simulation)), OldDataShtName, 2, False
                    
                    'Swop column name in series
                    SwopDataInSeries True, line, CStr(DataCol1(Graph)), OldDataCol1, 2, True
                    SwopDataInSeries True, line, CStr(DataCol2(Graph)), OldDataCol2, 2, True
                Next line
                
                'Update name of last chart
                Set LastChrt = NewChrt
            End With
        Next Graph
    Next simulation
End Sub


Sub SeriesPaste()
'Pastes the contents of the clipboard into the series collection of the active chart

    Dim Chrt As Chart
    Dim entry As String
    Dim NwSeries As Series
    Dim WkSht As Worksheet
    
    'Check Workbook actually contains a chart. Flag error if it doesn't
    If Charts.Count = 0 Then
        MsgBox "The Active Workbook has no Charts to paste into.", vbOKOnly + vbCritical, "SeriesPaste"
        Exit Sub
    End If
    
    'Check Clipboard contains text. Flag error if not
    If Application.ClipboardFormats(1) <> xlClipboardFormatText Then
        MsgBox "The Clipboard does not contain text." & vbCr & "Only text entries can be pasted into the SeriesCollection.", vbOKOnly + vbCritical, "SeriesPaste"
        Exit Sub
    End If
    
    'Check active sheet is a chart; if not ActiveChart will fail
    If ActiveSheet.Type = xlWorksheet Then
        MsgBox "The ""SeriesPaste"" macro can only be run when a Chart is active.", vbOKOnly + vbCritical, "SeriesPaste"
        Exit Sub
    End If
    
    'Get object for active chart
    Set Chrt = ActiveChart
    
    'Create new sheet, empty clipboard into it, take note of clipboard contents and delete new sheet
    Set WkSht = Chrt.Parent.Worksheets.Add
    With WkSht
        .Range("A1").Select
        .Paste
        entry = .Range("A1").Text 'Use text not value as later will return an error if Excel clipboard contents not in a form Excel recognises
        Application.DisplayAlerts = False 'Prevents box asking for deletion confirmation
        .Delete
        Application.DisplayAlerts = True
    End With
    
    'Check if "entry" contains something likely to be a chart series. Flag error if not
    If Left(entry, 8) <> "=SERIES(" Or Right(entry, 1) <> ")" Or Len(entry) < 13 Then
        MsgBox "The text you are trying to past into the SeriesCollection does not appear to be in the right format.", vbOKOnly + vbCritical, "SeriesPaste"
        Exit Sub
    End If
        
    'Add new series to chart. On charts with more than one y-axis, add to the first axis
    With Chrt.ChartGroups(1)
        Set NwSeries = .SeriesCollection.NewSeries
        NwSeries.Formula = entry
        NwSeries.PlotOrder = .SeriesCollection.Count 'Make the new series the last one in the list (for current ChartGroup / y-axis)
    End With
    
    'Ensure chart is still active
    Chrt.Activate
End Sub


Sub SeriesSort()
'Sorts the series (in each group) on the active chart into alphanumeric order by name
'Uses bubble sort algorithm
    
    Dim ChangeMade As Boolean
    Dim Chrt As Chart
    Dim Grp As ChartGroup
    Dim n As Integer
    Dim obj As Variant
    
    'Check active sheet is a chart; if not ActiveChart will fail
    If ActiveSheet.Type = xlWorksheet Then
        MsgBox "The ""SeriesSort"" macro can only be run when a Chart is active.", vbOKOnly + vbCritical, "SeriesSort"
        Exit Sub
    End If
    
    'Get the object for the active chart
    Set Chrt = ActiveChart
    
    'Sort series in each chart group in turn
    For Each Grp In Chrt.ChartGroups
'        With Grp 'This line works in Excel 2000, but inexplicibly fails with 2010
        If Is2010 Then 'Horride fudge to make this work in Excel 2010. Will crash if have 2 y-axes
            Set obj = Chrt
        Else
            Set obj = Grp
        End If
        With obj  'End of horrid fudge
            'Bubble sort on current group
            Do
                ChangeMade = False
                For n = 1 To .SeriesCollection.Count - 1
                    If .SeriesCollection(n).Name > .SeriesCollection(n + 1).Name Then
                        ChangeMade = True
                        .SeriesCollection(n + 1).PlotOrder = n
                    End If
                Next n
            Loop While ChangeMade
        End With
    Next Grp
End Sub


Private Sub SwopDataInSeries(colSwop As Boolean, line As Series, NewTxt As String, OldTxt As String, nSwop As Integer, ySwop As Boolean)
'Replaces text OldTxt with NewTxt in the chart series Line. ySwop=True for exchange in y part only. colSwop=true changes column entry. Attempts to make up to nSwop changes.

    Dim entry As String
    Dim i As Integer
    Dim InsertStart As Integer
    Dim LenEntry As Integer
    Dim LenOldTxt As Integer
    Dim NewEntry As String
    Dim Start As Integer
    
    LenOldTxt = Len(OldTxt)
    With line
        entry = .Formula 'Get series entry for line
        LenEntry = Len(entry)
        Start = InStr(entry, ",") 'Find first comma. Use to avoid searching for text to replace in name
        
        'If want to limit replacement to y entry, move start to y entry of series
        If ySwop Then
            Start = InStr(Start + 1, entry, ",")
        End If
        
        'If want to limit replacement to column rather than sheet name, move start to after !
        If colSwop Then
            Start = InStr(Start + 1, entry, "!")
        End If
            
        For i = 1 To nSwop
            InsertStart = InStr(Start, entry, OldTxt) - 1 'Position before start of insertion
            If InsertStart < 0 Then 'OldTxt not found
                If i = 1 Then 'No changes found so exit sub
                    Exit Sub
                Else 'At least one change found so exit loop and make change
                    Exit For
                End If
            End If
            NewEntry = Left(entry, InsertStart) & NewTxt & Right(entry, LenEntry - InsertStart - LenOldTxt)  'Insert NewTxt in place of OldTxt
            
            'Set variables up for second replacement
            entry = NewEntry
            LenEntry = Len(entry)
            Start = InsertStart + 2
        Next i

        'Write in new series
        .Formula = NewEntry
    End With
End Sub


Sub Sub_SuperscriptAxes()
'Attemps to sub/super-script axes titles and add degree signs as required
    
    Dim Ax As Axis
    Dim Chrt As Chart
    
    'Check active sheet is a chart; if not ActiveChart will fail
    If ActiveSheet.Type = xlWorksheet Then
        MsgBox "The ""Sub_SuperscriptAxes"" macro can only be run when a Chart is active.", vbOKOnly + vbCritical, "Sub_SuperscriptAxes"
        Exit Sub
    End If
    
    'Loop through all axes and sub/super-script if has a title
    Set Chrt = ActiveChart
    For Each Ax In Chrt.Axes
        If Ax.HasTitle Then
            ChartFormat.AutoDegreeAdd Ax
            ChartFormat.AutoSuperScript Ax.AxisTitle 'Important for superscript to be before subscript
            ChartFormat.AutoSubScript Ax.AxisTitle
        End If
    Next Ax
                
    'Centre x and y-axes. Sub/super-scripting changes length, so recentring required
     xAxisCentre2 Chrt
     yAxisCentre2 Chrt
End Sub

Sub xAxisCentre()
'Vertically centres any vertical axes on the active chart

    'Check active sheet is a chart; if not ActiveChart will fail
    If ActiveSheet.Type = xlWorksheet Then
        MsgBox "The ""xAxisCentre"" macro can only be run when a Chart is active.", vbOKOnly + vbCritical, "xAxisCentre"
        Exit Sub
    End If
    
    xAxisCentre2 ActiveChart
    
End Sub

Sub xAxisCentre2(Chrt As Chart)
'Horizonally centres any x-axes present with the PlotArea

    'Attempt to centre all axes. Only vertical ones are centred
    With Chrt
        CentreAxisH .Axes(xlValue)
        CentreAxisH .Axes(xlCategory)
        If .Axes.Count > 2 Then 'Have second x-axis?
            CentreAxisH .Axes(xlValue, xlSecondary)
        End If
    End With
End Sub


Sub yAxisCentre()
'Vertically centres any vertical axes on the active chart

    'Check active sheet is a chart; if not ActiveChart will fail
    If ActiveSheet.Type = xlWorksheet Then
        MsgBox "The ""yAxisCentre"" macro can only be run when a Chart is active.", vbOKOnly + vbCritical, "yAxisCentre"
        Exit Sub
    End If
    
    yAxisCentre2 ActiveChart
    
End Sub


Sub yAxisCentre2(Chrt As Chart)
'Vertically centres any y-axes present with the PlotArea

    'Attempt to centre all axes. Only vertical ones are centred
    With Chrt
        CentreAxisV .Axes(xlValue)
        CentreAxisV .Axes(xlCategory)
        If .Axes.Count > 2 Then 'Have second y-axis
            CentreAxisV .Axes(xlValue, xlSecondary)
        End If
    End With
End Sub

Function Is2010() As Boolean
'Returns true if Excel 2007 or later is being used

    Is2010 = Application.Version >= 12 ' Excel 2007 is version 12, so Is2010 true for Excel 2007, Excel 2010 and any subsequent versions

End Function
