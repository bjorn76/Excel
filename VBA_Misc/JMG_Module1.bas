Attribute VB_Name = "JMG_Module1"
Sub GroupColumns()
'
' GroupColumns Macro
' Will group columns in the HDD Eval summary sheets (as of 2016-10-13)
'
    If Sheets("Main").Range("K5").Value <> "CO2" Then  ' Summary Sheet from before May 2016
        Columns("D:E").Group
        Columns("G:M").Group
        Columns("O:T").Group
        Columns("V:AA").Group
        Columns("AC:AH").Group
        Columns("AJ:AL").Group
        Columns("AN:AP").Group
        Columns("AS:AS").Group
        Columns("AU:AU").Group
        Columns("AW:BD").Group
        ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
    Else
        Columns("D:E").Group
        Columns("G:N").Group
        Columns("P:V").Group
        Columns("X:AD").Group
        Columns("AF:AL").Group
        Columns("AN:AP").Group
        Columns("AR:AT").Group
        Columns("AW:AW").Group
        Columns("AY:AY").Group
        Columns("BB:BI").Group
        ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
    End If
End Sub

Sub ProtectSheets()
'
' LockSheets Macro
'
    For i = 1 To Sheets.Count
        Sheets(i).Select
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Next i
End Sub

Sub UnprotectSheets()
'
' Unprotect Macro
'
'
    For i = 1 To Sheets.Count
        Sheets(i).Select
        ActiveSheet.Unprotect
    Next i
End Sub
Sub CopyInputSheets()
Attribute CopyInputSheets.VB_Description = "Copy input sheets described in the ""Info"" sheet to simulation data "
Attribute CopyInputSheets.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CopyInputSheets Macro
' Copy input sheets described in the "Import" sheet to simulation data
'

' Jonas Edvardsson 2013-07-01

    MyWorkbook = ActiveWorkbook.Name
    Sheets("Import").Select
    SheetNoImport = ActiveSheet.Index
    LastRow = ActiveSheet.UsedRange.Rows.Count
'    keepcolumns = InputBox("How many columns in the input file? Old models = 27 columns, new models =28 columns", "Select format", "28")
    keepcolumns = 28
    old_filelocation = ""
    For i = 1 To LastRow
        rowno = i + 2
        filelocationcell = "F" & rowno
        originalsheetcell = "G" & rowno
        renamedsheetcell = "H" & rowno
        alreadyimportedcell = "A" & rowno
        filelocation = Range(filelocationcell).Value
        originalsheet = Range(originalsheetcell).Value
        renamedsheet = Range(renamedsheetcell).Value
        alreadyimported = Range(alreadyimportedcell).Value
        
        If alreadyimported Or filelocation = "" Or originalsheet = "" Then
            
        Else
            
            If filelocation <> oldfilelocation Then
               If oldfilelocation <> "" Then
                  Workbooks(sFilename).Close
                End If
                Workbooks.Open Filename:=filelocation
                inputfile = ActiveWorkbook.Name
                oldfilelocation = filelocation
            Else
                Workbooks(inputfile).Activate
            End If
'            Sheets(originalsheet).Select
            Application.CutCopyMode = False
            
'           Delete sheet with original name first
            Application.DisplayAlerts = False
            On Error Resume Next
               Workbooks(MyWorkbook).Sheets(originalsheet).Delete
               Application.DisplayAlerts = True
            On Error GoTo 0
            Sheets(originalsheet).Copy After:=Workbooks(MyWorkbook).Sheets(SheetNoImport + i - 1)
            sFilename = Mid(filelocation, InStrRev(filelocation, "\") + 1, Len(filelocation))
'            Workbooks(sFileName).Close
            Workbooks(MyWorkbook).Activate
            Sheets(originalsheet).Select
            If keepcolumns = 27 Then
              Range("AB:RR").Delete
            ElseIf keepcolumns = 28 Then
              Range("AC:RR").Delete
            End If
            Application.DisplayAlerts = False
            If originalsheet <> renamedsheet Then
                On Error Resume Next
                Sheets(renamedsheet).Delete
                Application.DisplayAlerts = True
                On Error GoTo 0
                Sheets(originalsheet).Name = renamedsheet
            End If
            Set rng1 = Range("A:A").Find("Cycle feed", , xlValues, xlWhole)
            If rng1 Is Nothing Then
                Sheets(renamedsheet).Tab.ColorIndex = 4
            Else
                Sheets(renamedsheet).Tab.ColorIndex = 3
            End If
            Workbooks(MyWorkbook).Activate
            Sheets("Import").Select
            Range("I" & rowno).Value = Format(Date, "YYYY-MM-DD") & " " & Time$
            Range(alreadyimportedcell).Value = "1"
            
        End If
    Next i
    Workbooks(sFilename).Close
        
End Sub

Sub ListWorkSheetNames()
    
    StartCell = ActiveCell.Select
    For i = 1 To Sheets.Count
        c = ActiveCell.Column
        r = ActiveCell.Row
        Cells(r + i - 1, c).Value = Sheets(i).Name
    Next i
End Sub

Sub RenameSheetNames()
    StartCell = ActiveCell.Select
    r = ActiveCell.Row
    c_old = ActiveCell.Column
    c_new = c_old + 1
    changednames = 0
    Ready = False
    While Not Ready
        oldname = Cells(r, c_old).Value
        newname = Cells(r, c_new).Value
        If oldname = "" And newname = "" Then
            Ready = True
        Else
            If oldname <> newname Then
                Sheets(oldname).Name = newname
                changednames = changednames + 1
            End If
            r = r + 1
        End If
    Wend
    MsgBox ("Changed " & changednames & " sheet names")
End Sub

Sub ResetSource(Optional usedefault As Boolean = False)
'
' Reset_source Macro
'

   Dim mychart As Chart
   Dim shp As Shape
   Set mychart = Application.ActiveWorkbook.ActiveChart
   If ActiveWorkbook.ActiveSheet.Type = -4169 Then
      mysheetno = ActiveSheet.Index
      If mysheetno > 1 Then
         usesheetno = mysheetno - 1
      ElseIf ActiveWorkbook.Sheets.Count > 1 Then
         usesheetno = mysheetno + 1
      End If
      If ActiveWorkbook.Sheets(usesheetno).Type = -4169 Then
         usesheetno = mysheetno
      End If
      inputsheet = Application.ActiveWorkbook.Sheets(usesheetno).Name
   Else
      inputsheet = Application.ActiveWorkbook.ActiveSheet.Name
   End If
   If Not usedefault Then
       inputsheet = InputBox("Enter name of new data sheet", "", inputsheet)
   End If
   If inputsheet <> "" And Not (mychart Is Nothing) Then
       For x = 1 To mychart.SeriesCollection.Count
          With mychart.SeriesCollection(x)
    '        Debug.Print mychart.SeriesCollection(x).Formula
            originalformula = mychart.SeriesCollection(x).Formula
            firstcitationmark = InStr(originalformula, "'")
            secondcitationmark = InStr(firstcitationmark + 1, originalformula, "'")
            If firstcitationmark = 0 Then
               firstcitationmark = InStr(originalformula, "(")
               secondcitationmark = InStr(originalformula, "!")
               If Mid(originalformula, firstcitationmark + 1, 1) = "," Then
                  firstcitationmark = firstcitationmark + 1
               End If
               addcitation = True
            Else
               addcitation = False
            End If
            replacestring = Mid(originalformula, firstcitationmark + 1, secondcitationmark - firstcitationmark - 1)
            If addcitation Then
               inputsheet_mod = Chr(39) & inputsheet & Chr(39)
            Else
               inputsheet_mod = inputsheet
            End If
            newformula = Replace(originalformula, replacestring, inputsheet_mod)
            mychart.SeriesCollection(x).Formula = newformula
     '       .Interior.Color = RGB(x * 75, 50, x * 50)
          End With
       Next x
       If mychart.HasTitle Then
          mychart.ChartTitle.Formula = Replace(mychart.ChartTitle.Formula, replacestring, inputsheet_mod)
       End If
       For Each shp In mychart.Shapes
           If shp.Type = msoTextBox Then
              shp.Select
              Selection.Formula = Replace(Selection.Formula, replacestring, inputsheet_mod)
            End If
       Next
    End If

End Sub
Sub ResetSingleSource()
   ResetSource (False)
End Sub
Sub ResetAllSources()
    startsheet = ActiveSheet.Index
    For i = 1 To Sheets.Count
        Sheets(i).Activate
        ResetSource (True)
    Next i
    Sheets(startsheet).Activate
End Sub

Sub ChartsToPresentation()
 ' If error then you must set a VBA reference to Microsoft PowerPoint Object Library
 ' This is done under Tools --> References (in this program)
 Dim PPT As PowerPoint.Application
 ' Dim PPApp As PowerPoint.Application
 Dim PPPres As PowerPoint.Presentation
 Dim PPSlide As PowerPoint.Slide
 Dim PresentationFileName As Variant
 Dim SlideCount As Long
 Dim iCht As Integer
 Const Factor = 28.346
 ' Reference existing instance of PowerPoint
 On Error Resume Next
Set PPApp = GetObject(, "Powerpoint.Application")
If PPApp Is Nothing Then
   Set PPApp = CreateObject("PowerPoint.Application")
   On Error GoTo 0
   PPApp.Visible = True
   Set PPPres = PPApp.Presentations.Add
Else
    ' Reference active presentation
    Set PPPres = PPApp.ActivePresentation
End If

 PPApp.ActiveWindow.ViewType = ppViewSlide
 For iSheet = 1 To Application.ActiveWorkbook.Worksheets.Count
    Application.ActiveWorkbook.Worksheets(iSheet).Select
     For iCht = 1 To ActiveSheet.ChartObjects.Count
     ' copy chart as a picture
         ActiveSheet.ChartObjects(iCht).Chart.CopyPicture _
         Appearance:=xlScreen, Size:=xlScreen, Format:=xlPicture
         ' Add a new slide and paste in the chart
        
         SlideCount = PPPres.Slides.Count
         Set PPSlide = PPPres.Slides.Add(SlideCount + 1, ppLayoutBlank)
         PPApp.ActiveWindow.View.GotoSlide PPSlide.SlideIndex
         With PPSlide
     ' paste and select the chart picture
            .Shapes.Paste.Select
     ' align the chart
            With PPApp.ActiveWindow.Selection.ShapeRange
                .Left = 3 * Factor
                .Top = 3 * Factor
                .ScaleHeight 0.7, msoTrue
                .ZOrder msoSendToBack
            End With
            
        End With
        PPSlide.NotesPage.Shapes(2).TextFrame.TextRange.Text = "Graph from " & ActiveWorkbook.FullNameURLEncoded & "#" & ActiveSheet.Name
    Next
 ' Clean up
 Next
 
 For iCht = 1 To Application.ActiveWorkbook.Charts.Count
    Application.ActiveWorkbook.Charts(iCht).Select
    ActiveChart.CopyPicture Appearance:=xlScreen, Size:=xlScreen, Format:=xlPicture
    SlideCount = PPPres.Slides.Count
    Set PPSlide = PPPres.Slides.Add(SlideCount + 1, ppLayoutBlank)
    PPApp.ActiveWindow.View.GotoSlide PPSlide.SlideIndex
    With PPSlide
' paste and select the chart picture
       .Shapes.Paste.Select
' align the chart
        With PPApp.ActiveWindow.Selection.ShapeRange
            .Left = 3 * Factor
            .Top = 3 * Factor
            .ScaleHeight 0.7, msoTrue
            .ZOrder msoSendToBack
        End With
    End With
    PPSlide.NotesPage.Shapes(2).TextFrame.TextRange.Text = "Graph from " & ActiveWorkbook.FullNameURLEncoded & "#" & ActiveSheet.Name
Next

PPApp.ActiveWindow.ViewType = ppViewNormal
 Set PPSlide = Nothing
 Set PPPres = Nothing
 Set PPApp = Nothing
  
End Sub

Sub ChartToPresentation()
' If error then you must set a VBA reference to Microsoft PowerPoint Object Library
 ' This is done under Tools --> References (in this program)
 Dim PPT As PowerPoint.Application
 ' Dim PPApp As PowerPoint.Application
 Dim PPPres As PowerPoint.Presentation
 Dim PPSlide As PowerPoint.Slide
 Dim PresentationFileName As Variant
 Dim SlideCount As Long
 Dim iCht As Integer
 Dim iSheet As String
 
 Const Factor = 28.346
 ' Reference existing instance of PowerPoint
    iSheet = ActiveSheet.Name
    TypeOfSheet = ActiveSheet.Type
         
    On Error Resume Next
    Set PPApp = GetObject(, "Powerpoint.Application")
    If PPApp Is Nothing Then
       Set PPApp = CreateObject("PowerPoint.Application")
       On Error GoTo 0
       PPApp.Visible = True
       Set PPPres = PPApp.Presentations.Add
       SlideCount = PPPres.Slides.Count
'       Set PPSlide = PPPres.Slides.Add(1, ppLayoutBlank)
       
    Else
        ' Reference active presentation
        Set PPPres = PPApp.ActivePresentation
    End If
    PPApp.ActiveWindow.ViewType = ppViewSlide
    
    Select Case TypeOfSheet
    
        Case xlChart, -4169
            Application.ActiveWorkbook.Charts(iSheet).Select
            ActiveChart.CopyPicture Appearance:=xlScreen, Size:=xlScreen, Format:=xlPicture
            SlideCount = PPPres.Slides.Count
            Set PPSlide = PPPres.Slides.Add(SlideCount + 1, ppLayoutBlank)
            PPApp.ActiveWindow.View.GotoSlide PPSlide.SlideIndex
            With PPSlide
        ' paste and select the chart picture
               .Shapes.Paste.Select
        ' align the chart
                With PPApp.ActiveWindow.Selection.ShapeRange
                    .Left = 3 * Factor
                    .Top = 3 * Factor
                    .ScaleHeight 0.7, msoTrue
                    .ZOrder msoSendToBack
                End With
            End With
        
        
        Case xlWorkbook, -4167
    
            For iCht = 1 To Application.Sheets(iSheet).ChartObjects.Count
                 ' copy chart as a picture
                ActiveSheet.ChartObjects(iCht).Chart.CopyPicture _
                Appearance:=xlScreen, Size:=xlScreen, Format:=xlPicture
                     ' Add a new slide and paste in the chart
                    
                SlideCount = PPPres.Slides.Count
                Set PPSlide = PPPres.Slides.Add(SlideCount + 1, ppLayoutBlank)
                PPApp.ActiveWindow.View.GotoSlide PPSlide.SlideIndex
                With PPSlide
                 ' paste and select the chart picture
                    .Shapes.Paste.Select
                 ' align the chart
                    With PPApp.ActiveWindow.Selection.ShapeRange
                        .Left = 3 * Factor
                        .Top = 3 * Factor
                        .ScaleHeight 0.7, msoTrue
                        .ZOrder msoSendToBack
                    End With
                End With
            Next
            
    End Select
    
 ' Clean up
    PPSlide.NotesPage.Shapes(2).TextFrame.TextRange.Text = "Graph from " & ActiveWorkbook.FullNameURLEncoded & "#" & ActiveSheet.Name
    PPApp.ActiveWindow.ViewType = ppViewNormal
    Set PPSlide = Nothing
    Set PPPres = Nothing
    Set PPApp = Nothing
End Sub

Function WorksheetExists(wsName As String) As Boolean
    Dim ws As Worksheet
    Dim ret As Boolean
    ret = False
    wsName = UCase(wsName)
    For Each ws In ActiveWorkbook.Sheets
        If UCase(ws.Name) = wsName Then
            ret = True
            Exit For
        End If
    Next
    WorksheetExists = ret
End Function

Sub ReplaceColumns()
   Load frmReplaceStrings
   frmReplaceStrings.RefreshData
End Sub

Sub EvalToPivot()
'
' EvalToPivot Macro
'

'
    Dim ws As Worksheet
    
    ' Copy "Main" to the new workbook
    Sheets("Main").Select
    Sheets("Main").Copy
    
    ' Make a "SystemDescription' worksheet
    Columns("D:D").Select
    Selection.Copy
    Set ws = Sheets.Add(After:=ActiveSheet)
    ws.Name = "SystemDescription"
    Sheets("SystemDescription").Select
    With ActiveWorkbook.Sheets("SystemDescription").Tab
        .Color = 15773696
        .TintAndShade = 0
    End With
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.Range("$A$1:$A$5000").RemoveDuplicates Columns:=1, Header:=xlNo
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Catalysts"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Description"
    Range("A2:B3").Select
    Selection.Delete Shift:=xlUp
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "System 1"
    Selection.AutoFill Destination:=Range("B2:B100")
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "External name"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "SCR 1"
    Selection.AutoFill Destination:=Range("C2:C100")
    
    Sheets("Main").Select
    With ActiveWorkbook.Sheets("Main").Tab
        .Color = 10498160
        .TintAndShade = 0
    End With

    
    If Sheets("Main").Range("K5").Value = "CO2" Then  ' Summary Sheet from May 2016
        Rows("7:7").Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("A7").Select
        ActiveCell.FormulaR1C1 = "Test #"
        Range("B7").Select
        ActiveCell.FormulaR1C1 = "Date"
        Range("C7").Select
        ActiveCell.FormulaR1C1 = "Cycle"
        Range("D7").Select
        ActiveCell.FormulaR1C1 = "Test object"
        Range("E7").Select
        ActiveCell.FormulaR1C1 = "Fuel"
        Range("F7").Select
        ActiveCell.FormulaR1C1 = "Comments"
        Range("G7").Select
        ActiveCell.FormulaR1C1 = "Comments2"
        Range("H7").Select
        ActiveCell.FormulaR1C1 = "R1 HC"
        Range("I7").Select
        ActiveCell.FormulaR1C1 = "R1 NOx"
        Range("J7").Select
        ActiveCell.FormulaR1C1 = "R1 CO"
        Range("K7").Select
        ActiveCell.FormulaR1C1 = "R1 CO2"
        Range("L7").Select
        ActiveCell.FormulaR1C1 = "R1 N2O"
        Range("M7").Select
        ActiveCell.FormulaR1C1 = "R1 NO2/NOx"
        Range("N7").Select
        ActiveCell.FormulaR1C1 = "R1 NH3 avg"
        Range("O7").Select
        ActiveCell.FormulaR1C1 = "R1 NH3 max"
        Range("P7").Select
        ActiveCell.FormulaR1C1 = "R2 HC"
        Range("Q7").Select
        ActiveCell.FormulaR1C1 = "R2 NOx"
        Range("R7").Select
        ActiveCell.FormulaR1C1 = "R2 CO"
        Range("S7").Select
        ActiveCell.FormulaR1C1 = "R2 CO2"
        Range("T7").Select
        ActiveCell.FormulaR1C1 = "R2 N2O"
        Range("U7").Select
        ActiveCell.FormulaR1C1 = "R2 NO2/NOx"
        Range("V7").Select
        ActiveCell.FormulaR1C1 = "R2 NH3 avg"
        Range("W7").Select
        ActiveCell.FormulaR1C1 = "R2 NH3 max"
        Range("X7").Select
        ActiveCell.FormulaR1C1 = "R3 HC"
        Range("Y7").Select
        ActiveCell.FormulaR1C1 = "R3 NOx"
        Range("Z7").Select
        ActiveCell.FormulaR1C1 = "R3 CO"
        Range("AA7").Select
        ActiveCell.FormulaR1C1 = "R3 CO2"
        Range("AB7").Select
        ActiveCell.FormulaR1C1 = "R3 N2O"
        Range("AC7").Select
        ActiveCell.FormulaR1C1 = "R3 NO2/NOx"
        Range("AD7").Select
        ActiveCell.FormulaR1C1 = "R3 NH3 avg"
        Range("AE7").Select
        ActiveCell.FormulaR1C1 = "R3 NH3 max"
        Range("AF7").Select
        ActiveCell.FormulaR1C1 = "R4 HC"
        Range("AG7").Select
        ActiveCell.FormulaR1C1 = "R4 NOx"
        Range("AH7").Select
        ActiveCell.FormulaR1C1 = "R4 CO"
        Range("AI7").Select
        ActiveCell.FormulaR1C1 = "R4 CO2"
        Range("AJ7").Select
        ActiveCell.FormulaR1C1 = "R4 N2O"
        Range("AK7").Select
        ActiveCell.FormulaR1C1 = "R4 NO2/NOx"
        Range("AL7").Select
        ActiveCell.FormulaR1C1 = "R4 NH3 avg"
        Range("AM7").Select
        ActiveCell.FormulaR1C1 = "R4 NH3 max"
        Range("AN7").Select
        ActiveCell.FormulaR1C1 = "EO BP"
        Range("AO7").Select
        ActiveCell.FormulaR1C1 = "dP1"
        Range("AP7").Select
        ActiveCell.FormulaR1C1 = "dP2"
        Range("AQ7").Select
        ActiveCell.FormulaR1C1 = "dP3"
        Range("AR7").Select
        ActiveCell.FormulaR1C1 = "HC Conv"
        Range("AS7").Select
        ActiveCell.FormulaR1C1 = "NO Conv"
        Range("AT7").Select
        ActiveCell.FormulaR1C1 = "NOx conv"
        Range("AU7").Select
        ActiveCell.FormulaR1C1 = "CO conv"
        Range("AV7").Select
        ActiveCell.FormulaR1C1 = "ANR"
        Range("AW7").Select
        ActiveCell.FormulaR1C1 = "NH3 avg"
        Range("AX7").Select
        ActiveCell.FormulaR1C1 = "NH3 max"
        Range("AY7").Select
        ActiveCell.FormulaR1C1 = "PM"
        Range("AZ7").Select
        ActiveCell.FormulaR1C1 = "PN"
        Range("BA7").Select
        ActiveCell.FormulaR1C1 = "BSFC"
        Range("BB7").Select
        ActiveCell.FormulaR1C1 = "LO HC 80%"
        Range("BC7").Select
        ActiveCell.FormulaR1C1 = "LO CO 80%"
        Range("BD7").Select
        ActiveCell.FormulaR1C1 = "LO NOx 50%"
        Range("BE7").Select
        ActiveCell.FormulaR1C1 = "LO NO 50%"
        Range("BF7").Select
        ActiveCell.FormulaR1C1 = "LO HC Max"
        Range("BG7").Select
        ActiveCell.FormulaR1C1 = "LO CO Max"
        Range("BH7").Select
        ActiveCell.FormulaR1C1 = "LO NOx Max"
        Range("BI7").Select
        ActiveCell.FormulaR1C1 = "LO NO Max"
        Range("BJ7").Select
        ActiveCell.FormulaR1C1 = "LO NO Max Temp"
        Range("BK7").Select
        ActiveCell.FormulaR1C1 = "Time"
        Range("BL7").Select
        ActiveCell.FormulaR1C1 = "Quality"
    
    ElseIf Sheets("Main").Range("G5").Value = "" Then  ' New Format of Summary Sheet
        Rows("7:7").Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("A7").Select
        ActiveCell.FormulaR1C1 = "Test #"
        Range("B7").Select
        ActiveCell.FormulaR1C1 = "Date"
        Range("C7").Select
        ActiveCell.FormulaR1C1 = "Cycle"
        Range("D7").Select
        ActiveCell.FormulaR1C1 = "Test object"
        Range("E7").Select
        ActiveCell.FormulaR1C1 = "Fuel"
        Range("F7").Select
        ActiveCell.FormulaR1C1 = "Comments"
        Range("G7").Select
        ActiveCell.FormulaR1C1 = "Comments2"
        Range("H7").Select
        ActiveCell.FormulaR1C1 = "R1 HC"
        Range("I7").Select
        ActiveCell.FormulaR1C1 = "R1 NOx"
        Range("J7").Select
        ActiveCell.FormulaR1C1 = "R1 CO"
        Range("K7").Select
        ActiveCell.FormulaR1C1 = "R1 N2O"
        Range("L7").Select
        ActiveCell.FormulaR1C1 = "R1 NO2/NOx"
        Range("M7").Select
        ActiveCell.FormulaR1C1 = "R1 NH3 avg"
        Range("N7").Select
        ActiveCell.FormulaR1C1 = "R1 NH3 max"
        Range("O7").Select
        ActiveCell.FormulaR1C1 = "R2 HC"
        Range("P7").Select
        ActiveCell.FormulaR1C1 = "R2 NOx"
        Range("Q7").Select
        ActiveCell.FormulaR1C1 = "R2 CO"
        Range("R7").Select
        ActiveCell.FormulaR1C1 = "R2 N2O"
        Range("S7").Select
        ActiveCell.FormulaR1C1 = "R2 NO2/NOx"
        Range("T7").Select
        ActiveCell.FormulaR1C1 = "R2 NH3 avg"
        Range("U7").Select
        ActiveCell.FormulaR1C1 = "R2 NH3 max"
        Range("V7").Select
        ActiveCell.FormulaR1C1 = "R3 HC"
        Range("W7").Select
        ActiveCell.FormulaR1C1 = "R3 NOx"
        Range("X7").Select
        ActiveCell.FormulaR1C1 = "R3 CO"
        Range("Y7").Select
        ActiveCell.FormulaR1C1 = "R3 N2O"
        Range("Z7").Select
        ActiveCell.FormulaR1C1 = "R3 NO2/NOx"
        Range("AA7").Select
        ActiveCell.FormulaR1C1 = "R3 NH3 avg"
        Range("AB7").Select
        ActiveCell.FormulaR1C1 = "R3 NH3 max"
        Range("AC7").Select
        ActiveCell.FormulaR1C1 = "R4 HC"
        Range("AD7").Select
        ActiveCell.FormulaR1C1 = "R4 NOx"
        Range("AE7").Select
        ActiveCell.FormulaR1C1 = "R4 CO"
        Range("AF7").Select
        ActiveCell.FormulaR1C1 = "R4 N2O"
        Range("AG7").Select
        ActiveCell.FormulaR1C1 = "R4 NO2/NOx"
        Range("AH7").Select
        ActiveCell.FormulaR1C1 = "R4 NH3 avg"
        Range("AI7").Select
        ActiveCell.FormulaR1C1 = "R4 NH3 max"
        Range("AJ7").Select
        ActiveCell.FormulaR1C1 = "EO BP"
        Range("AK7").Select
        ActiveCell.FormulaR1C1 = "dP1"
        Range("AL7").Select
        ActiveCell.FormulaR1C1 = "dP2"
        Range("AM7").Select
        ActiveCell.FormulaR1C1 = "dP3"
        Range("AN7").Select
        ActiveCell.FormulaR1C1 = "HC Conv"
        Range("AO7").Select
        ActiveCell.FormulaR1C1 = "NO Conv"
        Range("AP7").Select
        ActiveCell.FormulaR1C1 = "NOx conv"
        Range("AQ7").Select
        ActiveCell.FormulaR1C1 = "CO conv"
        Range("AR7").Select
        ActiveCell.FormulaR1C1 = "ANR"
        Range("AS7").Select
        ActiveCell.FormulaR1C1 = "NH3 avg"
        Range("AT7").Select
        ActiveCell.FormulaR1C1 = "NH3 max"
        Range("AU7").Select
        ActiveCell.FormulaR1C1 = "PM"
        Range("AV7").Select
        ActiveCell.FormulaR1C1 = "PN"
        Range("AW7").Select
        ActiveCell.FormulaR1C1 = "LO HC 80%"
        Range("AX7").Select
        ActiveCell.FormulaR1C1 = "LO CO 80%"
        Range("AY7").Select
        ActiveCell.FormulaR1C1 = "LO NOx 50%"
        Range("AZ7").Select
        ActiveCell.FormulaR1C1 = "LO NO 50%"
        Range("BA7").Select
        ActiveCell.FormulaR1C1 = "LO HC Max"
        Range("BB7").Select
        ActiveCell.FormulaR1C1 = "LO CO Max"
        Range("BC7").Select
        ActiveCell.FormulaR1C1 = "LO NOx Max"
        Range("BD7").Select
        ActiveCell.FormulaR1C1 = "LO NO Max"
        Range("BE7").Select
        ActiveCell.FormulaR1C1 = "LO NO Max Temp"
        Range("BF7").Select
        ActiveCell.FormulaR1C1 = "Time"
    Else     ' Old format of Summary sheet
        Rows("7:7").Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("A7").Select
        ActiveCell.FormulaR1C1 = "Test #"
        Range("B7").Select
        ActiveCell.FormulaR1C1 = "Date"
        Range("C7").Select
        ActiveCell.FormulaR1C1 = "Cycle"
        Range("D7").Select
        ActiveCell.FormulaR1C1 = "Test object"
        Range("E7").Select
        ActiveCell.FormulaR1C1 = "Fuel"
        Range("F7").Select
        ActiveCell.FormulaR1C1 = "Comments"
        Range("G7").Select
        ActiveCell.FormulaR1C1 = "PM"
        Range("H7").Select
        ActiveCell.FormulaR1C1 = "EO HC"
        Range("I7").Select
        ActiveCell.FormulaR1C1 = "EO NOx"
        Range("J7").Select
        ActiveCell.FormulaR1C1 = "EO CO"
        Range("K7").Select
        ActiveCell.FormulaR1C1 = "TP HC"
        Range("L7").Select
        ActiveCell.FormulaR1C1 = "TP NOx"
        Range("M7").Select
        ActiveCell.FormulaR1C1 = "TP CO"
        Range("N7").Select
        ActiveCell.FormulaR1C1 = "HC Conv"
        Range("O7").Select
        ActiveCell.FormulaR1C1 = "NOx conv"
        Range("P7").Select
        ActiveCell.FormulaR1C1 = "CO conv"
        Range("Q7").Select
        ActiveCell.FormulaR1C1 = "NH3 avg"
        Range("R7").Select
        ActiveCell.FormulaR1C1 = "NH3 max"
        Range("S7").Select
        ActiveCell.FormulaR1C1 = "N2O avg"
        Range("T7").Select
        ActiveCell.FormulaR1C1 = "N2O max"
        Range("U7").Select
        ActiveCell.FormulaR1C1 = "NO2/NOx"
        Range("V7").Select
        ActiveCell.FormulaR1C1 = "EO BP"
        Range("W7").Select
        ActiveCell.FormulaR1C1 = "dP1"
        Range("X7").Select
        ActiveCell.FormulaR1C1 = "dP2"
        Range("Y7").Select
        ActiveCell.FormulaR1C1 = "dP3"
        Range("Z7").Select
        ActiveCell.FormulaR1C1 = "LO HC 80%"
        Range("AA7").Select
        ActiveCell.FormulaR1C1 = "LO CO 80%"
        Range("AB7").Select
        ActiveCell.FormulaR1C1 = "LO NOx 50%"
        Range("AC7").Select
        ActiveCell.FormulaR1C1 = "LO NO 50%"
        Range("AD7").Select
        ActiveCell.FormulaR1C1 = "LO HC Max"
        Range("AE7").Select
        ActiveCell.FormulaR1C1 = "LO CO Max"
        Range("AF7").Select
        ActiveCell.FormulaR1C1 = "LO NOx Max"
        Range("AG7").Select
        ActiveCell.FormulaR1C1 = "LO NO Max"
        Range("AH7").Select
        ActiveCell.FormulaR1C1 = "LO NO Max Temp"
        Range("AI7").Select
        ActiveCell.FormulaR1C1 = "Time"
    End If
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A7").Select
    ActiveCell.FormulaR1C1 = "Test"
    Range("A6").Select
    ActiveCell.FormulaR1C1 = "Type of test (WHTC/FTP/USwing/..)"
    Range("A5").Select
    ActiveCell.FormulaR1C1 = "Add more columns to the right of Column A (then automatically included in data for pivot table)"
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B7").Select
    ActiveCell.FormulaR1C1 = "Valid"
    Range("B6").Select
    ActiveCell.FormulaR1C1 = "'1/0"
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C7").Select
    ActiveCell.FormulaR1C1 = "ANR"
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D7").Select
    ActiveCell.FormulaR1C1 = "System"
    Range("D8").Select
    ActiveCell.Formula = "=VLOOKUP($H8,'SystemDescription'!$A$1:$AZ$500,2,0)"
    Selection.AutoFill Destination:=Range("D8:D500")
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E7").Select
    ActiveCell.FormulaR1C1 = "Repeat"
    Range("E6").Select
    ActiveCell.FormulaR1C1 = "1,2,3 (Repeated WHTCs..)"
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("F7").Select
    ActiveCell.FormulaR1C1 = "Duplicate"
    Range("F6").Select
    ActiveCell.FormulaR1C1 = "1,2,3,..100? (during ageing, other duplicates)"
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("G7").Select
    ActiveCell.FormulaR1C1 = "Catalyst state"
    Range("G6").Select
    ActiveCell.FormulaR1C1 = "Cond/Aged/.."
    Range("A7").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.AutoFilter

'   Freeze top 7 rows
    ActiveWindow.FreezePanes = False
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 7
    End With
    ActiveWindow.FreezePanes = True
    

    
    ActiveWorkbook.Names.Add Name:="PivotInput", RefersToR1C1:= _
        Selection
    
    Set ws = Sheets.Add(Before:=ActiveSheet)
    ws.Name = "PivotChart"

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "PivotInput", Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:="PivotChart!R1C1", TableName:="PivotTable1", DefaultVersion _
        :=xlPivotTableVersion14
    Sheets("PivotChart").Select
    Cells(1, 1).Select
    ActiveSheet.Shapes.AddChart.Select
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Cycle")
        .Orientation = xlRowField
        .Position = 1
    End With
        With ActiveSheet.PivotTables("PivotTable1").PivotFields("System")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("NOx Conv"), "Count of NOx conv", xlCount

    ActiveSheet.ChartObjects("Chart 1").Activate
    Selection.Placement = xlFreeFloating

End Sub

Sub ExcelTestsToMatlab()
'
' ExcelTestsToMatlab Macro
'

    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
   
    sourceCol = 1
    RowCount = Cells(Rows.Count, sourceCol).End(xlUp).Row

    'for every row, find the first blank cell and select it
    For currentRow = RowCount To 1 Step -1
        currentRowValue = Cells(currentRow, sourceCol).Value
        If IsEmpty(currentRowValue) Or currentRowValue = "" Then
            ActiveSheet.Rows(currentRow).Delete
        End If
    Next
    RowCount = Cells(Rows.Count, sourceCol).End(xlUp).Row

    Range("B1").Select
    ActiveCell.FormulaR1C1 = "=""'""&RC[-1]&""',"""
    Range("B1").Select
    If RowCount > 1 Then
        Selection.AutoFill Destination:=Range("B1:B" & RowCount)
    End If
    Range("B1:B" & RowCount).Select
    Selection.Copy
    ActiveWindow.Close False
    MsgBox ("Names are ready to be pasted in Matlab.")
End Sub

Sub UpdateResultSheet()
    
Dim RAW_RowCnt, RES_RowCnt, i, j, x, y As Integer
Dim NoOfNew, NoOfReplace As Integer
Dim XPPath, XP, RES_FileName, TempFileName As String
Dim RAW_IDInfo(0 To 1000), RAW_TimeStamp(0 To 1000), RES_IDInfo(0 To 1000), RES_TimeStamp(0 To 1000) As String
Dim RAW_RemPos(0 To 1000), RES_RemPos(0 To 1000), TempPos As Integer
Dim HypPath, HypFileName As String
Dim Result As Byte
Dim ReturnCode As VbMsgBoxResult
     
On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    XPPath = ActiveWorkbook.Path
    XP = Mid(XPPath, Len(XPPath) - 8, 4)
    RES_FileName = ActiveWorkbook.Name
    If Range("A1").Value <> "Shared result sheet" Then
       ReturnCode = MsgBox("The macro will only work for sheet containing the string" & vbCrLf & "Shared result sheet" & vbCrLf & "in cell A1." & vbCrLf & "Do you want to add the string to the document?", vbYesNo)
       If ReturnCode = vbYes Then
          Range("A1").Value = "Shared result sheet"
       Else
          Error (100)
       End If
    End If
       
    
    Range("D3").Value = "XP" + XP + " RESULTS" 'Adds header to the RESULTAT-file

'**** OPEN AND SEARCHING ****
     Workbooks.OpenText Filename:=XPPath + "\XP" + XP + "RAW.txt", Origin _
        :=1251, StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote _
        , ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:= _
        False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers _
        :=True
        
    RAW_RowCnt = ActiveSheet.UsedRange.Rows.Count 'Reads the Test # in xxxxRAW.txt
    For i = 0 To RAW_RowCnt - 1
      RAW_IDInfo(i) = Range("A1").Offset(i, 0).Value
      RAW_TimeStamp(i) = Range("BF1").Offset(i, 0).Value
    Next i
    
    Windows(RES_FileName).Activate 'Reads the Test # in RESULTAT.xls
    RES_RowCnt = Range("A65000").End(xlUp).Row - 6
    For j = 0 To RES_RowCnt - 1
      RES_IDInfo(j) = Range("A1").Offset(j + 6, 0).Value
      RES_TimeStamp(j) = Range("BF1").Offset(j + 6, 0).Value
    Next j
  
'**** COMPARES ****
    x = 0
    NoOfNew = 0
    NoOfReplace = 0
    For i = RAW_RowCnt - 1 To 0 Step -1 'Compares RAW and RESULTAT to find the missing tests
      y = -1
      For j = RES_RowCnt - 1 To -1 Step -1
        If (j = -1) Then 'If RAW not found in RESULTAT then remember the position in xxxxRAW.txt
          If (y = -1) Then
            NoOfNew = NoOfNew + 1
           Else
            NoOfReplace = NoOfReplace + 1
          End If
          RES_RemPos(x) = y
          RAW_RemPos(x) = i
          x = x + 1
          Exit For
        End If
        If (RAW_IDInfo(i) = RES_IDInfo(j)) Then
         'If (RES_TimeStamp(j) <> "") Then
          If (RAW_TimeStamp(i) = RES_TimeStamp(j)) Or (RES_TimeStamp(j) = "") Then
           Exit For
          Else
           y = j
          End If
        ' End If
        End If
      Next j
    Next i

'**** COPIES ****
   If (NoOfNew <> 0) Or (NoOfReplace <> 0) Then 'Adds the missing results from RAW to RESULTAT
     For i = 0 To x - 1
      Windows("XP" + XP + "RAW.txt").Activate
      TempFileName = Range("A1").Offset(RAW_RemPos(i), 0).Value
      Range(Cells(RAW_RemPos(i) + 1, 1), Cells(RAW_RemPos(i) + 1, 65)).Cut '//2009-10-21 AP
      Windows(RES_FileName).Activate
      
      If (RES_RemPos(i) = -1) Then
       TempPos = RES_RowCnt + 6 + NoOfNew - i
       Result = vbOK
      Else
       Result = MsgBox("Replace " + TempFileName + "?", vbOKCancel)
       TempPos = RES_RemPos(i) + 7
      End If
       
      If Result = vbOK Then
       Range(Cells(TempPos, 1), Cells(TempPos, 65)).Select '//2009-10-21 AP
       ActiveSheet.Paste
      
       Range("A1").Offset(TempPos - 1, 0).Select 'Add Hyperlinks //2010-05-17 AP
       LinkFriendlyName = Selection.Value
       HypPath = Left(XPPath, Len(XPPath) - 4)
      
       HypFileName = Range("A1").Offset(TempPos - 1, 0).Value + ".xlsm"
       HypFull = HypPath & HypFileName
       ActiveCell.Formula = "=HYPERLINK(""" & HypFull & """,""" & LinkFriendlyName & """)"
              
'       ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:= _
'        HypPath + HypFileName, TextToDisplay:= _
'         Left(HypFileName, InStr(HypFileName, ".") - 1)
      Else: NoOfReplace = NoOfReplace - 1
      End If
     Next i
     
'**** COSMETIC **** '//2010-05-17 AP
    Range("H:K").NumberFormat = "0.000"
    Range("L:N").NumberFormat = "0.0"
    
    Range("O:R").NumberFormat = "0.000"
    Range("S:U").NumberFormat = "0.0"
    
    Range("V:Y").NumberFormat = "0.000"
    Range("Z:AB").NumberFormat = "0.0"
    
    Range("AC:AF").NumberFormat = "0.000"
    Range("AG:AT").NumberFormat = "0.0"
    Range("AW:BE").NumberFormat = "0.0"
            
  
                
    Range("B:B").NumberFormat = "m/d/yyyy"
    
    Range("A:B").HorizontalAlignment = xlCenter
    Range("E:E").HorizontalAlignment = xlCenter
    Range("G:BD").HorizontalAlignment = xlCenter
   End If
       
'**** END ****
    Workbooks("XP" + XP + "RAW.txt").Close SaveChanges:=False
    MsgBox CStr(NoOfNew) + " new result(s) has been added." + Chr(10) + CStr(NoOfReplace) + " result(s) has been replaced."
    Range("A1").Offset(RES_RowCnt + 6, 0).Select
    Application.ScreenUpdating = True
Exit Sub

ErrorHandler:
    MsgBox ("Error, the macro will stop!")
    Application.ScreenUpdating = True
End Sub

Sub ClearExcessRowsAndColumns()
    Dim ar As Range, r As Long, c As Long, tr As Long, tc As Long, x As Range
    Dim wksWks As Worksheet, ur As Range, arCount As Integer, i As Integer
    Dim blProtCont As Boolean, blProtScen As Boolean, blProtDO As Boolean
    Dim shp As Shape

    If ActiveWorkbook Is Nothing Then Exit Sub

    On Error Resume Next
    For Each wksWks In ActiveWorkbook.Worksheets
        Err.Clear
        Set ur = Nothing
        'Store worksheet protection settings and unprotect if protected.
        blProtCont = wksWks.ProtectContents
        blProtDO = wksWks.ProtectDrawingObjects
        blProtScen = wksWks.ProtectScenarios
        wksWks.Unprotect ""
        If Err.Number = 1004 Then
            Err.Clear
            MsgBox "'" & wksWks.Name & _
                   "' is protected with a password and cannot be checked." _
                 , vbInformation
        Else
            Application.StatusBar = "Checking " & wksWks.Name & _
                                    ", Please Wait..."
            r = 0
            c = 0

            'Determine if the sheet contains both formulas and constants
            Set ur = Union(wksWks.UsedRange.SpecialCells(xlCellTypeConstants), _
                           wksWks.UsedRange.SpecialCells(xlCellTypeFormulas))
            'If both fails, try constants only
            If Err.Number = 1004 Then
                Err.Clear
                Set ur = wksWks.UsedRange.SpecialCells(xlCellTypeConstants)
            End If
            'If constants fails then set it to formulas
            If Err.Number = 1004 Then
                Err.Clear
                Set ur = wksWks.UsedRange.SpecialCells(xlCellTypeFormulas)
            End If
            'If there is still an error then the worksheet is empty
            If Err.Number <> 0 Then
                Err.Clear
                If wksWks.UsedRange.Address <> "$A$1" Then
                    wksWks.UsedRange.EntireRow.Hidden = False
                    wksWks.UsedRange.EntireColumn.Hidden = False
                    wksWks.UsedRange.EntireRow.RowHeight = _
                    IIf(wksWks.StandardHeight <> 12.75, 12.75, 13)
                    wksWks.UsedRange.EntireColumn.ColumnWidth = 10
                    wksWks.UsedRange.EntireRow.Clear
                    'Reset column width which can also _
                     cause the lastcell to be innacurate
                    wksWks.UsedRange.EntireColumn.ColumnWidth = _
                    wksWks.StandardWidth
                    'Reset row height which can also cause the _
                     lastcell to be innacurate
                    If wksWks.StandardHeight < 1 Then
                        wksWks.UsedRange.EntireRow.RowHeight = 12.75
                    Else
                        wksWks.UsedRange.EntireRow.RowHeight = _
                        wksWks.StandardHeight
                    End If
                Else
                    Set ur = Nothing
                End If
            End If
            'On Error GoTo 0
            If Not ur Is Nothing Then
                arCount = ur.Areas.Count
                'determine the last column and row that contains data or formula
                For Each ar In ur.Areas
                    i = i + 1
                    tr = ar.Range("A1").Row + ar.Rows.Count - 1
                    tc = ar.Range("A1").Column + ar.Columns.Count - 1
                    If tc > c Then c = tc
                    If tr > r Then r = tr
                Next
                'Determine the area covered by shapes
                'so we don't remove shading behind shapes
                For Each shp In wksWks.Shapes
                    tr = shp.BottomRightCell.Row
                    tc = shp.BottomRightCell.Column
                    If tc > c Then c = tc
                    If tr > r Then r = tr
                Next
                Application.StatusBar = "Clearing Excess Cells in " & _
                                        wksWks.Name & ", Please Wait..."
                If r < wksWks.Rows.Count Then
                    Set ur = wksWks.Rows(r + 1 & ":" & wksWks.Rows.Count)
                    ur.EntireRow.Hidden = False
                    ur.EntireRow.RowHeight = IIf(wksWks.StandardHeight <> 12.75, _
                                                 12.75, 13)
                    'Reset row height which can also cause the _
                     lastcell to be innacurate
                    If wksWks.StandardHeight < 1 Then
                        ur.RowHeight = 12.75
                    Else
                        ur.RowHeight = wksWks.StandardHeight
                    End If
                    Set x = ur.Dependents
                    If 1 = 0 Then
                        ur.Clear
                    Else
                        ur.Delete
                    End If
                End If
                If c < wksWks.Columns.Count Then
                    Set ur = wksWks.Range(wksWks.Cells(1, c + 1), _
                                          wksWks.Cells(1, wksWks.Columns.Count)).EntireColumn
                    ur.EntireColumn.Hidden = False
                    ur.ColumnWidth = 18

                    'Reset column width which can _
                     also cause the lastcell to be innacurate
                    ur.EntireColumn.ColumnWidth = _
                    wksWks.StandardWidth

                    Set x = ur.Dependents
                    If Err.Number = 0 Then
                        ur.Clear
                    Else
                        ur.Delete
                    End If
                End If
            End If
        End If
        'Reset protection.
        wksWks.Protect "", blProtDO, blProtCont, blProtScen
        Err.Clear
    Next
    Application.StatusBar = False
    MsgBox "'" & ActiveWorkbook.Name & _
           "' has been cleared of excess formatting." & Chr(13) & _
           "You must save the file to keep the changes.", vbInformation
End Sub


