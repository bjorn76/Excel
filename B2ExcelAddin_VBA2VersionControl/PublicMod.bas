Attribute VB_Name = "PublicMod"
Option Explicit

Sub Main()
  Dim ct As Chart
  Dim rv As Integer
      
  'Workbook
  rv = MsgBox(Application.ActiveChart.Name & " Do you want to change the look?", vbOKCancel, "ActiveChart is...")
  If rv = 2 Then Exit Sub
  
  Set ct = Application.ActiveChart
  ChartSeriesFixedLook ct
  'ChartScaleUswing ct
  'ChartLables ct, "NOx Conversion U-Swing", "Temp(°C)", "NOx Conv.(%)"
  PrivateMod.ExportPNG
  
  
End Sub


Sub MainTest()
 Dim rv As Integer
 Dim R1, R2 As Range
 

rv = MsgBox(Application.ActiveChart.Name & "    Proceed? ", vbOKCancel, "MainTest")
  If rv = 2 Then Exit Sub
  'MsgBox rv
  'MsgBox ActiveChart.FullSeriesCollection(5).Name
  'MsgBox CStr(ActiveChart.FullSeriesCollection(5).XValues(1))
  'MsgBox
  'Set R1 = ActiveChart.FullSeriesCollection(1).XValues '= "='Sys6'!$B$3:$B$11"
  'MsgBox R1.Count
  'MsgBox ActiveChart.FullSeriesCollection(1).N
  '"=SERIES('Sys6'!$A$1,'Sys6'!$B$3:$B$11,'Sys6'!$C$3:$C$11,5)"
   
End Sub

Sub MainTest2()
 Dim rv As Integer
 Dim R1, R2 As Range
 

rv = MsgBox(Application.ActiveChart.Name & "    Proceed? ", vbOKCancel, "MainTest")
  
End Sub
