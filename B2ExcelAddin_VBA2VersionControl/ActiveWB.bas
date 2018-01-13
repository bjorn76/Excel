Attribute VB_Name = "ActiveWB"
Option Explicit

Sub ExportPNG()
'Export all charts as png pict in active workbook. Saved in a subfolder
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



