Attribute VB_Name = "VBA_Templates"
Option Explicit

'
'
'


Private Sub MyGreataSub()
  On Error GoTo EH:
  MsgBox "Great to run!"
  
Exit Sub
EH:
  MsgBox Err.Description
  
End Sub
