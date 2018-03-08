Attribute VB_Name = "ExcelSettings"
Option Explicit

'Functions in this module shall verify and change excel settings:
' # User confirmation
' # Act on selected och active chart(s)

Sub StatusbarOn()
    Application.DisplayStatusBar = True

    
End Sub


Sub StatusbarOff()
    Application.DisplayStatusBar = False
End Sub

Sub ShowSettings()
Dim s As String ' KeySettings
Dim i As Integer


  's = s & "x: " & Application.x & vbCrLf
  s = "Decimal Separator: " & Application.DecimalSeparator & vbCrLf
  s = s & "DisplayStatusBar: " & Application.DisplayStatusBar & vbCrLf
  s = s & "Addins Count: " & Application.AddIns2.Count & vbCrLf
  
  For i = 1 To Application.AddIns2.Count
      s = s & Application.AddIns2(i).Name & vbCrLf
      If Application.AddIns2(i).Installed Then
        s = s & "Activated and installed in:" & vbCrLf
      Else
        s = s & "NOT activated but intalled in:" & vbCrLf
      End If
      
      s = s & Application.AddIns2(i).Path & vbCrLf
  Next
  
  
  If Application.AddIns2("Analysis ToolPak").Installed = True Then
    s = s & "Analysis ToolPak add-in is installed" & vbCrLf
  Else
    s = s & "Analysis ToolPak add-in is NOT installed" & vbCrLf
  End If
  
  
  MsgBox s
End Sub

