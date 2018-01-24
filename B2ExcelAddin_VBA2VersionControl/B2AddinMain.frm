VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} B2AddinMain 
   Caption         =   "Main (B2 Addin)"
   ClientHeight    =   1605
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11190
   OleObjectBlob   =   "B2AddinMain.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "B2AddinMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cb01_Click()

End Sub

Private Sub btn01_Click()
Dim wb As Workbook
    
  cbSelActiveWB.Clear
  For Each wb In Application.Workbooks
    cbSelActiveWB.AddItem wb.Name
  Next
 
End Sub

Private Sub btnExportPNG_Click()
  ActiveWB.ExportPNG
End Sub

Private Sub cbSelActiveWB_Change()
  On Error GoTo EHand:
  Application.Workbooks(cbSelActiveWB.Value).Activate
    
  
  Exit Sub
EHand:
  MsgBox ("Workbook " & cbSelActiveWB.Value & " not open")

End Sub
