VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} B2AddinMain 
   Caption         =   "Main (B2 Addin)"
   ClientHeight    =   1860
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11190
   OleObjectBlob   =   "B2AddinMain.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "B2AddinMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Main from of B2Addin
' is Structured i Pages with diffrent themes
'
' Theme init


Const COMPMACROPATH = "C:\XP\JMR_CompMacro\"
Const COMPMACROFILE = "JMR_Comparison_Macro_B2_V002.xlsm"




Private Sub btn01_Click()
Dim wb As Workbook
    
  cbSelActiveWB.Clear
  For Each wb In Application.Workbooks
    cbSelActiveWB.AddItem wb.Name
  Next
  cbSelActiveWB.Value = ActiveWorkbook.Name
  
  
  
 
End Sub



Private Sub btnExportPNG_Click()
  ActiveWB.ExportPNG
End Sub


Private Sub btnFont_Click()
  ActiveChart.SelectChartFont
End Sub

Private Sub btnInitSheetsCombo_Click()
  Dim sht As Worksheet
   cbSelectActiveSheet.Clear
   For Each sht In ActiveWorkbook.Worksheets
     cbSelectActiveSheet.AddItem sht.Name
   Next
   cbSelectActiveSheet.Name = ActiveSheet.Name
   
   
   
End Sub

Private Sub btnLineW_Click()
  ActiveChart.LineW

End Sub

Private Sub btnMain_Click()
  PublicMod.Main
End Sub

Private Sub btnOpenCompMac_Click()
  Dim wb As Workbook
  'Set wb = Application.Workbooks.Open("C:\XP\JMR_CompMacro\JMR_Comparison_Macro_B2_V002.xlsm")
  Set wb = Application.Workbooks.Open(COMPMACROPATH & COMPMACROFILE)
  wb.Activate
  wb.Windows(1).ActivateNext
  
End Sub

Private Sub btnStartMac_Click()
  'Application.Run "JMR_Comparison_Macro_B2_V002.xlsm!Start_Macro"
  Application.Run COMPMACROFILE & "!Start_Macro"
End Sub

Private Sub btnAddDataSet_Click()
  Application.Run COMPMACROFILE & "!Start_AddDataset"
  
  
End Sub




Private Sub cbSelActiveWB_Change()
  On Error GoTo EHand:
  Application.Workbooks(cbSelActiveWB.Value).Activate
    
  
  Exit Sub
EHand:
  MsgBox ("Workbook " & cbSelActiveWB.Value & " not open")

End Sub

Private Sub cbSelActiveWB_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Dim wb As Workbook
  cbSelActiveWB.Clear
  For Each wb In Application.Workbooks
    cbSelActiveWB.AddItem wb.Name
  Next
End Sub

Private Sub cbSelectActiveSheet_Change()
     On Error GoTo EHand:
  Application.Workbooks(cbSelActiveWB.Value).Activate
  ActiveWorkbook.Sheets(cbSelectActiveSheet.Value).Activate
  'btnInitSheetsCombo_Click
  
  Exit Sub
EHand:
  MsgBox ("Workbook " & cbSelActiveWB.Value & " not open")

End Sub

