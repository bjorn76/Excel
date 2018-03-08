VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} B2AddinMain 
   Caption         =   "Main (B2 Addin)"
   ClientHeight    =   2385
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
' Moduel is sorted like the forms pages with diffrent themes:
' Form and controls init +
' Workbook and Sheet ComboBoxes events
' Workbook commands
' Chart Commands
' JMR compmacro control
' Module Export to enable Version controll of VBA code.



'http://www.cpearson.com/excel/SuppressChangeInForms.htm
Public EnableEvents As Boolean



Const COMPMACROPATH = "C:\XP\JMR_CompMacro\"
Const COMPMACROFILE = "JMR_Comparison_Macro_B2_V003.xlsm"




'***********************************************************
'Form Event Handlers
'***********************************************************

Private Sub UserForm_Initialize()
  Me.EnableEvents = False
  
  InitWBcombo
  InitSheetsCombo
  'MsgBox "UserForm init completed" 'Debug line
  Me.EnableEvents = True
End Sub


'***********************************************************
'Combo Events
'**********************************************************
Private Sub cbSelActiveWB_Change()
  On Error GoTo EHand:
  If Me.EnableEvents = False Then
    Exit Sub
  End If
  
  Unload Me 'To change active wb
  If cbSelActiveWB.Value <> "" Then
    Application.Workbooks(cbSelActiveWB.Value).Activate
  End If
  B2AddinMain.Show

  Exit Sub
EHand:
  MsgBox ("Workbook " & cbSelActiveWB.Value & " not open")

End Sub


Private Sub cbSelectActiveSheet_Change()
  On Error GoTo EHand:
  If Me.EnableEvents = False Then
    Exit Sub
  End If
  
  Application.Workbooks(cbSelActiveWB.Value).Activate
  ActiveWorkbook.Sheets(cbSelectActiveSheet.Value).Activate
  
  Exit Sub
EHand:
  MsgBox ("Workbook " & cbSelActiveWB.Value & " not open")

End Sub





'***********************************************************
'Init page
'***********************************************************

Private Sub btnInitSheetsCombo_Click()
  InitSheetsCombo
End Sub

Private Sub InitSheetsCombo()
  Dim sht As Worksheet
   Me.EnableEvents = False
   
   cbSelectActiveSheet.Clear
   'For Each sht In ActiveWorkbook.Worksheets
   For Each sht In Workbooks(cbSelActiveWB.Value).Worksheets
     cbSelectActiveSheet.AddItem sht.Name
   Next
   cbSelectActiveSheet.Value = ActiveSheet.Name
   
   Me.EnableEvents = True
   
End Sub


Private Sub btn01_Click()
  InitWBcombo
End Sub

Private Sub InitWBcombo()
Dim wb As Workbook
    
  cbSelActiveWB.Clear
  For Each wb In Application.Workbooks
    cbSelActiveWB.AddItem wb.Name
  Next
  cbSelActiveWB.Value = ActiveWorkbook.Name
End Sub

Private Sub btnCloseForm_Click()
  'unload B2AddinMain
  Unload Me
End Sub


'***********************************************************
'WorkBook
'***********************************************************


Private Sub btnExportPNG_Click()
  ExportPNG
End Sub

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

















'***********************************************************
'JMR Compmacro
'***********************************************************

Private Sub EnableMacroButtons(enbl As Boolean)
   btnAddDataSet.Enabled = enbl
  btnStartMac.Enabled = enbl
  btnRemoveDataset.Enabled = enbl
  btnAddChart.Enabled = enbl
End Sub

Private Sub btnOpenCompMac_Click()
  Dim wb As Workbook
  Dim wbWin As Window
  Dim i As Integer
  
  Application.ScreenUpdating = False
     
  For i = 1 To Application.Windows.Count
    'MsgBox Application.Windows(i).Caption & " " & COMPMACROFILE ' Debug
    If (Application.Windows(i).Caption = COMPMACROFILE) Then
      Application.Windows(i).WindowState = xlMinimized
      'Application.Windows(i).Visible = False ' more radical option to line above
    End If
    i = i + 1
   Next
  
  EnableMacroButtons (True)
     
  Application.ScreenUpdating = True
    
End Sub

Private Sub btnCloseCompMac_Click()
  On Error GoTo Errh:
  Application.Workbooks(COMPMACROFILE).Close SaveChanges:=True
  EnableMacroButtons (False)
Exit Sub
Errh:
  If Err.Number <> 9 Then ' 9 file already closed
   MsgBox Err.Description & " (" & Err.Number & ")"
  End If
  
End Sub

Private Sub btnStartMac_Click()
  'Application.Run "JMR_Comparison_Macro_B2_V002.xlsm!Start_Macro"
  Application.Run COMPMACROFILE & "!Start_Macro"
  btnStartMac.ControlTipText = COMPMACROFILE 'Works??????
End Sub

Private Sub btnAddDataSet_Click()
  Application.Run COMPMACROFILE & "!Start_AddDataset"
End Sub


Private Sub btnRemoveDataset_Click()
    Application.Run COMPMACROFILE & "!Start_RemoveDataset"
End Sub

Private Sub btnAddChart_Click()
  Application.Run COMPMACROFILE & "!Start_AddChart"
End Sub
'===========================================================




'***********************************************************
'Chart
'***********************************************************





Private Sub btnChartExpPNG_Click()
  ChartExpPNG
  
End Sub

Private Sub btnClustColFormat_Click()
  SetCol
  
End Sub




Private Sub btnExportSourceFiles_Click()
  VBA2VersionControl.ExportSourceFiles
End Sub

Private Sub btnFont_Click()
  mChartSelected.SelectChartFont
End Sub

Private Sub btnImportSourceFile_Click()
  MsgBox "to do"
  
End Sub



Private Sub btnLine_Click()
 mChartSelected.SetLine

End Sub

Private Sub btnLineColor_Click()
  mChartSelected.SetLineColor
End Sub

Private Sub btnLineW_Click()
  mChartSelected.LineW

End Sub





Private Sub btnScale_Click()
  mChartSelected.SetY1Scale
  
End Sub







