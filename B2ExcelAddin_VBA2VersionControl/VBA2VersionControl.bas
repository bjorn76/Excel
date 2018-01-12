Attribute VB_Name = "VBA2VersionControl"
Option Explicit

' This module enables all other modules in current VBA projekt to be
' imported and imported to a folder automatically. Then the files could be
' version control with a system free of choice such as SVN GIT etc.

'To make the VBE automation in this module to work you need to:
' 1) Enable Reference to:  Microsoft Visual Basic for Applications Extensibility 5.3
'    (Tools reference... from VBA editor)
' 2) Enable "Trust access to the VBA project object model"
'    (Fore Excel 2010 found in: <File> <options> <Trust center> <Trust center settings...> <Macro settings)
' 3) Import this module to the VBAProject you want to version control


'To read more:
' http://www.cpearson.com/excel/vbe.aspx


Public Const ThisModule = "VBA2VersionControl" 'used to exclude form import export
Public Const SCprefix As String = "B2" ' SourceControlprefix, Only Addins or workbooks prefixed with these letters will be exported/imported to sub folder

Public Sub ExportSourceFiles()
    Dim objVBproj As VBProject
    Dim pf As String, i As Integer
        
    pf = SCprefix
    i = 0
    For Each objVBproj In Application.VBE.VBProjects
      i = i + 1
      If Left(objVBproj.Name, Len(pf)) = pf Then
        MsgBox (sourcepath(objVBproj)) ' Debug
        Call ExportSourceFilesTo(sourcepath(objVBproj), i)
      End If
    Next
    
    
    
End Sub

Public Sub ImportSourceFiles()
    Dim objVBproj As VBProject
    Dim i As Integer, pf As String
    
    pf = SCprefix
    i = 0
    If MsgBox("Import VBA modules? Will write", vbOKCancel, "Import from file") = 2 Then
      Exit Sub
    End If
    
    
    
    For Each objVBproj In Application.VBE.VBProjects
      i = i + 1
      
      If Left(objVBproj.Name, Len(pf)) = pf Then
        'MsgBox (sourcepath(objVBproj)) ' Debug
        RemoveAllModules (i) '(objVBproj)
        Call ImportSourceFilesFrom(sourcepath(objVBproj), i)
      End If
    Next
      
End Sub
    
Public Sub SPmessage() ' Debugg func
   
  
    Dim objVBproj As VBProject
    For Each objVBproj In Application.VBE.VBProjects
      If Left(objVBproj.Name, 2) = SCprefix Then
        MsgBox (sourcepath(objVBproj))
      End If
    Next
    
  
End Sub


Private Function sourcepath(objVBproj As VBProject) As String
    
    Dim sPath As String 'path and file name
    Dim sFilename, sFileNoExt, sFoldername As String
   
    
    
   
    'sPath = Application.ActiveWorkbook.FullName ' Debug
    'sPath = Application.VBE.ActiveVBProject.Filename
    sPath = objVBproj.Filename
    

    

    
    
    'sPath = "B2ExcelAddin.xlam"
    
    'MsgBox sPath 'Debug
    
    sFilename = Mid(Mid(sPath, InStrRev(sPath, "/") + 1), InStrRev(sPath, "\") + 1)
    'MsgBox sFilename 'Debug
    
    sFileNoExt = Mid(sFilename, 1, InStrRev(sFilename, ".") - 1)
    ' MsgBox sFileNoExt 'Debug
    
    sFoldername = Left(sPath, Len(sPath) - Len(sFilename))
    ' MsgBox sFoldername Debug
    
    sourcepath = (sFoldername & sFileNoExt & "_" & ThisModule & "\")
    
End Function



Private Sub ExportSourceFilesTo(destPath As String, i As Integer)
 
  Dim component As VBComponent

    'Make folder if it doesn't exist
    If Dir(destPath, vbDirectory) = "" Then
        MkDir destPath
        'MsgBox "Folder " & destPath & " created"
    Else
        'MsgBox "Folder exist" 'Debug
    End If


'For Each component In Application.VBE.ActiveVBProject.VBComponents
For Each component In Application.VBE.VBProjects(i).VBComponents
    If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Then
        'component.Export (destPath &amp; component.Name &amp; ToFileExtension(component.Type))
        'component.Export (destPath)
        component.Export (destPath & component.Name & ToFileExtension(component.Type))
    End If
Next
 
End Sub
 
Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String
    Select Case vbeComponentType
    Case vbext_ComponentType.vbext_ct_ClassModule
      ToFileExtension = ".cls"
    Case vbext_ComponentType.vbext_ct_StdModule
      ToFileExtension = ".bas"
    Case vbext_ComponentType.vbext_ct_MSForm
      ToFileExtension = ".frm"
    Case vbext_ComponentType.vbext_ct_ActiveXDesigner
      Case vbext_ComponentType.vbext_ct_Document
    Case Else
      ToFileExtension = vbNullString
    End Select

End Function

Public Sub RAM() 'Debug func
  RemoveAllModules (2) '(Application.VBE.projects(2))
End Sub

Private Sub RemoveAllModules(i As Integer)  '(ByRef prj As VBProject)
Dim project As VBProject
Dim comp As VBComponent

Set project = Application.VBE.VBProjects(i)

For Each comp In project.VBComponents
  If Not comp.Name = ThisModule And (comp.Type = vbext_ct_ClassModule Or comp.Type = vbext_ct_StdModule) Then
     project.VBComponents.Remove comp
     'MsgBox comp.Name
  End If
Next
End Sub




Private Sub ImportSourceFilesFrom(sourcepath As String, pIndex As Integer)
Dim file As String
file = Dir(sourcepath)
    While (file <> vbNullString)
      If Not file = (ThisModule & ".bas") Then
        Application.VBE.VBProjects(pIndex).VBComponents.Import sourcepath & file
      End If
      file = Dir
    Wend
End Sub



