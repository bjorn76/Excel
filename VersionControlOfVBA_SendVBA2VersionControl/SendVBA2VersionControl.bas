Attribute VB_Name = "SendVBA2VersionControl"

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


Public Const ThisModule = "SendVBA2VersionControl" 'used to exclude form import export

Public Sub ExportSourceFiles()
    ExportSourceFilesTo (sourcepath())
End Sub

    
Private Function sourcepath() As String
    
    Dim sPath As String 'path and file name
    Dim sFilename, sFileNoExt, sFoldername As String
    
    sPath = Application.VBE.ActiveVBProject.Filename
    'MsgBox sPath
    
    sFilename = Mid(Mid(sPath, InStrRev(sPath, "/") + 1), InStrRev(sPath, "\") + 1)
    'MsgBox sFilename
    
    sFileNoExt = Mid(sFilename, 1, InStrRev(sFilename, ".") - 1)
    'MsgBox sFileNoExt
    
    sFoldername = Left(sPath, Len(sPath) - Len(sFilename))
    'MsgBox sFoldername
    
    sourcepath = (sFoldername & sFileNoExt & "_" & ThisModule & "\")
    
End Function





Private Sub ExportSourceFilesTo(destPath As String)
 
Dim component As VBComponent

'Make folder if it doesn't exist
If Dir(destPath, vbDirectory) = "" Then
    MkDir destPath
    MsgBox "Folder " & destPath & " created"
Else
    MsgBox "Folder exist"
End If


For Each component In Application.VBE.ActiveVBProject.VBComponents
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


Private Sub RemoveAllModules()
Dim project As VBProject
Set project = Application.VBE.ActiveVBProject
 
Dim comp As VBComponent
For Each comp In project.VBComponents
  If Not comp.Name = ThisModule And (comp.Type = vbext_ct_ClassModule Or comp.Type = vbext_ct_StdModule) Then
     project.VBComponents.Remove comp
     'MsgBox comp.Name
  End If
Next
End Sub


Public Sub ImportSourceFiles()
  RemoveAllModules
  ImportSourceFilesFrom (sourcepath())
End Sub

Private Sub ImportSourceFilesFrom(sourcepath As String)
Dim file As String
file = Dir(sourcepath)
    While (file <> vbNullString)
      If Not file = (ThisModule & ".bas") Then
        Application.VBE.ActiveVBProject.VBComponents.Import sourcepath & file
      End If
      file = Dir
    Wend
End Sub


