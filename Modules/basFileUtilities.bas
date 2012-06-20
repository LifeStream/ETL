Attribute VB_Name = "basFileUtilities"
Option Explicit

Public Function fileExists(strFile As String) As Boolean
  Dim fs As New FileSystemObject
  'Set fs = CreateObject("Scripting.FileSystemObject")
  
  fileExists = fs.fileExists(strFile)
  
  Set fs = Nothing
End Function

Public Function getFileList(Optional dirPath As String = "C:\") As Files
  Dim fs As New FileSystemObject
  
  Set getFileList = fs.GetFolder(dirPath).Files
  
  Set fs = Nothing
End Function

Public Function dirExists(strDir As String, _
                          Optional createNewDir As Boolean = False) As Boolean
  Dim fs As New FileSystemObject
  'Set fs = CreateObject("Scripting.FileSystemObject")
  
  dirExists = fs.FolderExists(strDir)
  
  If Not (dirExists) And createNewDir Then
    MkDir strDir
  End If
  
  Set fs = Nothing
End Function

Public Function loadFileStream(strFilePath As String) As String
  Dim ioStream As String
  Dim ioBuffer As String * 1
  Dim ioNum As Long
  
  ioNum = FreeFile
  ioStream = ""
  ioBuffer = ""
    
  If fileExists(strFilePath) Then
    Open strFilePath For Binary As #ioNum
    
    Do
      Get #ioNum, , ioBuffer
      ioStream = ioStream & ioBuffer
    Loop Until EOF(ioNum)
    
    Close #ioNum
  End If
  
  loadFileStream = ioStream
End Function

Public Sub reportStatus(svReport As String)
  frmMain.lblStatus.Caption = svReport
End Sub
