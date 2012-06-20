Attribute VB_Name = "basInitializerTerminator"
Option Explicit

Public Sub Main()
  'Prevents this Application from running multiple instances
  If App.PrevInstance Then Exit Sub
  
  With App
    appVersion = .Major & "." & .Minor & "." & .Revision
  End With
    
  'frmMain.loadForm
  Load frmSysTray
  
  setupConfig
  
  executeETL
End Sub

Public Sub subTerminator()
  'Not the mainTerminator, only subTerminator
  'Cannot end progam here. Must be ended by unloading the sysTray Form.
  ' See frmSysTray.mnuPopup_Click  - Case "Exit"
  
  Unload frmMain
  DoEvents
  
End Sub

Public Sub setupConfig()
  inProgressETL = True
  
  loadConfig App.Path & "\config.xml"
  loadETLs colSettings("data")
  
  inProgressETL = False
End Sub
