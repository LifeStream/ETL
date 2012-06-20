Attribute VB_Name = "xmlConfig"
Option Explicit

Public Sub loadConfig(svFileName As String)
  Dim xmlConfig As New MSXML2.DOMDocument60
  Dim xmlRoot As MSXML2.IXMLDOMNode
  
  Set xmlConfig = loadXmlFile(App.Path & "\config.xml")
  Set xmlRoot = xmlConfig.documentElement
  
  colODBCs.loadXML GetNode("odbc", xmlRoot)
  
  loadSettings xmlRoot
  
End Sub

Private Sub loadSettings(xmlRoot As MSXML2.IXMLDOMNode)
  Dim cSetting As clsSetting
  
  colSettings.loadXML GetNode("settings", xmlRoot)

  'Default queue time to 300 seconds (5 mins) if not present
  If colSettings("queue") Is Nothing Then
    Set cSetting = colSettings.Add("queue")
    cSetting = 300
  End If
  
  'Default data location to App.path \data if not present
  If colSettings("data") Is Nothing Then
    Set cSetting = colSettings.Add("data")
    cSetting = App.Path & "\data"
  End If
  
  Set cSetting = colSettings("data")
  
  If Mid(cSetting, 2, 2) <> ":\" Then
    cSetting = App.Path & "\" & cSetting
  End If
  
End Sub




