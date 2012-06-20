VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim colSettings As New Collection

Public Sub loadXML(odbcNode As MSXML2.IXMLDOMNode)
  Dim xNode As MSXML2.IXMLDOMNode
  Dim cSetting As clsSetting
  
  For Each xNode In odbcNode.childNodes
    If xNode.nodeType = NODE_ELEMENT Then
      Set cSetting = Add(xNode.nodeName)
      cSetting = xNode.nodeTypedValue
    End If
  Next xNode
  
End Sub

Public Function Add(Key As String, _
           Optional before As String = "", _
           Optional after As String = "") As clsSetting
           
  Dim cSetting As New clsSetting
  
  cSetting.Name = Key
  
  If before <> "" Then
    colSettings.Add cSetting, Key, before
  
  ElseIf after <> "" Then
    colSettings.Add cSetting, Key, , after
  
  Else
    colSettings.Add cSetting, Key
  
  End If
  
  Set Add = cSetting
    
End Function

Public Function Count() As Long
  Count = colSettings.Count
End Function

Public Sub Remove(ByVal Index As Variant)
  colSettings.Remove Index
End Sub

Public Function Item(ByVal Index As Variant) As clsSetting
Attribute Item.VB_UserMemId = 0
  Dim cSetting As clsSetting
  
  Set Item = Nothing
  
  For Each cSetting In colSettings
    If cSetting.Name = Index Then
      Set Item = colSettings.Item(Index)
      Exit Function
    End If
    
  Next cSetting

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colSettings.[_NewEnum]
End Function

