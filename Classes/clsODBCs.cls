VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODBCs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim colODBCs As New Collection

Public Sub loadXML(odbcNode As MSXML2.IXMLDOMNode)
  Dim xNode As MSXML2.IXMLDOMNode
  Dim sNode As MSXML2.IXMLDOMNode
  Dim cODBC As clsODBC
  
  For Each xNode In odbcNode.childNodes
    If xNode.nodeType = NODE_ELEMENT Then
      Set cODBC = Add(xNode.nodeName)
      
      For Each sNode In xNode.childNodes
        Select Case sNode.nodeName
          Case "db_type"
            cODBC.DB_Type = sNode.nodeTypedValue
            
          Case "encapsulation"
            loadEncapsulation sNode, cODBC
            
          Case "connection"
            cODBC.Connection = createConnection(sNode)
            
        End Select
      Next sNode
            
    End If
  Next xNode
  
End Sub

Private Sub loadEncapsulation(sNode As MSXML2.IXMLDOMNode, cODBC As clsODBC)
  Dim eNode As MSXML2.IXMLDOMNode
  
  For Each eNode In sNode.childNodes
    cODBC.AddEncapsulation eNode.nodeName, eNode.nodeTypedValue
  Next eNode
  
End Sub

Public Function Add(Key As String, _
           Optional before As String = "", _
           Optional after As String = "") As clsODBC
           
  Dim cODBC As New clsODBC
  
  cODBC.Name = Key
  
  If before <> "" Then
    colODBCs.Add cODBC, Key, before
  
  ElseIf after <> "" Then
    colODBCs.Add cODBC, Key, , after
  
  Else
    colODBCs.Add cODBC, Key
  
  End If
  
  Set Add = cODBC
    
End Function

Public Function Count() As Long
  Count = colODBCs.Count
End Function

Public Sub Remove(ByVal Index As Variant)
  colODBCs.Remove Index
End Sub

Public Function Item(ByVal Index As Variant) As clsODBC
Attribute Item.VB_UserMemId = 0
  Dim cODBC As clsODBC
  
  Set Item = Nothing
  
  For Each cODBC In colODBCs
    If cODBC.Name = Index Then
      Set Item = colODBCs.Item(Index)
      Exit Function
    End If
    
  Next cODBC

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colODBCs.[_NewEnum]
End Function

Private Function createConnection(cnNode As MSXML2.IXMLDOMNode) As String
  Dim xNode As MSXML2.IXMLDOMNode
  Dim xAttr As MSXML2.IXMLDOMAttribute
  Dim svODBC As String
  Dim svName As String
  
  svODBC = ""
  
  For Each xNode In cnNode.childNodes
    If xNode.nodeType = NODE_ELEMENT Then
    
      Set xAttr = GetAttribute("name", xNode)
              
      If xAttr Is Nothing Then
        svName = xNode.nodeName
        
      Else
        svName = xAttr.nodeTypedValue
        
      End If
      
      svODBC = svODBC & svName & "=" & xNode.nodeTypedValue & ";"
      
      
    End If
  Next xNode

  createConnection = svODBC
End Function
