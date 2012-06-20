Attribute VB_Name = "xmlTools"
Option Explicit

Public Function loadXmlFile(strFilePath As String, _
                   Optional defaultXML As String = "") As MSXML2.DOMDocument60
                            
  'Dim xmlFile As New MSXML2.DOMDocument40
  Dim xmlFile As New MSXML2.DOMDocument60
  
  'Dim objXML  As New MSXML2.DOMDocument26
    
  If fileExists(strFilePath) Then
    xmlFile.loadXML loadFileStream(strFilePath)
    
  Else
    xmlFile.loadXML defaultXML
    
  End If

  Set loadXmlFile = xmlFile
End Function

Function GetNode(nodeName As String, parentNode As MSXML2.IXMLDOMNode) As MSXML2.IXMLDOMNode
  Dim xNode As MSXML2.IXMLDOMNode

  Set GetNode = Nothing

  For Each xNode In parentNode.childNodes
    If xNode.nodeName = nodeName Then
      Set GetNode = xNode
      Exit Function
    End If
  Next xNode

End Function

Function GetAttribute(attributeName As String, parentNode As MSXML2.IXMLDOMNode) As MSXML2.IXMLDOMAttribute
  Dim xAttr As MSXML2.IXMLDOMAttribute

  Set GetAttribute = Nothing

  For Each xAttr In parentNode.Attributes
    If xAttr.nodeName = attributeName Then
      Set GetAttribute = xAttr
      Exit Function
    End If
  Next xAttr

End Function

Public Function GetChildNodes(node As MSXML2.IXMLDOMNode) As MSXML2.IXMLDOMNode
' Return all the child nodes of a given node
' returns child nodes for a given MSXML2.IXMLDOMNode
 Set GetChildNodes = node.childNodes
End Function


Function GetRootNode(xmlDoc As Object) As MSXML2.IXMLDOMNode
' Get the root node of a XML Document
' returns root node
 Set GetRootNode = xmlDoc.documentElement
End Function


Function LoadError(xmlDoc As Object) As Boolean
' Check if a parsing error occurred when loading a XML document
' checks if a xml file load error occurred
 LoadError = (xmlDoc.parseError.errorCode <> 0)
End Function


