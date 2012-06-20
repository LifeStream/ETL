Attribute VB_Name = "xmlETL"
Option Explicit

Public Sub loadETLs(svFilePath As String)
  Dim fETL As File
  
  If dirExists(svFilePath) Then
    For Each fETL In getFileList(svFilePath)
      If Right(fETL.Name, 4) = ".xml" Then
        loadETL fETL.Path
      End If
    Next fETL
  End If
    
End Sub

Public Sub loadETL(svFileName As String)
  Dim xmlETL As New MSXML2.DOMDocument60
  Dim xmlRoot As MSXML2.IXMLDOMNode
      
  Set xmlETL = loadXmlFile(svFileName)
  Set xmlRoot = xmlETL.documentElement
  
  loadDatabases xmlRoot
  mapFields xmlRoot
  mapData xmlRoot

End Sub

Private Sub loadDatabases(xmlRoot As MSXML2.IXMLDOMNode)
  Dim rNode As MSXML2.IXMLDOMNode
  Dim dNode As MSXML2.IXMLDOMNode
  
  For Each rNode In xmlRoot.childNodes
    If rNode.nodeType = NODE_ELEMENT And rNode.nodeName = "record" Then
      For Each dNode In rNode.childNodes
        If dNode.nodeType = NODE_ELEMENT Then
        
          If colETLs(dNode.nodeName) Is Nothing Then colETLs.AddDatabase dNode.nodeName
          
        End If
      Next dNode
    
    End If
  Next rNode

End Sub

Private Sub mapFields(xmlRoot As MSXML2.IXMLDOMNode)
  Dim rNode As MSXML2.IXMLDOMNode
  Dim dNode As MSXML2.IXMLDOMNode
  Dim svGUID As String
  
  Dim cnDb As clsDatabase
  Dim cnRec As clsRecord
  
  For Each rNode In xmlRoot.childNodes
    If rNode.nodeType = NODE_ELEMENT And rNode.nodeName = "record" Then
      
      svGUID = GetGUID
    
      For Each dNode In rNode.childNodes
        If dNode.nodeType = NODE_ELEMENT Then
          Set cnDb = colETLs(dNode.nodeName)
          Set cnRec = loadMap(svGUID, cnDb, dNode)
          
        End If
          
      Next dNode
    
    End If
  Next rNode
  
End Sub

Private Function loadMap(svGUID As String, cDB As clsDatabase, rNode As MSXML2.IXMLDOMNode) As clsRecord
  Dim dNode As MSXML2.IXMLDOMNode
  Dim xAttr As MSXML2.IXMLDOMAttribute
  
  Dim cRec As clsRecord
  Dim cField As clsField
  
  Set cRec = cDB.AddRecord(svGUID)
  
  
  For Each dNode In rNode.childNodes
    If dNode.nodeType = NODE_ELEMENT Then
      Select Case dNode.nodeName
        Case "table"
          cRec.Table = dNode.nodeTypedValue
          
        Case "map"
          Set xAttr = GetAttribute("id", dNode)
                  
          If Not (xAttr Is Nothing) Then
            Set cField = cRec.AddField(xAttr.nodeTypedValue)
            
            cField.Name = dNode.nodeTypedValue
            
            Set xAttr = GetAttribute("type", dNode)
            If Not (xAttr Is Nothing) Then
              Select Case LCase(xAttr.nodeTypedValue)
                Case "pri_key"
                  cField.PrimaryKey = True
                  
                Case "ts_updated"
                  cField.Updated_TS = True
                  
              End Select
            End If
            
            Set xAttr = GetAttribute("dataEnc", dNode)
            If Not (xAttr Is Nothing) Then
              cField.DataEncapsulation = xAttr.nodeTypedValue
            End If
            
            Set xAttr = GetAttribute("dataMap", dNode)
            If Not (xAttr Is Nothing) Then
              cField.DataMap = xAttr.nodeTypedValue
            End If
            
          End If
        
      End Select
      
    End If
      
  Next dNode
  
End Function

Private Sub mapData(xmlRoot As MSXML2.IXMLDOMNode)
  Dim rNode As MSXML2.IXMLDOMNode
  Dim dNode As MSXML2.IXMLDOMNode
  
  'Dim sGUID As String
  
  Dim cDataMap As clsDataMap
  
  For Each rNode In xmlRoot.childNodes
    If rNode.nodeType = NODE_ELEMENT And rNode.nodeName = "dataMap" Then
    
      For Each dNode In rNode.childNodes
        If dNode.nodeType = NODE_ELEMENT Then
          loadDataMap dNode
                  
        End If
          
      Next dNode
    
    End If
  Next rNode
  
End Sub

Private Function loadDataMap(dNode As MSXML2.IXMLDOMNode) As clsDataMap
  Dim sNode As MSXML2.IXMLDOMNode
  Dim mAttr As MSXML2.IXMLDOMAttribute
  
  
  Dim cDataMap As clsDataMap
  Dim cDataMapSet As clsDataMapSet
  Dim cDataSet As clsDataSet
  
  Set cDataMap = colDataMaps.Add(dNode.nodeName)
  
  For Each sNode In dNode.childNodes
    Set cDataMapSet = cDataMap.AddDataMapSet(GetGUID)
    
    For Each mAttr In sNode.Attributes
      Set cDataSet = cDataMapSet.AddDataSet(mAttr.nodeName)
      cDataSet.Value = mAttr.nodeTypedValue
      
    Next mAttr
    
  Next sNode
  
End Function
