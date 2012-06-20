VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim vKey As String
Dim colDataMapSets As New Collection

Public Property Get Key() As String
  Key = vKey
End Property

Public Property Let Key(ByVal vNewValue As String)
  vKey = vNewValue
End Property

Public Function AddDataMapSet(Key As String, _
                     Optional before As String = "", _
                     Optional after As String = "") As clsDataMapSet
           
  Dim cDataMapSet As New clsDataMapSet
  
  cDataMapSet.Key = Key
  
  If before <> "" Then
    colDataMapSets.Add cDataMapSet, Key, before
  
  ElseIf after <> "" Then
    colDataMapSets.Add cDataMapSet, Key, , after
  
  Else
    colDataMapSets.Add cDataMapSet, Key
  
  End If
  
  Set AddDataMapSet = cDataMapSet
    
End Function

Public Function DataMapSetCount() As Long
  DataMapSetCount = colDataMapSets.Count
End Function

Public Sub RemoveDataMapSet(ByVal Index As Variant)
  colDataMapSets.Remove Index
End Sub

Public Function DataMapSetItem(ByVal Index As Variant) As clsDataMapSet
  Dim cDataMapSet As clsDataMapSet
  
  Set DataMapSetItem = Nothing
  
  For Each cDataMapSet In colDataMapSets
    If cDataMapSet.Key = Index Then
      Set DataMapSetItem = colDataMapSets.Item(Index)
      Exit Function
    End If
    
  Next cDataMapSet

End Function

Public Function NewEnum() As IUnknown
  Set NewEnum = colDataMapSets.[_NewEnum]
End Function

Public Function findMappedValue(cSyncField As clsSyncField) As String
  Dim cDataMapSet As clsDataMapSet
  Dim cDataSet As clsDataSet
  
  findMappedValue = ""
  
  For Each cDataMapSet In colDataMapSets
    Set cDataSet = cDataMapSet.DataSetItem(cSyncField.Parent.DatabaseQuery)
    
    If cDataSet.Value = cSyncField.Value Then
      Set cDataSet = cDataMapSet.DataSetItem(cSyncField.Parent.DatabaseUpdate)
      
      findMappedValue = cDataSet.Value
      Exit Function
      
    End If
    
  Next cDataMapSet
  
End Function
