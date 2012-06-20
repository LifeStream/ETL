VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataMapSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim vKey As String
Dim colDataSets As New Collection

Public Property Get Key() As String
  Key = vKey
End Property

Public Property Let Key(ByVal vNewValue As String)
  vKey = vNewValue
End Property

Public Function AddDataSet(Key As String, _
                  Optional before As String = "", _
                  Optional after As String = "") As clsDataSet
           
  Dim cDataSet As New clsDataSet
  
  cDataSet.Database = Key
  
  If before <> "" Then
    colDataSets.Add cDataSet, Key, before
  
  ElseIf after <> "" Then
    colDataSets.Add cDataSet, Key, , after
  
  Else
    colDataSets.Add cDataSet, Key
  
  End If
  
  Set AddDataSet = cDataSet
    
End Function

Public Function DataSetCount() As Long
  DataSetCount = colDataSets.Count
End Function

Public Sub RemoveDataSet(ByVal Index As Variant)
  colDataSets.Remove Index
End Sub

Public Function DataSetItem(ByVal Index As Variant) As clsDataSet
  Dim cDataSet As clsDataSet
  
  Set DataSetItem = Nothing
  
  For Each cDataSet In colDataSets
    If cDataSet.Database = Index Then
      Set DataSetItem = colDataSets.Item(Index)
      Exit Function
    End If
    
  Next cDataSet

End Function

Public Function NewEnum() As IUnknown
  Set NewEnum = colDataSets.[_NewEnum]
End Function




