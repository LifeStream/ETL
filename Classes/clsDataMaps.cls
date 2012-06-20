VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataMaps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim colDataMaps As New Collection

Public Function Add(Key As String, _
           Optional before As String = "", _
           Optional after As String = "") As clsDataMap
           
  Dim cDataMap As New clsDataMap
  
  cDataMap.Key = Key
  
  If before <> "" Then
    colDataMaps.Add cDataMap, Key, before
  
  ElseIf after <> "" Then
    colDataMaps.Add cDataMap, Key, , after
  
  Else
    colDataMaps.Add cDataMap, Key
  
  End If
  
  Set Add = cDataMap
    
End Function

Public Function Count() As Long
  Count = colDataMaps.Count
End Function

Public Sub Remove(ByVal Index As Variant)
  colDataMaps.Remove Index
End Sub

Public Function Item(ByVal Index As Variant) As clsDataMap
Attribute Item.VB_UserMemId = 0
  Dim cDataMap As clsDataMap
  
  Set Item = Nothing
  
  For Each cDataMap In colDataMaps
    If cDataMap.Key = Index Then
      Set Item = colDataMaps.Item(Index)
      Exit Function
    End If
    
  Next cDataMap

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colDataMaps.[_NewEnum]
End Function

Public Function reMap(cSyncField As clsSyncField) As String
  Dim cDataMap As clsDataMap
    
  Set cDataMap = colDataMaps(cSyncField.DataMap)

  reMap = cDataMap.findMappedValue(cSyncField)

End Function
