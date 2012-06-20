VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim vName As String
Dim colRecs As New Collection

Public Property Get Name() As String
  Name = vName
End Property

Public Property Let Name(vNewValue As String)
  vName = vNewValue
End Property

Public Function AddRecord(Key As String, _
                 Optional before As String = "", _
                 Optional after As String = "") As clsRecord
           
  Dim cRec As New clsRecord
  
  cRec.Key = Key
  
  If before <> "" Then
    colRecs.Add cRec, Key, before
  
  ElseIf after <> "" Then
    colRecs.Add cRec, Key, , after
  
  Else
    colRecs.Add cRec, Key
  
  End If
  
  cRec.Parent = Me
  Set AddRecord = cRec
    
End Function

Public Function RecordCount() As Long
  RecordCount = colRecs.Count
End Function

Public Sub RemoveRecord(ByVal Index As Variant)
  colRecs.Remove Index
End Sub

Public Function RecordItem(ByVal Index As Variant) As clsRecord
Attribute RecordItem.VB_UserMemId = 0
  Dim cRec As clsRecord
  
  Set RecordItem = Nothing
  
  For Each cRec In colRecs
    If cRec.Key = Index Then
      Set RecordItem = colRecs.Item(Index)
      Exit Function
    End If
    
  Next cRec

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colRecs.[_NewEnum]
End Function

