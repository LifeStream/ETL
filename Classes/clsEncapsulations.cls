VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEncapsulations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim colEncapulations As New Collection

Public Function Add(Key As String, _
           Optional before As String = "", _
           Optional after As String = "") As clsEncapsulation
           
  Dim cEncap As New clsEncapsulation
  
  cEncap.Name = Key
  
  If before <> "" Then
    colEncapulations.Add cEncap, Key, before
  
  ElseIf after <> "" Then
    colEncapulations.Add cEncap, Key, , after
  
  Else
    colEncapulations.Add cEncap, Key
  
  End If
  
  Set Add = cEncap
    
End Function

Public Function Count() As Long
  Count = colEncapulations.Count
End Function

Public Sub Remove(ByVal Index As Variant)
  colEncapulations.Remove Index
End Sub

Public Function Item(ByVal Index As Variant) As clsEncapsulation
Attribute Item.VB_UserMemId = 0
  Dim cEncap As clsEncapsulation
  
  Set Item = Nothing
  
  For Each cEncap In colEncapulations
    If cEncap.Name = Index Then
      Set Item = colEncapulations.Item(Index)
      Exit Function
    End If
    
  Next cEncap

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = 0
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colEncapulations.[_NewEnum]
End Function


