VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsETL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim colDBs As New Collection

Public Function AddDatabase(Key As String, _
           Optional before As String = "", _
           Optional after As String = "") As clsDatabase
           
  Dim cDB As New clsDatabase
  
  cDB.Name = Key
  
  If before <> "" Then
    colDBs.Add cDB, Key, before
  
  ElseIf after <> "" Then
    colDBs.Add cDB, Key, , after
  
  Else
    colDBs.Add cDB, Key
  
  End If
  
  Set AddDatabase = cDB
    
End Function

Public Function DatabaseCount() As Long
  DatabaseCount = colDBs.Count
End Function

Public Sub RemoveDatabase(ByVal Index As Variant)
  colDBs.Remove Index
End Sub

Public Function DatabaseItem(ByVal Index As Variant) As clsDatabase
Attribute DatabaseItem.VB_UserMemId = 0
  Dim cDB As clsDatabase
  
  Set DatabaseItem = Nothing
  
  For Each cDB In colDBs
    If cDB.Name = Index Then
      Set DatabaseItem = colDBs.Item(Index)
      Exit Function
    End If
    
  Next cDB

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colDBs.[_NewEnum]
End Function


