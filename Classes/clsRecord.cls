VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim vKey As String
Dim vTable As String
Dim vParent As clsDatabase
Dim colFields As New Collection

Public Property Get Key() As String
  Key = vKey
End Property

Public Property Let Key(ByVal vNewValue As String)
  vKey = vNewValue
End Property

Public Property Get Table() As String
  Table = vTable
End Property

Public Property Let Table(ByVal vNewValue As String)
  vTable = vNewValue
End Property

Public Property Get Parent() As clsDatabase
  Set Parent = vParent
End Property

Public Property Let Parent(ByVal vNewValue As clsDatabase)
  Set vParent = vNewValue
End Property

Public Function AddField(Key As String, _
                Optional before As String = "", _
                Optional after As String = "") As clsField
           
  Dim cField As New clsField
  
  cField.KeyID = Key
  
  'Defaults
  cField.PrimaryKey = False
  cField.Updated_TS = False
  
  
  If before <> "" Then
    colFields.Add cField, Key, before
  
  ElseIf after <> "" Then
    colFields.Add cField, Key, , after
  
  Else
    colFields.Add cField, Key
  
  End If
  
  cField.Parent = Me
  Set AddField = cField
    
End Function

Public Function FieldCount() As Long
  FieldCount = colFields.Count
End Function

Public Sub RemoveField(ByVal Index As Variant)
  colFields.Remove Index
End Sub

Public Function FieldItem(ByVal Index As Variant) As clsField
Attribute FieldItem.VB_UserMemId = 0
  Dim cField As clsField
  
  Set FieldItem = Nothing
  
  For Each cField In colFields
    If cField.KeyID = Index Then
      Set FieldItem = colFields.Item(Index)
      Exit Function
    End If
    
  Next cField

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colFields.[_NewEnum]
End Function

Public Function SyncClone() As clsSyncRecord
  Dim sClone As New clsSyncRecord
  Dim fClone As New clsSyncField
  Dim cField As clsField
  
  With sClone
    .DatabaseQuery = vParent.Name
    .TableQuery = vTable
      
    For Each cField In colFields
      Set fClone = .AddSyncFieldObject(cField.SyncClone)
      fClone.Parent = sClone
    Next cField
    
  End With
  
  Set SyncClone = sClone
End Function

