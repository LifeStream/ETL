VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSyncRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim vDbQuery As String
Dim vDbUpdate As String
Dim vTableQuery As String
Dim vTableUpdate As String
Dim colSyncFields As New Collection

Public Property Get DatabaseQuery() As String
  DatabaseQuery = vDbQuery
End Property

Public Property Let DatabaseQuery(ByVal vNewValue As String)
  vDbQuery = vNewValue
End Property

Public Property Get DatabaseUpdate() As String
  DatabaseUpdate = vDbUpdate
End Property

Public Property Let DatabaseUpdate(ByVal vNewValue As String)
  vDbUpdate = vNewValue
End Property

Public Property Get TableQuery() As String
  TableQuery = vTableQuery
End Property

Public Property Let TableQuery(ByVal vNewValue As String)
  vTableQuery = vNewValue
End Property

Public Property Get TableUpdate() As String
  TableUpdate = vTableUpdate
End Property

Public Property Let TableUpdate(ByVal vNewValue As String)
  vTableUpdate = vNewValue
End Property

Public Function AddSyncField(Key As String, _
                    Optional before As String = "", _
                    Optional after As String = "") As clsSyncField
           
  Dim cSyncField As New clsSyncField
  
  cSyncField.KeyID = Key
  
  'Defaults
  cSyncField.PrimaryKey = False
  cSyncField.Updated_TS = False
  
  
  If before <> "" Then
    colSyncFields.Add cSyncField, Key, before
  
  ElseIf after <> "" Then
    colSyncFields.Add cSyncField, Key, , after
  
  Else
    colSyncFields.Add cSyncField, Key
  
  End If
  
  Set AddSyncField = cSyncField
    
End Function

Public Function AddSyncFieldObject(obj As clsSyncField, _
                          Optional before As String = "", _
                          Optional after As String = "") As clsSyncField
             
  
  If before <> "" Then
    colSyncFields.Add obj, obj.KeyID, before
  
  ElseIf after <> "" Then
    colSyncFields.Add obj, obj.KeyID, , after
  
  Else
    colSyncFields.Add obj, obj.KeyID
  
  End If
  
  Set AddSyncFieldObject = obj
    
End Function

Public Function SyncFieldCount() As Long
  SyncFieldCount = colSyncFields.Count
End Function

Public Sub RemoveSyncField(ByVal Index As Variant)
  colSyncFields.Remove Index
End Sub

Public Function SyncFieldItem(ByVal Index As Variant) As clsSyncField
Attribute SyncFieldItem.VB_UserMemId = 0
  Dim cSyncField As clsSyncField
  
  Set SyncFieldItem = Nothing
  
  For Each cSyncField In colSyncFields
    If cSyncField.KeyID = Index Then
      Set SyncFieldItem = colSyncFields.Item(Index)
      Exit Function
    End If
    
  Next cSyncField

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colSyncFields.[_NewEnum]
End Function

Public Sub Map(lRec As clsRecord)
  Dim cSyncField As clsSyncField
  Dim cField As clsField
  
  vTableUpdate = lRec.Table
  vDbUpdate = lRec.Parent.Name
  
  For Each cSyncField In colSyncFields
    Set cField = lRec(cSyncField.KeyID)
    
    If cField Is Nothing Then
      colSyncFields.Remove cSyncField.KeyID
      
    Else
      With cSyncField
        .NameUpdate = cField.Name
        .DataEncapsulationUpdate = cField.DataEncapsulation
      End With
      
    End If
    
  Next cSyncField

End Sub

Public Sub LoadValues(rsQuery As ADODB.Recordset)
  Dim cSyncField As clsSyncField
  
  Reset
    
  For Each cSyncField In colSyncFields
    cSyncField.Value = Nz(rsQuery(cSyncField.NameQuery))
  Next cSyncField
  
End Sub

Public Sub Reset()
  Dim objItem As clsSyncField
    
  For Each objItem In colSyncFields
    objItem.Value = ""
  Next objItem

End Sub

Public Function getPrimaryKey() As clsSyncField
  Dim sField As clsSyncField

  Set getPrimaryKey = Nothing

  For Each sField In colSyncFields
    If sField.PrimaryKey Then
      Set getPrimaryKey = sField
      Exit Function
    End If
    
  Next sField
  
End Function

Public Function getUpdated_TS() As clsSyncField
  Dim sField As clsSyncField

  Set getUpdated_TS = Nothing

  For Each sField In colSyncFields
    If sField.Updated_TS Then
      Set getUpdated_TS = sField
      Exit Function
    End If
    
  Next sField
  
End Function


