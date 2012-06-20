VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim vKeyID As String
Dim vName As String
Dim vDataEnc As String
Dim vDataMap As String
Dim vPriKey As Boolean
Dim vUpdatedTS As Boolean
Dim vParent As clsRecord

Public Property Get KeyID() As String
Attribute KeyID.VB_UserMemId = 0
  KeyID = vKeyID
End Property

Public Property Let KeyID(ByVal vNewValue As String)
  vKeyID = vNewValue
End Property

Public Property Get Name() As String
  Name = vName
End Property

Public Property Let Name(ByVal vNewValue As String)
  vName = vNewValue
End Property

Public Property Get DataEncapsulation() As String
  DataEncapsulation = vDataEnc
End Property

Public Property Let DataEncapsulation(ByVal vNewValue As String)
  vDataEnc = vNewValue
End Property

Public Property Get DataMap() As String
  DataMap = vDataMap
End Property

Public Property Let DataMap(ByVal vNewValue As String)
  vDataMap = vNewValue
End Property


Public Property Get PrimaryKey() As Boolean
  PrimaryKey = vPriKey
End Property

Public Property Let PrimaryKey(ByVal vNewValue As Boolean)
  vPriKey = vNewValue
End Property

Public Property Get Updated_TS() As Boolean
  Updated_TS = vUpdatedTS
End Property

Public Property Let Updated_TS(ByVal vNewValue As Boolean)
  vUpdatedTS = vNewValue
End Property

Public Property Get Parent() As clsRecord
  Set Parent = vParent
End Property

Public Property Let Parent(ByVal vNewValue As clsRecord)
  Set vParent = vNewValue
End Property

Public Function SyncClone() As clsSyncField
  Dim cSyncField As New clsSyncField
  
  With cSyncField
    .KeyID = vKeyID
    .NameQuery = vName
    .DataMap = vDataMap
    .PrimaryKey = vPriKey
    .Updated_TS = vUpdatedTS
  End With
  
  Set SyncClone = cSyncField
End Function

