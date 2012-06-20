VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSyncField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim vKeyID As String
Dim vNameQuery As String
Dim vNameUpdate As String
Dim vDataEncUpdate As String
Dim vDataMap As String
Dim vPriKey As Boolean
Dim vUpdatedTS As Boolean
Dim vValue As String
Dim vParent As clsSyncRecord

Public Property Get KeyID() As String
  KeyID = vKeyID
End Property

Public Property Let KeyID(ByVal vNewValue As String)
  vKeyID = vNewValue
End Property

Public Property Get NameQuery() As String
  NameQuery = vNameQuery
End Property

Public Property Let NameQuery(ByVal vNewValue As String)
  vNameQuery = vNewValue
End Property

Public Property Get NameUpdate() As String
  NameUpdate = vNameUpdate
End Property

Public Property Let NameUpdate(ByVal vNewValue As String)
  vNameUpdate = vNewValue
End Property

Public Property Get DataEncapsulationUpdate() As String
  DataEncapsulationUpdate = vDataEncUpdate
End Property

Public Property Let DataEncapsulationUpdate(ByVal vNewValue As String)
  vDataEncUpdate = vNewValue
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

Public Property Get Value() As String
  Value = vValue
End Property

Public Property Let Value(ByVal vNewValue As String)
  vValue = vNewValue
End Property

Public Property Get Parent() As clsSyncRecord
  Set Parent = vParent
End Property

Public Property Let Parent(ByVal vNewValue As clsSyncRecord)
  Set vParent = vNewValue
End Property



