VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim vDatabase As String
Dim vValue As String

Public Property Get Database() As String
  Database = vDatabase
End Property

Public Property Let Database(ByVal vNewValue As String)
  vDatabase = vNewValue
End Property

Public Property Get Value() As String
  Value = vValue
End Property

Public Property Let Value(ByVal vNewValue As String)
  vValue = vNewValue
End Property

