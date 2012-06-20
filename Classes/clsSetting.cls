VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim vName As String
Dim vValue As String

Public Property Get Name() As String
  Name = vName
End Property

Public Property Let Name(vNewValue As String)
  vName = vNewValue
End Property

Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
  Value = vValue
End Property

Public Property Let Value(vNewValue As String)
  vValue = vNewValue
End Property

