VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEncapsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim vName As String
Dim vPrefix As String
Dim vSuffix As String

Public Property Get Name() As String
  Name = vName
End Property

Public Property Let Name(vNewValue As String)
  vName = vNewValue
End Property

Public Property Get Encapsulation() As String
  Encapsulation = vPrefix & vSuffix
End Property

Public Property Let Encapsulation(vNewValue As String)
  vPrefix = Left(vNewValue, 1)
  vSuffix = Right(vNewValue, 1)
End Property

Public Property Get Prefix() As String
  Prefix = vPrefix
End Property

Public Property Get Suffix() As String
  Suffix = vSuffix
End Property

