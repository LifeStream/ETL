VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim vName As String
Dim vConnection As String
Dim vEncapsulations As clsEncapsulations
Dim vDB_Type As String

Dim vDB As ADODB.Connection

Private Sub Class_Initialize()
  Set vEncapsulations = New clsEncapsulations
  
End Sub

Public Property Get Name() As String
  Name = vName
End Property

Public Property Let Name(vNewValue As String)
  vName = vNewValue
End Property

Public Property Get Connection() As String
  Connection = vConnection
End Property

Public Property Let Connection(vNewValue As String)
  vConnection = vNewValue
End Property

Public Property Get DB_Type() As String
  DB_Type = vDB_Type
End Property

Public Property Let DB_Type(vNewValue As String)
  vDB_Type = vNewValue
End Property

Public Function AddEncapsulation(vName As String, vEncapsulation As String) As clsEncapsulation
  Dim cEncap As clsEncapsulation
  
  Set cEncap = vEncapsulations.Add(vName)
  cEncap.Encapsulation = vEncapsulation
  
  Set AddEncapsulation = cEncap
End Function

Public Function Encapsulation(vName As String) As clsEncapsulation
  Set Encapsulation = vEncapsulations(vName)
  
End Function

Private Function getEncap(vName As String) As clsEncapsulation
  Set getEncap = vEncapsulations(vName)

End Function

Public Property Get DB() As ADODB.Connection
Attribute DB.VB_UserMemId = 0
  Set DB = vDB
End Property

Public Function Connect() As Boolean
  Dim cADODB As New ADODB.Connection

  Connect = False
  
  Select Case vDB_Type
    Case "dao"
      'Set vDB = OpenDatabase(vConnection)
      'Set cADODB = Nothing
      Set vDB = cADODB
      vDB.Open vConnection
      
    Case "adodb"
      Set vDB = cADODB
      vDB.Open vConnection
      
  End Select
  
  DoEvents
  
  Connect = True
  
End Function

Public Sub Disconnect()

  Dim cADODB As New ADODB.Connection
  
  Select Case vDB_Type
    Case "dao"
      'Set vDB = OpenDatabase(vConnection)
      'Set cADODB = Nothing
      Set vDB = cADODB
      If vDB.State = adStateOpen Then vDB.Close
      
    Case "adodb"
      Set vDB = cADODB
      If vDB.State = adStateOpen Then vDB.Close
      
  End Select
  
  DoEvents
  
  Set vDB = Nothing
  DoEvents
  
End Sub

Public Function QueryRec(sRec As clsSyncRecord) As ADODB.Recordset
  Dim sqlStr As String
  Dim eObj As clsEncapsulation
  
  Dim sField As clsSyncField
  
  Set sField = sRec.getPrimaryKey
  
  sqlStr = "SELECT * " & _
           "FROM " & getEncap("object").Prefix & sRec.TableUpdate & getEncap("object").Suffix & " " & _
           "WHERE " & getEncap("object").Prefix & sField.NameUpdate & getEncap("object").Suffix & " = " & _
                      getEncap(sField.DataEncapsulationUpdate).Prefix & sField.Value & _
                      getEncap(sField.DataEncapsulationUpdate).Suffix & ";"
           
  Set QueryRec = Execute(sqlStr)
  
End Function

Public Function QueryAll(sTable As String) As ADODB.Recordset
  Dim sqlStr As String
  Dim eObj As clsEncapsulation
  
  sqlStr = "SELECT * " & _
           "FROM " & getEncap("object").Prefix & sTable & getEncap("object").Suffix & ";"
           
  Set QueryAll = Execute(sqlStr)
  
End Function

Public Sub InsertInto(sRec As clsSyncRecord)
  Dim sqlStr As String

  sqlStr = "INSERT INTO " & getEncap("object").Prefix & sRec.TableUpdate & getEncap("object").Suffix & " " & _
           InsertIntoValueSet(sRec) & ";"
  
  Execute sqlStr
  
End Sub

Public Sub Update(sRec As clsSyncRecord)
  Dim sqlStr As String
  
  sqlStr = "UPDATE " & getEncap("object").Prefix & sRec.TableUpdate & getEncap("object").Suffix & " " & _
           "SET " & UpdateValueSet(sRec) & " " & _
           "WHERE " & UpdateWhereCriteria(sRec) & ";"
           
  Execute sqlStr
  
End Sub


Public Function Execute(sqlStr As String) As ADODB.Recordset
  
  Set Execute = vDB.Execute(sqlStr)
  
  DoEvents
End Function

Private Function InsertIntoValueSet(sRec As clsSyncRecord) As String
  Dim cSyncField As clsSyncField
  Dim colFields As String
  Dim colValues As String
  
  Dim trValue As String
  
  colFields = ""
  colValues = ""
  
  For Each cSyncField In sRec
    With cSyncField
      If Nz(.Value) <> "" Then
      
        If Nz(.DataMap) <> "" Then
          trValue = colDataMaps.reMap(cSyncField)
          
        Else
          trValue = .Value
          
        End If
                
        colFields = colFields & getEncap("object").Prefix & .NameUpdate & getEncap("object").Suffix & ", "
            
        colValues = colValues & getEncap(.DataEncapsulationUpdate).Prefix & _
                    trValue & getEncap(.DataEncapsulationUpdate).Suffix & ", "
      End If
    End With
  Next cSyncField
  
  colFields = "(" & Left(colFields, Len(colFields) - 2) & ")"
  colValues = "(" & Left(colValues, Len(colValues) - 2) & ")"
  
  InsertIntoValueSet = colFields & " VALUES " & colValues
  
End Function


Private Function UpdateValueSet(sRec As clsSyncRecord) As String
  Dim cSyncField As clsSyncField
  Dim sqlSet As String
  Dim trValue As String
  
  sqlSet = ""
  
  For Each cSyncField In sRec
    With cSyncField
      If Not (.PrimaryKey) Then
        If Nz(.Value) <> "" Then
          If Nz(.DataMap) <> "" Then
            trValue = colDataMaps.reMap(cSyncField)
            
          Else
            trValue = .Value
            
          End If
                  
          sqlSet = sqlSet & getEncap("object").Prefix & .NameUpdate & getEncap("object").Suffix & "=" & _
                            getEncap(.DataEncapsulationUpdate).Prefix & trValue & getEncap(.DataEncapsulationUpdate).Suffix & ", "
        End If
      End If
    End With
  Next cSyncField
  
  sqlSet = Left(sqlSet, Len(sqlSet) - 2)
  
  UpdateValueSet = sqlSet
  
End Function

Private Function UpdateWhereCriteria(sRec As clsSyncRecord) As String
  Dim mLink As clsSyncField
  Dim vCriteria As String
  
  vCriteria = UpdateWhereCriteriaItem(sRec.getPrimaryKey)
  vCriteria = vCriteria & UpdateWhereCriteriaItem(sRec.getUpdated_TS)
    
  vCriteria = Left(vCriteria, Len(vCriteria) - 5)
  UpdateWhereCriteria = vCriteria
  
End Function

Private Function UpdateWhereCriteriaItem(sField As clsSyncField) As String
  Dim vCriteria As String
    
  vCriteria = getEncap("object").Prefix & sField.NameUpdate & getEncap("object").Suffix & "=" & _
              getEncap(sField.DataEncapsulationUpdate).Prefix & _
              sField.Value & _
              getEncap(sField.DataEncapsulationUpdate).Suffix & " AND "
  
  UpdateWhereCriteriaItem = vCriteria
  
End Function

