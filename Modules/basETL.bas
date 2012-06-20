Attribute VB_Name = "basETL"
Option Explicit

Public Sub executeETL()
  

  If Not (inProgressETL) Then
    inProgressETL = True
    If Not (IsFormLoaded("frmMain")) Then frmMain.Show
    
    establishODBCs
    startETL
    closeODBCs
    
    reportStatus "ETL Complete"
    
    inProgressETL = False

  End If


End Sub

Private Function establishODBCs() As Boolean
  Dim cODBC As clsODBC
  
  establishODBCs = False
  
  For Each cODBC In colODBCs
    reportStatus "Connecting to DB: " & cODBC.Name
    cODBC.Connect
    
  Next cODBC
  
  establishODBCs = True
  
End Function

Private Sub startETL()
  Dim cODBC As clsODBC
  Dim cDB As clsDatabase
  Dim cRec As clsRecord
  
  Dim uODBC As clsODBC
  Dim uDb As clsDatabase
  Dim uRec As clsRecord
  
  Dim sRec As clsSyncRecord
  
  Dim rQuery As ADODB.Recordset
  Dim sqlStr As String
  
  reportStatus "Starting ETL"
  
  For Each cDB In colETLs
    reportStatus "Synchronizing DB: " & cDB.Name
    
    Set cODBC = colODBCs(cDB.Name)
    
    For Each cRec In cDB
      Set rQuery = cODBC.QueryAll(cRec.Table)
      Set sRec = cRec.SyncClone
      
      If Not (rQuery.EOF) Then
          
        For Each uDb In colETLs
          If uDb.Name <> cDB.Name Then
            Set uODBC = colODBCs(uDb.Name)
          
            Set uRec = uDb(cRec.Key)
            sRec.Map uRec
                          
            rQuery.MoveFirst
            
            Do
              sRec.LoadValues rQuery
              
              If uODBC.QueryRec(sRec).EOF Then
                uODBC.InsertInto sRec
                
              Else
                uODBC.Update sRec
              
              End If
              
              rQuery.MoveNext
            Loop Until rQuery.EOF
            
            
            
          End If
        Next uDb
          
      End If
      
    Next cRec
  Next cDB
  
End Sub

Private Function closeODBCs() As Boolean
  Dim cODBC As clsODBC
    
  For Each cODBC In colODBCs
    reportStatus "Disconnecting from DB: " & cODBC.Name
    cODBC.Disconnect
    
  Next cODBC
    
End Function
