Attribute VB_Name = "basFormUtilities"
Option Explicit

Public Function IsFormLoaded(ByVal frmName As String) As Boolean
  Dim frm As Form
  
  IsFormLoaded = False
  
  For Each frm In Forms
    If LCase(frm.Name) = LCase(frmName) Then
      IsFormLoaded = True
      Exit For
    End If
    
  Next frm

End Function
