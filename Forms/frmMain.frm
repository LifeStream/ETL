VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Wanchic's ETL"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4920
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblStatus 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4635
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub loadForm()
  
  'vLoginDate = Now()
  
  'If vClientID = 0 Then
  '  chdMenuAdmin.initLevel
    
  'Else
  '  Select Case str_Client_ClientDefined_Menus_ItemsByMenuID(vMenuID)
  '    Case "1"
  '      chdMenuClient.initLevel
  '
  '    Case "3"
  '      chdMenuEntry.initLevel
  '
  '    Case Else
  '      mainTerminator
  '
  '  End Select
  '
  'End If

  'load Form Values
  
  Me.Show
End Sub
