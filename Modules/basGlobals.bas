Attribute VB_Name = "basGlobals"
Option Explicit

Public Const appName = "Wanchic's ETL"

Public appVersion As String

Public colODBCs As New clsODBCs
Public colSettings As New clsSettings
Public colETLs As New clsETL
Public colDataMaps As New clsDataMaps

Public inProgressETL As Boolean
Public stopWatch As Long

