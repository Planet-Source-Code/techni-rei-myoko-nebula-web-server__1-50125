VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGlobalASPObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mini As clsMiniHini
Public Sub lockunlock()
    'GNDN
End Sub
Public Function ContentsCount() As Long
    Dim temp As String
    temp = mini.EnumKeys("Application")
    ContentsCount = countwords(temp, "&") + 1
End Function

Public Property Let AppVar(varname As String, value As String)
    mini.setitem "Application", varname, CStr(value)
End Property
Public Property Get AppVar(varname As String) As String
    AppVar = mini.getitem("Application", varname)
End Property

Private Sub Class_Initialize()
    Set mini = New clsMiniHini
End Sub

Private Sub Class_Terminate()
    Set mini = Nothing
End Sub

