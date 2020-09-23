VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsASPobjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'All the code needed to run (basic) Asp pages (converted to vbscript)
Dim mini As clsMiniHini, tempstr As String, tempstr2 As String 'asp requires 2 buffers. one that can be cleared, and one thats supposed to have been sent already

'Replaces Request, Response, Session, Apllication, Server, Error ASP objects
Public Property Let buffer(state As Boolean)
    'GNDN. Buffer is always on
End Property
Public Property Let CacheControl(state As Boolean)
    'GNDN. CacheControl is always public
End Property

Public Function ServerVariables(Key) As String
    ServerVariables = formatHTML(mini.getitem("ServerVariables", CStr(Key)))
End Function
Public Function Form(Key) As String
    Form = formatHTML(mini.getitem("Form", CStr(Key)))
End Function

Public Sub ClearQuery()
    mini.Clear
    tempstr = Empty
    tempstr2 = Empty
End Sub
Public Sub Clear()
    tempstr = Empty
End Sub
Public Sub flush()
    append tempstr2, vbNewLine, tempstr
End Sub

Public Sub SetQueryString(query As String, Optional Section As String = "QueryString")
    If query = Empty Then Exit Sub
    Dim tempstr() As String, temp As Long
    If InStr(query, "&") = 0 Then
        AddQuery Section, query
    Else
        tempstr = Split(query, "&")
        For temp = 0 To UBound(tempstr)
            AddQuery Section, tempstr(temp)
        Next
    End If
End Sub
Public Sub AddServerVariable(Key As String, value As String)
    mini.setitem "ServerVariables", Key, value
End Sub
Public Sub AddQuery(Section As String, KeyValue As String)
    If InStr(KeyValue, "=") = 0 Then
        mini.setitem Section, KeyValue
    Else
        mini.setitem Section, Left(KeyValue, InStr(KeyValue, "=") - 1), Right(KeyValue, Len(KeyValue) - InStr(KeyValue, "="))
    End If
End Sub
Public Function GetKey(Key) As String
    GetKey = formatHTML(mini.getitem("Form", CStr(Key), mini.getitem("QueryString", CStr(Key), Empty)))
End Function


Public Function Querystring(Optional Key) As String
    Dim temp As Long, tempstr As String
    If Len(CStr(Key)) = 0 Then
        Querystring = mini.EnumKeys("QueryString")
    Else
        Querystring = formatHTML(mini.getitem("QueryString", CStr(Key)))
    End If
End Function

 Public Function ServerCreateObject(name As String) As Object
    On Error GoTo err:
    Set ServerCreateObject = CreateObject(name)
    Exit Function
err:
    Set ServerCreateObject = Nothing
End Function

Public Function MapPath(sPath As String) As String
    MapPath = chkpath(Querystring("APPL_PHYSICAL_PATH"), sPath)
End Function

Public Sub sWrite(Optional text)
    append tempstr, vbNewLine, CStr(text)
End Sub
Public Function HTMLCODE() As String
    HTMLCODE = tempstr2
End Function
Public Sub processRequest(text As String)
    On Error Resume Next
    Dim tempstr() As String, temp As Long, mode As Boolean
    tempstr = Split(text, vbNewLine)
    AddQuery "ServerVariables", "REQUEST_METHOD=" & Left(tempstr(0), InStr(tempstr(0), " ") - 1)
    For temp = 1 To UBound(tempstr) 'First line is get and http
        If mode = False Then
            If Len(tempstr(temp)) = 0 Then
                mode = True
            Else
                AddQuery "ServerVariables", Left(tempstr(temp), InStr(tempstr(temp), ":") - 1) & "=" & Right(tempstr(temp), Len(tempstr(temp)) - InStr(tempstr(temp), " "))
            End If
        Else
            If Len(tempstr(temp)) > 0 Then SetQueryString tempstr(temp), "Form"
        End If
    Next
End Sub


Private Sub Class_Initialize()
    Set mini = New clsMiniHini
End Sub

Private Sub Class_Terminate()
    Set mini = Nothing
End Sub
