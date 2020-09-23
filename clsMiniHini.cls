VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMiniHini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type item
    Section As String
    Key As String
    value As String
End Type
Dim grid() As item, item_count As Long
Public Sub Clear()
    redimpreserve 0
End Sub
Private Function itemhandle(Section As String, Key As String) As Long
    Dim temp As Long
    For temp = 1 To item_count
        If StrComp(Section, grid(temp).Section, vbTextCompare) = 0 Then
            If StrComp(Key, grid(temp).Key, vbTextCompare) = 0 Then
                itemhandle = temp
            End If
        End If
    Next
End Function

Private Sub redimpreserve(items As Long)
    On Error Resume Next
    Select Case items
        Case 0: ReDim grid(0)
        Case 1: ReDim grid(1 To 1)
        Case Else: ReDim Preserve grid(1 To items)
    End Select
    DoEvents
    item_count = items
End Sub
Private Function countchars(text As String, char As String) As Long
    countchars = (Len(text) - Len(Replace(text, char, Empty))) / Len(char)
End Function
Public Function itemexists(Section As String, Key As String) As Boolean
    itemexists = itemhandle(Section, Key) > 0
End Function
Private Function appendto(destination As String, delimeter As String, text As String) As String
    If Len(destination) = 0 Then
        appendto = text
    Else
        appendto = destination & delimeter & text
    End If
End Function
Public Function getitem(Section As String, Key As String, Optional default As String) As String
    Dim temp As Long, temp2 As Long
    getitem = default
    temp = itemhandle(Section, Key)
    If temp > 0 Then getitem = grid(temp).value
End Function
    
Public Function setitem(Section As String, Key As String, Optional value As String)
    Dim temp As Long
    temp = itemhandle(Section, Key)
    If temp = 0 Then
        redimpreserve item_count + 1
        temp = item_count
        grid(temp).Section = Section
        grid(temp).Key = Key
    End If
    grid(temp).value = value
    'Debug.Print Section & "." & Key & "=" & Value
End Function
Public Sub deleteitem(Section As String, Key As String)
    Dim temp As Long
    temp = itemhandle(Section, Key)
    If item_count = 1 And temp = 1 Then
        redimpreserve 0
    Else
        If temp = item_count Then
            redimpreserve item_count - 1
        Else
            grid(temp) = grid(item_count)
            redimpreserve item_count - 1
        End If
    End If
End Sub
Public Function LoadMini(Filename As String) As Boolean
    On Error Resume Next
    Dim temp As Long, tempstr As String, tempfile As Long
    If fileexists(Filename) Then
        tempfile = FreeFile
        Open Filename For Input As #tempfile
            Do Until EOF(tempfile)
                Line Input #tempfile, tempstr
                setitem Left(tempstr, InStr(tempstr, ".") - 1), Mid(tempstr, InStr(tempstr, ".") + 1, InStr(tempstr, "=") - InStr(tempstr, ".") - 1), Right(tempstr, Len(tempstr) - InStr(tempstr, "="))
                'Debug.Print "Adding: " & Left(tempstr, InStr(tempstr, ".") - 1) & vbTab & Mid(tempstr, InStr(tempstr, ".") + 1, InStr(tempstr, "=") - InStr(tempstr, ".") - 1) & vbTab & Right(tempstr, Len(tempstr) - InStr(tempstr, "="))
            Loop
        Close #tempfile
    End If
    LoadMini = True
End Function
Public Function SaveMini(Filename As String) As Boolean
    On Error Resume Next
    Dim temp As Long, tempfile As Long
    tempfile = FreeFile
    Open Filename For Output As #tempfile
        For temp = 1 To item_count
            Print #tempfile, grid(temp).Section & "." & grid(temp).Key & "=" & grid(temp).value
        Next
    Close #tempfile
    SaveMini = True
End Function

Public Function KeyCount(Section As String) As Long
    Dim temp As Long, count As Long
    For temp = 1 To item_count
        If StrComp(grid(temp).Section, Section, vbTextCompare) = 0 Then count = count + 1
    Next
    KeyCount = count
End Function

Public Function GetKeyIndex(Section As String, Index As Long) As String
    Dim temp As Long, count As Long
    For temp = 1 To item_count
        If StrComp(grid(temp).Section, Section, vbTextCompare) = 0 Then
            count = count + 1
            If count = Index Then
                GetKeyIndex = grid(temp).value
                Exit Function
            End If
        End If
    Next
End Function

Public Function EnumKeys(Section As String, Optional MainDelimeter As String = "&", Optional SecondaryDelimeter As String = "=") As String
    Dim temp As Long, tempstr As String
    For temp = 1 To item_count
        If StrComp(grid(temp).Section, Section, vbTextCompare) = 0 Then
            tempstr = appendto(appendto(tempstr, MainDelimeter, grid(temp).Key), SecondaryDelimeter, grid(temp).value)
        End If
    Next
    EnumKeys = tempstr
End Function

Public Sub RenameSection(Section As String, name As String)
    Dim temp As Long
    If Not SectionExists(Section) Then
    For temp = 1 To item_count
        If StrComp(grid(temp).Section, Section, vbTextCompare) = 0 Then
            grid(temp).Section = name
        End If
    Next
    End If
End Sub
Public Function RenameKey(Section As String, Key As String, name As String) As Boolean
    Dim temp As Long
    temp = itemhandle(Section, name)
    If temp = 0 Then
        temp = itemhandle(Section, Key)
        If temp <> 0 Then
            grid(temp).Key = name
        End If
    End If
End Function
Public Function SectionExists(Section As String) As Boolean
    Dim temp As Long
    For temp = 1 To item_count
        If StrComp(grid(temp).Section, Section, vbTextCompare) = 0 Then
            SectionExists = True
            Exit Function
        End If
    Next
End Function

Public Sub DeleteSection(Section As String)
    Dim temp As Long
    For temp = item_count To 1 Step -1
        If StrComp(grid(temp).Section, Section, vbTextCompare) = 0 Then
            deleteitem Section, grid(temp).Key
        End If
    Next
End Sub
Public Function EnumSections(Optional delimeter As String = "|") As String
    Dim temp As Long, tempstring As String
    If item_count > 0 Then tempstring = grid(1).Section & delimeter
    For temp = 2 To item_count
        If Not containsword(tempstring, grid(temp).Section) Then tempstring = tempstring & grid(temp).Section & delimeter
    Next
    EnumSections = Left(tempstring, Len(tempstring) - 1)
End Function
