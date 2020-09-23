Attribute VB_Name = "ASPParser"
Option Explicit

Public Function ASP2VBS(ByVal HTMLCODE As String, Optional path As String) As String
    Dim tempstr As String, temp As Long, temp2 As Long, Filename As String, tempbuff() As String, temp3 As Long
    If path = Empty Then path = App.path
    Do Until Len(HTMLCODE) = 0
        Select Case Left(HTMLCODE, 2)
            Case "<%" 'is ASP
                temp = InStr(1, HTMLCODE, "%>")
                If Left(HTMLCODE, 3) = "<%=" Then ' is a <%= variable %>
                    append tempstr, vbNewLine, "Response.write " & Mid(HTMLCODE, 4, temp - 4)
                Else 'isnt
                    append tempstr, vbNewLine, Mid(HTMLCODE, 3, temp - 3)
                End If
                HTMLCODE = Right(HTMLCODE, Len(HTMLCODE) - temp - 1)
            Case "<!" 'is an include
                temp = InStr(1, HTMLCODE, "-->")
                Filename = chkpath(path, addfrom(Left(HTMLCODE, temp + 2), "File"))
                Filename = ASP2VBS(loadfile(Filename), Left(Filename, InStrRev(Filename, "\") - 1))
                append tempstr, vbNewLine, Filename
                HTMLCODE = Right(HTMLCODE, Len(HTMLCODE) - temp - 2)
            Case Else 'is text
                temp = InStr(1, HTMLCODE, "<%") - 1
                temp2 = InStr(1, HTMLCODE, "<!") - 1
                If temp2 > 0 And temp2 < temp Then temp = temp2
                If temp <= 0 Then temp = Len(HTMLCODE)
                If Replace(Trim(Left(HTMLCODE, temp)), vbNewLine, Empty) <> Empty Then
                    If InStr(Left(HTMLCODE, temp), vbNewLine) > 0 Then
                        tempbuff = Split(Replace(Left(HTMLCODE, temp), """", """"""), vbNewLine)
                        For temp3 = 0 To UBound(tempbuff)
                            append tempstr, vbNewLine, "Response.write """ & tempbuff(temp3) & """"
                        Next
                    Else
                        append tempstr, vbNewLine, "Response.write """ & Replace(Left(HTMLCODE, temp), """", """""") & """"
                    End If
                End If
                HTMLCODE = Right(HTMLCODE, Len(HTMLCODE) - temp)
        End Select
    Loop
    ASP2VBS = tempstr
End Function
Public Function loadfile(Filename As String) As String
    On Error Resume Next
    If FileLen(Filename) = 0 Then Exit Function
    Dim temp As Long, tempstr As String, tempstr2 As String
    temp = FreeFile
    If Dir(Filename) <> Filename Then
        Open Filename For Input As temp
            Do Until EOF(temp)
                Line Input #temp, tempstr
                append tempstr2, vbNewLine, tempstr
            Loop
            loadfile = tempstr2
        Close temp
    End If
End Function
Public Sub append(destination As String, delimeter As String, text As String)
    If Len(destination) = 0 Then
        destination = text
    Else
        destination = destination & delimeter & text
    End If
End Sub
Public Function addfrom(content As String, tag As String) As String
    Dim temp As Long, location As Long, temp2 As Long
    If LCase(tag) <> "node" Then
    
    location = InStr(1, content, tag, vbTextCompare)
    If location > 0 Then
    location = InStr(location, content, "=") + 1
    Select Case Mid(content, location, 1)
        Case """", "'"
            location = location + 1
            temp = InStr(location, content, """")
            If temp = 0 Then temp = InStr(location, content, "'")
            temp2 = InStr(location, content, ">")
        Case Else
            temp = InStr(location, content, " ")
            temp2 = InStr(location, content, ">")
    End Select
    If temp2 < temp And temp2 > 0 Then temp = temp2
    If temp = 0 Then temp = InStr(location, content, ">")
    If temp = 0 Then temp = Len(content)
    addfrom = Mid(content, location, temp - location)
    End If
    Else
        addfrom = removebrackets(content, "<", ">")
    End If
End Function
Public Function removetext(text As String, start As Long, finish As Long, Optional exclusive As Boolean = True) As String
    If exclusive = True Then
        removetext = Left(text, start - 1) & Right(text, Len(text) - finish)
    Else
        removetext = Mid(text, start, finish - start)
    End If
End Function
Public Function removebrackets(ByVal text As String, leftb As String, rightb As String) As String
    Do While InStr(text, leftb) > 0 And InStr(text, rightb) > InStr(text, leftb)
        text = removetext(text, InStr(text, leftb), InStr(text, rightb))
    Loop
    removebrackets = text
End Function
Public Function chkpath(ByVal basehref As String, ByVal URL As String) As String
'Debug.Print basehref & " " & URL
chkpath = basehref
Const goback As String = "..\"
Const slash As String = "\"
Dim spoth As Long
URL = Replace(URL, "/", "\")
If Left(URL, 1) = slash Then URL = Right(URL, Len(URL) - 1)
If Right(basehref, 1) = slash And Len(basehref) > 3 Then basehref = Left(basehref, Len(basehref) - 1)
If LCase(URL) <> LCase(basehref) And URL <> Empty And basehref <> Empty Then
If URL Like "?:*" Then 'is absolute
    chkpath = URL
Else
    If containsword(URL, goback) Then 'is relative
        If containsword(Right(basehref, Len(basehref) - 3), slash) = True Then
            For spoth = 1 To countwords(URL, goback)
                If countwords(basehref, slash) > 0 Then
                    URL = Right(URL, Len(URL) - Len(goback))
                    basehref = Left(basehref, InStrRev(basehref, slash) - 1)
                Else
                    URL = Replace(URL, goback, "")
                End If
            Next
        Else
            URL = Replace(URL, goback, "")
        End If
        If Right(basehref, 1) <> slash Then chkpath = basehref & slash & URL Else chkpath = basehref & URL
    Else 'is additive
        If Right(basehref, 1) <> slash Then chkpath = basehref & slash & URL Else chkpath = basehref & URL
    End If
End If
End If
End Function
Public Function countwords(phrase As String, word As String) As Long
    countwords = (Len(phrase) - Len(Replace(phrase, word, Empty))) / Len(word)
End Function
Public Function containsword(phrase As String, word As String) As Boolean
    containsword = InStr(1, phrase, word, vbTextCompare)
End Function
Public Function GetRelativePath(sBase As String, sFile As String)
sBase = LCase(sBase) 'must end with a slash
sFile = LCase(sFile)
    Dim Base() As String, File() As String
    Dim i As Integer, NewTreeStart As Long, sRel As String
    If Left(sBase, 3) <> Left(sFile, 3) Then
        GetRelativePath = sFile
        Exit Function
    End If
    Base = Split(sBase, "\")
    File = Split(sFile, "\")
    While Base(i) = File(i)
        i = i + 1
    Wend
    If i = UBound(Base) Then
        While i <= UBound(File)
            sRel = sRel + File(i) + "\"
            i = i + 1
        Wend
        GetRelativePath = Left(sRel, Len(sRel) - 1)
        Exit Function
    End If
    NewTreeStart = i
    While i < UBound(Base)
        sRel = sRel & "..\"
        i = i + 1
    Wend
    While NewTreeStart <= UBound(File)
        sRel = sRel & File(NewTreeStart) + "\"
        NewTreeStart = NewTreeStart + 1
    Wend
    GetRelativePath = Left(sRel, Len(sRel) - 1)
End Function
Public Function uniquefilename(Filename As String) As String
    Dim temp1 As String, temp2 As String, temp3 As Long
    uniquefilename = Filename
    
    If fileexists(Filename) Then
        Dim count As Long
        count = 1
        temp3 = InStrRev(Filename, ".")
        temp1 = Filename
        If temp3 > 0 Then
            temp1 = Left(Filename, temp3 - 1)
            temp2 = Right(Filename, Len(Filename) - temp3 + 1)
        End If
        Do Until fileexists(temp1 & " (" & count & ")" & temp2) = False
            count = count + 1
        Loop
        uniquefilename = temp1 & " (" & count & ")" & temp2
    End If
End Function
Public Sub Savefile(Filename As String, text As String, Optional purge As Boolean = True)
    On Error Resume Next
    Dim temp As Long
    temp = FreeFile
    If purge Then
        Open Filename For Output As #temp
    Else
        Open Filename For Append As #temp
    End If
        Print #temp, text
    Close #temp
End Sub
Public Function fileexists(Filename As String) As Boolean
    On Error Resume Next
    fileexists = Dir(Filename, vbNormal + vbHidden + vbSystem) <> Empty And Filename <> Empty
End Function
