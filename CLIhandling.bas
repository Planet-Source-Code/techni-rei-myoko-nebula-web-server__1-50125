Attribute VB_Name = "CLIhandling"
Option Explicit
Public txtmain As textbox, formmain As Form, endask As String, canecho As Boolean, stopsearching As Boolean, helpfound As Boolean
Public lastchar As Long, cwd As String, currentprompt As String, cwdrive As String, mode As dosmode, spoth As Long, spoth2 As Long
Public history() As String, tempstr() As String, pathlist() As String, histitems As Long, curritem As Long, stopped As Boolean

Public Enum dosmode
    dosLocal = 0
    dosRemot = 1
    dosFTP = 2
    dosQuestion = 3
    dosPassword = 4
    dospaused = 5
End Enum

Public Function echo(Optional ByVal text As String = Empty, Optional newline As Boolean = True, Optional startonnewline As Boolean = False, Optional sendanyway As Boolean = True)
Select Case LCase(text)
    Case "on": canecho = True
    Case "off": canecho = False
    Case Else
        If startonnewline Then text = vbNewLine & text
        If newline Then text = text & vbNewLine
        txtmain.text = txtmain.text & text
        txtmain.SelStart = Len(txtmain.text)
        lastchar = lastchar + Len(text)
        DoEvents
End Select
echo = text
End Function
Public Function getnext(expression As String, Index As Long, type2check As Boolean) As Long
Dim count As Long
If expression <> Empty And Index > 0 And Index <= Len(expression) Then
    count = Index + 1
    If Mid(expression, Index, 1) = """" Then
    If type2check = False Then count = count + 1
        Do Until Mid(expression, count, 1) = """" Or count = Len(expression)
            count = count + 1
        Loop
        getnext = count
    Else
        Do Until isalphanumeric(Mid(expression, count, 1)) = type2check Or count = Len(expression)
            count = count + 1
        Loop
        getnext = count - 1
    End If
Else
    getnext = 1
End If
End Function
Public Function isalphanumeric(text As String) As Boolean
Const chars As String = "_.$:~'\*-+=<>"
isalphanumeric = False
text = Left(LCase(text), 1)
If text >= "a" And text <= "z" Then isalphanumeric = True
If text >= "0" And text <= "9" Then isalphanumeric = True
If Replace(chars, text, Empty) <> chars Then isalphanumeric = True
End Function
Public Sub splitparameters(command As String, stringarray)
    command = command & " " 'padding so it works with 1 digit end strings
    Dim count As Long, count2 As Long, tempstr As String
    count = 0
    Do Until count = Len(command) Or count2 = Len(command)
            count = getnext(command, count, True)
            count2 = getnext(command, count, False)
            If tempstr <> Empty Then tempstr = tempstr & "|"
            If count > 0 Then tempstr = tempstr & Trim(Mid(command, count, count2 - count + 1))
            count = count2 + 1
    Loop
    If isalphanumeric(Right(tempstr, 1)) = False Then tempstr = Left(tempstr, InStrRev(tempstr, "|") - 1)
    stringarray = Split(tempstr, "|")
End Sub

Public Function allelse(Index As Long)
Dim tempstr2 As String, count As Long
For count = Index + 1 To UBound(tempstr)
    If tempstr2 <> Empty Then tempstr2 = tempstr2 & " "
    tempstr2 = tempstr2 & tempstr(count)
Next
allelse = Trim(Replace(tempstr2, ": ", ":"))
End Function

Public Function getparam(stringarray, number As Long, Optional default As String) As String
    If number > UBound(stringarray) Then
        getparam = default
    Else
        getparam = stringarray(number)
    End If
End Function

Public Function keydown(ByRef KeyCode As Integer) As Integer
On Error Resume Next
    Select Case KeyCode
    Case vbKeyDelete
        If txtmain.SelLength = 0 Then txtmain.SelLength = 1
        txtmain.SelText = Empty
        KeyCode = vbNull
    Case vbKeyRight
        If txtmain.SelStart < Len(txtmain.text) Then
            txtmain.SelStart = txtmain.SelStart + 1
            KeyCode = vbNull
        End If
    Case vbKeyBack, vbKeyLeft
        If txtmain.SelStart <= lastchar Then
            KeyCode = vbNull
        Else
            If KeyCode = vbKeyLeft Then txtmain.SelStart = txtmain.SelStart - 1: KeyCode = vbNull
            If KeyCode = vbKeyBack Then
                If txtmain.SelLength > 0 Then
                    txtmain.SelText = Empty
                Else
                    txtmain.SelStart = txtmain.SelStart - 1
                    txtmain.SelLength = 1
                    txtmain.SelText = Empty
                End If
                KeyCode = vbNull
            End If
        End If
    Case vbKeyReturn
        If mode = dosLocal Or mode = dosRemot Then
            txtmain.SelStart = Len(txtmain.text)
            'If mode = dosRemot And wscmain.State = 7 Then wscmain.SendData Chr(vbKeyReturn)
            txtmain = txtmain & vbNewLine
            If Len(txtmain.text) - lastchar > 0 And mode = dosLocal Then
                CLI Mid(txtmain.text, lastchar + 1, Len(txtmain.text) - lastchar)
            End If
            lastchar = txtmain.SelStart
            If mode = dosLocal Then echo prompt(currentprompt), False
            KeyCode = vbNull
        End If
        If mode = dosPassword Or mode = dosQuestion Then
            txtmain.SelStart = Len(txtmain.text)
            endask = Mid(txtmain.text, lastchar + 1, Len(txtmain.text) - lastchar)
            lastchar = txtmain.SelStart
            'MsgBox endask & vbNewLine & lastchar - 1 & vbNewLine & Len(txtmain.text) - lastchar & vbNewLine & Len(txtmain.text)
            KeyCode = vbNull
        End If
    Case vbKeyUp 'up one entry in history list (unless at ubound/current)
        If curritem < histitems Then curritem = curritem + 1
        processhistitem
    Case vbKeyDown 'down one entry in the history list (unless at 0)
        If curritem > 0 Then curritem = curritem - 1
        processhistitem
    Case vbKeyPageUp, vbKeyPageDown
    Case vbKeyHome
        txtmain.SelStart = lastchar
        KeyCode = vbNull
    Case 3, 19
        stopped = True
    Case Else
        'If wscmain.State = sckConnected Then wscmain.SendData Chr(KeyCode)
        If mode = dosPassword Then 'needs work
            KeyCode = vbKeyMultiply
        End If
End Select
keydown = KeyCode
End Function

Public Sub addhistitem(item As String)
    Dim count As Long
    
    If histitems < Val(GetSetting(App.EXEName, "Main", "Max History", "200")) Then
        histitems = histitems + 1
        ReDim Preserve history(histitems)
    Else
        For count = LBound(history) To UBound(history) - 1
            history(count) = history(count + 1)
        Next
    End If
    curritem = histitems + 1
    history(histitems) = item
End Sub
Public Sub processhistitem()
If histitems > 0 Then
    If curritem <= histitems Then
        txtmain.SelStart = lastchar
        txtmain.SelLength = Len(txtmain.text) - lastchar
        txtmain.SelText = history(curritem)
    Else
        txtmain.SelStart = lastchar
        txtmain.SelLength = Len(txtmain.text) - lastchar
        txtmain.SelText = Empty
    End If
End If
End Sub

Public Sub onUnload(Form As Form)
    txtmain = Empty
    Call SaveSetting(App.EXEName, "Main", "WindowState", Form.WindowState)
    Form.WindowState = 0
    Call SaveSetting(App.EXEName, "Main", "Width", Form.width)
    Call SaveSetting(App.EXEName, "Main", "Height", Form.height)
    Call SaveSetting(App.EXEName, "Main", "Top", Form.Top)
    Call SaveSetting(App.EXEName, "Main", "Left", Form.Left)
    Call SaveSetting(App.EXEName, "Main", "Prompt", currentprompt)
    Call SaveSetting(App.EXEName, "Main", "Last Used", Date)
    End
End Sub
Public Sub onLoad(Form As Form, textbox As textbox)
    Form.WindowState = GetSetting(App.EXEName, "Main", "WindowState", Form.WindowState)
    Form.width = GetSetting(App.EXEName, "Main", "Width", Form.width)
    Form.height = GetSetting(App.EXEName, "Main", "Height", Form.height)
    Form.Top = GetSetting(App.EXEName, "Main", "Top", Form.Top)
    Form.Left = GetSetting(App.EXEName, "Main", "Left", Form.Left)
    currentprompt = GetSetting(App.EXEName, "Main", "Prompt", defaultprompt)
    Set txtmain = textbox
    echo GetSetting(App.EXEName, "Main", "MOTD", defaultMOTD)
    Set formmain = Form
    CLI "initialize"
    If command <> Empty Then CLI command
    echo prompt(currentprompt), False
End Sub
Public Function TextFormat(data As String, length As Byte, Optional rightalign As Boolean = False, Optional spacechar As String = " ", Optional moveover As Boolean = False) As String
On Error Resume Next
Dim count As Long
If moveover = True And Len(data) > length Then
    Do Until length * count >= Len(data)
        count = count + 1
    Loop
    length = length * count
End If
If Len(data) > length Then
    If rightalign = False Then 'left
        TextFormat = Left(data, length)
    Else
        TextFormat = Right(data, length)
    End If
Else
    If rightalign = False Then 'left
        TextFormat = data & String(length - Len(data), spacechar)
    Else
        TextFormat = String(length - Len(data), spacechar) & data
    End If
End If
End Function
Public Sub helpitem(name As String, description As String)
If islike(getparam(tempstr, 1, Empty), name) = True Or LCase(getparam(tempstr, 1, Empty)) = Empty Then
echo TextFormat(name, maxlen) & ": " & description
helpfound = True
End If
End Sub
Public Sub subitem(name As String, description As String)
If islike(LCase(getparam(tempstr, 1, Empty)), name) Then
echo String(maxlen + 2, " ") & description
End If
End Sub
Public Function pluralize(text As String, count As Long, Optional alt As String) As String
If alt = Empty Then pluralize = IIf(count = 1, text, text & "s")
If alt <> Empty Then pluralize = IIf(count = 1, text, alt)
End Function
Public Function ask(Optional question As String = Empty, Optional maskchar As String = Empty, Optional masklength As Double = 0) As String
echo question, False
If masklength > 0 Then txtmain.MaxLength = Len(txtmain.text) + masklength
endask = Empty
mode = dosQuestion
lastchar = Len(txtmain.text)
Do Until endask <> Empty
    DoEvents
Loop
mode = dosLocal
ask = Trim(endask)
txtmain.MaxLength = 0
echo
End Function
Public Function extract(ByVal Filename As String) As String
    extract = Right(Filename, Len(Filename) - InStrRev(Filename, "\"))
End Function
Public Sub pause(Optional prompt As String = Empty)
echo prompt, False
Dim temp As Long
temp = txtmain.SelStart
Do Until txtmain.SelStart <> temp
    DoEvents
Loop
echo
End Sub


