Attribute VB_Name = "WinSockHandler"
Option Explicit
Public maxconnections As Long
Public Sub CONNECT(sock As Winsock, ip As String, Optional port As Long = 23)
On Error GoTo err:
sock.LocalPort = port
echo "Connecting to: " & ip & " on port " & port
sock.CONNECT ip
Exit Sub
err:
echo "Error> " & err.description
End Sub
Public Sub LISTEN(sock As Winsock, Optional port As Long = 23)
    On Error GoTo err:
    With sock
        .Close
        .LocalPort = port
        .LISTEN
        echo .LocalIP & " is listening on port " & .LocalPort
    End With
    Exit Sub
err:
    echo err.description & " (port " & port & " is being used already)"
End Sub
Public Sub SEND(sock As Winsock, ByVal text As String, Optional block As Long = 1024, Optional closewhendone As Boolean)
    If Len(text) > block Then
        Do While Len(text) > 0
            If Len(text) < block Then block = Len(text)
            If sock.state = sckConnected Then sock.SendData Left(text, block)
            text = Right(text, Len(text) - block)
        Loop
    Else
        If sock.state = sckConnected Then sock.SendData text
    End If
    If closewhendone Then
        Do Until sock.tag <> Empty
            DoEvents
        Loop
        sock.Close
    End If
End Sub
Public Function getfirstfreesock(sockarray) As Long
    getfirstfreesock = sockarray.LBound - 1
    Dim temp As Long
    For temp = sockarray.LBound To sockarray.UBound
        If sockarray(temp).state = 0 Then
            getfirstfreesock = temp
            Exit Function
        End If
    Next
    temp = sockarray.UBound + 1
    
    If temp <= maxconnections Then
        Load sockarray(temp)
        Do Until sockarray.UBound = temp
            DoEvents
        Loop
        getfirstfreesock = temp
    End If
End Function
Public Function AcceptRequest(requestID As Long, sockarray) As Boolean
    Dim temp As Long
    temp = getfirstfreesock(sockarray)
    If temp >= sockarray.LBound Then
        sockarray(temp).Accept requestID
        echo sockarray(temp).RemoteHostIP & " has connected at " & Now
        AcceptRequest = True
    Else
        echo "Request " & requestID & " was rejected due to lack of free connections (" & maxconnections & ")"
    End If
End Function
Public Function SendFile(Filename As String, sock As Winsock, Optional closewhendone As Boolean = True)
    On Error GoTo err:
    Filename = Replace(Filename, "/", "\")
    Filename = Replace(Filename, "%20", " ")
    echo "Sending " & Filename & " to " & sock.RemoteHostIP
    Dim tempfile As Long, filebin As String, filesize As Long, sentsize As Long, temp As Long, issending As Boolean
    Const buffer = 1024
    filesize = FileLen(Filename)
    tempfile = FreeFile
    Open Filename For Binary As #tempfile 'open filename
    filebin = Space(buffer) 'create 1024 byte buffer
    issending = True
    Do Until issending = False
        Get #tempfile, , filebin
        sentsize = sentsize + Len(filebin)
        sock.tag = Empty
        If sentsize > filesize Then
            temp = sentsize - filesize
            SEND sock, Left(filebin, Len(filebin) - temp)
            issending = False
        Else
            SEND sock, filebin
        End If
        Do Until sock.tag <> Empty
            DoEvents
        Loop
        If sock.tag = "False" Then
            closewhendone = True
            GoTo err
        End If
    Loop
    Close tempfile
err:
    If closewhendone Then sock.Close
End Function

Public Function gettag(tag As String, data As String) As String
    Dim tempstr() As String, temp As Long, temp2 As Long
    tag = LCase(tag)
    tempstr = Split(data, vbNewLine)
    For temp = 0 To UBound(tempstr)
    If tag = "method" Then
        If StrComp(Left(tempstr(temp), temp2 - 1), "GET", vbTextCompare) = 0 Then gettag = "Get": Exit Function
        If StrComp(Left(tempstr(temp), temp2 - 1), "POST", vbTextCompare) = 0 Then gettag = "Post": Exit Function
    Else
        If tag = "get" Or tag = "http" Or tag = "post" Then
            temp2 = InStr(tempstr(temp), " ")
            If temp2 > 0 Then
                If StrComp(Left(tempstr(temp), temp2 - 1), "GET", vbTextCompare) = 0 Or StrComp(Left(tempstr(temp), temp2 - 1), "POST", vbTextCompare) = 0 Then
                    If tag = "get" Or tag = "post" Then
                        gettag = Mid(tempstr(temp), temp2 + 1, InStrRev(tempstr(temp), " HTTP", , vbTextCompare) - temp2 - 1)
                    Else
                        temp2 = InStrRev(tempstr(temp), "HTTP", , vbTextCompare) + Len("HTTP")
                        gettag = Right(tempstr(temp), Len(tempstr(temp)) - temp2)
                    End If
                End If
                
            End If
        Else
            temp2 = InStr(tempstr(temp), ":")
            If temp2 > 0 Then
                If StrComp(Left(tempstr(temp), temp2 - 1), tag, vbTextCompare) = 0 Then
                    gettag = Right(tempstr(temp), Len(tempstr(temp)) - temp2 - 1)
                    Exit Function
                End If
            End If
        End If
    End If
    Next
End Function

Public Sub test()
    Dim temp As String
    temp = Clipboard.GetText
    Debug.Print gettag("http", temp)
End Sub
