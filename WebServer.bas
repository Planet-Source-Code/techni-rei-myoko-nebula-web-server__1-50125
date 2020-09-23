Attribute VB_Name = "WebServer"
Option Explicit
Public Const defaultMOTD = "Welcome to Nebula", maxlen As Long = 20, defaultprompt As String = "Console> "
Public Const imageextentions As String = "*.bmp;*.gif;*.jpg;*.jpeg;*.jpe;*.jfif;*.png"
Public Const UniqueIcons As String = "*.scr;*.exe;*.ico;*.lnk;*.cpl;*.msc"
Public clsASP As ClsASPobjects, clsServer As clsMiniHini, clsglobal As clsGlobalASPObjects
Public winServer As Winsock, winconns, pic As PictureBox, filebox As FileListBox, dirbox As DirListBox
Public Function allsetup() As Boolean
    Dim temp As String, buffer As Boolean
   
    temp = getvar("LocalPort")
    buffer = Len(temp) > 0 And Val(temp) > 0
    
    temp = getvar("MaxConnections")
    buffer = buffer And Len(temp) > 0 And Val(temp) > 0
    
    temp = getvar("Error404")
    buffer = buffer And Len(temp) > 0 And fileexists(temp)
    
    temp = AliasPath("NeoCache")
    buffer = buffer And Len(temp) > 0
    
    temp = AliasPath("WWWRoot")
    buffer = buffer And Len(temp) > 0
    
    allsetup = buffer
End Function

Public Sub initialize(ASPobjects As ClsASPobjects, minihini As clsMiniHini, Server As Winsock, slaves, globalcls As clsGlobalASPObjects)
     Set clsASP = ASPobjects
     Set clsServer = minihini
     Set winServer = Server
     Set winconns = slaves
     Set clsglobal = globalcls
End Sub
Public Function getvar(varname As String, Optional default) As String
    getvar = clsServer.getitem("Main", varname, CStr(default))
End Function
Public Sub setvar(varname As String, Optional value)
    clsServer.setitem "Main", varname, CStr(value)
End Sub
Public Function MapPath(ByVal GETrequest As String, Optional root As String = "WWWroot") As String
    Dim temp As String
    If Left(GETrequest, 1) = "/" Then GETrequest = Right(GETrequest, Len(GETrequest) - 1)
    GETrequest = Replace(GETrequest, "/", "\")
    
    'Get which alias is requested, then the filename
    If InStr(GETrequest, "\") > 0 Then
        If AliasExists(Left(GETrequest, InStr(GETrequest, "\") - 1)) Then
            temp = getpath(AliasPath(Left(GETrequest, InStr(GETrequest, "\") - 1)), Right(GETrequest, Len(GETrequest) - InStr(GETrequest, "\")))
        Else
            temp = getpath(AliasPath(root), GETrequest)
        End If
    Else
        If AliasExists(GETrequest) Then
            temp = getpath(AliasPath(GETrequest), Empty)
        Else
            temp = getpath(AliasPath(root), GETrequest)
        End If
    End If
    
    If Not fileexists(temp) Then If Not isadir(temp) Then temp = getvar("Error404")
    temp = neocache(temp) 'check for a cached version of the file
    
    'If you dont allow browsing outside alias'd dirs, and the filename isnt in an alias'd dir, give it the 404
    'The alias is rechecked (since the beginning) because the user may have used a relative path name to escape the alias
    If getvar("AllowOutside", "True") = "False" Then
        If Len(getalias(temp)) = 0 Then temp = getvar("Error404")
    End If
    
    MapPath = temp
End Function
Public Function getpath(ByVal path As String, ByVal Filename As String) As String
    Dim temp As Long, alias As String 'if the file is empty, find the first existing default
    If Right(path, 1) = "\" Then path = Left(path, Len(path) - 1)
    Filename = formatHTML(Filename)
    path = formatHTML(path)
    getpath = chkpath(path, Filename)
    alias = getalias(path)

    If AliasProperty(alias, "AllowBrowse") = "False" Then
    'if allowbrowse is true, then you dont give the default file
        If Len(Filename) = 0 Or Filename = "\" Then
            For temp = 1 To getvar("DefaultCount", 0)
                If fileexists(chkpath(path, clsServer.getitem("Default", CStr(temp)))) Then
                    getpath = chkpath(path, clsServer.getitem("Default", CStr(temp)))
                    Exit Function
                End If
            Next
            getpath = path
        End If
    Else
        getpath = chkpath(path, Filename)
    End If
End Function
Public Function neocacheexists(Filename As String)
    neocacheexists = fileexists(neo(Filename))
End Function
Public Function neocache(Filename As String) As String
    neocache = Filename
    Dim temp As Long, File As String, file2 As String
    If getvar("keepasneos") = "True" And isASP(Filename) Then 'if using neocache and the file is asp
        File = Filename
        If Left(File, 1) = "/" Then File = Right(File, Len(File) - 1) 'trim leading slash
        file2 = neo(File) 'convert to neo cache filename
        If fileexists(file2) Then 'if the neo cache file exists, and its older than the original, use it instead
            If DateDiff("s", FileDateTime(File), FileDateTime(file2)) >= 0 Then neocache = file2
        End If
    End If
End Function
Public Function neo(ByVal Filename As String) As String
    Filename = Left(Filename, InStr(Filename, ".")) & "neo"
    Filename = Replace(Filename, "/", "%5C")
    Filename = Replace(Filename, "\", "%5C")
    Filename = Replace(Filename, ":", "%3A")
    neo = chkpath(AliasPath("NeoCache"), Filename)
End Function
Public Sub SetAlias(alias As String, path As String, Optional AllowBrowse As Boolean = True, Optional AllowRemote As Boolean = True, Optional AllowAsp As Boolean = True)
    With clsServer
        .setitem "Alias", alias, path
        .setitem alias, "AllowBrowse", CStr(AllowBrowse)
        .setitem alias, "AllowRemote", CStr(AllowRemote)
        .setitem alias, "AllowAsp", CStr(AllowAsp)
    End With
End Sub
Public Function enumAliases() As String
    Dim tempstr() As String, temp As Long, temp2 As String
    temp2 = clsServer.EnumKeys("Alias")
    If InStr(temp2, "&") > 0 Then
        tempstr = Split(temp2, "&")
        temp2 = Empty
        For temp = 0 To UBound(tempstr)
            temp2 = temp2 & IIf(Len(temp2) = 0, Empty, "&") & Left(tempstr(temp), InStr(tempstr(temp), "=") - 1)
        Next
        enumAliases = temp2
    Else
        If Len(temp2) > 0 Then enumAliases = Left(temp2, InStr(temp2, "=") - 1)
    End If
End Function
Public Function AliasPath(alias As String) As String
    AliasPath = clsServer.getitem("Alias", alias, chkpath(App.path, alias))
End Function
Public Function AliasExists(alias As String) As Boolean
    AliasExists = clsServer.itemexists("Alias", alias)
End Function
Public Function AliasProperty(alias As String, property As String) As String
If StrComp(alias, "WWWRoot", vbTextCompare) = 0 Then 'Forced parameters
    If StrComp(property, "AllowBrowse", vbTextCompare) = 0 Then AliasProperty = "False": Exit Function 'WWWRoot.Allowbrowse=False
    If StrComp(property, "AllowRemote", vbTextCompare) = 0 Then AliasProperty = "True": Exit Function  'WWWRoot.Allowbrowse=True
End If

    AliasProperty = clsServer.getitem(alias, property)
End Function
Public Sub SetAliasProperty(alias As String, property As String, value)
    clsServer.setitem alias, property, CStr(value)
End Sub
Public Function validalias(alias As String) As Boolean
    If InStr(alias, "&") + InStr(alias, " ") > 0 > 0 Then Exit Function
    Select Case LCase(alias)
        Case "alias", "main", "default", Empty
        Case Else: validalias = True
    End Select
End Function

Public Sub setlocalport(port As Long)
    setvar "LocalPort", port
    LISTEN winServer, port
End Sub
Public Function isASP(Filename As String) As Boolean
'    isASP = containsword(Filename, ".asp")
    isASP = StrComp(Right(Filename, 4), ".asp", vbTextCompare) = 0
End Function
Public Function isNeo(Filename As String) As Boolean
    'isNeo = containsword(Filename, ".neo")
    isNeo = StrComp(Right(Filename, 4), ".neo", vbTextCompare) = 0
End Function
Public Function ASP2NEO(ByVal ASPCODE As String, Filename As String) As String
        ASPCODE = ASP2VBS(ASPCODE, Left(Filename, InStrRev(Filename, "\")))
        ASPCODE = Replace(ASPCODE, "Response.write", "ASP.sWrite", , , vbTextCompare)
        ASPCODE = Replace(ASPCODE, "Server.CreateObject", "ASP.ServerCreateObject", , , vbTextCompare)
        ASPCODE = Replace(ASPCODE, "Server.", "ASP.", , , vbTextCompare)
        
        ASPCODE = Replace(ASPCODE, "Application.lock", "System.lockunlock", , , vbTextCompare)
        ASPCODE = Replace(ASPCODE, "Application.unlock", "System.lockunlock", , , vbTextCompare)
        ASPCODE = Replace(ASPCODE, "Application.Contents.Count", "System.ContentsCount", , , vbTextCompare)
        ASPCODE = Replace(ASPCODE, "Application.", "System.", , , vbTextCompare)
        
        ASPCODE = Replace(ASPCODE, "Request.", "ASP.", , , vbTextCompare)
        ASPCODE = Replace(ASPCODE, "Request(", "ASP.Form(", , , vbTextCompare)
        ASPCODE = Replace(ASPCODE, "Session.", "ASP.", , , vbTextCompare)
        ASPCODE = Replace(ASPCODE, "Response.", "ASP.", , , vbTextCompare)
        ASP2NEO = ASPCODE
End Function
Public Function ExecuteASP(ByVal Filename As String, sock As Winsock, msscript As ScriptControl, clsASP As ClsASPobjects, HTTPRequest As String, Optional makenative As Boolean = True) As String
'    On Error GoTo err:
    Dim ASPCODE As String, Querystring As String
    If InStr(InStrRev(Filename, "\"), Filename, "?") > 0 Then
        Querystring = Right(Filename, Len(Filename) - InStrRev(Filename, "?"))
        Filename = Left(Filename, InStrRev(Filename, "?") - 1)
    End If
    msscript.Reset
    msscript.AddObject "ASP", clsASP
    msscript.AddObject "System", clsglobal
    'Convert the ASP code to accept one object instead of multiple (reduce complexity on my side)
    ASPCODE = loadfile(Filename)
    
    If makenative Then ASPCODE = ASP2NEO(ASPCODE, Filename) 'may make cacheing in the future, so conversion wont be neccesary all the time

    'Add some server variables
    With clsASP
        .processRequest HTTPRequest
        .AddServerVariable "APPL_PHYSICAL_PATH", Left(Filename, InStrRev(Filename, "\"))
        .AddServerVariable "LOCAL_ADDR", sock.LocalIP
        .AddServerVariable "PATH_TRANSLATED", Filename
        .AddServerVariable "QUERY_STRING", Querystring
        .AddServerVariable "REMOTE_ADDR", sock.RemoteHostIP
        .AddServerVariable "REMOTE_HOST", sock.RemoteHost
        .AddServerVariable "SERVER_NAME", sock.LocalHostName
        .AddServerVariable "SERVER_PORT", sock.LocalPort
        .AddServerVariable "SERVER_SOFTWARE", "Nebula version " & App.Major & "." & App.Minor & "." & App.Revision
        .SetQueryString Querystring 'Set the querystring
    
        'Run ASP code
        msscript.AddCode ASPCODE
        msscript.ExecuteStatement "ASP.Flush"
        ExecuteASP = .HTMLCODE
    
        'Set clsASP = Nothing 'destroy ASP server
        .ClearQuery
    End With
    Exit Function
'err:
'    ExecuteASP = "An error occurred on line " & msscript.Error.Line & ", it was: " & msscript.Error.description
End Function


Public Sub listdefaults(list As ListBox, Optional setorget As Boolean = True)
    Dim temp As Long
    If Not setorget Then 'set
        list.Clear
        For temp = 1 To Val(getvar("DefaultCount", 0))
            list.AddItem clsServer.getitem("Default", Val(temp))
        Next
    Else
        setvar "DefaultCount", list.ListCount
        For temp = 0 To list.ListCount - 1
            clsServer.setitem "Default", CStr(temp + 1), list.list(temp)
        Next
    End If
End Sub
Public Sub listaliases(list As ListBox)
    Dim temp As Long, tempstr() As String, temp2 As String
    temp2 = enumAliases
    list.Clear
    If InStr(temp2, "&") = 0 Then
        list.AddItem temp2, 0
    Else
        tempstr = Split(temp2, "&")
        For temp = 0 To UBound(tempstr)
            list.AddItem tempstr(temp), temp
        Next
    End If
    list.Refresh
End Sub
Public Function getalias(path As String) As String
    Dim temp As Long, tempstr() As String, temp2 As String
    temp2 = enumAliases
    If InStr(temp2, "&") = 0 Then
        If StrComp(Left(path, Len(AliasPath(temp2))), AliasPath(temp2), vbTextCompare) = 0 Then
            getalias = temp2
        End If
    Else
        tempstr = Split(temp2, "&")
        For temp = 0 To UBound(tempstr)
            If StrComp(Left(path, Len(AliasPath(tempstr(temp)))), AliasPath(tempstr(temp)), vbTextCompare) = 0 Then
                getalias = tempstr(temp)
                Exit Function
            End If
        Next
    End If
End Function

Public Sub GlobalExecute(command As String, script As ScriptControl, clsASP As ClsASPobjects)
        On Error Resume Next
        Dim Filename As String, ASPCODE As String
        Filename = MapPath("global.asa")
        If fileexists(Filename) Then
            script.Reset
            script.AddObject "ASP", clsASP
            script.AddObject "System", clsglobal
            
            ASPCODE = loadfile(Filename)
            ASPCODE = ASP2NEO(ASPCODE, Filename)
            
            script.AddCode ASPCODE
            script.ExecuteStatement command
            script.ExecuteStatement "ASP.Flush"
            clsASP.ClearQuery
        End If
End Sub
Public Function islike(filter As String, expression As String) As Boolean
On Error Resume Next
Dim tempstr() As String, count As Long
If Replace(filter, ";", Empty) <> filter Then
tempstr = Split(filter, ";")
islike = False
For count = LBound(tempstr) To UBound(tempstr)
    If LCase(expression) Like LCase(tempstr(count)) Then islike = True
Next
Else
If expression Like filter Then islike = True Else islike = False
End If
End Function
Public Function isadir(Filename As String) As Boolean
On Error Resume Next
If Filename <> Empty Then isadir = (GetAttr(Filename) And vbDirectory) = vbDirectory
End Function
Public Function GetSize(ByVal size, Optional Bytes As String = "B", Optional KB As String = "K", Optional MB As String = "M", Optional GB As String = "G") As String
    Select Case Val(size)
        Case 0 To 1023
            GetSize = Val(size) & Bytes
        Case 1024 To 1048576
            GetSize = Round(Val(size) / 1024, 2) & KB
        Case 1048576 To 1073741824
            GetSize = Round(Val(size) / 1048576, 2) & MB
        Case Is > 1073741824
            GetSize = Round(Val(size) / 1073741824, 2) & GB
    End Select
End Function

Public Function makelink(href As String, Optional Caption As String) As String
    If Len(Caption) = 0 Then Caption = href
    makelink = "<a href=""" & href & """>" & Caption & "</a>"
End Function
Public Function maketable(text As String, Optional seperator As String = "TR") As String
    maketable = "<" & seperator & ">" & text & "</" & seperator & ">"
End Function
Public Function makeimage(src As String) As String
    makeimage = "<IMG SRC=""" & src & """ border=0>"
End Function

Public Function makehref(name As String, iconname As String, dirs2root As Long, Optional alias As String, Optional ByVal virtualpath As String, Optional lastchar As Boolean) As String
    If Not lastchar Then virtualpath = virtualpath & "\" Else virtualpath = Empty
    If Len(alias) > 0 Then
        makehref = makelink(virtualpath & name, makeimage(string2(dirs2root, "../") & geticon(iconname)) & name)
    Else 'virtualpath & "/" &
        makehref = makelink(virtualpath & name, name)
    End If
End Function
Public Function string2(ByVal number As Long, text As String) As String
    Dim temp As String
    Do While number > 0
        temp = temp & text
        number = number - 1
    Loop
    string2 = temp
End Function

Public Function makefile(ByVal Filename As String, alias As String, dirs2root As Long, virtualpath As String, lastchar As Boolean) As String
    Dim temp(0 To 4) As String '0 = icon, 1 = name, 2 = size, 3 = type, 4 = date modified
    temp(1) = Right(Filename, Len(Filename) - InStrRev(Filename, "\")) 'Name
    If isadir(Filename) Or Filename = "../" Then
        If Filename = "../" Then
            temp(0) = makelink(IIf(lastchar, "../", Empty), makeimage(string2(dirs2root, "../") & geticon(".Folder")) & "[Parent Directory]")
        Else
            temp(0) = makehref(temp(1), ".Folder", dirs2root, alias, virtualpath, lastchar)
        End If
        temp(3) = "Folder"
    Else
        temp(0) = makehref(temp(1), Filename, dirs2root, alias, virtualpath, lastchar)
        temp(2) = GetSize(FileLen(Filename), " Bytes", " KB", " MB", " GB")
        temp(3) = filetype(Filename)
    End If
    If Filename <> "../" Then
        temp(4) = FileDateTime(Filename)
    Else
        temp(4) = chkpath(AliasPath(alias), "..\")
        If Right(temp(4), 1) = "\" And Len(temp(4)) > 3 Then temp(4) = Left(temp(4), Len(temp(4)) - 1)
        temp(4) = FileDateTime(temp(4))
    End If
    makefile = maketable(maketable(temp(0), "TD") & maketable(temp(2), "TD") & maketable(temp(3), "TD") & maketable(temp(4), "TD"))
End Function

Public Function generatefilelist(path As String, dirlist As DirListBox, filelist As FileListBox, Optional debugmode As Boolean, Optional HTTPRequest As String) As String
    Dim temp As Long, alias As String, alias_path As String, virtual_path As String, virtual_path2 As String, dirs2root As Long, tempstr As String, lastchar As Boolean
    Const colheaders As String = "<TR><TD>Name</Td><TD>Size</TD><TD>Type</TD><TD>Date Modified</TD></TR>"
    Const msg As String = "<TR><TD><Center><B>You do not have permission to view "
    Const msg2 As String = "</TD></TR>"
    lastchar = Replace(Right(HTTPRequest, 1), "/", "\") = "\"

    If Len(path) > 0 Then 'generate dir list, WWWRoot isn't browsable anyway so dont handle it
        alias = getalias(path)
        If Len(alias) > 0 Then alias_path = AliasPath(alias)
        virtual_path = path
        If Len(alias_path) > 0 Then
            virtual_path = Replace(path, AliasPath(alias), alias, , , vbTextCompare)
            
            dirs2root = countwords(virtual_path, "\") + 1
            
            If Right(virtual_path, 1) = "\" Then
                virtual_path = Left(virtual_path, Len(virtual_path) - 1)
            End If
            If InStr(virtual_path, "\") > 0 Then
                virtual_path2 = Right(virtual_path, Len(virtual_path) - InStrRev(virtual_path, "\"))
            Else
                virtual_path2 = virtual_path
            End If
            
            'If lastchar Then virtual_path2 = Empty
        End If
        tempstr = "<Title>Browsing: " & Replace(virtual_path, "\", "/") & "</TITLE><TABLE width=100% >" & colheaders
        
        'make this link absolute string2(dirs2root, "../")
        append tempstr, vbNewLine, makefile("../", alias, dirs2root, virtual_path2, lastchar)
        
        dirlist.path = path
        filelist.path = path
        For temp = 0 To dirlist.ListCount - 1
            append tempstr, vbNewLine, makefile(dirlist.list(temp), alias, dirs2root, virtual_path2, lastchar)
        Next
        For temp = 0 To filelist.ListCount - 1
            append tempstr, vbNewLine, makefile(chkpath(path, filelist.list(temp)), alias, dirs2root, virtual_path2, lastchar)
        Next
    
    End If
endfunction:
    generatefilelist = tempstr & "</TABLE>"
End Function

Public Function geticon(ByVal Filename As String) As String
    Dim count As String, isunique As Boolean
    If islike(UniqueIcons, Filename) Then 'Or (isadir(filename) And fileexists(chkpath(filename, "desktop.ini"))) Then
        isunique = True
    Else
        If isadir(Filename) = True Then Filename = ".Folder"  'is a normal folder
        isunique = False
    End If
    count = searchicon(Filename, isunique)  'search by extention
    If Len(count) = 0 Then count = createicon(Filename, isunique)
    geticon = Replace(Replace(count, AliasPath("Thumbcache"), "Thumbcache", , , vbTextCompare), "\", "/")
End Function
Public Function searchicon(Filename As String, Optional isunique As Boolean) As String
    Dim temp As String
    temp = cachename(Filename, 16, 16, isunique)
    If fileexists(temp) Then searchicon = temp
End Function
Public Function formatHTML(ByVal text As String) As String
    text = Replace(text, "+", " ")
    Do While InStr(text, "%") > 0
        text = Replace(text, Mid(text, InStr(text, "%"), 3), Chr(Dec(Mid(text, InStr(text, "%"), 3))))
    Loop
    formatHTML = text
End Function
Public Function createicon(Filename As String, isunique As Boolean) As String
    On Error Resume Next
    Dim temp As String
    temp = cachename(Filename, 16, 16, isunique)
    
    With pic
        .Cls
        .AutoRedraw = True
        drawfileicon Filename, SmallIcon, .hDC, 0, 0
        SavePicture .Image, temp
    End With
    createicon = temp
End Function
Public Function reformat(ByVal Filename As String, Optional special As String = "%") As String
    Filename = Replace(Filename, "/", special & "5C")
    Filename = Replace(Filename, "\", special & "5C")
    Filename = Replace(Filename, ":", special & "3A")
    reformat = Filename
End Function

Public Function cachename(ByVal Filename As String, width As Long, height As Long, Optional isunique As Boolean, Optional extention As String = "bmp")
    Filename = reformat(Filename, "-")
    If isunique Then
        Filename = Filename & "(" & width & "," & height & ")." & extention
    Else
        Filename = Right(Filename, Len(Filename) - InStrRev(Filename, ".")) & "(" & width & "," & height & ")." & extention
    End If
    cachename = Replace(AliasPath("Thumbcache") & "\" & Filename, "\\", "\")
End Function
Public Function Hex2Dec(ByVal text As String) As Long
    text = Left(UCase(text), 1)
    If text = "%" Then Exit Function
    If Asc(text) >= 48 And Asc(text) <= 57 Then
        Hex2Dec = Asc(text) - 48
    Else
        Hex2Dec = Asc(text) - 55
    End If
End Function
Public Function Dec(ByVal text As String) As Long
    Dim temp As Long, temp2 As Long
    temp = 1
    Do While Len(text) > 0
        temp2 = temp2 + Hex2Dec(Right(text, 1)) * temp
        text = Left(text, Len(text) - 1)
        temp = temp * 16
    Loop
    Dec = temp2
End Function
Public Sub CLI(command As String)
    'GNDN
End Sub
Public Function prompt(current As String)
    'GNDN
End Function
