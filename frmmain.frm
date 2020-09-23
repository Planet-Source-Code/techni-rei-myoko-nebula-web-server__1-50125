VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nebula"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11160
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   11160
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin ASPWebServer.CinTitlebar CinSystem 
      Height          =   1455
      Left            =   8400
      TabIndex        =   39
      Top             =   360
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2566
      Caption         =   "System Objects"
      TitleBar        =   -1  'True
      Backcolor       =   -2147483643
      Begin VB.FileListBox Filmain 
         Height          =   285
         Left            =   960
         TabIndex        =   41
         Top             =   360
         Width           =   375
      End
      Begin VB.DirListBox Dirmain 
         Height          =   315
         Left            =   600
         TabIndex        =   40
         Top             =   360
         Width           =   375
      End
      Begin MSWinsockLib.Winsock winsockServer 
         Left            =   120
         Top             =   600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   23
      End
      Begin MSWinsockLib.Winsock winsockslaves 
         Index           =   0
         Left            =   120
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSScriptControlCtl.ScriptControl scriptmain 
         Index           =   0
         Left            =   1440
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         AllowUI         =   -1  'True
      End
      Begin VB.PictureBox picmain 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   240
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   43
         ToolTipText     =   "Do not resize or rename me"
         Top             =   480
         Width           =   240
      End
      Begin VB.Label lblwarning 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "These are required. Dont Delete or rename"
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   960
         Width           =   1935
      End
   End
   Begin VB.PictureBox pichelp 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   2
      Left            =   3480
      Picture         =   "frmmain.frx":0E42
      ScaleHeight     =   330
      ScaleWidth      =   255
      TabIndex        =   38
      Top             =   4920
      Width           =   255
   End
   Begin VB.PictureBox pichelp 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   1
      Left            =   3480
      Picture         =   "frmmain.frx":12FC
      ScaleHeight     =   330
      ScaleWidth      =   255
      TabIndex        =   37
      Top             =   3000
      Width           =   255
   End
   Begin VB.PictureBox pichelp 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   0
      Left            =   3480
      Picture         =   "frmmain.frx":17B6
      ScaleHeight     =   330
      ScaleWidth      =   255
      TabIndex        =   36
      Top             =   120
      Width           =   255
   End
   Begin ASPWebServer.CinTitlebar Cinmain 
      Height          =   6900
      Index           =   1
      Left            =   3960
      TabIndex        =   10
      Tag             =   "Current Activity"
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   12171
      Caption         =   "Current Activity"
      Begin VB.TextBox txtdisplay 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   120
         Width           =   6855
      End
   End
   Begin ASPWebServer.CinTitlebar Cinmain 
      Height          =   1860
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3281
      Caption         =   "Default Files"
      TitleBar        =   -1  'True
      Begin VB.CommandButton cmdbrowse 
         Caption         =   "ý"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   3240
         TabIndex        =   35
         ToolTipText     =   "Remove from the priority list"
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdbrowse 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   3240
         TabIndex        =   34
         ToolTipText     =   "Move up the priority list"
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton cmdbrowse 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   3240
         TabIndex        =   33
         ToolTipText     =   "Move up the priority list"
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdbrowse 
         Caption         =   "þ"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3240
         TabIndex        =   15
         ToolTipText     =   "Add to the bottom of the priority list"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtopt 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   3015
      End
      Begin VB.ListBox lstmain 
         ForeColor       =   &H80000007&
         Height          =   1035
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   3015
      End
   End
   Begin ASPWebServer.CinTitlebar Cinmain 
      Height          =   2100
      Index           =   3
      Left            =   120
      TabIndex        =   16
      Top             =   4920
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3704
      Caption         =   "Directories"
      TitleBar        =   -1  'True
      Begin VB.CommandButton cmdbrowse 
         Caption         =   "Set"
         Height          =   285
         Index           =   5
         Left            =   3240
         TabIndex        =   32
         ToolTipText     =   "Add/Set the selected alias"
         Top             =   510
         Width           =   375
      End
      Begin VB.TextBox txtopt 
         Height          =   285
         Index           =   4
         Left            =   1800
         TabIndex        =   30
         Tag             =   "localport"
         Top             =   510
         Width           =   1335
      End
      Begin VB.CheckBox Chkopt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow ASP execution"
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   22
         Tag             =   "browsing"
         ToolTipText     =   "Allow the execution of ASP pages"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CheckBox Chkopt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow remote viewing"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   21
         Tag             =   "remote"
         ToolTipText     =   "Allow viewers from other sites to go directly here"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cmdbrowse 
         Caption         =   "..."
         Height          =   285
         Index           =   3
         Left            =   3240
         TabIndex        =   20
         ToolTipText     =   "Browse"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtopt 
         Height          =   285
         Index           =   3
         Left            =   720
         TabIndex        =   19
         Top             =   1680
         Width           =   2415
      End
      Begin VB.CheckBox Chkopt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow dir browsing"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   18
         Tag             =   "browsing"
         ToolTipText     =   "Allow browsing of this directory"
         Top             =   840
         Width           =   1815
      End
      Begin VB.ListBox lstmain 
         ForeColor       =   &H80000007&
         Height          =   1230
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblopt 
         BackStyle       =   0  'Transparent
         Caption         =   "Alias:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   31
         Top             =   255
         Width           =   855
      End
      Begin VB.Label lblopt 
         BackStyle       =   0  'Transparent
         Caption         =   "Path:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   1710
         Width           =   495
      End
   End
   Begin ASPWebServer.CinTitlebar Cinmain 
      Height          =   2775
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4895
      Caption         =   "Main Options"
      TitleBar        =   -1  'True
      Begin VB.CommandButton cmdbrowse 
         Caption         =   "..."
         Height          =   285
         Index           =   7
         Left            =   3240
         TabIndex        =   28
         ToolTipText     =   "Browse"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtopt 
         Height          =   285
         Index           =   6
         Left            =   1080
         TabIndex        =   27
         Tag             =   "logfile"
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton cmdbrowse 
         Caption         =   "Set"
         Height          =   285
         Index           =   6
         Left            =   3240
         TabIndex        =   25
         ToolTipText     =   "Browse"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtopt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   2040
         TabIndex        =   24
         Tag             =   "localport"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdstart 
         Caption         =   "Start"
         Height          =   375
         Left            =   600
         TabIndex        =   9
         Top             =   2280
         Width           =   2535
      End
      Begin VB.CheckBox Chkopt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cache converted ASP files"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Tag             =   "keepasneos"
         Top             =   2040
         Width           =   3375
      End
      Begin VB.CheckBox Chkopt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow browsing outside aliased directories"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Tag             =   "Outside"
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txtopt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   5
         Tag             =   "localport"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdbrowse 
         Caption         =   "Set"
         Height          =   285
         Index           =   1
         Left            =   3240
         TabIndex        =   4
         ToolTipText     =   "Browse"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtopt 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   2
         Tag             =   "logfile"
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton cmdbrowse 
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   3240
         TabIndex        =   1
         ToolTipText     =   "Browse"
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblopt 
         BackStyle       =   0  'Transparent
         Caption         =   "Error 404:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   29
         Top             =   765
         Width           =   855
      End
      Begin VB.Label lblopt 
         BackStyle       =   0  'Transparent
         Caption         =   "Max Connections:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   26
         Top             =   1455
         Width           =   1695
      End
      Begin VB.Label lblopt 
         BackStyle       =   0  'Transparent
         Caption         =   "Local Port:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   1095
         Width           =   975
      End
      Begin VB.Label lblopt 
         BackStyle       =   0  'Transparent
         Caption         =   "Log File:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "All events recorded will be logged here"
         Top             =   405
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cansetyet As Boolean

Private Sub Chkopt_Click(Index As Integer)
If Index = 4 Then echo "This only checks the file requested to see if it's up to date. If you use Includes alot, you should not use this"
If Index = 2 Or Index = 4 Then setmainoptions
If StrComp(txtopt(4), "WWWRoot", vbTextCompare) = 0 And (Index = 1 Or Index = 0) Then Chkopt(Index).value = IIf(Index = 0, vbUnchecked, vbChecked)
If Index = 0 Or Index = 1 Or Index = 3 Then cmdbrowse_Click 5
End Sub



Public Sub cmdbrowse_Click(Index As Integer)
Select Case Index 'all index's accounted for
    Case 0, 7, 8 'select logfile, 404, firewall
        InitOpen "*.*", "Please select a file", AliasPath("WWWRoot")
        cmdbrowse(0).tag = Open_File(Me.hWnd)
        If Len(cmdbrowse(0).tag) > 0 Then
            If Index = 7 Or Index = 8 Then Index = Index - 1
            txtopt(Index).text = cmdbrowse(0).tag
        End If
    Case 1, 6: setmainoptions 'change the main options
    Case 2 'add a default
        If Not itemexists(lstmain(0), txtopt(2)) Then
            lstmain(0).AddItem txtopt(2)
            listdefaults lstmain(0)
        End If
        txtopt(2) = Empty
    Case 3 'Browse for alias path
        txtopt(3).tag = BrowseForFolder(Me.hWnd, "Please select a folder for " & IIf(Len(txtopt(4)) > 0, txtopt(4), "the new alias"))
        If Len(txtopt(3).tag) > 0 Then txtopt(3).text = txtopt(3).tag
    Case 4: switchitems lstmain(0), lstmain(0).ListIndex, lstmain(0).ListIndex - 1 'move up priority list
    Case 5 'Set/add alias parameters
        If validalias(txtopt(4)) Then
            cmdbrowse(5).tag = getalias(txtopt(3))
            
            If cleardir(Not isadir(txtopt(3)), "The directory doesn't exist") Then Exit Sub
            If cleardir(Len(txtopt(3)) < 4, "Can not share the root directory of a drive. That is extremely unsafe.") Then Exit Sub
            If cleardir(Len(cmdbrowse(5).tag) > 0 And StrComp(txtopt(4), cmdbrowse(5).tag, vbTextCompare) <> 0, "Can not use " & txtopt(3) & " in more than one alias. It's already used in " & cmdbrowse(5).tag) Then Exit Sub
            
            'If StrComp(txtopt(4), "WWWRoot", vbTextCompare) = 0 And Chkopt(0).value = vbChecked Then Chkopt(0).value = vbUnchecked
            SetAlias txtopt(4), txtopt(3), Chkopt(0).value = vbChecked, Chkopt(1).value = vbChecked, Chkopt(3).value = vbChecked
            listaliases lstmain(1)
        Else
            echo txtopt(4) & " is a reserved alias or invalid. Please select another"
        End If
    Case 9: switchitems lstmain(0), lstmain(0).ListIndex, lstmain(0).ListIndex + 1 'move down priority list
    Case 10 'Remove selected default
        If lstmain(0).ListCount > 2 And lstmain(0).ListIndex > -1 Then
            lstmain(0).RemoveItem lstmain(0).ListIndex
            listdefaults lstmain(1), True
        End If
End Select
End Sub
Public Function cleardir(condition As Boolean, reason As String) As Boolean
    If condition Then
        echo reason
        cleardir = True
        txtopt(3) = Empty
    End If
End Function
Public Sub switchitems(list As ListBox, item1 As Long, item2 As Long)
    If item1 > -1 And item2 > -1 Then
        If item1 <= list.ListCount - 1 And item2 <= list.ListCount - 1 Then
            list.tag = list.list(item1)
            list.list(item1) = list.list(item2)
            list.list(item2) = list.tag
            list.tag = Empty
            listdefaults list
        End If
    End If
End Sub
Public Function itemexists(list As ListBox, item As String) As Boolean
    Dim temp As Long
    For temp = 0 To list.ListCount
        If StrComp(list.list(temp), item, vbTextCompare) = 0 Then
            itemexists = True
            Exit Function
        End If
    Next
End Function

Public Sub cmdstart_Click()
If allsetup Then
    If cmdstart.Caption = "Start" Then
        LISTEN Me.winsockServer, Val(getvar("LocalPort"))
        cmdstart.Caption = "Stop"
        GlobalExecute "Application_OnStart()", Me.scriptmain(0), clsASP
    Else
        cmdstart.Caption = "Start"
        winsockServer.Close
        winsockServer_Close
        GlobalExecute "Application_OnStop()", Me.scriptmain(0), clsASP
    End If
Else
    echo "Can not start server yet. Setup is incomplete or missing"
    echo "I require the following:"
    
    If Not AliasExists("WWWRoot") Then echo "You must create an alias named WWWRoot (Root directory)"
    If Not AliasExists("Neocache") Then echo "You must create an alias named Neocache (Storing converted ASP)"
    If Not AliasExists("Thumbcache") Then echo "You must create an alias named Thumbcache (Storing icons for directory browsing)"
    
    If Len(getvar("LocalPort")) = 0 Or Val(getvar("LocalPort")) <= 0 Then echo "Local port needs to be a number above zero"
    If Len(getvar("MaxConnections")) = 0 Or Val(getvar("MaxConnections")) <= 0 Then echo "Maximum Connections needs to be a number above zero"
    
    If Not fileexists(getvar("Error404")) Then echo "I require an error 404 page"
End If
End Sub


Private Sub Form_Load()
    Set txtmain = Me.txtdisplay 'used for echoing
    Set pic = picmain 'used for icon extraction
    Set filebox = Me.Filmain
    Set dirbox = Me.Dirmain
    initialize New ClsASPobjects, New clsMiniHini, Me.winsockServer, Me.winsockslaves, New clsGlobalASPObjects
    
    If Len(command) = 0 Then
        clsServer.LoadMini chkpath(App.path, "Server Settings.Nebula")
    Else
        clsServer.LoadMini getfromquotes(command)
    End If
    
    getmainoptions
    If lstmain(1).ListCount > 0 Then
        lstmain(1).ListIndex = 0
        lstmain_Click 1
    End If
    
    'onLoad Me, txtmain
    echo "Capabilities  : Standard HTML pages, Partial ASP compatability, GET and POST method, global.asa, Application, Directory Browsing"
    echo "Working on it : Cookies, Redirection, Session"
    
    cmdstart_Click
    
End Sub
Public Function getfromquotes(ByVal text As String) As String
    If Left(text, 1) = """" Then text = Right(text, Len(text) - 1)
    If Right(text, 1) = """" Then text = Left(text, Len(text) - 1)
    getfromquotes = text
End Function
Public Sub getmainoptions() 'get settings from minihini and put into the controls
    txtopt(0) = getvar("Logfile", chkpath(App.path, "logfile.txt"))
    txtopt(6) = getvar("Error404")
    txtopt(1) = getvar("LocalPort", 80)
    txtopt(5) = getvar("MaxConnections", 100)
    Chkopt(2) = txt2chk(getvar("AllowOutside", "False"))
    Chkopt(4) = txt2chk(getvar("NeoCache", "True"))
    listdefaults lstmain(0), False
    listaliases lstmain(1)
    cansetyet = True
    maxconnections = Val(getvar("MaxConnections", 100))
    cmdstart.Caption = IIf(Me.winsockServer.state = sckListening, "Stop", "Start")
End Sub
Public Sub setmainoptions() 'extract settings from controls, and put into minihini
If cansetyet Then
    setvar "Logfile", txtopt(0)
    setvar "Error404", txtopt(6)
    
    If IsNumeric(txtopt(1)) Then
        If Trim(txtopt(1)) <> getvar("LocalPort") Then
            setvar "LocalPort", txtopt(1)
            setlocalport Val(txtopt(1))
        End If
    Else
        echo "LocalPort has to be a number. " & txtopt(1) & " is not a number"
    End If
    
    If IsNumeric(txtopt(5)) Then
        setvar "MaxConnections", txtopt(5)
        maxconnections = Val(txtopt(5))
    Else
        echo "Maximum connections has to be a number. " & txtopt(5) & " is not a number"
    End If
    
    setvar "AllowOutside", Chkopt(2).value = vbChecked
    setvar "NeoCache", Chkopt(4).value = vbChecked
End If
End Sub
Public Function txt2chk(text As String) As Long
    txt2chk = IIf(text = "True", vbChecked, vbUnchecked)
End Function
Private Sub Form_Unload(Cancel As Integer)
    If cmdstart.Caption = "Stop" Then cmdstart_Click
    Savefile getvar("logfile"), txtdisplay, False
    
    If Len(command) = 0 Then
        clsServer.SaveMini chkpath(App.path, "Server Settings.Nebula")
    Else
        clsServer.SaveMini getfromquotes(command)
    End If
    'onUnload Me
End Sub

Private Sub lstmain_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 1 And KeyCode = 46 And lstmain(1).ListIndex > -1 Then
        Dim temp As String
        temp = lstmain(1).list(lstmain(1).ListIndex)
        clsServer.deleteitem "Alias", temp
        clsServer.DeleteSection temp
        lstmain(1).RemoveItem lstmain(1).ListIndex
        If lstmain(1).ListCount > 0 Then lstmain(1).ListIndex = 0
    End If
End Sub

Private Sub pichelp_Click(Index As Integer)
Select Case Index
    Case 0
        echo "General Assistance : Main Options"
        echo "Logfile        : Events recorded (In the big textbox on ther right) will be saved to this file upon closing the program"
        echo "Error 404      : When a visitor requests a file that doesn't exist or that they have no rights to, this file will be given to them instead"
        echo "Firewall       : A VBScript file with a function allowIP(ip as string) as boolean. When a visitor attempts to connect, this function will determine whether or not you let them"
        echo "Local Port     : Which port the server listens to. By default, browsers attempt to connect to 80"
        echo "Max Connections: How many visitors can connect at a given time"
        echo "Allow Browsing outside aliased directories : If a visitor attempts to browse outside the directories you've choosen, they'll be sent to the 404 page instead"
        echo "Cache converted ASP files : Technically, this program doesn't execute ASP files, but rather converts them into VBScript files. If cacheing is turned on, then this program only needs to convert when the original ASP is changed"
        echo "Start/Stop     : Toggle the status of your server"
    Case 1
        echo "General Assistance : Default Files"
        echo "When a visitor requests a directory, but doesn't give a filename, one must be choosen for them"
        echo "If directory allows browsing, a list of files will be given instead"
        echo "Otherwise, the first existing default file found will be given"
        echo "You can change the order using the arrow buttons, the higher on the list, the higher priority"
    Case 2
        echo "General Assistance : Directories"
        echo "When a visitor requests a file, the program has to find where it actually is on the hard drive."
        echo "Since the file can be in one of many directories, the program takes the part of the file before the first slash"
        echo "and checks to see if the alias exists. If there is no slash, or alias, then it gives the default alias (WWWRoot)"
        echo "WWWRoot    : is the default directory, and is required"
        echo "Neocache   : is where the converted ASP files are stored, and is required"
        echo "Thumbcache : is where thumbnails and file icons are stores, and is not required unless you allow directory browsing"
        echo "Main, Default, and Alias are reserved names and can not be used"
End Select
echo
End Sub

Private Sub winsockslaves_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
    Dim temp As String, temp2 As String, neocachetemp As String, temp3 As String, temp4 As Long
    winconns(Index).GetData temp
    temp2 = temp
    temp = Replace(temp, vbNewLine & vbNewLine, vbNewLine)
    temp = gettag("get", temp)
    temp3 = temp
    temp = MapPath(temp)   '       chkpath(homedir,)
    
    If isASP(temp) Or isNeo(temp) Then
        'echo "An asp file was requested (" & temp & ") The result will not be echoed"
        'The asp is saved to a file cause for some reason it wont sent right sent otherwise
        If 1 = 2 Then
        If getvar("NeoCache", "True") = "True" Then
            neocachetemp = neocache(temp)
            If StrComp(neocachetemp, temp, vbTextCompare) = 0 Then
            'check if the neocache is up to date. if not, update it,
                Savefile neocachetemp, ASP2NEO(loadfile(temp), temp)
            End If
            temp = neocachetemp 'give the neo cache instead
        End If
        End If
        
        temp4 = Me.scriptmain.UBound + 1
        Load scriptmain(temp4)
        temp2 = ExecuteASP(temp, winconns.item(Index), Me.scriptmain(temp4), clsASP, temp2, isASP(temp))
        Unload scriptmain(temp4)
        
        temp = uniquefilename(neo(temp))
        Savefile temp, temp2
        SendFile temp, winsockslaves.item(Index)
        DoEvents
        Kill temp
    Else
        'echo "An asp file was not requested"
        If isadir(temp) Then 'make a dir list
            temp2 = uniquefilename(chkpath(AliasPath("neocache"), "Dir.html"))
            Savefile temp2, generatefilelist(temp, dirbox, filebox, False, temp3)
            temp = temp2
            temp2 = "DELETEME"
        End If
        SendFile temp, winsockslaves.item(Index)
        If temp2 = "DELETEME" Then Kill temp
    End If
End Sub

Private Sub winsockslaves_Error(Index As Integer, ByVal number As Integer, description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
winsockslaves(Index).tag = False
End Sub

Private Sub winsockslaves_SendComplete(Index As Integer)
winsockslaves(Index).tag = True
End Sub

Public Sub winsockServer_Close()
    echo "Server was halted at " & Now
    cmdstart.Caption = "Start"
End Sub

Private Sub winsockServer_ConnectionRequest(ByVal requestID As Long)
    AcceptRequest requestID, Me.winsockslaves
End Sub

Private Sub winsockServer_Error(ByVal number As Integer, description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
winsockServer.tag = False
End Sub

Private Sub winsockServer_SendComplete()
winsockServer.tag = True
End Sub

Public Sub lstmain_Click(Index As Integer)
    Dim temp As Long
    If lstmain(1).ListIndex > -1 And Index = 1 Then
        temp = lstmain(1).ListIndex
        txtopt(4) = lstmain(1).list(temp)
        txtopt(3) = AliasPath(txtopt(4))
        Chkopt(0) = IIf(AliasProperty(txtopt(4), "AllowBrowse") = "True", vbChecked, vbUnchecked)
        Chkopt(1) = IIf(AliasProperty(txtopt(4), "AllowRemote") = "True", vbChecked, vbUnchecked)
        Chkopt(3) = IIf(AliasProperty(txtopt(4), "AllowAsp") = "True", vbChecked, vbUnchecked)
        
        Chkopt(0).Enabled = StrComp(txtopt(4), "WWWRoot", vbTextCompare) <> 0
        Chkopt(1).Enabled = Chkopt(0).Enabled
        lstmain(1).ListIndex = temp
    End If
End Sub

Private Sub txtopt_Change(Index As Integer)
Select Case Index
    Case 0, 6, 7: setmainoptions
    Case 4
        Chkopt(0).Enabled = StrComp(txtopt(4), "WWWRoot", vbTextCompare) <> 0
        Chkopt(1).Enabled = Chkopt(0).Enabled
End Select
End Sub
