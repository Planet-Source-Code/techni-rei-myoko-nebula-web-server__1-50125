VERSION 5.00
Begin VB.UserControl CinTitlebar 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1980
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1560
   ScaleWidth      =   1980
   ToolboxBitmap   =   "CinTitlebar.ctx":0000
   Begin VB.Label lblcaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cinaria"
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   15
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Image imgmain 
      Height          =   315
      Index           =   11
      Left            =   240
      Picture         =   "CinTitlebar.ctx":0312
      Stretch         =   -1  'True
      Top             =   15
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgmain 
      Height          =   120
      Index           =   10
      Left            =   240
      Picture         =   "CinTitlebar.ctx":034A
      Stretch         =   -1  'True
      Top             =   255
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image imgmain 
      Height          =   330
      Index           =   5
      Left            =   1080
      Picture         =   "CinTitlebar.ctx":03A0
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgmain 
      Height          =   120
      Index           =   0
      Left            =   240
      Picture         =   "CinTitlebar.ctx":03D8
      Stretch         =   -1  'True
      Top             =   -30
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape shpsec 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      Top             =   15
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image imgmain 
      Height          =   1275
      Index           =   6
      Left            =   1920
      Picture         =   "CinTitlebar.ctx":042E
      Stretch         =   -1  'True
      Top             =   240
      Width           =   45
   End
   Begin VB.Image imgmain 
      Height          =   120
      Index           =   9
      Left            =   1920
      Picture         =   "CinTitlebar.ctx":0466
      Top             =   1440
      Width           =   45
   End
   Begin VB.Image imgmain 
      Height          =   120
      Index           =   8
      Left            =   120
      Picture         =   "CinTitlebar.ctx":04B6
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Image imgmain 
      Height          =   120
      Index           =   7
      Left            =   0
      Picture         =   "CinTitlebar.ctx":0501
      Top             =   1440
      Width           =   135
   End
   Begin VB.Image imgmain 
      Height          =   1200
      Index           =   4
      Left            =   0
      Picture         =   "CinTitlebar.ctx":05B7
      Stretch         =   -1  'True
      Top             =   240
      Width           =   135
   End
   Begin VB.Image imgmain 
      Height          =   120
      Index           =   3
      Left            =   1920
      Picture         =   "CinTitlebar.ctx":0628
      Top             =   120
      Width           =   45
   End
   Begin VB.Image imgmain 
      Height          =   120
      Index           =   2
      Left            =   120
      Picture         =   "CinTitlebar.ctx":0679
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1815
   End
   Begin VB.Image imgmain 
      Height          =   120
      Index           =   1
      Left            =   0
      Picture         =   "CinTitlebar.ctx":06CF
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape shpmain 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   1335
      Left            =   45
      Top             =   165
      Width           =   1905
   End
End
Attribute VB_Name = "CinTitlebar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim showtitlebar As Boolean
Public Property Let Caption(text As String)
    If text <> lblcaption.Caption Then
        lblcaption.tag = text
        If Len(text) = 0 Then TitleBar = False Else Refresh
    End If
End Property
Public Property Get Caption() As String
    Caption = lblcaption.tag
End Property
Public Property Let TitleBar(Visible As Boolean)
    If Visible <> showtitlebar Then
        showtitlebar = Visible
        lblcaption.Visible = Visible
        imgmain(0).Visible = Visible
        imgmain(5).Visible = Visible
        imgmain(10).Visible = Visible
        imgmain(11).Visible = Visible
        shpsec.Visible = Visible
        Refresh
    End If
End Property
Public Property Get TitleBar() As Boolean
    TitleBar = showtitlebar
End Property
Public Property Get Backcolor() As Long
    Backcolor = shpmain.Backcolor
End Property
Public Property Let Backcolor(Color As Long)
    shpmain.Backcolor = Color
    shpsec.Backcolor = Color
End Property
Public Sub Refresh()
With UserControl
    lblcaption.Caption = trimto(lblcaption.tag, .width - 660)
    imgmain(2).width = .width - 165
    imgmain(8).width = imgmain(2).width
    imgmain(3).Left = .width - 60
    imgmain(9).Left = imgmain(3).Left
    imgmain(6).Left = imgmain(3).Left
    shpmain.width = .width - 90
    imgmain(7).Top = .height - 120
    imgmain(8).Top = imgmain(7).Top
    imgmain(9).Top = imgmain(7).Top
    
    imgmain(1).Top = IIf(showtitlebar, 120, 0) 'top corner
    imgmain(2).Top = imgmain(1).Top ' top bar
    imgmain(3).Top = imgmain(1).Top  'right corner
    
    imgmain(4).Top = IIf(showtitlebar, 240, 120) 'left side
    imgmain(6).Top = imgmain(4).Top 'right side
    
    imgmain(4).height = .height - IIf(showtitlebar, 360, 240) 'left side
    imgmain(6).height = .height - IIf(showtitlebar, 285, 165) 'right side
        
    shpmain.Top = IIf(showtitlebar, 165, 45) 'middle
    shpmain.height = .height - IIf(showtitlebar, 225, 105)
    
    imgmain(5).Left = lblcaption.width + lblcaption.Left + 75
    imgmain(0).width = imgmain(5).Left - imgmain(0).Left
    imgmain(10).width = imgmain(0).width
    shpsec.width = imgmain(0).width
End With
End Sub

Private Sub UserControl_Resize()
    Refresh
End Sub

Private Function trimto(text As String, width As Long) As String
    Dim temp As Long
    temp = Len(text)
    trimto = text
    If TextWidth(text) > width Then
        Do Until temp = 0 Or TextWidth(Left(text, temp) & "...") <= width
            temp = temp - 1
        Loop
        If temp = 0 Then trimto = Empty Else trimto = Left(text, temp) & "..."
    End If
End Function

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lblcaption.tag = PropBag.ReadProperty("Caption", UserControl.name)
    TitleBar = PropBag.ReadProperty("TitleBar", False)
    Backcolor = PropBag.ReadProperty("BackColor", vbWhite)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", lblcaption.tag, UserControl.name)
    Call PropBag.WriteProperty("TitleBar", showtitlebar, False)
    Call PropBag.WriteProperty("Backcolor", shpmain.Backcolor, vbWhite)
End Sub

