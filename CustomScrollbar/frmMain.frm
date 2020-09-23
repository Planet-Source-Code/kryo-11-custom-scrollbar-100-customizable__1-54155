VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Scroll Example"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   ScaleHeight     =   264
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   169
   StartUpPosition =   3  'Windows Default
   Begin prjScrollbars.ctlScrollbar ctlScrollbar1 
      Height          =   1575
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2778
      StateOver       =   2
      Max             =   100
      TrackColor      =   -2147483628
   End
   Begin prjScrollbars.ctlScrollbar ctlScrollbar2 
      Height          =   1575
      Left            =   420
      TabIndex        =   1
      Top             =   60
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2778
      StateNormal     =   1
      StateDown       =   5
      Max             =   100
      TrackColor      =   -2147483628
   End
   Begin prjScrollbars.ctlScrollbar ctlScrollbar3 
      Height          =   1575
      Left            =   780
      TabIndex        =   2
      Top             =   60
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2778
      StateNormal     =   2
      StateDown       =   5
      Max             =   100
      TrackColor      =   -2147483628
   End
   Begin prjScrollbars.ctlScrollbar ctlScrollbar4 
      Height          =   1575
      Left            =   1140
      TabIndex        =   3
      Top             =   60
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2778
      StateNormal     =   3
      StateDown       =   5
      Max             =   100
      TrackColor      =   -2147483628
   End
   Begin prjScrollbars.ctlScrollbar ctlScrollbar5 
      Height          =   1575
      Left            =   1500
      TabIndex        =   4
      Top             =   60
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2778
      StateNormal     =   4
      StateDown       =   5
      Max             =   100
      TrackColor      =   -2147483628
   End
   Begin prjScrollbars.ctlScrollbar ctlScrollbar9 
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   1680
      Width           =   2415
      _ExtentX        =   3625
      _ExtentY        =   1614
      StateNormal     =   3
      Max             =   100
      TrackColor      =   -2147483628
      Orientation     =   1
   End
   Begin prjScrollbars.ctlScrollbar ctlScrollbar11 
      Height          =   1575
      Left            =   1860
      TabIndex        =   6
      Top             =   60
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2778
      StateNormal     =   5
      StateDown       =   5
      Max             =   100
      TrackColor      =   -2147483628
   End
   Begin prjScrollbars.ctlScrollbar ctlScrollbar6 
      Height          =   1575
      Left            =   2220
      TabIndex        =   7
      Top             =   60
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2778
      StateNormal     =   5
      StateDown       =   5
      Max             =   100
      TrackColor      =   -2147483628
      UsePictureUp    =   -1  'True
      UsePictureDown  =   -1  'True
      UsePictureThumb =   -1  'True
      UpButtonNormal  =   "frmMain.frx":0000
      UpButtonHover   =   "frmMain.frx":03C6
      UpButtonDown    =   "frmMain.frx":078C
      DownButtonNormal=   "frmMain.frx":0B52
      DownButtonHover =   "frmMain.frx":0F18
      DownButtonDown  =   "frmMain.frx":12DE
      ThumbButtonNormal=   "frmMain.frx":16A4
      ThumbButtonHover=   "frmMain.frx":1A6A
      ThumbButtonDown =   "frmMain.frx":1E30
      BackgroundPicture=   "frmMain.frx":21F6
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   $"frmMain.frx":22E4
      Height          =   1950
      Left            =   60
      TabIndex        =   8
      Top             =   1980
      Width           =   2445
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    Me.Cls
    Me.Print "Min = " & ctlScrollbar1.Min
    Me.Print "Max = " & ctlScrollbar1.Max
    Me.Print "Value = " & ctlScrollbar1.Value
End Sub
