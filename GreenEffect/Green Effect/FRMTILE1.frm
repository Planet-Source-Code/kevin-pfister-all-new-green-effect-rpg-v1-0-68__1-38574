VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmGreenEffect 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Green Effect"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   Icon            =   "FRMTILE1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   474
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBlack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   9360
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   113
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1815
      Left            =   0
      TabIndex        =   107
      Top             =   4920
      Width           =   7095
      Begin VB.Label LblMessage 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to Green Effect"
         BeginProperty Font 
            Name            =   "Lucida Blackletter"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   109
         Top             =   180
         Width           =   6675
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         Height          =   1635
         Left            =   120
         Top             =   120
         Width           =   6915
      End
      Begin VB.Label LblText 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "This rpg was created by kevin pfister, in 2002. please vote for this program and have fun playing."
         BeginProperty Font 
            Name            =   "Lucida Blackletter"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1095
         Left            =   240
         TabIndex        =   108
         Top             =   480
         Width           =   6675
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4935
      Left            =   5040
      TabIndex        =   99
      Top             =   0
      Width           =   2055
      Begin MSComDlg.CommonDialog CD1 
         Left            =   1320
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblmenu 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Menu"
         BeginProperty Font 
            Name            =   "Lucida Blackletter"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   112
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Lucida Blackletter"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   111
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label LblItems 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Item(s)"
         BeginProperty Font 
            Name            =   "Lucida Blackletter"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   110
         Top             =   480
         Width           =   1575
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         Height          =   4695
         Left            =   120
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label LblPos 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Lucida Blackletter"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   300
         TabIndex        =   106
         Top             =   3540
         Width           =   1575
      End
      Begin VB.Label LblAble 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Lucida Blackletter"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   105
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Shape ProgressBack 
         BorderColor     =   &H0000C000&
         Height          =   195
         Index           =   1
         Left            =   300
         Top             =   3300
         Width           =   1575
      End
      Begin VB.Shape ProgressFore 
         BorderColor     =   &H0000FFFF&
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   1
         Left            =   300
         Top             =   3300
         Width           =   795
      End
      Begin VB.Shape ProgressBack 
         BorderColor     =   &H0000C000&
         Height          =   195
         Index           =   2
         Left            =   300
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Shape ProgressFore 
         BorderColor     =   &H0000FFFF&
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   2
         Left            =   300
         Top             =   2760
         Width           =   795
      End
      Begin VB.Label LblArmour 
         BackStyle       =   0  'Transparent
         Caption         =   "Armour"
         BeginProperty Font 
            Name            =   "Lucida Blackletter"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   300
         TabIndex        =   104
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Shape ProgressBack 
         BorderColor     =   &H0000C000&
         Height          =   195
         Index           =   3
         Left            =   300
         Top             =   2220
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Weapon"
         BeginProperty Font 
            Name            =   "Lucida Blackletter"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   300
         TabIndex        =   103
         Top             =   2460
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Health"
         BeginProperty Font 
            Name            =   "Lucida Blackletter"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   300
         TabIndex        =   102
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Shape ProgressFore 
         BorderColor     =   &H0000FFFF&
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   3
         Left            =   300
         Top             =   2220
         Width           =   795
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Castra(s)"
         BeginProperty Font 
            Name            =   "Lucida Blackletter"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   300
         TabIndex        =   101
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblMoney 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Lucida Blackletter"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   300
         TabIndex        =   100
         Top             =   1620
         Width           =   1515
      End
   End
   Begin VB.PictureBox PBD 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   7320
      Picture         =   "FRMTILE1.frx":0442
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   98
      Top             =   8640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PBD 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   6720
      Picture         =   "FRMTILE1.frx":1168
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   97
      Top             =   8640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PBU 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   7320
      Picture         =   "FRMTILE1.frx":1E8E
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   96
      Top             =   8040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PBU 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   6720
      Picture         =   "FRMTILE1.frx":2BB4
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   95
      Top             =   8040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PBR 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   7320
      Picture         =   "FRMTILE1.frx":38DA
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   94
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PBL 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   7320
      Picture         =   "FRMTILE1.frx":4600
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   93
      Top             =   6840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PBR 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   6720
      Picture         =   "FRMTILE1.frx":5326
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   92
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PBL 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   6720
      Picture         =   "FRMTILE1.frx":604C
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   91
      Top             =   6840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   7800
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   337
      TabIndex        =   90
      Top             =   5640
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.PictureBox PicSupportRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1320
      Picture         =   "FRMTILE1.frx":6D72
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   89
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicMachine 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   7800
      Picture         =   "FRMTILE1.frx":7190
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   88
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox PicGrass 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   7170
      Picture         =   "FRMTILE1.frx":7E62
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   87
      Top             =   120
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicSand 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   7170
      Picture         =   "FRMTILE1.frx":838E
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   86
      Top             =   660
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicWater 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   7170
      Picture         =   "FRMTILE1.frx":8894
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   85
      Top             =   1200
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicTree 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   7710
      Picture         =   "FRMTILE1.frx":8D0C
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   84
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicFlowers 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   7710
      Picture         =   "FRMTILE1.frx":925B
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   83
      Top             =   660
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicBotLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   7170
      Picture         =   "FRMTILE1.frx":9782
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   82
      Top             =   2820
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicStopleft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   7710
      Picture         =   "FRMTILE1.frx":9CB6
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   81
      Top             =   2820
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicStopRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   8250
      Picture         =   "FRMTILE1.frx":A1E0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   80
      Top             =   660
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicBotRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   7170
      Picture         =   "FRMTILE1.frx":A706
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   79
      Top             =   3360
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicFenceLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   7710
      Picture         =   "FRMTILE1.frx":AC3F
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   78
      Top             =   3360
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicAcross 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   8250
      Picture         =   "FRMTILE1.frx":B180
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   77
      Top             =   120
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicFenceRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   7170
      Picture         =   "FRMTILE1.frx":B6B1
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   76
      Top             =   1740
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicStopLeftUp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   7170
      Picture         =   "FRMTILE1.frx":BBEC
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   75
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicStopRightUp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   7710
      Picture         =   "FRMTILE1.frx":C12B
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   74
      Top             =   2280
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicTopRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   7710
      Picture         =   "FRMTILE1.frx":C667
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   73
      Top             =   1200
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicTopLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   7710
      Picture         =   "FRMTILE1.frx":CB94
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   72
      Top             =   1740
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicRock 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   8250
      Picture         =   "FRMTILE1.frx":D0BE
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   71
      Top             =   1200
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicCom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   7200
      Picture         =   "FRMTILE1.frx":D554
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   70
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox PicDead 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   8250
      Picture         =   "FRMTILE1.frx":DA32
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   69
      Top             =   1740
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicDead 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   8250
      Picture         =   "FRMTILE1.frx":DE55
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   68
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicFire 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   8250
      Picture         =   "FRMTILE1.frx":E275
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   67
      Top             =   2820
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicFade 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   8790
      Picture         =   "FRMTILE1.frx":E70F
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   66
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicFade 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   8790
      Picture         =   "FRMTILE1.frx":EB62
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   65
      Top             =   660
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicFade 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   8790
      Picture         =   "FRMTILE1.frx":EFB9
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   64
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicFade 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   8790
      Picture         =   "FRMTILE1.frx":F40F
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   63
      Top             =   1740
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicFade 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   8790
      Picture         =   "FRMTILE1.frx":F860
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   62
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicFade 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   5
      Left            =   8790
      Picture         =   "FRMTILE1.frx":FCAA
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   61
      Top             =   2820
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicMud 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   8250
      Picture         =   "FRMTILE1.frx":100F6
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   60
      Top             =   3360
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicMud 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   8790
      Picture         =   "FRMTILE1.frx":1058B
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   59
      Top             =   3360
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicField 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   7170
      Picture         =   "FRMTILE1.frx":10A20
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   58
      Top             =   3900
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicField 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   7710
      Picture         =   "FRMTILE1.frx":10E9A
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   57
      Top             =   3900
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicLGrass 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   8250
      Picture         =   "FRMTILE1.frx":11314
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   56
      Top             =   3900
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicLGrass 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   8790
      Picture         =   "FRMTILE1.frx":118A3
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   55
      Top             =   3900
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicRoof 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   7170
      Picture         =   "FRMTILE1.frx":11E32
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   54
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicGrass 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   120
      Picture         =   "FRMTILE1.frx":1225D
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   53
      Top             =   6840
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicPath 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   720
      Picture         =   "FRMTILE1.frx":12789
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   52
      Top             =   6840
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicSand 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   120
      Picture         =   "FRMTILE1.frx":12BC3
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   51
      Top             =   7440
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicRock 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   720
      Picture         =   "FRMTILE1.frx":130C9
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   50
      Top             =   7440
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicWater 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   1320
      Picture         =   "FRMTILE1.frx":1355F
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   49
      Top             =   6840
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicTree 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   1920
      Picture         =   "FRMTILE1.frx":139D7
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   48
      Top             =   6840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicWell 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1920
      Picture         =   "FRMTILE1.frx":13F26
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   47
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicTopLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   2520
      Picture         =   "FRMTILE1.frx":14469
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   46
      Top             =   8040
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicTopRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   3120
      Picture         =   "FRMTILE1.frx":14993
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   45
      Top             =   8040
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicStopRightUp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   2520
      Picture         =   "FRMTILE1.frx":14EC0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   44
      Top             =   8640
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicChest 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   3120
      Picture         =   "FRMTILE1.frx":153FC
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   43
      Top             =   8640
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicStopLeftUp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   1920
      Picture         =   "FRMTILE1.frx":158F4
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   42
      Top             =   8640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicFenceRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   1920
      Picture         =   "FRMTILE1.frx":15E33
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   41
      Top             =   8040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicAcross 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   1320
      Picture         =   "FRMTILE1.frx":1636E
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   40
      Top             =   8040
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicFenceLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   720
      Picture         =   "FRMTILE1.frx":1689F
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   39
      Top             =   8640
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicBotRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   120
      Picture         =   "FRMTILE1.frx":16DE0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   38
      Top             =   8640
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicStopRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   1320
      Picture         =   "FRMTILE1.frx":17319
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   37
      Top             =   8640
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicStopleft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   720
      Picture         =   "FRMTILE1.frx":1783F
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   36
      Top             =   8040
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicBotLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   120
      Picture         =   "FRMTILE1.frx":17D69
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   35
      Top             =   8040
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicFlowers 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   3720
      Picture         =   "FRMTILE1.frx":1829D
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   34
      Top             =   8040
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   2520
      Picture         =   "FRMTILE1.frx":187C4
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   33
      Top             =   6840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   2520
      Picture         =   "FRMTILE1.frx":18C00
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   32
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   3120
      Picture         =   "FRMTILE1.frx":1904C
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   31
      Top             =   6840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   3120
      Picture         =   "FRMTILE1.frx":19499
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   30
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicUp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   3720
      Picture         =   "FRMTILE1.frx":198D0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   29
      Top             =   6840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicUp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   3720
      Picture         =   "FRMTILE1.frx":19D23
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   28
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicDown 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   4320
      Picture         =   "FRMTILE1.frx":1A176
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   27
      Top             =   6840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicDown 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   4320
      Picture         =   "FRMTILE1.frx":1A5DD
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   26
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicDirt 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3720
      Picture         =   "FRMTILE1.frx":1AA47
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   25
      Top             =   8640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicBrick 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4320
      Picture         =   "FRMTILE1.frx":1AE74
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   24
      Top             =   8040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicDoor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4320
      Picture         =   "FRMTILE1.frx":1B241
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   23
      Top             =   8640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicStool 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4920
      Picture         =   "FRMTILE1.frx":1B628
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   22
      Top             =   8040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicWindow 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4920
      Picture         =   "FRMTILE1.frx":1BA8E
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   21
      Top             =   8640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicPerson1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4920
      Picture         =   "FRMTILE1.frx":1BEE1
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   20
      Top             =   6840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicPerson2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4920
      Picture         =   "FRMTILE1.frx":1C332
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   19
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicSign 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5520
      Picture         =   "FRMTILE1.frx":1C788
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   18
      Top             =   6840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicBCase 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5520
      Picture         =   "FRMTILE1.frx":1CB9A
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   17
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicCase 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5520
      Picture         =   "FRMTILE1.frx":1D003
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   16
      Top             =   8040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicInn 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      Picture         =   "FRMTILE1.frx":1D40C
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   15
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicArmor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      Picture         =   "FRMTILE1.frx":1D859
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   14
      Top             =   6840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicCarpetTopRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4920
      Picture         =   "FRMTILE1.frx":1DCC9
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   13
      Top             =   9240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicCarpetTopLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4320
      Picture         =   "FRMTILE1.frx":1E17C
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   12
      Top             =   9240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicCarpetTop 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3720
      Picture         =   "FRMTILE1.frx":1E635
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   11
      Top             =   9240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicBottomLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   1320
      Picture         =   "FRMTILE1.frx":1EA6D
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   10
      Top             =   9240
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicCarpet 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   120
      Picture         =   "FRMTILE1.frx":1EF2F
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   9
      Top             =   9240
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicCarpetBottom 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   720
      Picture         =   "FRMTILE1.frx":1F28C
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   8
      Top             =   9240
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicBottomRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1920
      Picture         =   "FRMTILE1.frx":1F6BC
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   7
      Top             =   9240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicCarpetRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   3120
      Picture         =   "FRMTILE1.frx":1FB7C
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   6
      Top             =   9240
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicCarpetLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   2520
      Picture         =   "FRMTILE1.frx":1FFD1
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   5
      Top             =   9240
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicStepsLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5520
      Picture         =   "FRMTILE1.frx":2042F
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   4
      Top             =   8640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicSteps 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      Picture         =   "FRMTILE1.frx":2089A
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   3
      Top             =   8040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicStepsRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      Picture         =   "FRMTILE1.frx":20C20
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   2
      Top             =   8640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicBed 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5520
      Picture         =   "FRMTILE1.frx":2108C
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   1
      Top             =   9240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicSupportLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      Picture         =   "FRMTILE1.frx":214F5
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   0
      Top             =   9240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer TmrDay 
      Interval        =   60000
      Left            =   4560
      Top             =   600
   End
   Begin VB.Timer TmrPlayer 
      Interval        =   75
      Left            =   4560
      Top             =   120
   End
   Begin VB.Menu mnupopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuGE 
         Caption         =   "Green Effect Menu"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnusave 
         Caption         =   "Save Game"
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "Load Game"
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuoptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "FrmGreenEffect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Green Effect RPG By Kevin Pfister
'~FrmGreenEffect~

'Green Effect is not currently completed, if you can think of any suggestions or ways
'of improving the game, please send me an email at: Yet_Another_Idiot@Hotmail.com
'I have spent a long time working on this game and hope to continue, please vote
'or leave comments for this program to show how well recieved and how much you would
'like a new version to be made.

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As Integer
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Type GEHouse
    Name As String
    X As Integer
    Y As Integer
    Key As Integer
    KeyName As String
End Type

Private Type TMonster
    Anim As Integer
    X As Integer
    Y As Integer
    LastAction As Integer
    TrackNodes(1 To 10, 1 To 10) As Integer
    NodesToDo(1 To 10, 1 To 10) As Integer
    WayPointsX() As Integer
    WayPointsY() As Integer
    WayPoints As Integer
    CurrentWayPoint As Integer
    WayPointsOn As Boolean
End Type

Private Type PUObj
    Name As String
    Desc As String
    Extra As String
    Type As String
End Type

Dim ObjGrid(1 To 200, 1 To 200) As PUObj

Dim House() As GEHouse
Dim Houses As Integer
Dim MapHouse(1 To 200, 1 To 200) As Integer

Dim Str_UnCompressedMap(1 To 200) As String                'The Map array
Dim Str_CompressedMap(1 To 200) As String           'The Map array

Dim Str_Map(1 To 10, 1 To 10) As String         'The Current Display Int_TrackNodes
Dim Str_ExtraNames(1 To 250) As String          'The Characters Name
Dim Str_ExtraText(1 To 250) As String          'What the characters say
Dim Str_HouseName(1 To 250) As String           'The name of the House (Not really Needed)
    
Dim Int_ObjectLocation(1 To 10, 1 To 10) As Boolean
Dim Int_PassableLand(1 To 10, 1 To 10) As Integer
Dim Lng_AlphaBlendDc(1 To 10, 1 To 10) As Long
Dim Lng_MapDcs(1 To 10, 1 To 10) As Long
Dim Int_ExtraPlayers(1 To 200, 1 To 200) As Integer   'The Player array
Dim Int_ExtraPlayerType(1 To 800) As Integer         'The Characters Type
Dim Int_ExtraPlayerX(1 To 250) As Integer              'The X Position of the Character
Dim Int_ExtraPlayerY(1 To 250) As Integer             'The Y Position of the Character
Dim BuildingX As Integer    'The Current X Building
Dim BuildingY As Integer    'The Current Y Building

Dim Color As Long
Dim Color2 As Long            'The Colours
Dim R As Integer, G As Integer, B As Integer    'The (R)ed (G)reen (B)lue values
Dim R1 As Integer, G1 As Integer, B1 As Integer
Dim R2 As Integer, G2 As Integer, B2 As Integer     'The Second (R)ed (G)reen (B)lue values

Dim ScreenX As Integer      'The Current X Screen
Dim ScreenY As Integer      'The Current Y Screen
Dim PlayerX As Integer      'The Current X Pos of the player
Dim PlayerY As Integer      'The Current Y Pos of the Player
Dim PlayerProgress As Integer
Dim Anim As Integer         'The Current Animation of the Player
Dim InBuilding As Boolean   'Is the Player in a building?
Dim LastAction

Dim EnterX As Integer
Dim EnterY As Integer

Dim Int_TrackNodes(1 To 10, 1 To 10) As Integer
Dim Int_NodesToDo(1 To 10, 1 To 10) As Integer
Dim Int_EndTrackX As Integer
Dim Int_EndTrackY As Integer
Dim Int_CurrentTrackX As Integer
Dim Int_CurrentTrackY As Integer
Dim Bln_TrackInProgress As Boolean
Dim Int_MonsterAnim As Integer
Dim Int_MonsterLastPos As Integer

Dim Day As Integer
Dim AIDelayNo As Integer

Public Lng_Modifer As Long
Public NiceGradient As Boolean     'Show the Nice AlphaBlendDcing of the Tiles?
Public AIDelay As Integer

Private Sub Form_Load()
    On Error Resume Next

    Dim StrText As String
    
    Castras = 100
    PlayerHealth = 100
    PlayerProgress = 1
    Day = 0             'Set the value of the day (0 = Day , 1 = Night)
    Lng_Modifer = 2     'Route Modifer(Higher = More Random)
    Call Startup_WAA    'Start up the (W)eapons (A)nd (A)rmour
    LastAction = 2      'Set the last action (Up Down Left Right)
    Randomize Timer     'Used in the Random Chat Sub
    PlayerX = 4         'The Default X Position of the player
    PlayerY = 4         'The Default Y Position of the player
    ScreenX = 1         'The Default X Screen Position
    ScreenY = 1         'The Default Y Screen Position
    Anim = 1            'The Default Animation image
    InBuilding = False      'Set the character not to be in a building
    EnterX = PlayerX        'Set the X point of entry
    EnterY = PlayerY        'Set the Y point of entry
    AIDelay = 2
    
    'These try to make the form visible
    FrmGreenEffect.Show
    FrmGreenEffect.Visible = True
    
    NiceGradient = True 'False = Faster less graphics... True = Slower Better Grpahics
    Call TempMonsters
    
    Call DoTimeTravel   'This is the startup, it shows the player travelling through the time machine
    
    Keys(1).KeyVal = 1
    Keys(1).Name = "General Key"
    KeyCount = 1
    
    'The following are the codes for the map
    'A = Wall
    'B = Door
    'C = Chest
    'D = Dirt
    'E = Grass Rocks
    'F = Flowers
    'G = Grass
    'H = Stool
    'I = Window
    'J = Person1
    'K = Person2
    'L = Book Case
    'M = Case
    'N = Carpet
    'O = Carpet Top Left
    'P = Path
    'Q = Carpet Top
    'R = Rock
    'S = Sand
    'T = Tree
    'U = Carpet Top Right
    'V = Carpet Right
    'W = Water
    'X = Carpet Bottom
    'Y = Carpet Bottom Left
    'Z = Carpet Left
    '1 = Bottom Left Fence
    '2 = Bottom Right Fence
    '3 = Stop at left Fence
    '4 = Left Fence
    '5 = Horizontal Fence
    '6 = Stop at right fence
    '7 = Right Fence
    '8 = Stop at Upper Left
    '9 = Top Left Fence
    '0 = Stop at Upper Right
    '/ = Top Right Fence
    '! = Steps Left
    '% = Steps
    '£ = Steps Right
    '$ = Bed
    '^ = Carpet Bottom Right
    '& = Support Left
    '* = Support Right
    
    '... there are more codes
    
    
    Call OutofBuilding  'The computer believes that the time travel was in a building this unloads it and loads the normal map

    'Open the Hus File to load the positions of the doors
    Open App.Path + "\GE\GEffect.Hus" For Input As #1
    'Contains the locations of the entry points of the doors
    Input #1, StrText  'Input the number of doors
    ToNum = Val(StrText)   'Set the Number
    ReDim House(1 To ToNum)
    'Loop to collect the door info
    For Extras = 1 To ToNum
        'Collect the three pieces of Data
        For GetInfo = 1 To 5
            Input #1, StrText  'input data
            If Mid(StrText, 1, 6) = "##Name" Then
                House(Extras).Name = Mid(StrText, 8)     'This is not really need, may use this later?
            ElseIf Mid(StrText, 1, 6) = "##XPos" Then
               House(Extras).X = Val(Mid(StrText, 8))    'Save X position
            ElseIf Mid(StrText, 1, 6) = "##YPos" Then
                House(Extras).Y = Val(Mid$(StrText, 8))   'Save Y Position
            ElseIf Mid(StrText, 1, 5) = "##Key" Then
                MapHouse(House(Extras).X, House(Extras).Y) = Extras
                House(Extras).Key = Val(Mid$(StrText, 7))
            ElseIf Mid(StrText, 1, 7) = "##KName" Then
                House(Extras).KeyName = Mid(StrText, 9)
            End If
        Next
    Next
    
    Close
    
    PlayerX = 5 'Set the X position of the player
    PlayerY = 3 'Set the Y position of the player
    Call RenderMap   'Call the main Drawing Engine
    Call DoFadeIn   'Make the Player Fade In
    Anim = 1 - Anim 'Invert the animation
    Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PBU(Anim).hDC, 0, 0, vbSrcPaint)
    Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PicUp(Anim).hDC, 0, 0, vbSrcAnd)
    FrmGreenEffect.Refresh
End Sub

Sub DoProgress()
    On Error Resume Next

    If ProgressFore(1).Width <> Int((ProgressBack(1).Width / Armour.max) * Armour.current) Then
        ProgressFore(1).Width = Int((ProgressBack(1).Width / Armour.max) * Armour.current)
    End If
    If Weapon(PlayerWeapon).max <> -1 Then
        If ProgressFore(2).Width <> Int((ProgressBack(2).Width / Weapon(PlayerWeapon).max) * Weapon(PlayerWeapon).current) Then
            ProgressFore(2).Width = Int((ProgressBack(2).Width / Weapon(PlayerWeapon).max) * Weapon(PlayerWeapon).current)
        End If
    Else
        If ProgressFore(2).Width <> Int((ProgressBack(2).Width / 100) * 100) Then
            ProgressFore(2).Width = Int((ProgressBack(2).Width / 100) * 100)
        End If
    End If
    If ProgressFore(3).Width <> Int((ProgressBack(3).Width / 100) * PlayerHealth) Then
        ProgressFore(3).Width = Int((ProgressBack(3).Width / 100) * PlayerHealth)
    End If
    If lblMoney.Caption <> Castras Then
        lblMoney.Caption = Castras
    End If
End Sub

Sub RenderMap()     'The Drawing Part, paints the images in the Map array to the Screen
    On Error Resume Next
    
    '--------------------------------------------------
    'Description:
    'The main tile engine, draws the tiles to the back
    'buffer. Also can alpha blend the tiles together
    'giving a more realistic effect.
    '--------------------------------------------------
    
    Dim InnerLoop As Integer
    Dim OuterLoop As Integer
    Dim OutFade As Integer
    Dim InFade As Integer
    Dim TempDc As Long
    Dim Dc1 As Long
    Dim Dc2 As Long
    Dim X As Integer
    Dim Y As Integer
    Dim Colour As Long
    Dim Colour1 As Long
    Dim Percent As Integer
    Dim InOuterLoop As Integer
    Dim InInnerLoop As Integer
    
    
    TmrPlayer.Enabled = False   'Stop the Player Movement Timer
    PicBuffer.Cls   'Clear the backbuffer
    
    For OuterLoop = 1 To 10
        For InnerLoop = 1 To 10
            
            Str_Map(InnerLoop, OuterLoop) = Mid$(Str_UnCompressedMap(ScreenY * 10 - 10 + OuterLoop), ScreenX * 10 - 10 + InnerLoop, 1) 'Mid$(Map(OuterLoop), InnerLoop, 1)
            Int_ObjectLocation(InnerLoop, OuterLoop) = 0    'Clear random Money
            Int_PassableLand(InnerLoop, OuterLoop) = 0    'Clears Int_PassableLand objects
            
            'Below selects the right hdc's for the map code and also sets if it is impassable or not
            'TempDc is the Hdc of the tile
            'Lng_AlphaBlendDc is the Hdc of the tile which is for alpha blending,
            'instead of using the same tile it uses the one which the background is from:
            'ie. Tree and flowers both have a grass background
            
            Select Case Str_Map(InnerLoop, OuterLoop)
            Case "G"    'Grass
                TempDc = PicGrass(Day).hDC
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicGrass(Day).hDC
            Case "W"    'Water
                TempDc = PicWater(Day).hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 2
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicWater(Day).hDC
            Case "S"    'Sand
                TempDc = PicSand(Day).hDC
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicSand(Day).hDC
            Case "R"    'Rock
                TempDc = PicRock(Day).hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicRock(Day).hDC
            Case "T"    'Tree
                TempDc = PicTree(Day).hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicGrass(Day).hDC
            Case "E"    'Well
                TempDc = PicWell.hDC
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicGrass(Day).hDC
            Case "C"    'A Chest
                TempDc = PicChest.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicChest.hDC
            Case "P"    'Path
                TempDc = PicPath.hDC
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicPath.hDC
            Case "F"    'Flowers
                TempDc = PicFlowers(Day).hDC
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicGrass(Day).hDC
            Case "D"    'Dirt (Path)
                TempDc = PicDirt.hDC
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicDirt.hDC
            Case "A"    'Brick
                TempDc = PicBrick.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicBrick.hDC
            Case "B"    'Door
                TempDc = PicDoor.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicDoor.hDC
            Case "H"    'Stool
                TempDc = PicStool.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicPath.hDC
            Case "I"    'Window
                TempDc = PicWindow.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicWindow.hDC
            Case "1"    'Fence Bottom Left
                TempDc = PicBotLeft(Day).hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicGrass(Day).hDC
            Case "2"    'Fence Bottom Right
                TempDc = PicBotRight(Day).hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicGrass(Day).hDC
            Case "3"    'Fence Left Stop
                TempDc = PicStopleft(Day).hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicGrass(Day).hDC
            Case "4"    'Fence Vertical Left
                TempDc = PicFenceLeft(Day).hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicGrass(Day).hDC
            Case "5"    'Fence Horizontal
                TempDc = PicAcross(Day).hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicGrass(Day).hDC
            Case "6"    'Fence right stop
                TempDc = PicStopRight(Day).hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicGrass(Day).hDC
            Case "7"    'Fence vertical Right
                TempDc = PicFenceRight(Day).hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicGrass(Day).hDC
            Case "8"    'Fence Stop Vertical Left
                TempDc = PicStopLeftUp(Day).hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicGrass(Day).hDC
            Case "9"    'Fence Top Left Corner
                TempDc = PicTopLeft(Day).hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicGrass(Day).hDC
            Case "0"    'Fence Stop Vertical Right
                TempDc = PicStopRightUp(Day).hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicGrass(Day).hDC
            Case "/"    'Fence Top Right Corner
                TempDc = PicTopRight(Day).hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicGrass(Day).hDC
            Case "L"    'Bookcase
                TempDc = PicBCase.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicBCase.hDC
            Case "M"    'Case
                TempDc = PicCase.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicCase.hDC
            Case "N"    'Center Carpet
                TempDc = PicCarpet.hDC
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicCarpet.hDC
            Case "O"    'Top Left Carpet
                TempDc = PicCarpetTopLeft.hDC
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicCarpet.hDC
            Case "U"    'Top Right Carpet
                TempDc = PicCarpetTopRight.hDC
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicCarpet.hDC
            Case "V"    'Right Carpet
                TempDc = PicCarpetRight.hDC
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicCarpet.hDC
            Case "^"    'Bottom Right Carpet
                TempDc = PicBottomRight.hDC
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicCarpet.hDC
            Case "X"    'Bottom Horizontal Carpet
                TempDc = PicCarpetBottom.hDC
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicCarpet.hDC
            Case "Y"    'Bottom Left of carpet
                TempDc = PicBottomLeft.hDC
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicCarpet.hDC
            Case "Z"    'Left of carpet
                TempDc = PicCarpetLeft.hDC
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicCarpet.hDC
            Case "Q"    'Top Horizontal of carpet
                TempDc = PicCarpetTop.hDC
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicCarpet.hDC
            Case "!"    'Left Side of Steps
                TempDc = PicStepsLeft.hDC
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicStepsLeft.hDC
            Case "%"    'Center of Steps
                TempDc = PicSteps.hDC
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicSteps.hDC
            Case "£"    'Right side of steps
                TempDc = PicStepsRight.hDC
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicStepsRight.hDC
            Case "$"    'Bed
                TempDc = PicBed.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicBed.hDC
            Case "&"    'Building Support Left
                TempDc = PicSupportLeft.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicSupportLeft.hDC
            Case "*"    'Building Support Right
                TempDc = PicSupportRight.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicSupportRight.hDC
            Case "}"    'Building Roof
                TempDc = PicRoof.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicRoof.hDC
            Case ";"    'Long Grass
                TempDc = PicLGrass(Day).hDC
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicLGrass(Day).hDC
            Case ":"    'Field
                TempDc = PicField(Day).hDC
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicField(Day).hDC
            Case "'"    'Mud
                TempDc = PicMud(Day).hDC
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicMud(Day).hDC
            Case "["    'FirePlace
                TempDc = PicFire.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicFire.hDC
            Case "]"    'The Inn Sign
                TempDc = PicInn.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicInn.hDC
            Case "{"    'The Armoury sign
                TempDc = PicArmor.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = PicArmor.hDC
            Case "`"
                TempDc = picBlack.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 2
                Lng_AlphaBlendDc(InnerLoop, OuterLoop) = picBlack.hDC
            End Select
            'Once the information is gathered about its Hdc's and its Passablity
            'Draw it to the backbuffer
            Call BitBlt(PicBuffer.hDC, InnerLoop * 30 - 9, OuterLoop * 30 - 9, 30, 30, TempDc, 0, 0, vbSrcCopy)
            Lng_MapDcs(InnerLoop, OuterLoop) = TempDc
        Next
    Next
    
    Int_ObjectLocation(Rnd * 9 + 1, Rnd * 9 + 1) = 1 'Random Placement of Money

    If InBuilding = False And NiceGradient = True Then  'If there is need for the nice gradient then draw it
        For OutFade = 1 To 10   'Loop to Lng_AlphaBlendDc the colours of the Images
            For InFade = 1 To 10
                If InFade - 1 > 0 Then
                    If Lng_AlphaBlendDc(InFade, OutFade) <> Lng_AlphaBlendDc(InFade - 1, OutFade) Then
                        'This is the Vertical Fade Part
                        Dc1 = Lng_MapDcs(InFade - 1, OutFade)
                        Dc2 = Lng_MapDcs(InFade, OutFade)
                        For Y = 0 To 29
                            Colour = GetPixel(Dc1, 26, Y)
                            Colour1 = GetPixel(Dc2, 1, Y)
                            Call GetRgb(Colour, R1, G1, B1)
                            Call GetRgb(Colour1, R2, G2, B2)
                            Percent = 20
                            SetPixelV PicBuffer.hDC, InFade * 30 - 7, OutFade * 30 - 9 + Y, RGB(R2 - (R2 / 100) * Percent + (R1 / 100) * Percent, G2 - (G2 / 100) * Percent + (G1 / 100) * Percent, B2 - (B2 / 100) * Percent + (B1 / 100) * Percent)
                        Next
                        For Y = 0 To 29
                            Colour = GetPixel(Dc1, 27, Y)
                            Colour1 = GetPixel(Dc2, 2, Y)
                            Call GetRgb(Colour, R1, G1, B1)
                            Call GetRgb(Colour1, R2, G2, B2)
                            Percent = 40
                            SetPixelV PicBuffer.hDC, InFade * 30 - 8, OutFade * 30 - 9 + Y, RGB(R2 - (R2 / 100) * Percent + (R1 / 100) * Percent, G2 - (G2 / 100) * Percent + (G1 / 100) * Percent, B2 - (B2 / 100) * Percent + (B1 / 100) * Percent)
                        Next
                        For Y = 0 To 29
                            Colour = GetPixel(Dc1, 28, Y)
                            Colour1 = GetPixel(Dc2, 3, Y)
                            Call GetRgb(Colour, R1, G1, B1)
                            Call GetRgb(Colour1, R2, G2, B2)
                            Percent = 60
                            SetPixelV PicBuffer.hDC, InFade * 30 - 9, OutFade * 30 - 9 + Y, RGB(R2 - (R2 / 100) * Percent + (R1 / 100) * Percent, G2 - (G2 / 100) * Percent + (G1 / 100) * Percent, B2 - (B2 / 100) * Percent + (B1 / 100) * Percent)
                        Next
                        For Y = 0 To 29
                            Colour = GetPixel(Dc1, 29, Y)
                            Colour1 = GetPixel(Dc2, 4, Y)
                            Call GetRgb(Colour, R1, G1, B1)
                            Call GetRgb(Colour1, R2, G2, B2)
                            Percent = 80
                            SetPixelV PicBuffer.hDC, InFade * 30 - 10, OutFade * 30 - 9 + Y, RGB(R2 - (R2 / 100) * Percent + (R1 / 100) * Percent, G2 - (G2 / 100) * Percent + (G1 / 100) * Percent, B2 - (B2 / 100) * Percent + (B1 / 100) * Percent)
                        Next
                    End If
                End If
                If OutFade - 1 > 0 Then
                    If Lng_AlphaBlendDc(InFade, OutFade) <> Lng_AlphaBlendDc(InFade, OutFade - 1) Then
                        'This is the Horizontal fade part
                        Dc1 = Lng_MapDcs(InFade, OutFade - 1)
                        Dc2 = Lng_MapDcs(InFade, OutFade)
                        For X = 0 To 29
                            Colour = GetPixel(Dc1, X, 26)
                            Colour1 = GetPixel(Dc2, X, 1)
                            Call GetRgb(Colour, R1, G1, B1)
                            Call GetRgb(Colour1, R2, G2, B2)
                            Percent = 20
                            SetPixelV PicBuffer.hDC, InFade * 30 - 9 + X, OutFade * 30 - 7, RGB(R2 - (R2 / 100) * Percent + (R1 / 100) * Percent, G2 - (G2 / 100) * Percent + (G1 / 100) * Percent, B2 - (B2 / 100) * Percent + (B1 / 100) * Percent)
                        Next
                        For X = 0 To 29
                            Colour = GetPixel(Dc1, X, 27)
                            Colour1 = GetPixel(Dc2, X, 2)
                            Call GetRgb(Colour, R1, G1, B1)
                            Call GetRgb(Colour1, R2, G2, B2)
                            Percent = 40
                            SetPixelV PicBuffer.hDC, InFade * 30 - 9 + X, OutFade * 30 - 8, RGB(R2 - (R2 / 100) * Percent + (R1 / 100) * Percent, G2 - (G2 / 100) * Percent + (G1 / 100) * Percent, B2 - (B2 / 100) * Percent + (B1 / 100) * Percent)
                        Next
                        For X = 0 To 29
                            Colour = GetPixel(Dc1, X, 28)
                            Colour1 = GetPixel(Dc2, X, 3)
                            Call GetRgb(Colour, R1, G1, B1)
                            Call GetRgb(Colour1, R2, G2, B2)
                            Percent = 60
                            SetPixelV PicBuffer.hDC, InFade * 30 - 9 + X, OutFade * 30 - 9, RGB(R2 - (R2 / 100) * Percent + (R1 / 100) * Percent, G2 - (G2 / 100) * Percent + (G1 / 100) * Percent, B2 - (B2 / 100) * Percent + (B1 / 100) * Percent)
                        Next
                        For X = 0 To 29
                            Colour = GetPixel(Dc1, X, 29)
                            Colour1 = GetPixel(Dc2, X, 4)
                            Call GetRgb(Colour, R1, G1, B1)
                            Call GetRgb(Colour1, R2, G2, B2)
                            Percent = 80
                            SetPixelV PicBuffer.hDC, InFade * 30 - 9 + X, OutFade * 30 - 10, RGB(R2 - (R2 / 100) * Percent + (R1 / 100) * Percent, G2 - (G2 / 100) * Percent + (G1 / 100) * Percent, B2 - (B2 / 100) * Percent + (B1 / 100) * Percent)
                        Next
                    End If
                End If
            Next
        Next
    End If
    
    For OuterLoop = 1 To 10     'This Draws the People and Signs to the Screen, cuts out green for transparent look
        For InnerLoop = 1 To 10
            If Int_ExtraPlayers((ScreenX * 10 - 10 + InnerLoop), (ScreenY * 10 - 10 + OuterLoop)) > 0 Then
                Int_PassableLand(InnerLoop, OuterLoop - 1) = 1
                For InOuterLoop = 1 To 31
                    For InInnerLoop = 1 To 31
                        If Int_ExtraPlayerType(Int_ExtraPlayers((ScreenX * 10 - 10 + InnerLoop), (ScreenY * 10 - 10 + OuterLoop))) = 1 Then
                            Color = GetPixel(PicPerson1.hDC, InInnerLoop, InOuterLoop)
                        ElseIf Int_ExtraPlayerType(Int_ExtraPlayers((ScreenX * 10 - 10 + InnerLoop), (ScreenY * 10 - 10 + OuterLoop))) = 2 Then
                            Color = GetPixel(PicPerson2.hDC, InInnerLoop, InOuterLoop)
                        ElseIf Int_ExtraPlayerType(Int_ExtraPlayers((ScreenX * 10 - 10 + InnerLoop), (ScreenY * 10 - 10 + OuterLoop))) = 3 Then
                            Color = GetPixel(PicSign.hDC, InInnerLoop, InOuterLoop)
                        End If
                        If Color <> vbGreen Then
                            SetPixelV PicBuffer.hDC, InnerLoop * 30 - 9 + InInnerLoop, (OuterLoop - 1) * 30 - 9 + InOuterLoop, Color
                        End If
                    Next
                Next
                If Int_ExtraPlayerType(Int_ExtraPlayers((ScreenX * 10 - 10 + InnerLoop), (ScreenY * 10 - 10 + OuterLoop))) = 5 Then
                    If Day = 1 Then
                        Call LightMap(InnerLoop * 30 + 6, OuterLoop * 30 + 6, 75, Int_ExtraPlayerType(Int_ExtraPlayers((ScreenX * 10 - 10 + InnerLoop), (ScreenY * 10 - 10 + OuterLoop))) - 5, 150)
                    End If
                ElseIf Int_ExtraPlayerType(Int_ExtraPlayers((ScreenX * 10 - 10 + InnerLoop), (ScreenY * 10 - 10 + OuterLoop))) = 6 Then
                    If Day = 1 Then
                        Call LightMap(InnerLoop * 30 + 6, OuterLoop * 30 + 6, 75, Int_ExtraPlayerType(Int_ExtraPlayers((ScreenX * 10 - 10 + InnerLoop), (ScreenY * 10 - 10 + OuterLoop))) - 5, 150)
                    End If
                ElseIf Int_ExtraPlayerType(Int_ExtraPlayers((ScreenX * 10 - 10 + InnerLoop), (ScreenY * 10 - 10 + OuterLoop))) = 7 Then
                    If Day = 1 Then
                        Call LightMap(InnerLoop * 30 + 6, OuterLoop * 30 + 6, 75, Int_ExtraPlayerType(Int_ExtraPlayers((ScreenX * 10 - 10 + InnerLoop), (ScreenY * 10 - 10 + OuterLoop))) - 5, 150)
                    End If
                ElseIf Int_ExtraPlayerType(Int_ExtraPlayers((ScreenX * 10 - 10 + InnerLoop), (ScreenY * 10 - 10 + OuterLoop))) = 8 Then
                    If Day = 1 Then
                        Call LightMap(InnerLoop * 30 + 6, OuterLoop * 30 + 21, 75, Int_ExtraPlayerType(Int_ExtraPlayers((ScreenX * 10 - 10 + InnerLoop), (ScreenY * 10 - 10 + OuterLoop))) - 5, 150)
                    End If
                ElseIf Int_ExtraPlayerType(Int_ExtraPlayers((ScreenX * 10 - 10 + InnerLoop), (ScreenY * 10 - 10 + OuterLoop))) = 9 Then
                    If Day = 1 Then
                        Call LightMap(InnerLoop * 30 + 6, OuterLoop * 30 - 9, 75, Int_ExtraPlayerType(Int_ExtraPlayers((ScreenX * 10 - 10 + InnerLoop), (ScreenY * 10 - 10 + OuterLoop))) - 5, 150)
                    End If
                End If
            End If
        Next
    Next
    
    'These are for Testing reasons to see how well the tracking engine can perform
    Call ClearTrack
    If InBuilding = False Then
        Dim trackstart As Boolean
        trackstart = False
        Do
            Xpos = Rnd * 9 + 1
            Ypos = Rnd * 9 + 1
            If Int_PassableLand(Xpos, Ypos) = 0 Then
                Call StartTrack(Xpos, Ypos, PlayerX, PlayerY)
                trackstart = True
            End If
        Loop Until trackstart = True
    End If
    PicBuffer.Refresh
    Call BitBlt(FrmGreenEffect.hDC, 0, 0, PicBuffer.ScaleWidth, PicBuffer.ScaleHeight, PicBuffer.hDC, 0, 0, vbSrcCopy)
    FrmGreenEffect.Refresh
    TmrPlayer.Enabled = True
End Sub

Sub GetRgb(ByVal Color As Long, ByRef red As Integer, ByRef green As Integer, ByRef blue As Integer)
    
    '--------------------------------------------------
    'Gets the Red Green and Blue values from a long
    '--------------------------------------------------
    
    Dim temp As Long
    temp = (Color And 255)
    red = temp And 255
    temp = Int(Color / 256)
    green = temp And 255
    temp = Int(Color / 65536)
    blue = temp And 255
End Sub

Private Sub Form_Terminate()
    End 'Exit Game
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End 'Exit Game
End Sub

Private Sub lblmenu_Click()
    FrmGreenEffect.PopupMenu mnupopup
End Sub

Private Sub mnuexit_Click()
    End 'Exit Game
End Sub

Private Sub mnuLoad_Click()
    CD1.Filter = "Green Effect Saved Game (*.ges)|*.ges"
    CD1.ShowOpen
    Filename = CD1.Filename
    If Filename = "" Then Exit Sub
    If InBuilding = True Then
        Call OutofBuilding
    End If
    Open Filename For Input As #1
    Input #1, StrData
    KeyCount = Val(StrData)
    ReDim Keys(1 To KeyCount) As Key
    KeyCount = 0
    Do
        Input #1, StrData
        If Mid(StrData, 1, 8) = "#ScreenX" Then
            ScreenX = Val(Mid(StrData, 10))
        ElseIf Mid(StrData, 1, 8) = "#ScreenY" Then
            ScreenY = Val(Mid(StrData, 10))
        ElseIf Mid(StrData, 1, 8) = "#PlayerX" Then
            PlayerX = Val(Mid(StrData, 10))
        ElseIf Mid(StrData, 1, 8) = "#PlayerY" Then
            PlayerY = Val(Mid(StrData, 10))
        ElseIf Mid(StrData, 1, 7) = "#EnterX" Then
            EnterX = Val(Mid(StrData, 9))
        ElseIf Mid(StrData, 1, 7) = "#EnterY" Then
            EnterY = Val(Mid(StrData, 9))
        ElseIf Mid(StrData, 1, 4) = "#Day" Then
            Day = Val(Mid(StrData, 6))
        ElseIf Mid(StrData, 1, 11) = "#Lastaction" Then
            LastAction = Val(Mid(StrData, 13))
        ElseIf Mid(StrData, 1, 8) = "#Castras" Then
            Castras = Val(Mid(StrData, 10))
        ElseIf Mid(StrData, 1, 11) = "#ArmourName" Then
            Armour.Name = Val(Mid(StrData, 13))
        ElseIf Mid(StrData, 1, 10) = "#ArmourMax" Then
            Armour.max = Val(Mid(StrData, 12))
        ElseIf Mid(StrData, 1, 11) = "#ArmourUsed" Then
            Armour.current = Val(Mid(StrData, 13))
        ElseIf Mid(StrData, 1, 12) = "#ArmourPrice" Then
            Armour.Price = Val(Mid(StrData, 14))
        ElseIf Mid(StrData, 1, 11) = "#WeaponName" Then
            NewWeap = NewWeap + 1
            Weapon(NewWeap).Name = Val(Mid(StrData, 13))
        ElseIf Mid(StrData, 1, 13) = "#WeaponAttack" Then
            Weapon(NewWeap).Attack = Val(Mid(StrData, 15))
        ElseIf Mid(StrData, 1, 10) = "#WeaponMax" Then
            Weapon(NewWeap).max = Val(Mid(StrData, 11))
        ElseIf Mid(StrData, 1, 11) = "#WeaponMiss" Then
            Weapon(NewWeap).Misschance = Val(Mid(StrData, 14))
        ElseIf Mid(StrData, 1, 12) = "#WeaponPrice" Then
            Weapon(NewWeap).Price = Val(Mid(StrData, 13))
        ElseIf Mid(StrData, 1, 11) = "#WeaponUsed" Then
            Weapon(NewWeap).current = Val(Mid(StrData, 12))
        ElseIf Mid(StrData, 1, 8) = "#KName" Then
            Keys(KeyCount).Name = Mid(StrData, 10)
        ElseIf Mid(StrData, 1, 7) = "#Key" Then
            KeyCount = KeyCount + 1
            Keys(KeyCount).KeyVal = Val(Mid(StrData, 9))
        End If
    Loop Until EOF(1)
    Close
    'Render the map with the loaded information
    Call RenderMap
End Sub

Private Sub mnuoptions_Click()
    'Show the options form
    frmOptions.Show
End Sub

Private Sub mnusave_Click()
    If InBuilding = True Then
        'At the moment this save routine can only save infomation about the co-ordinates
        'when outside a building, i am currently working on this...
        
        Call MsgBox("Please Exit Building Before Saving", vbInformation)
        Exit Sub
    End If
    
    CD1.Filter = "Green Effect Saved Game (*.ges)|*.ges"
    CD1.ShowSave
    Filename = CD1.Filename
    If Filename = "" Then Exit Sub
    Open Filename For Output As #1
    Print #1, Str(KeyCount)
    Print #1, "#Health" & Str(PlayerHealth)
    Print #1, "#ScreenX" & Str(ScreenX)
    Print #1, "#ScreenY" & Str(ScreenY)
    Print #1, "#PlayerX" & Str(PlayerX)
    Print #1, "#PlayerY" & Str(PlayerY)
    Print #1, "#EnterX" & Str(EnterX)
    Print #1, "#EnterY" & Str(EnterY)
    Print #1, "#Day" & Str(Day)
    Print #1, "#Lastaction" & Str(LastAction)
    Print #1, "#Castras" & Str(Castras)
    Print #1, "#ArmourName " & Armour.Name
    Print #1, "#ArmourMax" & Str(Armour.max)
    Print #1, "#ArmourUsed" & Str(Armour.current)
    Print #1, "#ArmourPrice" & Str(Armour.Price)
    For Weap = 1 To 10
        If Weapon(Weap).Name <> "" Then
            Print #1, "#WeaponName " & Weapon(Weap).Name
            Print #1, "#WeaponAttack" & Str(Weapon(Weap).Attack)
            Print #1, "#WeaponMax" & Str(Weapon(Weap).max)
            Print #1, "#WeaponMiss" & Str(Weapon(Weap).Misschance)
            Print #1, "#WeaponPrice" & Str(Weapon(Weap).Price)
            Print #1, "#WeaponUsed" & Str(Weapon(Weap).current)
        End If
    Next
    For SaveKey = 1 To KeyCount
        Print #1, "#Key" & Str(Keys(SaveKey).KeyVal)
        Print #1, "#KName " & Keys(SaveKey).Name
    Next
    Close
End Sub

Private Sub TmrDay_Timer()
    Day = 1 - Day   'Change from day to Night
    Call RenderMap  'Redraw the map to the new light condition
End Sub

Private Sub TmrPlayer_Timer()       'This is the player movement part
    '--------------------------------------------------
    'Description:
    'This is the main part of the program it checks to
    'see if there is any keypresses, if so it will draw
    'the new player position to the screen. It will
    'call the tracking sub routine if needed. If the
    'Player is trying to exit a building it will call
    'outofbuilding, and will check to see if there is
    'any messages...
    '--------------------------------------------------
    If GetAsyncKeyState(vbKeyDown) < 0 Then
        LastAction = 1
        If InBuilding = True Then
            Call CheckBuilding
        End If
        Anim = 1 - Anim 'Change the Walking animation
        If PlayerY + 1 < 11 Then
            If Int_PassableLand(PlayerX, PlayerY + 1) = 0 Then
                If Int_CurrentTrackX = PlayerX And Int_CurrentTrackY = PlayerY + 1 Then
                Else
                    PlayerY = PlayerY + 1
                End If
            ElseIf Int_PassableLand(PlayerX, PlayerY + 1) = 2 Then
                PlayerY = PlayerY + 1
                Call DoDie
            End If
        Else
            PlayerY = 1
            ScreenY = ScreenY + 1
            EnterX = PlayerX
            EnterY = PlayerY
            Call RenderMap   'Draw the New map
        End If
    End If
    If GetAsyncKeyState(vbKeyUp) < 0 Then
        LastAction = 2
        If InBuilding = False Then
            Call CheckForBuilding
        End If
        Anim = 1 - Anim 'Change the walking animation
        If PlayerY - 1 > 0 Then
            If Int_PassableLand(PlayerX, PlayerY - 1) = 0 Then
                If Int_CurrentTrackX = PlayerX And Int_CurrentTrackY = PlayerY - 1 Then
                Else
                    PlayerY = PlayerY - 1
                End If
            ElseIf Int_PassableLand(PlayerX, PlayerY - 1) = 2 Then
                PlayerY = PlayerY - 1
                Call DoDie
            End If
        Else
            PlayerY = 10
            ScreenY = ScreenY - 1
            EnterX = PlayerX
            EnterY = PlayerY
            Call RenderMap   'Draw the new map
        End If
    End If
    If GetAsyncKeyState(vbKeyLeft) < 0 Then
        LastAction = 3
        Anim = 1 - Anim 'Change the player walking animation
        If PlayerX - 1 > 0 Then
            If Int_PassableLand(PlayerX - 1, PlayerY) = 0 Then
                If Int_CurrentTrackX = PlayerX - 1 And Int_CurrentTrackY = PlayerY Then
                Else
                    PlayerX = PlayerX - 1
                End If
            ElseIf Int_PassableLand(PlayerX - 1, PlayerY) = 2 Then
                PlayerX = PlayerX - 1
                Call DoDie
            End If
        Else
            PlayerX = 10
            ScreenX = ScreenX - 1
            EnterX = PlayerX
            EnterY = PlayerY
            Call RenderMap   'Draw the new map
        End If
    End If
    If GetAsyncKeyState(vbKeyRight) < 0 Then
        LastAction = 4
        Anim = 1 - Anim 'Change the player walking animation
        If PlayerX + 1 < 11 Then
            If Int_PassableLand(PlayerX + 1, PlayerY) = 0 Then
                If Int_CurrentTrackX = PlayerX + 1 And Int_CurrentTrackY = PlayerY Then
                Else
                    PlayerX = PlayerX + 1
                End If
            ElseIf Int_PassableLand(PlayerX + 1, PlayerY) = 2 Then
                PlayerX = PlayerX + 1
                Call DoDie
            End If
        Else
            PlayerX = 1
            ScreenX = ScreenX + 1
            EnterX = PlayerX
            EnterY = PlayerY
            Call RenderMap   'Draw the new map
        End If
    End If
    If GetAsyncKeyState(32) < 0 Then
        'Only check if something is ahead
        If PlayerY - 1 > 0 Then
            'Check to see if something is ahead and if so, show the appreiate message
            If Str_Map(PlayerX, PlayerY - 1) = "L" Then
                Call DoBCase
            ElseIf Str_Map(PlayerX, PlayerY - 1) = "M" Then
                Call DoCase
            ElseIf Str_Map(PlayerX, PlayerY - 1) = "$" Then
                Call DoBed
            ElseIf Str_Map(PlayerX, PlayerY - 1) = "[" Then
                Call DoFplace
            End If
        End If
    End If
    If GetAsyncKeyState(65) < 0 Then    '(A)ttack
        If PlayerX = Int_CurrentTrackX + 1 Or PlayerX = Int_CurrentTrackX - 1 Or PlayerX = Int_CurrentTrackX Then
            If PlayerY = Int_CurrentTrackY + 1 Or PlayerY = Int_CurrentTrackY - 1 Or PlayerY = Int_CurrentTrackY Then
                Call AttackPTM(ScreenX, ScreenY)
            End If
        End If
    End If
    If Int_ObjectLocation(PlayerX, PlayerY) <> 0 Then
        'So Far if is just castras but will expand to keys, weapons, health...
        If Int_ObjectLocation(PlayerX, PlayerY) = 1 Then    'Castra
            Call DoCastra(PlayerX, PlayerY)
        End If
        Int_ObjectLocation(PlayerX, PlayerY) = 0
    End If
    LblPos.Caption = Str$(ScreenX * 10 - 10 + PlayerX) + "," + Str$(ScreenY * 10 - 10 + PlayerY)
    'Draw the backbuffer to the screen and then draw player and other objects(When required) to the screen
    Call BitBlt(FrmGreenEffect.hDC, 0, 0, PicBuffer.ScaleWidth, PicBuffer.ScaleHeight, PicBuffer.hDC, 0, 0, vbSrcCopy)
    If LastAction = 1 Then
        Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PBD(Anim).hDC, 0, 0, vbSrcPaint)
        Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PicDown(Anim).hDC, 0, 0, vbSrcAnd)
    ElseIf LastAction = 2 Then
        Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PBU(Anim).hDC, 0, 0, vbSrcPaint)
        Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PicUp(Anim).hDC, 0, 0, vbSrcAnd)
    ElseIf LastAction = 3 Then
        Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PBL(Anim).hDC, 0, 0, vbSrcPaint)
        Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PicLeft(Anim).hDC, 0, 0, vbSrcAnd)
    ElseIf LastAction = 4 Then
        Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PBR(Anim).hDC, 0, 0, vbSrcPaint)
        Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PicRight(Anim).hDC, 0, 0, vbSrcAnd)
    End If
    'Refresh the screen so that the new image is displayed
    Call DoProgress
    'Check the position to see if there is any messages that can be displayed
    Call CheckPos
    AIDelayNo = AIDelayNo + 1
    If Bln_TrackInProgress = True Then
        If AIDelayNo >= AIDelay Then
            If AIDelayNo > AIDelay Then
                AIDelayNo = 0
            End If
            Call ContinueTrack
        Else
            Call DrawMonster
        End If
    End If
    FrmGreenEffect.Refresh
End Sub

Sub CheckPos()  'Check the player position to see if there is a message waiting

    '--------------------------------------------------
    'Description:
    'Checks the current position for any messages
    '--------------------------------------------------
    Dim TmpNum1 As Integer
    Dim TmpNum2 As Integer
    
    LblAble.Caption = ""    'Clear the message label
    TmpNum1 = ScreenX * 10 - 10 + PlayerX
    TmpNum2 = ScreenY * 10 - 10 + PlayerY
    If Int_ExtraPlayers(TmpNum1, TmpNum2) > 0 Then
        If Int_ExtraPlayerType(Int_ExtraPlayers(TmpNum1, TmpNum2)) < 5 Or Int_ExtraPlayerType(Int_ExtraPlayers(TmpNum1, TmpNum2)) > 9 Then
            LblAble.Caption = "Message!"    'Show there is a message waiting
            If GetAsyncKeyState(32) < 0 Then    'If the Player presses space
                Call Chat(TmpNum1, TmpNum2)
            End If
        End If
    End If
End Sub

Sub CheckBuilding() 'If the player is in a building, check if he is exiting

    '--------------------------------------------------
    'Description:
    'Checks to see if the player is trying to exit a
    'building
    '--------------------------------------------------
    
    If ScreenX = 11 And ScreenY = 11 And PlayerX = 5 And PlayerY = 10 Then
        Call OutofBuilding  'Exit the building subroutine
    End If
End Sub

Sub CheckForBuilding()
    Dim HasEntered As Boolean
    
    '--------------------------------------------------
    'Description:
    'If the player is entering a building, check for
    'doors. Been Modified to Check to see if the player
    'has the right key for the door. If not tells the
    'player which key is needed
    '--------------------------------------------------
    
    HasEntered = False
    HouseVal = MapHouse((ScreenX * 10 - 10 + PlayerX), (ScreenY * 10 - 10 + PlayerY))
    If HouseVal > 0 Then
        For CanEnter = 1 To KeyCount
            If House(HouseVal).Key = Keys(CanEnter).KeyVal Then
                HasEntered = True
                Call BuildingMaps(HouseVal)
                Exit For
            End If
        Next
        If HasEntered = False Then
            LblMessage.Caption = "Can't Enter Building"
            LblText.Caption = "Haven't got the correct Key, you need the " & House(HouseVal).KeyName
            Call Wait(1 + (Len(LblText.Caption) / 20))  'Wait
            LblMessage.Caption = ""
            LblText.Caption = ""
        End If
    End If
End Sub

Sub BuildingMaps(ByVal Index As Integer)
    
    '--------------------------------------------------
    'Description:
    'Loads and sets the new maps for the building the
    'player enters
    '--------------------------------------------------
    
    TmrPlayer.Enabled = False
    BuildingX = (ScreenX * 10 - 10 + PlayerX)
    BuildingY = (ScreenY * 10 - 11 + PlayerY)
    
    Filename = App.Path + "\GE\GMap"
    
    File = Filename + Mid$(Str(Index), 2)
    
    Open File + ".Map" For Input As #1
    For OuterLoop = 1 To 200
        Input #1, Str_CompressedMap(OuterLoop)
    Next
    Close
    
    Call ClearData
    
    For OuterLoop = 1 To 200
        Str_UnCompressedMap(OuterLoop) = ""
        For InnerLoop = 1 To 200
            
            Int_ExtraPlayers(InnerLoop, OuterLoop) = 0
            
            If Mid(Str_CompressedMap(OuterLoop), InnerLoop, 1) = "(" Then
                If Mid(Str_CompressedMap(OuterLoop), InnerLoop + 2, 1) = ")" Then
                    For A = 1 To Val(Mid(Str_CompressedMap(OuterLoop), InnerLoop + 1, 1))
                        Str_UnCompressedMap(OuterLoop) = Str_UnCompressedMap(OuterLoop) + Mid(Str_CompressedMap(OuterLoop), InnerLoop + 3, 1)
                    Next
                    InnerLoop = InnerLoop + 3
                ElseIf Mid(Str_CompressedMap(OuterLoop), InnerLoop + 3, 1) = ")" Then
                    For A = 1 To Val(Mid(Str_CompressedMap(OuterLoop), InnerLoop + 1, 2))
                        Str_UnCompressedMap(OuterLoop) = Str_UnCompressedMap(OuterLoop) + Mid(Str_CompressedMap(OuterLoop), InnerLoop + 4, 1)
                    Next
                    InnerLoop = InnerLoop + 4
                ElseIf Mid(Str_CompressedMap(OuterLoop), InnerLoop + 4, 1) = ")" Then
                    For A = 1 To Val(Mid(Str_CompressedMap(OuterLoop), InnerLoop + 1, 3))
                        Str_UnCompressedMap(OuterLoop) = Str_UnCompressedMap(OuterLoop) + Mid(Str_CompressedMap(OuterLoop), InnerLoop + 5, 1)
                    Next
                    InnerLoop = InnerLoop + 5
                End If
            Else
                Str_UnCompressedMap(OuterLoop) = Str_UnCompressedMap(OuterLoop) + Mid(Str_CompressedMap(OuterLoop), InnerLoop, 1)
            End If
        Next
    Next
    
    Open File + ".Pls" For Input As #1
    CurrentPlayer = 0
    Input #1, Text
    ToNum = Val(Text)
    If ToNum <> 0 Then
        For Extras = 1 To ToNum
            For GetInfo = 1 To 5
                Input #1, Text
                If Mid(Text, 1, 6) = "##Name" Then
                    CurrentPlayer = CurrentPlayer + 1
                    Str_ExtraNames(CurrentPlayer) = Mid(Text, 8)
                ElseIf Mid(Text, 1, 6) = "##XPos" Then
                    Int_ExtraPlayerX(CurrentPlayer) = Val(Mid(Text, 8))
                ElseIf Mid(Text, 1, 6) = "##YPos" Then
                    Int_ExtraPlayerY(CurrentPlayer) = Val(Mid$(Text, 8))
                ElseIf Mid(Text, 1, 6) = "##Text" Then
                    Str_ExtraText(CurrentPlayer) = Mid(Text, 8)
                ElseIf Mid(Text, 1, 6) = "##Type" Then
                    Int_ExtraPlayerType(CurrentPlayer) = Val(Mid(Text, 8))
                    Int_ExtraPlayers(Int_ExtraPlayerX(CurrentPlayer), Int_ExtraPlayerY(CurrentPlayer)) = CurrentPlayer
                End If
            Next
        Next
    End If
    Close
    PlayerX = 5
    PlayerY = 11
    ScreenX = 11
    ScreenY = 11
    Anim = 1
    InBuilding = True
    Call RenderMap
    Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PBU(Anim).hDC, Counter, 0, vbSrcPaint)
    Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PicUp(Anim).hDC, Counter, 0, vbSrcAnd)
    TmrPlayer.Enabled = True
End Sub

Sub DoCastra(ByVal X As Integer, ByVal Y As Integer)        'Shows the User has found a castra(Money)

    '--------------------------------------------------
    'Description:
    'Shows the User has found a castra(Money)
    '--------------------------------------------------

    LblMessage.Caption = "Object Found..."
    LblText.Caption = "You have found a Castra"
    Castras = Castras + 1
    lblMoney.Caption = Castras
    Call Wait(1)
    LblMessage.Caption = ""
    LblText.Caption = ""
End Sub


Sub Chat(ByVal InnerLoop As Integer, ByVal OuterLoop As Integer)
    
    '--------------------------------------------------
    'Description:
    'Shows the messages when the player talks to a
    'person
    '--------------------------------------------------
    
    Dim Chat(1 To 100) As String
    If Str_ExtraNames(Int_ExtraPlayers(InnerLoop, OuterLoop)) = "SignPost" Then
        LblMessage.Caption = "Reading SignPost"
    Else
        LblMessage.Caption = "Talking to " + Str_ExtraNames(Int_ExtraPlayers(InnerLoop, OuterLoop))
    End If
    If Str_ExtraText(Int_ExtraPlayers(InnerLoop, OuterLoop)) = "Random" Then
        Chat(1) = "Hello"
        Chat(2) = "Go Away"
        Chat(3) = "Welcome Traveller"
        Chat(4) = "You are dressed very strange young'un"
        Chat(5) = "Welcome to Muncipium"
        Chat(6) = "Have a look around"
        Chat(7) = "You can buy things at the shop"
        Chat(8) = "My Name is " + Str_ExtraNames(Int_ExtraPlayers(InnerLoop, OuterLoop))
        Chat(9) = "You could do with some better armour!"
        Chat(10) = "You could do with a better Weapon!"
        Chat(11) = "Try not to fall in the water, you could drown..."
        Chat(12) = "Why not send comments to Kevin at Yet_Another_Idiot@Hotmail.com about how this could be improved"
        If Day = 1 Then
            Chat(13) = "Ooh its night, better get inside"
        Else
            Chat(13) = "What a nice day"
        End If
        RandomChat = Int(Rnd * 13) + 1
        LblText.Caption = Chat(RandomChat)
    Else
        LblText.Caption = Str_ExtraText(Int_ExtraPlayers(InnerLoop, OuterLoop))
    End If
    
    Call Wait(1 + (Len(LblText.Caption) / 20))  'Wait
    
    If Str_ExtraNames(Int_ExtraPlayers(InnerLoop, OuterLoop)) = "ShopKeeper" Then
        TmrPlayer.Enabled = False   'Disables movement
        FrmShop.Visible = False
        FrmShop.Show
        FrmShop.Left = FrmGreenEffect.Left + PlayerX * 30 + 24
        FrmShop.Top = FrmGreenEffect.Top + PlayerY * 30 + 24
        FrmShop.Visible = True
    End If
    
    LblMessage.Caption = ""
    LblText.Caption = ""
End Sub

Sub OutofBuilding()

    '--------------------------------------------------
    'Description:
    'Load the default map when coming out of a building
    '--------------------------------------------------
    Dim OuterLoop As Integer
    Dim InnerLoop As Integer
    Dim A As Integer
    Dim CurrentPlayer As Integer
    Dim StrText As String
    Dim ToNum As Integer
    Dim Extras As Integer
    Dim GetInfo As Integer
    
    TmrPlayer.Enabled = False
    InBuilding = False
    Open App.Path + "\GE\GEffect.Map" For Input As #1
    For OuterLoop = 1 To 200
        Input #1, Str_CompressedMap(OuterLoop)
    Next
    Close
    
    'Clear loaded character data
    Call ClearData
    
    For OuterLoop = 1 To 200
        Str_UnCompressedMap(OuterLoop) = ""
        For InnerLoop = 1 To 200
        
            Int_ExtraPlayers(InnerLoop, OuterLoop) = 0
            
            If Mid(Str_CompressedMap(OuterLoop), InnerLoop, 1) = "(" Then
                If Mid(Str_CompressedMap(OuterLoop), InnerLoop + 2, 1) = ")" Then
                    For A = 1 To Val(Mid(Str_CompressedMap(OuterLoop), InnerLoop + 1, 1))
                        Str_UnCompressedMap(OuterLoop) = Str_UnCompressedMap(OuterLoop) + Mid(Str_CompressedMap(OuterLoop), InnerLoop + 3, 1)
                    Next
                    InnerLoop = InnerLoop + 3
                ElseIf Mid(Str_CompressedMap(OuterLoop), InnerLoop + 3, 1) = ")" Then
                    For A = 1 To Val(Mid(Str_CompressedMap(OuterLoop), InnerLoop + 1, 2))
                        Str_UnCompressedMap(OuterLoop) = Str_UnCompressedMap(OuterLoop) + Mid(Str_CompressedMap(OuterLoop), InnerLoop + 4, 1)
                    Next
                    InnerLoop = InnerLoop + 4
                ElseIf Mid(Str_CompressedMap(OuterLoop), InnerLoop + 4, 1) = ")" Then
                    For A = 1 To Val(Mid(Str_CompressedMap(OuterLoop), InnerLoop + 1, 3))
                        Str_UnCompressedMap(OuterLoop) = Str_UnCompressedMap(OuterLoop) + Mid(Str_CompressedMap(OuterLoop), InnerLoop + 5, 1)
                    Next
                    InnerLoop = InnerLoop + 5
                End If
            Else
                Str_UnCompressedMap(OuterLoop) = Str_UnCompressedMap(OuterLoop) + Mid(Str_CompressedMap(OuterLoop), InnerLoop, 1)
            End If
        Next
    Next
    
    Open App.Path + "\GE\GEffect.Pls" For Input As #1
    CurrentPlayer = 0
    Input #1, StrText
    ToNum = Val(StrText)
    For Extras = 1 To ToNum
        For GetInfo = 1 To 5
            Input #1, Text
            If Mid(Text, 1, 6) = "##Name" Then
                CurrentPlayer = CurrentPlayer + 1
                Str_ExtraNames(CurrentPlayer) = Mid(Text, 8)
            ElseIf Mid(Text, 1, 6) = "##XPos" Then
                Int_ExtraPlayerX(CurrentPlayer) = Val(Mid(Text, 8))
            ElseIf Mid(Text, 1, 6) = "##YPos" Then
                Int_ExtraPlayerY(CurrentPlayer) = Val(Mid$(Text, 8))
            ElseIf Mid(Text, 1, 6) = "##Text" Then
                Str_ExtraText(CurrentPlayer) = Mid(Text, 8)
            ElseIf Mid(Text, 1, 6) = "##Type" Then
                Int_ExtraPlayerType(CurrentPlayer) = Val(Mid(Text, 8))
                Int_ExtraPlayers(Int_ExtraPlayerX(CurrentPlayer), Int_ExtraPlayerY(CurrentPlayer)) = CurrentPlayer
            End If
        Next
    Next
    Close
    ScreenX = Int(BuildingX / 10)
    ScreenY = Int(BuildingY / 10)
    PlayerX = BuildingX - (ScreenX * 10)
    PlayerY = BuildingY - (ScreenY * 10)
    ScreenX = ScreenX + 1
    ScreenY = ScreenY + 1
    Anim = 1
    BuildingX = 0
    BuildingY = 0
    'Render the map again
    Call RenderMap
    'Copy Mask to Screen
    Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PBU(Anim).hDC, 0, 0, vbSrcPaint)
    'Copy Picture to Screen
    Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PicUp(Anim).hDC, 0, 0, vbSrcAnd)
    TmrPlayer.Enabled = True
End Sub

Sub ClearData()

    '--------------------------------------------------
    'Description:
    'this clears the extra character data when loading
    'a map
    '--------------------------------------------------
    Dim ClearPlayer As Integer
    
    For ClearPlayer = 1 To 250
        Str_ExtraNames(ClearPlayer) = ""
        Int_ExtraPlayerX(ClearPlayer) = 0
        Int_ExtraPlayerY(ClearPlayer) = 0
        Int_ExtraPlayerType(ClearPlayer) = 0
        Str_ExtraText(ClearPlayer) = ""
    Next
End Sub

Sub DoTimeTravel()

    '--------------------------------------------------
    'Description:
    'This shows the startup scene where the user goes
    'through time
    '--------------------------------------------------
    
    Dim Filename As String
    Dim OuterLoop As Integer
    Dim InnerLoop As Integer
    Dim A As Integer
    Dim Walk As Integer
    
    TmrPlayer.Enabled = False
    BuildingX = (ScreenX * 10 - 10 + PlayerX)
    BuildingY = (ScreenY * 10 - 11 + PlayerY)
    
    Filename = App.Path + "\GE\GStart"
    
    
    Open Filename + ".Map" For Input As #1
    For OuterLoop = 1 To 200
        Input #1, Str_CompressedMap(OuterLoop)
    Next
    Close
    
    Call ClearData
    
    For OuterLoop = 1 To 200
        Str_UnCompressedMap(OuterLoop) = ""
        For InnerLoop = 1 To 200
            
            Int_ExtraPlayers(InnerLoop, OuterLoop) = 0
            
            If Mid(Str_CompressedMap(OuterLoop), InnerLoop, 1) = "(" Then
                If Mid(Str_CompressedMap(OuterLoop), InnerLoop + 2, 1) = ")" Then
                    For A = 1 To Val(Mid(Str_CompressedMap(OuterLoop), InnerLoop + 1, 1))
                        Str_UnCompressedMap(OuterLoop) = Str_UnCompressedMap(OuterLoop) + Mid(Str_CompressedMap(OuterLoop), InnerLoop + 3, 1)
                    Next
                    InnerLoop = InnerLoop + 3
                ElseIf Mid(Str_CompressedMap(OuterLoop), InnerLoop + 3, 1) = ")" Then
                    For A = 1 To Val(Mid(Str_CompressedMap(OuterLoop), InnerLoop + 1, 2))
                        Str_UnCompressedMap(OuterLoop) = Str_UnCompressedMap(OuterLoop) + Mid(Str_CompressedMap(OuterLoop), InnerLoop + 4, 1)
                    Next
                    InnerLoop = InnerLoop + 4
                ElseIf Mid(Str_CompressedMap(OuterLoop), InnerLoop + 4, 1) = ")" Then
                    For A = 1 To Val(Mid(Str_CompressedMap(OuterLoop), InnerLoop + 1, 3))
                        Str_UnCompressedMap(OuterLoop) = Str_UnCompressedMap(OuterLoop) + Mid(Str_CompressedMap(OuterLoop), InnerLoop + 5, 1)
                    Next
                    InnerLoop = InnerLoop + 5
                End If
            Else
                Str_UnCompressedMap(OuterLoop) = Str_UnCompressedMap(OuterLoop) + Mid(Str_CompressedMap(OuterLoop), InnerLoop, 1)
            End If
        Next
    Next
    
    PlayerX = 5
    PlayerY = 10
    ScreenX = 11
    ScreenY = 11
    Anim = 1
    InBuilding = True
    Call RenderMap
    TmrPlayer.Enabled = False
    
    For OuterLoop = 1 To PicMachine.ScaleHeight - 2
        For InnerLoop = 1 To PicMachine.ScaleWidth - 2
            Color = GetPixel(PicMachine.hDC, InnerLoop, OuterLoop)
            If Color <> vbGreen Then
                SetPixelV PicBuffer.hDC, (125 + InnerLoop), (6.6 + OuterLoop), Color
            End If
        Next
    Next
    For OuterLoop = 1 To PicCom.ScaleHeight - 2
        For InnerLoop = 1 To PicCom.ScaleWidth - 2
            Color = PicCom.Point(InnerLoop, OuterLoop)
            If Color <> vbGreen Then
                SetPixelV PicBuffer.hDC, 60 + InnerLoop, 40 + OuterLoop, Color
                SetPixelV PicBuffer.hDC, 60 + InnerLoop, 96 + OuterLoop, Color
                SetPixelV PicBuffer.hDC, 60 + InnerLoop, 151 + OuterLoop, Color
                SetPixelV PicBuffer.hDC, 60 + InnerLoop, 208 + OuterLoop, Color
                SetPixelV PicBuffer.hDC, 240 + InnerLoop, 40 + OuterLoop, Color
                SetPixelV PicBuffer.hDC, 240 + InnerLoop, 96 + OuterLoop, Color
                SetPixelV PicBuffer.hDC, 240 + InnerLoop, 151 + OuterLoop, Color
                SetPixelV PicBuffer.hDC, 240 + InnerLoop, 208 + OuterLoop, Color
            End If
        Next
    Next
    FrmGreenEffect.Refresh
    For Walk = 1 To 9
        Call Wait(0.2)  'Wait for .2 of a second
        Anim = 1 - Anim 'Change the walking animation
        PlayerY = PlayerY - 1
        Call BitBlt(FrmGreenEffect.hDC, 0, 0, PicBuffer.ScaleWidth, PicBuffer.ScaleHeight, PicBuffer.hDC, 0, 0, vbSrcCopy)
        Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PBU(Anim).hDC, 0, 0, vbSrcPaint)
        Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PicUp(Anim).hDC, 0, 0, vbSrcAnd)
        FrmGreenEffect.Refresh
    Next
    FrmGreenEffect.Refresh
    Call DoFadeOut
    TmrPlayer.Enabled = True
    FrmGreenEffect.Cls
End Sub

Sub DoDie()

    '--------------------------------------------------
    'Description:
    'This places a dead body at the angle of which
    'the player was facing
    'OutPut:
    'Draws a body at the Players X,Y in the correct
    'Direction
    '--------------------------------------------------

    DeadType = Int(Rnd * 2)
    If LastAction = 1 Then     'Check the Direction of the Death
        Call BitBlt(FrmGreenEffect.hDC, 0, 0, PicBuffer.ScaleWidth, PicBuffer.ScaleHeight, PicBuffer.hDC, 0, 0, vbSrcCopy)
        For OuterLoop = 1 To 32
            For InnerLoop = 1 To 32
                Color = GetPixel(PicDead(DeadType).hDC, InnerLoop, OuterLoop)
                If Color <> vbGreen Then
                    SetPixelV FrmGreenEffect.hDC, PlayerX * 30 - 11.6 + InnerLoop, PlayerY * 30 - 11.6 + 33 - OuterLoop, Color
                End If
            Next
        Next
    ElseIf LastAction = 2 Then
        Call BitBlt(FrmGreenEffect.hDC, 0, 0, PicBuffer.ScaleWidth, PicBuffer.ScaleHeight, PicBuffer.hDC, 0, 0, vbSrcCopy)
        For OuterLoop = 1 To 32
            For InnerLoop = 1 To 32
                Color = GetPixel(PicDead(DeadType).hDC, InnerLoop, OuterLoop)
                If Color <> vbGreen Then
                    SetPixelV FrmGreenEffect.hDC, PlayerX * 30 - 11.6 + InnerLoop, PlayerY * 30 - 11.6 + OuterLoop, Color
                End If
            Next
        Next
    ElseIf LastAction = 3 Then
        Call BitBlt(FrmGreenEffect.hDC, 0, 0, PicBuffer.ScaleWidth, PicBuffer.ScaleHeight, PicBuffer.hDC, 0, 0, vbSrcCopy)
        For OuterLoop = 1 To 32
            For InnerLoop = 1 To 32
                Color = GetPixel(PicDead(DeadType).hDC, InnerLoop, OuterLoop)
                If Color <> vbGreen Then
                    SetPixelV FrmGreenEffect.hDC, PlayerX * 30 - 11.6 + OuterLoop, PlayerY * 30 - 11.6 + InnerLoop, Color
                End If
            Next
        Next
    ElseIf LastAction = 4 Then
        Call BitBlt(FrmGreenEffect.hDC, 0, 0, PicBuffer.ScaleWidth, PicBuffer.ScaleHeight, PicBuffer.hDC, 0, 0, vbSrcCopy)
        For OuterLoop = 1 To 32
            For InnerLoop = 1 To 32
                Color = GetPixel(PicDead(DeadType).hDC, InnerLoop, OuterLoop)
                If Color <> vbGreen Then
                    SetPixelV FrmGreenEffect.hDC, PlayerX * 30 - 11.6 + 33 - OuterLoop, PlayerY * 30 - 11.6 + InnerLoop, Color
                End If
            Next
        Next
    End If
    If Bln_TrackInProgress = True Then
        Call DrawMonster
    Else
        FrmGreenEffect.Refresh
    End If
    Call Wait(2)  'Wait for 2 seconds
    
    Castras = Int(Castras / 2)
    PlayerHealth = 100
    
    Call RenderMap
    PlayerX = EnterX
    PlayerY = EnterY
    Call DoFadeIn
    Call BitBlt(FrmGreenEffect.hDC, 0, 0, PicBuffer.ScaleWidth, PicBuffer.ScaleHeight, PicBuffer.hDC, 0, 0, vbSrcCopy)
    Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PBU(Anim).hDC, 0, 0, vbSrcPaint)
    Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PicUp(Anim).hDC, 0, 0, vbSrcAnd)
    FrmGreenEffect.Refresh
End Sub

Sub DoFadeOut()

    '--------------------------------------------------
    'Description:
    'Do the Fade out effect when time travelling
    'OutPut:
    'Draws an animation of the fade body(For Time Travel)
    '--------------------------------------------------
    Dim FadeOut As Integer
    
    TmrPlayer.Enabled = False
    Call BitBlt(FrmGreenEffect.hDC, 0, 0, PicBuffer.ScaleWidth, PicBuffer.ScaleHeight, PicBuffer.hDC, 0, 0, vbSrcCopy)
    Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PBD(0).hDC, 0, 0, vbSrcPaint)
    Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PicFade(0).hDC, 0, 0, vbSrcAnd)
    FrmGreenEffect.Refresh
    For FadeOut = 0 To 5
        Call Wait(0.2)  'Wait for .2 of a second
        
        Call BitBlt(FrmGreenEffect.hDC, 0, 0, PicBuffer.ScaleWidth, PicBuffer.ScaleHeight, PicBuffer.hDC, 0, 0, vbSrcCopy)
        Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PBD(0).hDC, 0, 0, vbSrcPaint)
        Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PicFade(FadeOut).hDC, 0, 0, vbSrcAnd)
        FrmGreenEffect.Refresh
    Next
    TmrPlayer.Enabled = True
End Sub

Sub DoFadeIn()

    '--------------------------------------------------
    'Description:
    'Do the fade in effect after you died
    '--------------------------------------------------

    Dim FadeIn As Integer
    'Stop the player timer so it won't interfer with the fading
    TmrPlayer.Enabled = False
    Call BitBlt(FrmGreenEffect.hDC, 0, 0, PicBuffer.ScaleWidth, PicBuffer.ScaleHeight, PicBuffer.hDC, 0, 0, vbSrcCopy)
    Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PBU(0).hDC, 0, 0, vbSrcPaint)
    Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PicFade(5).hDC, 0, 0, vbSrcAnd)
    For FadeIn = 5 To 0 Step -1
        Call Wait(0.2)  'Wait for .2 of a second
        
        Call BitBlt(FrmGreenEffect.hDC, 0, 0, PicBuffer.ScaleWidth, PicBuffer.ScaleHeight, PicBuffer.hDC, 0, 0, vbSrcCopy)
        
        Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PBU(0).hDC, 0, 0, vbSrcPaint)
        Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PicFade(FadeIn).hDC, 0, 0, vbSrcAnd)
        If Bln_TrackInProgress = True Then
            Call DrawMonster
        Else
            FrmGreenEffect.Refresh
        End If
    Next
    DoEvents
    LastAction = 2
    TmrPlayer.Enabled = True
End Sub

Private Sub LblItems_Click()
    TmrPlayer.Enabled = False
    FrmWStats.Show
    Call WStatShow
End Sub

Sub LightMap(ByVal X As Integer, ByVal Y As Integer, ByVal Rad As Integer, ByVal Way As Integer, ByVal Inten As Integer)

    '--------------------------------------------------
    'Description:
    'Brightens a certain area mimicing light
    'Inputs:
    'X = X position on screen
    'Y = Y Position on screen
    'Rad = Radius of the circle
    'Way = Way the circle is facing
    'Inten = Intensity of the brightness
    'OutPuts:
    'Draws the bright patch on the screen
    '--------------------------------------------------
    
    If Way = 0 Then 'Whole Circle
        For OuterLoop = Y - Rad To Y + Rad
            For InnerLoop = X - Rad To X + Rad
                Colour = GetPixel(PicBuffer.hDC, InnerLoop, OuterLoop)
                Call GetRgb(Colour, R1, G1, B1)
                ColAdd = Sin((InnerLoop - (X - Rad)) / (Rad / 2)) + Sin((OuterLoop - (Y - Rad)) / (Rad / 2))
                If ColAdd > 1 Then
                    ColAdd = ColAdd * Inten - Inten
                    SetPixelV PicBuffer.hDC, InnerLoop, OuterLoop, RGB(R1 + ColAdd, G1 + ColAdd, B1 + ColAdd)
                End If
            Next
        Next
    ElseIf Way = 1 Then 'Left Side
        For OuterLoop = Y - Rad To Y + Rad
            For InnerLoop = X - Rad To X
                Colour = GetPixel(PicBuffer.hDC, InnerLoop, OuterLoop)
                Call GetRgb(Colour, R1, G1, B1)
                ColAdd = Sin((InnerLoop - (X - Rad)) / (Rad / 2)) + Sin((OuterLoop - (Y - Rad)) / (Rad / 2))
                If ColAdd > 1 Then
                    ColAdd = ColAdd * Inten - Inten
                    SetPixelV PicBuffer.hDC, InnerLoop, OuterLoop, RGB(R1 + ColAdd, G1 + ColAdd, B1 + ColAdd)
                End If
            Next
        Next
    ElseIf Way = 2 Then 'Right Side
        For OuterLoop = Y - Rad To Y + Rad
            For InnerLoop = X To X + Rad
                Colour = GetPixel(PicBuffer.hDC, InnerLoop, OuterLoop)
                Call GetRgb(Colour, R1, G1, B1)
                ColAdd = Sin((InnerLoop - (X - Rad)) / (Rad / 2)) + Sin((OuterLoop - (Y - Rad)) / (Rad / 2))
                If ColAdd > 1 Then
                    ColAdd = ColAdd * Inten - Inten
                    SetPixelV PicBuffer.hDC, InnerLoop, OuterLoop, RGB(R1 + ColAdd, G1 + ColAdd, B1 + ColAdd)
                End If
            Next
        Next
    ElseIf Way = 3 Then 'Top
        For OuterLoop = Y - Rad To Y
            For InnerLoop = X - Rad To X + Rad
                Colour = GetPixel(PicBuffer.hDC, InnerLoop, OuterLoop)
                Call GetRgb(Colour, R1, G1, B1)
                ColAdd = Sin((InnerLoop - (X - Rad)) / (Rad / 2)) + Sin((OuterLoop - (Y - Rad)) / (Rad / 2))
                If ColAdd > 1 Then
                    ColAdd = ColAdd * Inten - Inten
                    SetPixelV PicBuffer.hDC, InnerLoop, OuterLoop, RGB(R1 + ColAdd, G1 + ColAdd, B1 + ColAdd)
                End If
           Next
        Next
    ElseIf Way = 4 Then 'Bottom
        For OuterLoop = Y To Y + Rad
            For InnerLoop = X - Rad To X + Rad
                Colour = GetPixel(PicBuffer.hDC, InnerLoop, OuterLoop)
                Call GetRgb(Colour, R1, G1, B1)
                ColAdd = Sin((InnerLoop - (X - Rad)) / (Rad / 2)) + Sin((OuterLoop - (Y - Rad)) / (Rad / 2))
                If ColAdd > 1 Then
                    ColAdd = ColAdd * Inten - Inten
                    SetPixelV PicBuffer.hDC, InnerLoop, OuterLoop, RGB(R1 + ColAdd, G1 + ColAdd, B1 + ColAdd)
                End If
            Next
        Next
    End If
    PicBuffer.Refresh
End Sub

Sub StartTrack(ByVal StartingX As Integer, ByVal StartingY As Integer, ByVal EndX As Integer, ByVal EndY As Integer)
    Dim HaveToDo As Boolean
    StartX = StartingX  'Send the starting X position of the track into memory
    StartY = StartingY  'Send the starting Y position of the track into memory
    Int_EndTrackX = EndX
    Int_EndTrackY = EndY
    
    'As the end position would normally be the player, the tracking system works by
    'Making a route from the player to the tracker, the tracker then tracks by finding
    'the route which is lowest and when the number reaches 1 it has reached the player
    
    Int_NodesToDo(Int_EndTrackX, Int_EndTrackY) = 1 'This starts the node processing
    Int_TrackNodes(Int_EndTrackX, Int_EndTrackY) = 1    'Set the end value(Don't change)
    Bln_TrackInProgress = True  'Tell the computer that there is a track in progress
    'Only exit if there is a link between the end position and beginning position
    'or there are more nodes to be processed; in which case a route was not made
    Do
        HaveToDo = False    'Set to False, if a node needs processing it is set to true
        'if by the end of the loop it is still false exit the main loop and no route was made.
        
        'Go though the grid and check to see if there is a node which needs processing
        For X = 1 To 10
            For Y = 1 To 10
                 If Int_NodesToDo(X, Y) = 1 Then
                    'Check if a node needs processing if yes set Havetogo to True
                    HaveToDo = True
                    'Set the current position to that it does not need processing
                    Int_NodesToDo(X, Y) = 0
                    'Check to see if going up is within array bounds
                    If Y - 1 >= 1 Then
                        'Only Place new node if it is smaller than the which is already
                        'there or if it is 0 in which case it has not been used
                        If Int_TrackNodes(X, Y - 1) = 0 Or Int_TrackNodes(X, Y) + 1 < Int_TrackNodes(X, Y - 1) Then
                            'Check if the land is passable, proceed if it is
                            If Int_PassableLand(X, Y - 1) = 0 Then
                                'Different Land types give different speeds
                                If Str_Map(X, Y - 1) = "G" Then
                                    Int_TrackNodes(X, Y - 1) = Int_TrackNodes(X, Y) + Rnd * Lng_Modifer + 2   'Makes the movement more random
                                ElseIf Str_Map(X, Y - 1) = "P" Then
                                    Int_TrackNodes(X, Y - 1) = Int_TrackNodes(X, Y) + Rnd * Lng_Modifer + 1   'Makes the movement more random
                                Else
                                    Int_TrackNodes(X, Y - 1) = Int_TrackNodes(X, Y) + Rnd * Lng_Modifer + 3   'Makes the movement more random
                                End If
                                'Set the new node to have processing
                                Int_NodesToDo(X, Y - 1) = 1
                            End If
                        End If
                    End If
                    'check to see if going left is within array bounds
                    If X - 1 >= 1 Then
                        'Only Place new node if it is smaller than the which is already
                        'there or if it is 0 in which case it has not been used
                        If Int_TrackNodes(X - 1, Y) = 0 Or Int_TrackNodes(X, Y) + 1 < Int_TrackNodes(X - 1, Y) Then
                            'Check if the land is passable, proceed if it is
                            If Int_PassableLand(X - 1, Y) = 0 Then
                                'Different Land types give different speeds
                                If Str_Map(X - 1, Y) = "G" Then
                                    Int_TrackNodes(X - 1, Y) = Int_TrackNodes(X, Y) + Rnd * Lng_Modifer + 2   'Makes the movement more random
                                ElseIf Str_Map(X - 1, Y) = "P" Then
                                    Int_TrackNodes(X - 1, Y) = Int_TrackNodes(X, Y) + Rnd * Lng_Modifer + 1   'Makes the movement more random
                                Else
                                    Int_TrackNodes(X - 1, Y) = Int_TrackNodes(X, Y) + Rnd * Lng_Modifer + 3   'Makes the movement more random
                                End If
                                'Set the new node to have processing
                                Int_NodesToDo(X - 1, Y) = 1
                            End If
                        End If
                    End If
                    'check to see if going right is within array bounds
                    If X + 1 <= 10 Then
                        'Only Place new node if it is smaller than the which is already
                        'there or if it is 0 in which case it has not been used
                        If Int_TrackNodes(X + 1, Y) = 0 Or Int_TrackNodes(X, Y) + 1 < Int_TrackNodes(X + 1, Y) Then
                            'Check if the land is passable, proceed if it is
                            If Int_PassableLand(X + 1, Y) = 0 Then
                                'Different Land types give different speeds
                                If Str_Map(X + 1, Y) = "G" Then
                                    Int_TrackNodes(X + 1, Y) = Int_TrackNodes(X, Y) + Rnd * Lng_Modifer + 2   'Makes the movement more random
                                ElseIf Str_Map(X + 1, Y) = "P" Then
                                    Int_TrackNodes(X + 1, Y) = Int_TrackNodes(X, Y) + Rnd * Lng_Modifer + 1   'Makes the movement more random
                                Else
                                    Int_TrackNodes(X + 1, Y) = Int_TrackNodes(X, Y) + Rnd * Lng_Modifer + 3  'Makes the movement more random
                                End If
                                'Set the new node to have processing
                                Int_NodesToDo(X + 1, Y) = 1
                            End If
                        End If
                    End If
                    'Check to see if going down is within array bounds
                    If Y + 1 <= 10 Then
                        'Only Place new node if it is smaller than the which is already
                        'there or if it is 0 in which case it has not been used
                        If Int_TrackNodes(X, Y + 1) = 0 Or Int_TrackNodes(X, Y) + 1 < Int_TrackNodes(X, Y + 1) Then
                            'Check if the land is passable, proceed if it is
                            If Int_PassableLand(X, Y + 1) = 0 Then
                                'Different Land types give different speeds
                                If Str_Map(X, Y + 1) = "G" Then
                                    Int_TrackNodes(X, Y + 1) = Int_TrackNodes(X, Y) + Rnd * Lng_Modifer + 2
                                ElseIf Str_Map(X, Y + 1) = "P" Then
                                    Int_TrackNodes(X, Y + 1) = Int_TrackNodes(X, Y) + Rnd * Lng_Modifer + 1
                                Else
                                    Int_TrackNodes(X, Y + 1) = Int_TrackNodes(X, Y) + Rnd * Lng_Modifer + 3   'Makes the movement more random
                                End If
                                'Set the new node to have processing
                                Int_NodesToDo(X, Y + 1) = 1
                            End If
                        End If
                    End If
                End If
            Next
        Next
    Loop Until Int_TrackNodes(StartX, StartY) <> 0 Or HaveToDo = 0
    'Set the current position to the starting position
    Int_CurrentTrackX = StartX
    Int_CurrentTrackY = StartY
End Sub

Sub ClearTrack()
    'Clears the track nodes to perpare it for a new track
    For X = 1 To 10
        For Y = 1 To 10
            Int_NodesToDo(X, Y) = 0 'Clears nodes to be processed
            Int_TrackNodes(X, Y) = 0    'Clear the node values
        Next
    Next
    Int_CurrentTrackX = 1   'This Sets the current X position
    Int_CurrentTrackY = 1   'This sets the current Y position
    Int_EndTrackX = 0   'This clears the end X position
    Int_EndTrackY = 0   'This clears teh end Y Position
    Bln_TrackInProgress = False 'This tells the computer there is no tracking process running
End Sub

Sub ContinueTrack()
    If PlayerX <> Int_EndTrackX Or PlayerY <> Int_EndTrackY Then
        'This checks to see if the player has moved from his original position
        'If he has, recalculate from the new position
        
        'Store the Current Position because ClearTrack will clear it
        TempX = Int_CurrentTrackX
        TempY = Int_CurrentTrackY
        'End Current Track, Clear Nodes and Tracking Infomation
        Call ClearTrack
        'Start a new track as the current position as a source and the players new
        'position as the end point
        Call StartTrack(TempX, TempY, PlayerX, PlayerY)
    End If
    'This retrieves the current Node value, then it will see if it can lower the value
    'by going up, down, left or right. The lower the number the closer it is to the
    'Player. 1 would mean that it is at the end point
    CurValue = Int_TrackNodes(Int_CurrentTrackX, Int_CurrentTrackY)
    'Check to see if going up is within array bounds
    If Int_CurrentTrackY - 1 >= 1 Then
        'Check to see if the position has a lower value, if it is 0 the position was
        'not used and so is invalid
        If Int_TrackNodes(Int_CurrentTrackX, Int_CurrentTrackY - 1) < CurValue And Int_TrackNodes(Int_CurrentTrackX, Int_CurrentTrackY - 1) > 0 Then
            'Set the Choosen Path to 1(Up) and set the new Current Value to check from
            Choosen = 1
            CurValue = Int_TrackNodes(Int_CurrentTrackX, Int_CurrentTrackY - 1)
        End If
    End If
    'Check to see if going down is within array bounds
    If Int_CurrentTrackY + 1 <= 10 Then
        'Check to see if the position has a lower value, if it is 0 the position was
        'not used and so is invalid
        If Int_TrackNodes(Int_CurrentTrackX, Int_CurrentTrackY + 1) < CurValue And Int_TrackNodes(Int_CurrentTrackX, Int_CurrentTrackY + 1) > 0 Then
            'Set the Choosen Path to 2(Down) and set the new Current Value to check from
            Choosen = 2
            CurValue = Int_TrackNodes(Int_CurrentTrackX, Int_CurrentTrackY + 1)
        End If
    End If
    'Check to see if going left is within array bounds
    If Int_CurrentTrackX - 1 >= 1 Then
        'Check to see if the position has a lower value, if it is 0 the position was
        'not used and so is invalid
        If Int_TrackNodes(Int_CurrentTrackX - 1, Int_CurrentTrackY) < CurValue And Int_TrackNodes(Int_CurrentTrackX - 1, Int_CurrentTrackY) > 0 Then
            'Set the Choosen Path to 3(Left) and set the new Current Value to check from
            Choosen = 3
            CurValue = Int_TrackNodes(Int_CurrentTrackX - 1, Int_CurrentTrackY)
        End If
    End If
    'Check to see if going right is within array bounds
    If Int_CurrentTrackX + 1 <= 10 Then
        'Check to see if the position has a lower value, if it is 0 the position was
        'not used and so is invalid
        If Int_TrackNodes(Int_CurrentTrackX + 1, Int_CurrentTrackY) < CurValue And Int_TrackNodes(Int_CurrentTrackX + 1, Int_CurrentTrackY) > 0 Then
            'Set the Choosen Path to 4(Right) and set the new Current Value to check from
            Choosen = 4
            CurValue = Int_TrackNodes(Int_CurrentTrackX + 1, Int_CurrentTrackY)
        End If
    End If
    Int_MonsterAnim = 1 - Int_MonsterAnim
    If Choosen = 1 Then
        Int_CurrentTrackY = Int_CurrentTrackY - 1
        If CurValue = 1 Then
            'If the new position means that it is at the end position stop the track
            Int_CurrentTrackY = Int_CurrentTrackY + 1
            Call AttackMTP(ScreenX, ScreenY)
        End If
        Call BitBlt(FrmGreenEffect.hDC, Int_CurrentTrackX * 30 - 9, Int_CurrentTrackY * 30 - 9, 33, 33, PBU(Int_MonsterAnim).hDC, 0, 0, vbSrcPaint)
        Call BitBlt(FrmGreenEffect.hDC, Int_CurrentTrackX * 30 - 9, Int_CurrentTrackY * 30 - 9, 33, 33, PicUp(Int_MonsterAnim).hDC, 0, 0, vbSrcAnd)
        Int_MonsterLastPos = 1
    ElseIf Choosen = 2 Then
        Int_CurrentTrackY = Int_CurrentTrackY + 1
        If CurValue = 1 Then
            'If the new position means that it is at the end position stop the track
            Int_CurrentTrackY = Int_CurrentTrackY - 1
            Call AttackMTP(ScreenX, ScreenY)
        End If
        Call BitBlt(FrmGreenEffect.hDC, Int_CurrentTrackX * 30 - 9, Int_CurrentTrackY * 30 - 9, 33, 33, PBD(Int_MonsterAnim).hDC, 0, 0, vbSrcPaint)
        Call BitBlt(FrmGreenEffect.hDC, Int_CurrentTrackX * 30 - 9, Int_CurrentTrackY * 30 - 9, 33, 33, PicDown(Int_MonsterAnim).hDC, 0, 0, vbSrcAnd)
        Int_MonsterLastPos = 2
    ElseIf Choosen = 3 Then
        Int_CurrentTrackX = Int_CurrentTrackX - 1
        If CurValue = 1 Then
            'If the new position means that it is at the end position stop the track
            Int_CurrentTrackX = Int_CurrentTrackX + 1
            Call AttackMTP(ScreenX, ScreenY)
        End If
        Call BitBlt(FrmGreenEffect.hDC, Int_CurrentTrackX * 30 - 9, Int_CurrentTrackY * 30 - 9, 33, 33, PBL(Int_MonsterAnim).hDC, 0, 0, vbSrcPaint)
        Call BitBlt(FrmGreenEffect.hDC, Int_CurrentTrackX * 30 - 9, Int_CurrentTrackY * 30 - 9, 33, 33, PicLeft(Int_MonsterAnim).hDC, 0, 0, vbSrcAnd)
        Int_MonsterLastPos = 3
    ElseIf Choosen = 4 Then
        Int_CurrentTrackX = Int_CurrentTrackX + 1
        If CurValue = 1 Then
            'If the new position means that it is at the end position stop the track
            Int_CurrentTrackX = Int_CurrentTrackX - 1
            Call AttackMTP(ScreenX, ScreenY)
        End If
        Call BitBlt(FrmGreenEffect.hDC, Int_CurrentTrackX * 30 - 9, Int_CurrentTrackY * 30 - 9, 33, 33, PBR(Int_MonsterAnim).hDC, 0, 0, vbSrcPaint)
        Call BitBlt(FrmGreenEffect.hDC, Int_CurrentTrackX * 30 - 9, Int_CurrentTrackY * 30 - 9, 33, 33, PicRight(Int_MonsterAnim).hDC, 0, 0, vbSrcAnd)
        Int_MonsterLastPos = 4
    Else
        Int_MonsterAnim = 1 - Int_MonsterAnim
        Call BitBlt(FrmGreenEffect.hDC, Int_CurrentTrackX * 30 - 9, Int_CurrentTrackY * 30 - 9, 33, 33, PBD(Int_MonsterAnim).hDC, 0, 0, vbSrcPaint)
        Call BitBlt(FrmGreenEffect.hDC, Int_CurrentTrackX * 30 - 9, Int_CurrentTrackY * 30 - 9, 33, 33, PicDown(Int_MonsterAnim).hDC, 0, 0, vbSrcAnd)
        Int_MonsterLastPos = 2
    End If
    'Show the position of the tracking computer player (For debuging and testing
    'perposes only
    Label4.Caption = Str(Int_CurrentTrackX) & "," & Str(Int_CurrentTrackY)
    'Refresh the screen to make changes visible
    FrmGreenEffect.Refresh
End Sub

Sub DrawMonster()
    'This is only called if:
    '1. the monsters picture is clear from screen
    '2. The tracking system is then not called
    '3. But the monsters picture is still needed
    If Int_MonsterLastPos = 1 Then
        Call BitBlt(FrmGreenEffect.hDC, Int_CurrentTrackX * 30 - 9, Int_CurrentTrackY * 30 - 9, 33, 33, PBU(Int_MonsterAnim).hDC, 0, 0, vbSrcPaint)
        Call BitBlt(FrmGreenEffect.hDC, Int_CurrentTrackX * 30 - 9, Int_CurrentTrackY * 30 - 9, 33, 33, PicUp(Int_MonsterAnim).hDC, 0, 0, vbSrcAnd)
    ElseIf Int_MonsterLastPos = 2 Then
        Call BitBlt(FrmGreenEffect.hDC, Int_CurrentTrackX * 30 - 9, Int_CurrentTrackY * 30 - 9, 33, 33, PBD(Int_MonsterAnim).hDC, 0, 0, vbSrcPaint)
        Call BitBlt(FrmGreenEffect.hDC, Int_CurrentTrackX * 30 - 9, Int_CurrentTrackY * 30 - 9, 33, 33, PicDown(Int_MonsterAnim).hDC, 0, 0, vbSrcAnd)
    ElseIf Int_MonsterLastPos = 3 Then
        Call BitBlt(FrmGreenEffect.hDC, Int_CurrentTrackX * 30 - 9, Int_CurrentTrackY * 30 - 9, 33, 33, PBL(Int_MonsterAnim).hDC, 0, 0, vbSrcPaint)
        Call BitBlt(FrmGreenEffect.hDC, Int_CurrentTrackX * 30 - 9, Int_CurrentTrackY * 30 - 9, 33, 33, PicLeft(Int_MonsterAnim).hDC, 0, 0, vbSrcAnd)
    ElseIf Int_MonsterLastPos = 4 Then
        Call BitBlt(FrmGreenEffect.hDC, Int_CurrentTrackX * 30 - 9, Int_CurrentTrackY * 30 - 9, 33, 33, PBR(Int_MonsterAnim).hDC, 0, 0, vbSrcPaint)
        Call BitBlt(FrmGreenEffect.hDC, Int_CurrentTrackX * 30 - 9, Int_CurrentTrackY * 30 - 9, 33, 33, PicRight(Int_MonsterAnim).hDC, 0, 0, vbSrcAnd)
    Else
        Call BitBlt(FrmGreenEffect.hDC, Int_CurrentTrackX * 30 - 9, Int_CurrentTrackY * 30 - 9, 33, 33, PBD(Int_MonsterAnim).hDC, 0, 0, vbSrcPaint)
        Call BitBlt(FrmGreenEffect.hDC, Int_CurrentTrackX * 30 - 9, Int_CurrentTrackY * 30 - 9, 33, 33, PicDown(Int_MonsterAnim).hDC, 0, 0, vbSrcAnd)
    End If
    FrmGreenEffect.Refresh
End Sub

Sub Wait(ByVal Pause As Double)
    'Pauses the application for a certain length of time but allows processes to run in
    'the background
    BeforeTime = Timer
    Do
        DoEvents
    Loop Until Timer > BeforeTime + Pause
End Sub

Sub GetMoney(ByVal Amount As Integer)
    LblMessage.Caption = "You have killed " & Monster(ScreenX, ScreenY).Name
    LblText.Caption = "He was carrying" & Str(Amount) & " Castras"
    Call Wait(2)
    LblMessage.Caption = ""
    LblText.Caption = ""
    Castras = Castras + Amount
End Sub

Sub DrawDead()
    For OuterLoop = 1 To 32
        For InnerLoop = 1 To 32
            Color = GetPixel(PicDead(0).hDC, InnerLoop, OuterLoop)
            If Color <> vbGreen Then
                SetPixelV PicBuffer.hDC, Int_CurrentTrackX * 30 - 11.6 + InnerLoop, Int_CurrentTrackY * 30 - 9 + OuterLoop, Color
            End If
        Next
    Next
    PicBuffer.Refresh
    Call BitBlt(FrmGreenEffect.hDC, 0, 0, PicBuffer.ScaleWidth, PicBuffer.ScaleHeight, PicBuffer.hDC, 0, 0, vbSrcCopy)
    If Int_MonsterLastPos = 1 Then
        Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PBD(Int_MonsterAnim).hDC, 0, 0, vbSrcPaint)
        Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PicDown(Int_MonsterAnim).hDC, 0, 0, vbSrcAnd)
    ElseIf Int_MonsterLastPos = 2 Then
        Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PBU(Int_MonsterAnim).hDC, 0, 0, vbSrcPaint)
        Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PicUp(Int_MonsterAnim).hDC, 0, 0, vbSrcAnd)
    ElseIf Int_MonsterLastPos = 3 Then
        Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PBL(Int_MonsterAnim).hDC, 0, 0, vbSrcPaint)
        Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PicLeft(Int_MonsterAnim).hDC, 0, 0, vbSrcAnd)
    ElseIf Int_MonsterLastPos = 4 Then
        Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PBR(Int_MonsterAnim).hDC, 0, 0, vbSrcPaint)
        Call BitBlt(FrmGreenEffect.hDC, PlayerX * 30 - 9, PlayerY * 30 - 9, 33, 33, PicRight(Int_MonsterAnim).hDC, 0, 0, vbSrcAnd)
    End If
End Sub
