VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MapEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GreenEffect Map Editor"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8670
   Icon            =   "FrmMapEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   8670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdRndGrass 
      Caption         =   "Random Grass (Advanced)"
      Height          =   375
      Left            =   5400
      TabIndex        =   124
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Frame Frame3 
      Caption         =   "The Total Map"
      Height          =   3315
      Left            =   5400
      TabIndex        =   122
      Top             =   1920
      Width           =   3195
      Begin VB.PictureBox PicTotal 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   2955
         Left            =   120
         ScaleHeight     =   197
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   197
         TabIndex        =   123
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tiles"
      Height          =   1755
      Left            =   5400
      TabIndex        =   103
      Top             =   60
      Width           =   3195
      Begin VB.CommandButton CmdView 
         Caption         =   "Preview"
         Height          =   375
         Left            =   1140
         TabIndex        =   121
         Top             =   660
         Width           =   975
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   120
         TabIndex        =   107
         Top             =   660
         Width           =   975
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   1140
         TabIndex        =   106
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdLoad 
         Caption         =   "Load"
         Height          =   375
         Left            =   120
         TabIndex        =   105
         Top             =   240
         Width           =   975
      End
      Begin VB.PictureBox PicCopy 
         AutoRedraw      =   -1  'True
         Height          =   495
         Left            =   2400
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   104
         Top             =   360
         Width           =   495
      End
      Begin MSComctlLib.ProgressBar Pb 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   108
         Top             =   1080
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Max             =   200
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar Pb 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   109
         Top             =   1380
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Max             =   200
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tiles"
      Height          =   2835
      Left            =   120
      TabIndex        =   102
      Top             =   5340
      Width           =   5055
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   49
         Left            =   4440
         Picture         =   "FrmMapEditor.frx":0442
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   163
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   500
         Index           =   48
         Left            =   3960
         Picture         =   "FrmMapEditor.frx":08D7
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   162
         Top             =   2160
         Width           =   500
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   47
         Left            =   3480
         Picture         =   "FrmMapEditor.frx":0D51
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   161
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   46
         Left            =   3000
         Picture         =   "FrmMapEditor.frx":1A77
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   160
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   45
         Left            =   2520
         Picture         =   "FrmMapEditor.frx":2006
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   159
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   44
         Left            =   2040
         Picture         =   "FrmMapEditor.frx":2431
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   158
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   43
         Left            =   1560
         Picture         =   "FrmMapEditor.frx":28A1
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   157
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   42
         Left            =   1080
         Picture         =   "FrmMapEditor.frx":2CEE
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   156
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   500
         Index           =   41
         Left            =   600
         Picture         =   "FrmMapEditor.frx":3188
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   155
         Top             =   2160
         Width           =   500
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   40
         Left            =   120
         Picture         =   "FrmMapEditor.frx":35A6
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   154
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   39
         Left            =   4440
         Picture         =   "FrmMapEditor.frx":39CB
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   153
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   38
         Left            =   3960
         Picture         =   "FrmMapEditor.frx":3E37
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   152
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   37
         Left            =   3480
         Picture         =   "FrmMapEditor.frx":42A2
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   151
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   36
         Left            =   3000
         Picture         =   "FrmMapEditor.frx":4628
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   150
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   35
         Left            =   2520
         Picture         =   "FrmMapEditor.frx":4ADB
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   149
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   34
         Left            =   2040
         Picture         =   "FrmMapEditor.frx":4F94
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   148
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   33
         Left            =   1560
         Picture         =   "FrmMapEditor.frx":53CC
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   147
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   32
         Left            =   1080
         Picture         =   "FrmMapEditor.frx":5821
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   146
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   31
         Left            =   600
         Picture         =   "FrmMapEditor.frx":5C7F
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   145
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   500
         Index           =   30
         Left            =   120
         Picture         =   "FrmMapEditor.frx":613F
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   144
         Top             =   1680
         Width           =   500
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         Height          =   495
         Index           =   29
         Left            =   4440
         Picture         =   "FrmMapEditor.frx":656F
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   143
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         Height          =   495
         Index           =   28
         Left            =   3960
         Picture         =   "FrmMapEditor.frx":69D8
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   142
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         Height          =   495
         Index           =   27
         Left            =   3480
         Picture         =   "FrmMapEditor.frx":6D35
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   141
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         Height          =   495
         Index           =   26
         Left            =   3000
         Picture         =   "FrmMapEditor.frx":71F7
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   140
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         Height          =   495
         Index           =   25
         Left            =   2520
         Picture         =   "FrmMapEditor.frx":7600
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   139
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         Height          =   495
         Index           =   24
         Left            =   2040
         Picture         =   "FrmMapEditor.frx":7A69
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   138
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         Height          =   495
         Index           =   23
         Left            =   1560
         Picture         =   "FrmMapEditor.frx":7ECF
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   137
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         Height          =   495
         Index           =   22
         Left            =   1080
         Picture         =   "FrmMapEditor.frx":8322
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   136
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         Height          =   495
         Index           =   21
         Left            =   600
         Picture         =   "FrmMapEditor.frx":8709
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   135
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         Height          =   495
         Index           =   20
         Left            =   120
         Picture         =   "FrmMapEditor.frx":8AD6
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   134
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   500
         Index           =   19
         Left            =   4440
         Picture         =   "FrmMapEditor.frx":8F03
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   133
         Top             =   720
         Width           =   500
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   500
         Index           =   18
         Left            =   3960
         Picture         =   "FrmMapEditor.frx":9399
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   132
         Top             =   720
         Width           =   500
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   17
         Left            =   3480
         Picture         =   "FrmMapEditor.frx":9811
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   131
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   16
         Left            =   3000
         Picture         =   "FrmMapEditor.frx":9D3E
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   130
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   15
         Left            =   2520
         Picture         =   "FrmMapEditor.frx":A268
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   129
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   14
         Left            =   2040
         Picture         =   "FrmMapEditor.frx":A7AB
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   128
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   495
         Index           =   9
         Left            =   4440
         Picture         =   "FrmMapEditor.frx":ACD1
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   127
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   13
         Left            =   1560
         Picture         =   "FrmMapEditor.frx":B1D7
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   126
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   12
         Left            =   1080
         Picture         =   "FrmMapEditor.frx":B701
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   125
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   11
         Left            =   600
         Picture         =   "FrmMapEditor.frx":BC28
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   120
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   10
         Left            =   120
         Picture         =   "FrmMapEditor.frx":C120
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   119
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   8
         Left            =   3960
         Picture         =   "FrmMapEditor.frx":C659
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   118
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   7
         Left            =   3480
         Picture         =   "FrmMapEditor.frx":CBA8
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   117
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   6
         Left            =   3000
         Picture         =   "FrmMapEditor.frx":D0E4
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   116
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   5
         Left            =   2520
         Picture         =   "FrmMapEditor.frx":D623
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   115
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   4
         Left            =   2040
         Picture         =   "FrmMapEditor.frx":DB5E
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   114
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   3
         Left            =   1560
         Picture         =   "FrmMapEditor.frx":E09F
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   113
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   2
         Left            =   1080
         Picture         =   "FrmMapEditor.frx":E5CB
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   112
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   1
         Left            =   600
         Picture         =   "FrmMapEditor.frx":EA05
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   111
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox PicChoose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   495
         Index           =   0
         Left            =   120
         Picture         =   "FrmMapEditor.frx":EF39
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   110
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.VScrollBar VS 
      Height          =   4815
      Left            =   4980
      Max             =   19
      TabIndex        =   101
      Top             =   120
      Width           =   315
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   99
      Left            =   4440
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   100
      Top             =   4440
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   98
      Left            =   3960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   99
      Top             =   4440
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   97
      Left            =   3480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   98
      Top             =   4440
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   96
      Left            =   3000
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   97
      Top             =   4440
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   95
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   96
      Top             =   4440
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   94
      Left            =   2040
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   95
      Top             =   4440
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   93
      Left            =   1560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   94
      Top             =   4440
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   92
      Left            =   1080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   93
      Top             =   4440
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   91
      Left            =   600
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   92
      Top             =   4440
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   90
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   91
      Top             =   4440
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   89
      Left            =   4440
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   90
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   88
      Left            =   3960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   89
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   87
      Left            =   3480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   88
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   86
      Left            =   3000
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   87
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   85
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   86
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   84
      Left            =   2040
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   85
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   83
      Left            =   1560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   84
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   82
      Left            =   1080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   83
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   81
      Left            =   600
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   82
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   80
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   81
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   79
      Left            =   4440
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   80
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   78
      Left            =   3960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   79
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   77
      Left            =   3480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   78
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   76
      Left            =   3000
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   77
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   75
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   76
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   74
      Left            =   2040
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   75
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   73
      Left            =   1560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   74
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   72
      Left            =   1080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   73
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   71
      Left            =   600
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   72
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   70
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   71
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   69
      Left            =   4440
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   70
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   68
      Left            =   3960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   69
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   67
      Left            =   3480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   68
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   66
      Left            =   3000
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   67
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   65
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   66
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   64
      Left            =   2040
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   65
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   63
      Left            =   1560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   64
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   62
      Left            =   1080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   63
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   61
      Left            =   600
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   62
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   60
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   61
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   59
      Left            =   4440
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   60
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   58
      Left            =   3960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   59
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   57
      Left            =   3480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   58
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   56
      Left            =   3000
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   57
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   55
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   56
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   54
      Left            =   2040
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   55
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   53
      Left            =   1560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   54
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   52
      Left            =   1080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   53
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   51
      Left            =   600
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   52
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   50
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   51
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   49
      Left            =   4440
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   50
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   48
      Left            =   3960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   49
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   47
      Left            =   3480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   48
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   46
      Left            =   3000
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   47
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   45
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   46
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   44
      Left            =   2040
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   45
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   43
      Left            =   1560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   44
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   42
      Left            =   1080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   43
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   41
      Left            =   600
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   42
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   40
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   41
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   39
      Left            =   4440
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   40
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   38
      Left            =   3960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   39
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   37
      Left            =   3480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   38
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   36
      Left            =   3000
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   37
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   35
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   36
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   34
      Left            =   2040
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   35
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   33
      Left            =   1560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   34
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   32
      Left            =   1080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   33
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   31
      Left            =   600
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   32
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   30
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   31
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   29
      Left            =   4440
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   30
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   28
      Left            =   3960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   29
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   27
      Left            =   3480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   28
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   26
      Left            =   3000
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   27
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   25
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   26
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   24
      Left            =   2040
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   25
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   23
      Left            =   1560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   24
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   22
      Left            =   1080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   23
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   21
      Left            =   600
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   22
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   20
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   21
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   19
      Left            =   4440
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   20
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   18
      Left            =   3960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   19
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   17
      Left            =   3480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   18
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   16
      Left            =   3000
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   17
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   15
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   16
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   14
      Left            =   2040
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   15
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   13
      Left            =   1560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   14
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   12
      Left            =   1080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   13
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   11
      Left            =   600
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   12
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   10
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   11
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   9
      Left            =   4440
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   10
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   8
      Left            =   3960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   9
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   7
      Left            =   3480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   6
      Left            =   3000
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   5
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   6
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   4
      Left            =   2040
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   3
      Left            =   1560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   2
      Left            =   1080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   600
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox PicGrid 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   0
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   120
      Width           =   495
      Begin MSComDlg.CommonDialog CDialog 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.HScrollBar HS 
      Height          =   315
      Left            =   120
      Max             =   19
      TabIndex        =   0
      Top             =   4980
      Width           =   4815
   End
End
Attribute VB_Name = "MapEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As Integer

Dim X, Y
Dim Map(1 To 200, 1 To 200) As String

Dim TotalMap(1 To 200) As String
Dim CompressedMap(1 To 200) As String
Dim Compress(1 To 200) As String

Private Sub CmdClear_Click()
    Call DoClear
End Sub

Private Sub CmdLoad_Click()
    CDialog.Filter = "Green Effect Map Files (*.map)|*.map"
    CDialog.ShowOpen
    File$ = CDialog.FileName
    If File$ = "" Then Exit Sub
    Open File$ For Input As #1
    For OuterLoop = 1 To 200
        Input #1, CompressedMap(OuterLoop)
    Next
    Close
    
    For OuterLoop = 1 To 200
        Pb(0) = OuterLoop
        TotalMap(OuterLoop) = ""
        For InnerLoop = 1 To 200
            If Mid(CompressedMap(OuterLoop), InnerLoop, 1) = "(" Then
                If Mid(CompressedMap(OuterLoop), InnerLoop + 2, 1) = ")" Then
                    For A = 1 To Val(Mid(CompressedMap(OuterLoop), InnerLoop + 1, 1))
                        TotalMap(OuterLoop) = TotalMap(OuterLoop) & Mid(CompressedMap(OuterLoop), InnerLoop + 3, 1)
                    Next
                    InnerLoop = InnerLoop + 3
                ElseIf Mid(CompressedMap(OuterLoop), InnerLoop + 3, 1) = ")" Then
                    For A = 1 To Val(Mid(CompressedMap(OuterLoop), InnerLoop + 1, 2))
                        TotalMap(OuterLoop) = TotalMap(OuterLoop) & Mid(CompressedMap(OuterLoop), InnerLoop + 4, 1)
                    Next
                    InnerLoop = InnerLoop + 4
                ElseIf Mid(CompressedMap(OuterLoop), InnerLoop + 4, 1) = ")" Then
                    For A = 1 To Val(Mid(CompressedMap(OuterLoop), InnerLoop + 1, 3))
                        TotalMap(OuterLoop) = TotalMap(OuterLoop) & Mid(CompressedMap(OuterLoop), InnerLoop + 5, 1)
                    Next
                    InnerLoop = InnerLoop + 5
                End If
            Else
                TotalMap(OuterLoop) = TotalMap(OuterLoop) & Mid(CompressedMap(OuterLoop), InnerLoop, 1)
            End If
        Next
    Next
    
    For OuterLoop = 1 To 200
        For InnerLoop = 1 To 200
            Map(InnerLoop, OuterLoop) = Mid$(TotalMap(OuterLoop), InnerLoop, 1)
        Next
    Next
    HS = 0
    VS = 0
    CharToGrid
End Sub

Sub CharToGrid()
 For OuterLoop = 1 To 10
        For InnerLoop = 1 To 10
            If Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "P" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(2).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "G" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(3).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "S" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(9).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "C" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(11).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "F" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(12).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "E" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(15).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "W" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(18).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "R" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(19).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "T" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(8).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "D" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(20).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "1" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(1).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "2" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(10).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "3" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(14).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "4" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(4).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "5" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(0).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "6" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(13).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "7" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(5).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "8" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(6).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "9" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(16).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "0" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(7).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "/" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(17).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "A" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(21).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "B" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(22).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "H" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(24).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "I" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(23).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "L" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(25).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "M" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(26).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "N" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(28).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "O" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(35).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "Q" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(34).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "U" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(36).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "V" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(33).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "^" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(31).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "X" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(30).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "Y" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(27).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "Z" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(32).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "!" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(38).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "%" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(37).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "£" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(39).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "$" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(29).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "&" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(40).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "*" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(41).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "[" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(42).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "]" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(43).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "{" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(44).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "}" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(45).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = ";" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(46).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "`" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(47).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = ":" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(48).Picture
            ElseIf Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "'" Then
                PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(49).Picture
            End If
        Next
    Next
End Sub

Sub GridToChar()
    For OuterLoop = 1 To 10
        For InnerLoop = 1 To 10
            If PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(0) Then         'Across Fence
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "5"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(1) Then     'Bottom Left Fence
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "1"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(2) Then     'Path
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "P"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(3) Then     'Grass
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "G"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(4) Then     'Left Vert Fence
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "4"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(5) Then     'Right Vert Fence
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "7"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(6) Then     'Stop At upper Left Fence
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "8"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(7) Then     'Stop At Upper Right Fence
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "0"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(8) Then     'Tree
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "T"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(9) Then     'Sand
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "S"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(10) Then    'Bottom Right Fence
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "2"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(11) Then    'Chest
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "C"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(12) Then    'Flowers
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "F"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(13) Then    'Stop At right Fence
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "6"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(14) Then    'Stop at left fence
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "3"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(15) Then   'GrassRocks
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "E"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(16) Then    'Top Left Fence
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "9"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(17) Then    'Top Right Fence
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "/"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(18) Then   'Water
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "W"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(19) Then   'Rock
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "R"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(20) Then   'Dirt Cobbles
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "D"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(21) Then    'Wall
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "A"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(22) Then   'Door
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "B"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(24) Then   'Stool
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "H"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(23) Then   'Window
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "I"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(25) Then   'BookCase
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "L"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(26) Then   'Case
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "M"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(27) Then    'Bottom Left
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "Y"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(28) Then   'Carpet
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "N"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(29) Then   'Bed
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "$"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(30) Then
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "X"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(31) Then
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "^"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(32) Then
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "Z"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(34) Then
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "Q"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(33) Then
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "V"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(35) Then
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "O"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(36) Then
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "U"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(38) Then
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "!"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(37) Then
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "%"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(39) Then
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "£"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(40) Then
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "&"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(41) Then
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "*"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(42) Then
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "["
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(43) Then
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "]"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(44) Then
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "{"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(45).Picture Then
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "}"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(46).Picture Then
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = ";"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(47).Picture Then
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "`"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(48).Picture Then
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = ":"
            ElseIf PicGrid(OuterLoop * 10 - 11 + InnerLoop).Picture = PicChoose(49).Picture Then
                Map((X * 10) + InnerLoop, (Y * 10) + OuterLoop) = "'"
            End If
        Next
    Next
End Sub

Private Sub CmdRndGrass_Click()
    'Randomly makes grass into long grass
    Call GridToChar
    For OuterLoop = 1 To 200
        For InnerLoop = 1 To 200
            If Map(InnerLoop, OuterLoop) = "G" Then
                RanVal = Rnd * 10
                If RanVal < 2 Then
                    Map(InnerLoop, OuterLoop) = ";"
                End If
            End If
        Next
    Next
    Call CharToGrid
End Sub

Private Sub CmdSave_Click()
    Call DoSave
End Sub

Private Sub CmdView_Click()
    Call DrawMap
End Sub

Private Sub Form_Load()
    Call DoClear
End Sub

Private Sub HS_Change()
    Call GridToChar
    X = HS
    Call CharToGrid
End Sub

Private Sub PicChoose_Click(Index As Integer)
    PicCopy.Picture = PicChoose(Index).Picture
End Sub

Private Sub PicGrid_Click(Index As Integer)
    If PicCopy.Picture = PicChoose(22).Picture Then
        Ask = MsgBox("Would You Like to update the *.Hus file?", vbYesNoCancel)
        If Ask = vbYes Then
            PicGrid(Index).Picture = PicCopy.Picture
            CDialog.Filter = "Green Effect House File (*.hus)|*.hus"
            CDialog.ShowSave
            File$ = CDialog.FileName
            If File$ = "" Then Exit Sub
            Open File$ For Input As #1
            Open "C:\GEMap.tmp" For Output As #2
            Input #1, Num$
            HouseNumber = Val(Num$) + 1
            Print #2, Mid(Str(Val(Num$) + 1), 2)
            Do
                Input #1, Text$
                Print #2, Text$
            Loop Until EOF(1)
            HusName$ = InputBox("Name of house", "Name", "")
            Key = InputBox("Key Needed to Open Door", "Key", 1)
            KeyName$ = InputBox("Name of Key needed to open door", "KeyName", "General Key")
            Print #2, "##Name " & HusName$
            Y1 = Int((Index + 1) / 10) - 1
            X1 = Index - (Y1 * 10)
            Print #2, "##XPos" & Str((X * 10) - 9 + X1)
            Print #2, "##YPos" & Str((Y * 10) + Y1 + 3)
            Print #2, "##Key" & Str(Key)
            Print #2, "##KName " & KeyName$
            Close
            Kill File$
            FileCopy "C:\GEMap.tmp", File$
            Kill "C:\GEMap.tmp"
            Ask = MsgBox("Would You Like to Create a default Files", vbYesNo)
            If Ask = vbYes Then
                Folder$ = Mid(File$, 1, Len(File$) - 11)
                Open Folder$ & "GMap" & Mid(Str(HouseNumber), 2) & ".PLS" For Output As #1
                Print #1, "0"
                Close
                Open Folder$ & "GMap" & Mid(Str(HouseNumber), 2) & ".MAP" For Output As #1
                For PNo = 1 To 200
                    Print #1, "(200)G"
                Next
                Close
                Open Folder$ & "GMap" & Mid(Str(HouseNumber), 2) & ".OLF" For Output As #1
                Close
            End If
        ElseIf Ask = vbNo Then
            PicGrid(Index).Picture = PicCopy.Picture
        End If
    Else
        PicGrid(Index).Picture = PicCopy.Picture
    End If
End Sub

Private Sub VS_Change()
    Call GridToChar
    Y = VS
    Call CharToGrid
End Sub

Sub DoSave()
    GridToChar
    For OuterLoop = 1 To 200
        TotalMap(OuterLoop) = ""
        For InnerLoop = 1 To 200
            Compress(OuterLoop) = Compress(OuterLoop) & Map(InnerLoop, OuterLoop)
        Next
    Next
    
    For OuterLoop = 1 To 200
        Pb(0) = OuterLoop
        For InnerLoop = 1 To 200
            TxtText = String(200, Mid(Compress(OuterLoop), InnerLoop, 1))
            For Check = 200 To 4 Step -1
                If Mid(Compress(OuterLoop), InnerLoop, Check) = Mid(TxtText, 1, Check) Then
                    TotalMap(OuterLoop) = TotalMap(OuterLoop) & "(" & Mid(Str(Check), 2) & ")"
                    InnerLoop = InnerLoop + Check - 1
                    Check = 0
                End If
            Next
            If Check = 4 Then
                TotalMap(OuterLoop) = TotalMap(OuterLoop) & Mid(Compress(OuterLoop), InnerLoop, 1)
            End If
            TotalMap(OuterLoop) = TotalMap(OuterLoop) & Mid(Compress(OuterLoop), InnerLoop, 1)
            Pb(1) = InnerLoop
        Next
    Next
    CDialog.Filter = "Green Effect Map File (*.map)|*.map"
    CDialog.ShowSave
    File$ = CDialog.FileName
    If File$ = "" Then Exit Sub
    Open File$ For Output As #1
    For OuterLoop = 1 To 200
        Print #1, TotalMap(OuterLoop)
    Next
    Close
End Sub

Sub DoClear()
    PicCopy.Picture = PicChoose(3).Picture
    For SetGrass = 0 To 99
        PicGrid(SetGrass).Picture = PicCopy.Picture
    Next
    For OuterLoop = 1 To 200
        For InnerLoop = 1 To 200
            Map(InnerLoop, OuterLoop) = "G"
        Next
    Next
    X = 0
    Y = 0
End Sub

Sub DrawMap()
    GridToChar
    PicTotal.Cls
    For OuterLoop = 1 To 200
        For InnerLoop = 1 To 200
            If Map(InnerLoop, OuterLoop) = "G" Or Map(InnerLoop, OuterLoop) = ";" Then
                SetPixelV PicTotal.hDC, (PicTotal.ScaleWidth / 200) * InnerLoop, (PicTotal.ScaleHeight / 200) * OuterLoop, RGB(0, 175, 0)
            ElseIf Map(InnerLoop, OuterLoop) = "W" Then
                SetPixelV PicTotal.hDC, (PicTotal.ScaleWidth / 200) * InnerLoop, (PicTotal.ScaleHeight / 200) * OuterLoop, RGB(0, 0, 175)
            ElseIf Map(InnerLoop, OuterLoop) = "R" Then
                SetPixelV PicTotal.hDC, (PicTotal.ScaleWidth / 200) * InnerLoop, (PicTotal.ScaleHeight / 200) * OuterLoop, RGB(192, 192, 192)
            ElseIf Map(InnerLoop, OuterLoop) = "P" Then
                SetPixelV PicTotal.hDC, (PicTotal.ScaleWidth / 200) * InnerLoop, (PicTotal.ScaleHeight / 200) * OuterLoop, RGB(80, 80, 80)
            ElseIf Map(InnerLoop, OuterLoop) = "D" Then
                SetPixelV PicTotal.hDC, (PicTotal.ScaleWidth / 200) * InnerLoop, (PicTotal.ScaleHeight / 200) * OuterLoop, RGB(143, 106, 69)
            ElseIf Map(InnerLoop, OuterLoop) = "S" Then
                SetPixelV PicTotal.hDC, (PicTotal.ScaleWidth / 200) * InnerLoop, (PicTotal.ScaleHeight / 200) * OuterLoop, RGB(238, 197, 74)
            End If
        Next
    Next
    PicTotal.Refresh
End Sub
