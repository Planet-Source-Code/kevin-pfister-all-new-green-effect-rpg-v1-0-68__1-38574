VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PLS Editor"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   4980
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicBlack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3120
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   56
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicFire 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2520
      Picture         =   "FrmMain.frx":0442
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   55
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicBalloon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   7200
      Picture         =   "FrmMain.frx":08DC
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   54
      Top             =   6600
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox PicPerson 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   6720
      Picture         =   "FrmMain.frx":0A66
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   53
      Top             =   6600
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox PicSupportLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      Picture         =   "FrmMain.frx":0BF0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   52
      Top             =   8400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicBed 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5520
      Picture         =   "FrmMain.frx":1015
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   51
      Top             =   8400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicStepsRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      Picture         =   "FrmMain.frx":147E
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   50
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicSteps 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      Picture         =   "FrmMain.frx":18EA
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   49
      Top             =   7200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicStepsLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5520
      Picture         =   "FrmMain.frx":1C70
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   48
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicCarpetLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   2520
      Picture         =   "FrmMain.frx":20DB
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   47
      Top             =   8400
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicCarpetRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   3120
      Picture         =   "FrmMain.frx":2539
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   46
      Top             =   8400
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicBottomRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1920
      Picture         =   "FrmMain.frx":298E
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   45
      Top             =   8400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicCarpetBottom 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   720
      Picture         =   "FrmMain.frx":2E4E
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   44
      Top             =   8400
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicCarpet 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   120
      Picture         =   "FrmMain.frx":327E
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   43
      Top             =   8400
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicBottomLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   1320
      Picture         =   "FrmMain.frx":35DB
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   42
      Top             =   8400
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicCarpetTop 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3720
      Picture         =   "FrmMain.frx":3A9D
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   41
      Top             =   8400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicCarpetTopLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4320
      Picture         =   "FrmMain.frx":3ED5
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   40
      Top             =   8400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicCarpetTopRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4920
      Picture         =   "FrmMain.frx":438E
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   39
      Top             =   8400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicArmor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4320
      Picture         =   "FrmMain.frx":4841
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   38
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicInn 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      Picture         =   "FrmMain.frx":4CB1
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   37
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicCase 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5520
      Picture         =   "FrmMain.frx":50FE
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   36
      Top             =   7200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicBCase 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5520
      Picture         =   "FrmMain.frx":5507
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   35
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicWindow 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4920
      Picture         =   "FrmMain.frx":5970
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   34
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicStool 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4920
      Picture         =   "FrmMain.frx":5DC3
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   33
      Top             =   7200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicDoor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4320
      Picture         =   "FrmMain.frx":6229
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   32
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicBrick 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4320
      Picture         =   "FrmMain.frx":6610
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   31
      Top             =   7200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicDirt 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3720
      Picture         =   "FrmMain.frx":69DD
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   30
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicFlowers 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   3720
      Picture         =   "FrmMain.frx":6E0A
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   29
      Top             =   7200
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicBotLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   120
      Picture         =   "FrmMain.frx":7331
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   28
      Top             =   7200
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicStopleft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   720
      Picture         =   "FrmMain.frx":7865
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   27
      Top             =   7200
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicStopRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   1320
      Picture         =   "FrmMain.frx":7D8F
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   26
      Top             =   7800
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicBotRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   120
      Picture         =   "FrmMain.frx":82B5
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   25
      Top             =   7800
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicFenceLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   720
      Picture         =   "FrmMain.frx":87EE
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   24
      Top             =   7800
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicAcross 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   1320
      Picture         =   "FrmMain.frx":8D2F
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   23
      Top             =   7200
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicFenceRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1920
      Picture         =   "FrmMain.frx":9260
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   22
      Top             =   7200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicStopLeftUp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1920
      Picture         =   "FrmMain.frx":979B
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   21
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicChest 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   3120
      Picture         =   "FrmMain.frx":9CDA
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   20
      Top             =   7800
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicStopRightUp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   2520
      Picture         =   "FrmMain.frx":A1D2
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   19
      Top             =   7800
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicTopRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   3120
      Picture         =   "FrmMain.frx":A70E
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   18
      Top             =   7200
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicTopLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   2520
      Picture         =   "FrmMain.frx":AC3B
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   17
      Top             =   7200
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicWell 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1920
      Picture         =   "FrmMain.frx":B165
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   16
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicTree 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1920
      Picture         =   "FrmMain.frx":B6A8
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   15
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicWater 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   1320
      Picture         =   "FrmMain.frx":BBF7
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   14
      Top             =   6000
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicRock 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   720
      Picture         =   "FrmMain.frx":C06F
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   13
      Top             =   6600
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicSand 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   120
      Picture         =   "FrmMain.frx":C505
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   12
      Top             =   6600
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicPath 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   720
      Picture         =   "FrmMain.frx":CA0B
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   11
      Top             =   6000
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicGrass 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   120
      Picture         =   "FrmMain.frx":CE45
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   10
      Top             =   6000
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicRoof 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2520
      Picture         =   "FrmMain.frx":D371
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   9
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicLGrass 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3120
      Picture         =   "FrmMain.frx":D79C
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   8
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicField 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4920
      Picture         =   "FrmMain.frx":DD2B
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   7
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicMud 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   3720
      Picture         =   "FrmMain.frx":E1A5
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   6
      Top             =   6600
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicSupportRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1320
      Picture         =   "FrmMain.frx":E63A
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   5
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save File"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   5520
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   120
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdLoad 
      Caption         =   "load File"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   5040
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      Max             =   20
      Min             =   1
      TabIndex        =   2
      Top             =   4680
      Value           =   1
      Width           =   4575
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   4575
      Left            =   4680
      Max             =   20
      Min             =   1
      TabIndex        =   1
      Top             =   120
      Value           =   1
      Width           =   255
   End
   Begin VB.PictureBox PicBuffer 
      AutoRedraw      =   -1  'True
      Height          =   4575
      Left            =   120
      ScaleHeight     =   301
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   301
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Type PUPls
    Name As String
    Text As String
    Type As String
End Type

Dim PlsGrid(1 To 200, 1 To 200) As PUPls
Dim AOP As Integer

Dim Int_PassableLand(1 To 10, 1 To 10) As Integer
Dim CompressedMap(1 To 200) As String
Dim TotalMap(1 To 200) As String

'To add object click on the map, if that does not work move mouse around that area while
'still holding the button down

'Left button adds an object
'Right button removes an object

Private Sub CmdLoad_Click()
    CDialog.Filter = "People Location Files (*.PLS)|*.PLS"
    CDialog.ShowOpen
    File$ = CDialog.FileName
    If File$ = "" Then Exit Sub
    MapFile$ = Mid(File$, 1, Len(File$) - 3) & "MAP"
    Open MapFile$ For Input As #1
    For OuterLoop = 1 To 200
        Input #1, CompressedMap(OuterLoop)
    Next
    Close
    For OuterLoop = 1 To 200
        TotalMap(OuterLoop) = ""
        For InnerLoop = 1 To 200
            PlsGrid(InnerLoop, OuterLoop).Name = ""
            PlsGrid(InnerLoop, OuterLoop).Text = ""
            PlsGrid(InnerLoop, OuterLoop).Type = ""
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
    Open File$ For Input As #1
    Input #1, People$
    AOP = Val(People$)
    If AOP <> 0 Then
        Do
            Input #1, Name1$
            Input #1, XPos$
            Input #1, YPos$
            Input #1, Text$
            Input #1, Type1$
            X = Val(Mid(XPos$, 8))
            Y = Val(Mid(YPos$, 8))
            PlsGrid(X, Y).Name = Mid(Name1$, 8)
            PlsGrid(X, Y).Text = Mid(Text$, 8)
            PlsGrid(X, Y).Type = Mid(Type1$, 8)
        Loop Until EOF(1)
    End If
    Close
    Call Redraw
End Sub

Sub Redraw()
    For OuterLoop = 1 To 10
        For InnerLoop = 1 To 10
            Int_PassableLand(InnerLoop, OuterLoop) = 0    'Clears Int_PassableLand objects
            X = HScroll1.Value * 10 - 10 + InnerLoop
            Y = VScroll1.Value * 10 - 10 + OuterLoop
            Select Case Mid(TotalMap(Y), X, 1)
            Case "G"    'Grass
                TempDc = PicGrass.hDC
            Case "W"    'Water
                TempDc = PicWater.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 2
            Case "S"    'Sand
                TempDc = PicSand.hDC
            Case "R"    'Rock
                TempDc = PicRock.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case "T"    'Tree
                TempDc = PicTree.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case "E"    'Well
                TempDc = PicWell.hDC
            Case "C"    'A Chest
                TempDc = PicChest.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case "P"    'Path
                TempDc = PicPath.hDC
            Case "F"    'Flowers
                TempDc = PicFlowers.hDC
            Case "D"    'Dirt (Path)
                TempDc = PicDirt.hDC
            Case "A"    'Brick
                TempDc = PicBrick.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case "B"    'Door
                TempDc = PicDoor.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case "H"    'Stool
                TempDc = PicStool.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case "I"    'Window
                TempDc = PicWindow.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case "1"    'Fence Bottom Left
                TempDc = PicBotLeft.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case "2"    'Fence Bottom Right
                TempDc = PicBotRight.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case "3"    'Fence Left Stop
                TempDc = PicStopleft.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case "4"    'Fence Vertical Left
                TempDc = PicFenceLeft.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case "5"    'Fence Horizontal
                TempDc = PicAcross.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case "6"    'Fence right stop
                TempDc = PicStopRight.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case "7"    'Fence vertical Right
                TempDc = PicFenceRight.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case "8"    'Fence Stop Vertical Left
                TempDc = PicStopLeftUp.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case "9"    'Fence Top Left Corner
                TempDc = PicTopLeft.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case "0"    'Fence Stop Vertical Right
                TempDc = PicStopRightUp.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case "/"    'Fence Top Right Corner
                TempDc = PicTopRight.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case "L"    'Bookcase
                TempDc = PicBCase.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case "M"    'Case
                TempDc = PicCase.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case "N"    'Center Carpet
                TempDc = PicCarpet.hDC
            Case "O"    'Top Left Carpet
                TempDc = PicCarpetTopLeft.hDC
            Case "U"    'Top Right Carpet
                TempDc = PicCarpetTopRight.hDC
            Case "V"    'Right Carpet
                TempDc = PicCarpetRight.hDC
            Case "^"    'Bottom Right Carpet
                TempDc = PicBottomRight.hDC
            Case "X"    'Bottom Horizontal Carpet
                TempDc = PicCarpetBottom.hDC
            Case "Y"    'Bottom Left of carpet
                TempDc = PicBottomLeft.hDC
            Case "Z"    'Left of carpet
                TempDc = PicCarpetLeft.hDC
            Case "Q"    'Top Horizontal of carpet
                TempDc = PicCarpetTop.hDC
            Case "!"    'Left Side of Steps
                TempDc = PicStepsLeft.hDC
            Case "%"    'Center of Steps
                TempDc = PicSteps.hDC
            Case "£"    'Right side of steps
                TempDc = PicStepsRight.hDC
            Case "$"    'Bed
                TempDc = PicBed.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case "&"    'Building Support Left
                TempDc = PicSupportLeft.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case "*"    'Building Support Right
                TempDc = PicSupportRight.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case "}"    'Building Roof
                TempDc = PicRoof.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case ";"    'Long Grass
                TempDc = PicLGrass.hDC
            Case ":"    'Field
                TempDc = PicField.hDC
            Case "'"    'Mud
                TempDc = PicMud.hDC
            Case "["    'FirePlace
                TempDc = PicFire.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case "]"    'The Inn Sign
                TempDc = PicInn.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case "{"    'The Armoury sign
                TempDc = PicArmor.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 1
            Case "`"
                TempDc = PicBlack.hDC
                Int_PassableLand(InnerLoop, OuterLoop) = 2
            End Select
            'Once the information is gathered about its Hdc's and its Passablity
            'Draw it to the backbuffer
            Call BitBlt(PicBuffer.hDC, InnerLoop * 30 - 30, OuterLoop * 30 - 30, 30, 30, TempDc, 0, 0, vbSrcCopy)
            If PlsGrid(X, Y).Name <> "" Then
                If PlsGrid(X, Y).Type = "3" Or PlsGrid(X, Y).Type = "4" Then
                    Call BitBlt(PicBuffer.hDC, InnerLoop * 30 - 27, (OuterLoop - 1) * 30 - 26, 24, 22, PicBalloon.hDC, 0, 0, vbSrcCopy)
                Else
                    Call BitBlt(PicBuffer.hDC, InnerLoop * 30 - 27, (OuterLoop - 1) * 30 - 26, 24, 22, PicPerson.hDC, 0, 0, vbSrcCopy)
                End If
            End If
        Next
    Next
    PicBuffer.Refresh
End Sub

Private Sub CmdSave_Click()
    CDialog.Filter = "People Location Files (*.PLS)|*.PLS"
    CDialog.ShowSave
    File$ = CDialog.FileName
    If File$ = "" Then Exit Sub
    Open File$ For Output As #1
    Print #1, Mid(Str(AOP), 2)
    For OuterLoop = 1 To 200
        For InnerLoop = 1 To 200
            If PlsGrid(InnerLoop, OuterLoop).Name <> "" Then
                Print #1, "##Name " & PlsGrid(InnerLoop, OuterLoop).Name
                Print #1, "##XPos" & Str(InnerLoop)
                Print #1, "##YPos" & Str(OuterLoop)
                Print #1, "##Text " & PlsGrid(InnerLoop, OuterLoop).Text
                Print #1, "##Type " & PlsGrid(InnerLoop, OuterLoop).Type
            End If
        Next
    Next
    Close
End Sub

Private Sub HScroll1_Change()
    Call Redraw
End Sub

Private Sub PicBuffer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        X1 = HScroll1.Value * 10 - 9 + Int(10 / PicBuffer.ScaleWidth * X)
        Y1 = VScroll1.Value * 10 - 9 + Int(10 / PicBuffer.ScaleHeight * Y) + 1
        Ask = MsgBox("Add person to " & Mid(Str(X1), 2) & ":" & Mid(Str(Y1), 2) & "?", vbYesNo)
        If Ask = vbYes Then
            WasEmpty = PlsGrid(X1, Y1).Name
            PlsGrid(X1, Y1).Name = InputBox("Name of person ", "Name", PlsGrid(X1, Y1).Name)
            PlsGrid(X1, Y1).Text = InputBox("Text", "Description", PlsGrid(X1, Y1).Text)
            PlsGrid(X1, Y1).Type = InputBox("Type of person ", "Type", "1")
            If WasEmpty = "" Then
                AOP = AOP + 1
            End If
            Call Redraw
        End If
    ElseIf Button = 2 Then
        X1 = HScroll1.Value * 10 - 9 + Int(10 / PicBuffer.ScaleWidth * X)
        Y1 = VScroll1.Value * 10 - 9 + Int(10 / PicBuffer.ScaleHeight * Y) + 1
        Ask = MsgBox("Remove Person from " & Mid(Str(X1), 2) & ":" & Mid(Str(Y1), 2) & "?", vbYesNo)
        If Ask = vbYes Then
            If PlsGrid(X1, Y1).Name <> "" Then
                PlsGrid(X1, Y1).Name = ""
                PlsGrid(X1, Y1).Text = ""
                PlsGrid(X1, Y1).Type = ""
                AOP = AOP - 1
            End If
            Call Redraw
        End If
    End If
End Sub

Private Sub VScroll1_Change()
    Call Redraw
End Sub
