VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4920
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6150
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7435
      _Version        =   393216
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmOptions.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "SldSpeed"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "SldDay"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Graphics"
      TabPicture(1)   =   "frmOptions.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ChkNiceGrad"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "AI Players"
      TabPicture(2)   =   "frmOptions.frx":0044
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label4"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "SldAITurns"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "SldPath"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.CheckBox ChkNiceGrad 
         Caption         =   "Alpha Blend Tiles (Checked = Better graphics but slower performace)"
         Height          =   255
         Left            =   -74760
         TabIndex        =   14
         Top             =   600
         Width           =   5175
      End
      Begin MSComctlLib.Slider SldDay 
         Height          =   375
         Left            =   -74760
         TabIndex        =   10
         Top             =   1440
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   661
         _Version        =   393216
         Max             =   120
         SelStart        =   60
         TickFrequency   =   5
         Value           =   60
      End
      Begin MSComctlLib.Slider SldSpeed 
         Height          =   375
         Left            =   -74760
         TabIndex        =   11
         Top             =   480
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   661
         _Version        =   393216
         LargeChange     =   1
         Min             =   25
         Max             =   125
         SelStart        =   75
         TickFrequency   =   5
         Value           =   75
      End
      Begin MSComctlLib.Slider SldPath 
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   661
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin MSComctlLib.Slider SldAITurns 
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   661
         _Version        =   393216
         Min             =   1
         Max             =   5
         SelStart        =   1
         TickFrequency   =   5
         Value           =   1
      End
      Begin VB.Label Label4 
         Caption         =   "How many player turns is needed before AI moves 1 step"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1920
         Width           =   5415
      End
      Begin VB.Label Label1 
         Caption         =   "Adjust the scale to give a more realistic pathfinding ( Higher = More realistic but will take more time to reach it's target)"
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   5415
      End
      Begin VB.Label Label2 
         Caption         =   "Adjust the slider to determine the game speed (Lower = Faster, but needs faster computer to run)"
         Height          =   495
         Left            =   -74760
         TabIndex        =   13
         Top             =   960
         Width           =   5415
      End
      Begin VB.Label Label3 
         Caption         =   "Adjust the slider to determine the length of the days and nights, the lower the number the faster the days and nights will change"
         Height          =   495
         Left            =   -74760
         TabIndex        =   12
         Top             =   1920
         Width           =   5295
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   7
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   6
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2490
      TabIndex        =   0
      Top             =   4455
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdApply_Click()
    Call SaveSettings
End Sub

Private Sub cmdCancel_Click()
    FrmGreenEffect.TmrDay.Enabled = True
    FrmGreenEffect.TmrPlayer.Enabled = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call SaveSettings
    FrmGreenEffect.TmrDay.Enabled = True
    FrmGreenEffect.TmrPlayer.Enabled = True
    Unload Me
End Sub

Private Sub Form_Load()
    FrmGreenEffect.TmrDay.Enabled = False
    FrmGreenEffect.TmrPlayer.Enabled = False
    If FrmGreenEffect.NiceGradient = True Then
        ChkNiceGrad.Value = 1
    Else
        ChkNiceGrad.Value = 0
    End If
    SldPath.Value = FrmGreenEffect.Lng_Modifer
    SldSpeed.Value = FrmGreenEffect.TmrPlayer.Interval
    SldDay.Value = FrmGreenEffect.TmrDay.Interval / 1000
    SldAITurns.Value = FrmGreenEffect.AIDelay
End Sub

Sub SaveSettings()
    FrmGreenEffect.NiceGradient = CBool(ChkNiceGrad.Value)
    FrmGreenEffect.Lng_Modifer = SldPath.Value
    FrmGreenEffect.AIDelay = SldAITurns.Value
    FrmGreenEffect.TmrPlayer.Interval = SldSpeed.Value
    FrmGreenEffect.TmrDay.Interval = SldDay.Value * 1000
    FrmGreenEffect.RenderMap
End Sub
