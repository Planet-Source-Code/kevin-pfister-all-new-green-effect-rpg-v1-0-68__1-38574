VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmNewBody 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Draw a New Body"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   Icon            =   "FrmNewBody.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicMask 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3360
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   33
      Top             =   720
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Make Mask"
      Height          =   255
      Left            =   1920
      TabIndex        =   32
      Top             =   120
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3420
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   255
      Left            =   660
      TabIndex        =   31
      Top             =   300
      Width           =   1155
   End
   Begin VB.Frame Frame9 
      Caption         =   "Trouser Dark"
      Height          =   555
      Left            =   60
      TabIndex        =   18
      Top             =   4860
      Width           =   3135
      Begin VB.PictureBox Picture1 
         Height          =   315
         Index           =   5
         Left            =   60
         ScaleHeight     =   255
         ScaleWidth      =   1995
         TabIndex        =   30
         Top             =   180
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Browse..."
         Height          =   255
         Index           =   5
         Left            =   2160
         TabIndex        =   24
         Top             =   180
         Width           =   915
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Trouser Light"
      Height          =   555
      Left            =   60
      TabIndex        =   17
      Top             =   4260
      Width           =   3135
      Begin VB.PictureBox Picture1 
         Height          =   315
         Index           =   4
         Left            =   60
         ScaleHeight     =   255
         ScaleWidth      =   1995
         TabIndex        =   29
         Top             =   180
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Browse..."
         Height          =   255
         Index           =   4
         Left            =   2160
         TabIndex        =   23
         Top             =   180
         Width           =   915
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Shirt Dark"
      Height          =   555
      Left            =   60
      TabIndex        =   16
      Top             =   3660
      Width           =   3135
      Begin VB.PictureBox Picture1 
         Height          =   315
         Index           =   3
         Left            =   60
         ScaleHeight     =   255
         ScaleWidth      =   1995
         TabIndex        =   28
         Top             =   180
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Browse..."
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   22
         Top             =   180
         Width           =   915
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Shirt Light"
      Height          =   555
      Left            =   60
      TabIndex        =   15
      Top             =   3060
      Width           =   3135
      Begin VB.PictureBox Picture1 
         Height          =   315
         Index           =   2
         Left            =   60
         ScaleHeight     =   255
         ScaleWidth      =   1995
         TabIndex        =   27
         Top             =   180
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Browse..."
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   21
         Top             =   180
         Width           =   915
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Head Light"
      Height          =   555
      Left            =   60
      TabIndex        =   14
      Top             =   1860
      Width           =   3135
      Begin VB.PictureBox Picture1 
         Height          =   315
         Index           =   0
         Left            =   60
         ScaleHeight     =   255
         ScaleWidth      =   1995
         TabIndex        =   25
         Top             =   180
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Browse..."
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   19
         Top             =   180
         Width           =   915
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Head Dark"
      Height          =   555
      Left            =   60
      TabIndex        =   13
      Top             =   2460
      Width           =   3135
      Begin VB.PictureBox Picture1 
         Height          =   315
         Index           =   1
         Left            =   60
         ScaleHeight     =   255
         ScaleWidth      =   1995
         TabIndex        =   26
         Top             =   180
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Browse..."
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   20
         Top             =   180
         Width           =   915
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pictures"
      Height          =   855
      Left            =   60
      TabIndex        =   6
      Top             =   5640
      Width           =   3135
      Begin VB.PictureBox PicHead 
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   0
         Left            =   120
         Picture         =   "FrmNewBody.frx":0442
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   29
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox PicHead 
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   1
         Left            =   600
         Picture         =   "FrmNewBody.frx":07FA
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   29
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox PicHead 
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   2
         Left            =   1080
         Picture         =   "FrmNewBody.frx":0BC3
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   29
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox PicBody 
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   0
         Left            =   1560
         Picture         =   "FrmNewBody.frx":0F85
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   29
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox PicBody 
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   1
         Left            =   2040
         Picture         =   "FrmNewBody.frx":136A
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   29
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox PicBody 
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   2
         Left            =   2520
         Picture         =   "FrmNewBody.frx":1750
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   29
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Body Type"
      Height          =   555
      Left            =   60
      TabIndex        =   4
      Top             =   660
      Width           =   3135
      Begin MSComctlLib.Slider Slider2 
         Height          =   255
         Left            =   60
         TabIndex        =   5
         Top             =   240
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Max             =   2
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Head Type"
      Height          =   555
      Left            =   60
      TabIndex        =   2
      Top             =   1260
      Width           =   3135
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   240
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Max             =   2
      End
   End
   Begin VB.CommandButton CmdDraw 
      Caption         =   "Draw"
      Height          =   255
      Left            =   660
      TabIndex        =   1
      Top             =   60
      Width           =   1155
   End
   Begin VB.PictureBox PicCharacter 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   60
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   0
      Top             =   60
      Width           =   495
   End
End
Attribute VB_Name = "FrmNewBody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As Byte

Dim Red As Integer
Dim Green As Integer
Dim Blue As Integer
Dim Red1 As Integer
Dim Green1 As Integer
Dim Blue1 As Integer

Private Sub CmdDraw_Click()
    PicCharacter.Cls
    For OuterLoop = 1 To 33
        For innerLoop = 1 To 33
            Color1 = GetPixel(PicBody(Slider2.Value).hDC, innerLoop, OuterLoop)
            Color2 = GetPixel(PicHead(Slider1.Value).hDC, innerLoop, OuterLoop)
            GetRgb Color1, Red, Green, Blue
            GetRgb Color2, Red1, Green1, Blue1
            
            If Red = 0 And Green = 255 And Blue = 0 Then
                'Draw Nothing
            ElseIf Red = 0 And Green = 0 And Blue = 150 Then
                SetPixelV PicCharacter.hDC, innerLoop, OuterLoop, Picture1(3).BackColor
            ElseIf Red = 0 And Green = 0 And Blue = 100 Then
                SetPixelV PicCharacter.hDC, innerLoop, OuterLoop, Picture1(2).BackColor
            ElseIf Red = 104 And Green = 100 And Blue = 0 Then
                SetPixelV PicCharacter.hDC, innerLoop, OuterLoop, Picture1(5).BackColor
            ElseIf Red = 150 And Green = 150 And Blue = 0 Then
                SetPixelV PicCharacter.hDC, innerLoop, OuterLoop, Picture1(4).BackColor
            ElseIf Red <> 255 And Green <> 255 And Blue <> 255 Then
                SetPixelV PicCharacter.hDC, innerLoop, OuterLoop, RGB(Red, Green, Blue)
            End If
            
            If Red1 = 0 And Green1 = 255 And Blue1 = 0 Then
                'Draw Nothing
            ElseIf Red1 = 200 And Green1 = 0 And Blue1 = 0 Then
                SetPixelV PicCharacter.hDC, innerLoop, OuterLoop, Picture1(1).BackColor
            ElseIf Red1 = 150 And Green1 = 0 And Blue1 = 0 Then
                SetPixelV PicCharacter.hDC, innerLoop, OuterLoop, Picture1(0).BackColor
            Else
                SetPixelV PicCharacter.hDC, innerLoop, OuterLoop, RGB(Red1, Green1, Blue1)
            End If
        Next
    Next
End Sub

Sub GetRgb(ByVal Color As Long, ByRef Red As Integer, ByRef Green As Integer, ByRef Blue As Integer)
    Dim temp As Long
    
    temp = (Color And 255)
    Red = temp And 255
    
    temp = Int(Color / 256)
    Green = temp And 255
    
    temp = Int(Color / 65536)
    Blue = temp And 255
       
End Sub

Private Sub Command1_Click(Index As Integer)
    cd.ShowColor
    Picture1(Index).BackColor = cd.Color
End Sub

Private Sub Command2_Click()
    cd.ShowSave
    If cd.FileName <> "" Then
        Call SavePicture(PicCharacter.Image, cd.FileName & ".bmp")
        For X = 0 To 33
            For Y = 0 To 33
                Colour = GetPixel(PicCharacter.hDC, X, Y)
                If Colour <> vbWhite Then
                    SetPixelV PicCharacter.hDC, innerLoop, OuterLoop, Colour
                End If
            Next
        Next
        PicCharacter.Refresh
        Call SavePicture(PicCharacter.Image, cd.FileName & "Mask.bmp")
        PicMask.Refresh
    End If
End Sub

Private Sub Form_Load()
    Picture1(0).BackColor = RGB(152, 0, 0)
    Picture1(1).BackColor = RGB(200, 0, 0)
    Picture1(2).BackColor = RGB(0, 0, 152)
    Picture1(3).BackColor = RGB(0, 0, 200)
    Picture1(4).BackColor = RGB(152, 152, 0)
    Picture1(5).BackColor = RGB(104, 100, 0)
End Sub
