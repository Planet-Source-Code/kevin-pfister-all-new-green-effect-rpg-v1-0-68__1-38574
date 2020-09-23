VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tile Maker"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicNight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5280
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   10
      Top             =   2760
      Width           =   495
   End
   Begin VB.CheckBox ChkGrid 
      Caption         =   "Grid"
      Height          =   255
      Left            =   4320
      TabIndex        =   9
      Top             =   3360
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.PictureBox PicDay 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4440
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   7
      Top             =   2760
      Width           =   495
   End
   Begin MSComDlg.CommonDialog Cdialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox ChkDark 
      Caption         =   "Save Night Tile"
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton CmdColour 
      Caption         =   "Change Colour"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.PictureBox PicColour 
      Height          =   375
      Left            =   4320
      ScaleHeight     =   315
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton CmdOpen 
      Caption         =   "Open Tile"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save Tile"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox PicZoom 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4215
      ScaleWidth      =   4095
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "Actual Size"
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "*Note that the game only uses the first 30*30 pixels, the rest is for display and for the alpha blending of the tiles"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   5775
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As Byte
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

Dim TileGrid(0 To 33, 0 To 33) As Long

Private Sub ChkGrid_Click()
    Call Redraw
End Sub

Private Sub CmdColour_Click()
    Cdialog.ShowColor
    PicColour.BackColor = Cdialog.Color
End Sub

Private Sub CmdOpen_Click()
    Cdialog.ShowOpen
    FileName = Cdialog.FileName
    If FileName = "" Then Exit Sub
    PicDay.Picture = LoadPicture(FileName)
    For X = 0 To 33
        For Y = 0 To 33
            TileGrid(X, Y) = GetPixel(PicDay.hDC, X, Y)
        Next
    Next
    Call Redraw
End Sub

Private Sub CmdSave_Click()
    Cdialog.ShowSave
    FileName = Cdialog.FileName
    If FileName = "" Then Exit Sub
    Call SavePicture(PicDay.Image, FileName)
    If ChkDark.Value = 1 Then
        Cdialog.ShowSave
        FileName = Cdialog.FileName
        If FileName = "" Then Exit Sub
        Call SavePicture(PicNight.Image, FileName)
    End If
End Sub

Private Sub Form_Load()
    For X = 0 To 33
        For Y = 0 To 33
            TileGrid(X, Y) = vbWhite
        Next
    Next
    Call DrawGrid
End Sub

Sub DrawGrid()
    PicZoom.ForeColor = RGB(128, 128, 128)
    For X = 1 To 32
        PicZoom.Line (PicZoom.Width / 33 * X, 0)-(PicZoom.Width / 33 * X, PicZoom.Height)
    Next
    For Y = 1 To 32
        PicZoom.Line (0, PicZoom.Height / 33 * Y)-(PicZoom.Width, PicZoom.Height / 33 * Y)
    Next
    PicZoom.ForeColor = PicColour.BackColor
End Sub

Sub Redraw()
    Dim Red As Integer
    Dim Green As Integer
    Dim Blue As Integer
    PicZoom.Cls
    For X = 0 To 33
        For Y = 0 To 33
            PicZoom.Line (PicZoom.Width / 33 * X, PicZoom.Height / 33 * Y)-(PicZoom.Width / 33 * (X + 1), PicZoom.Height / 33 * (Y + 1)), TileGrid(X, Y), BF
        Next
    Next
    For X = 0 To 33
        For Y = 0 To 33
            Call SetPixelV(PicDay.hDC, X, Y, TileGrid(X, Y))
            GetRgb TileGrid(X, Y), Red, Green, Blue
            Call SetPixelV(PicNight.hDC, X, Y, RGB(Red / 100 * 75, Green / 100 * 75, Blue / 100 * 75))
        Next
    Next
    PicDay.Refresh
    PicNight.Refresh
    If ChkGrid = 1 Then
        Call DrawGrid
    End If
End Sub

Private Sub PicZoom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        X = Int(33 / PicZoom.Width * X)
        Y = Int(33 / PicZoom.Height * Y)
        TileGrid(X, Y) = PicColour.BackColor
        Call Redraw
    Else
        X = Int(33 / PicZoom.Width * X)
        Y = Int(33 / PicZoom.Height * Y)
        PicColour.BackColor = TileGrid(X, Y)
    End If
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
