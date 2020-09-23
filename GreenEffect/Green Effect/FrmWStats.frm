VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmWStats 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Items"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Items"
      Height          =   1935
      Left            =   3240
      TabIndex        =   5
      Top             =   2160
      Width           =   3015
      Begin MSComctlLib.TreeView TVItems 
         Height          =   1575
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   2778
         _Version        =   393217
         Indentation     =   176
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWStats.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWStats.frx":0454
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Keys"
      Height          =   1935
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   3015
      Begin MSComctlLib.TreeView TVKey 
         Height          =   1575
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   2778
         _Version        =   393217
         Indentation     =   176
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Weapons"
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin MSComctlLib.TreeView TVWeap 
         Height          =   3615
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   6376
         _Version        =   393217
         Indentation     =   176
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "FrmWStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Green Effect RPG By Kevin Pfister
'~FrmWStats~

'Description:
'Shows the weapons Status and allows user to select a default weapon

Private Sub cmdOK_Click()
    FrmGreenEffect.TmrPlayer.Enabled = True
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmGreenEffect.TmrPlayer.Enabled = True
End Sub

Private Sub TVWeap_DblClick()
    If Mid(TVWeap.SelectedItem.Text, 1, 5) = "Name:" Then
        ask = MsgBox("Set weapon to default", vbYesNo)
        If ask = vbYes Then
            Number = Val(Mid(TVWeap.SelectedItem.Key, 8))
            PlayerWeapon = Number
        End If
    End If
End Sub
