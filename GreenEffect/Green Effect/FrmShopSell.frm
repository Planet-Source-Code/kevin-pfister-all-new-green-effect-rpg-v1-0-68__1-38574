VERSION 5.00
Begin VB.Form FrmShopSell 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   8520
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fmeweapon 
      BackColor       =   &H8000000A&
      Caption         =   "Weapon 1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1275
      Index           =   0
      Left            =   120
      TabIndex        =   40
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton cmdsell 
         Caption         =   "Sell Weapon"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label lblname 
         BackColor       =   &H8000000A&
         Caption         =   "name:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblprice 
         BackColor       =   &H8000000A&
         Caption         =   "price:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   42
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame fmeweapon 
      BackColor       =   &H8000000A&
      Caption         =   "Weapon 2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1275
      Index           =   1
      Left            =   1800
      TabIndex        =   36
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton cmdsell 
         Caption         =   "Sell Weapon"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   37
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label lblprice 
         BackColor       =   &H8000000A&
         Caption         =   "price:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   39
         Top             =   495
         Width           =   1335
      End
      Begin VB.Label lblname 
         BackColor       =   &H8000000A&
         Caption         =   "name:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fmeweapon 
      BackColor       =   &H8000000A&
      Caption         =   "Weapon 3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1275
      Index           =   2
      Left            =   3480
      TabIndex        =   32
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton cmdsell 
         Caption         =   "Sell Weapon"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   33
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label lblprice 
         BackColor       =   &H8000000A&
         Caption         =   "price:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblname 
         BackColor       =   &H8000000A&
         Caption         =   "name:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fmeweapon 
      BackColor       =   &H8000000A&
      Caption         =   "Weapon 4"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1275
      Index           =   3
      Left            =   5160
      TabIndex        =   28
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton cmdsell 
         Caption         =   "Sell Weapon"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   29
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label lblprice 
         BackColor       =   &H8000000A&
         Caption         =   "price:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblname 
         BackColor       =   &H8000000A&
         Caption         =   "name:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fmeweapon 
      BackColor       =   &H8000000A&
      Caption         =   "Weapon 5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1275
      Index           =   4
      Left            =   6840
      TabIndex        =   24
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton cmdsell 
         Caption         =   "Sell Weapon"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   25
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label lblprice 
         BackColor       =   &H8000000A&
         Caption         =   "price:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblname 
         BackColor       =   &H8000000A&
         Caption         =   "name:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fmeweapon 
      BackColor       =   &H8000000A&
      Caption         =   "Weapon 6"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1275
      Index           =   5
      Left            =   120
      TabIndex        =   20
      Top             =   1440
      Width           =   1575
      Begin VB.CommandButton cmdsell 
         Caption         =   "Sell Weapon"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   21
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label lblprice 
         BackColor       =   &H8000000A&
         Caption         =   "price:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblname 
         BackColor       =   &H8000000A&
         Caption         =   "name:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fmeweapon 
      BackColor       =   &H8000000A&
      Caption         =   "Weapon 7"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1275
      Index           =   6
      Left            =   1800
      TabIndex        =   16
      Top             =   1440
      Width           =   1575
      Begin VB.CommandButton cmdsell 
         Caption         =   "Sell Weapon"
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   17
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label lblprice 
         BackColor       =   &H8000000A&
         Caption         =   "price:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblname 
         BackColor       =   &H8000000A&
         Caption         =   "name:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fmeweapon 
      BackColor       =   &H8000000A&
      Caption         =   "Weapon 8"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1275
      Index           =   7
      Left            =   3480
      TabIndex        =   12
      Top             =   1440
      Width           =   1575
      Begin VB.CommandButton cmdsell 
         Caption         =   "Sell Weapon"
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   13
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label lblprice 
         BackColor       =   &H8000000A&
         Caption         =   "price:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblname 
         BackColor       =   &H8000000A&
         Caption         =   "name:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fmeweapon 
      BackColor       =   &H8000000A&
      Caption         =   "Weapon 9"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1275
      Index           =   8
      Left            =   5160
      TabIndex        =   8
      Top             =   1440
      Width           =   1575
      Begin VB.CommandButton cmdsell 
         Caption         =   "Sell Weapon"
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   9
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label lblprice 
         BackColor       =   &H8000000A&
         Caption         =   "price:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblname 
         BackColor       =   &H8000000A&
         Caption         =   "name:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fmeweapon 
      BackColor       =   &H8000000A&
      Caption         =   "Weapon 10"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1275
      Index           =   9
      Left            =   6840
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
      Begin VB.CommandButton cmdsell 
         Caption         =   "Sell Weapon"
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   5
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label lblprice 
         BackColor       =   &H8000000A&
         Caption         =   "price:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblname 
         BackColor       =   &H8000000A&
         Caption         =   "name:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame FmeSellA 
      BackColor       =   &H8000000A&
      Caption         =   "Armour"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1275
      Left            =   3480
      TabIndex        =   0
      Top             =   2760
      Width           =   1575
      Begin VB.CommandButton CmdSellA 
         Caption         =   "Sell Armour"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label LblNameA 
         BackColor       =   &H8000000A&
         Caption         =   "Name:"
         DataSource      =   "&H80000004&"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblPriceA 
         BackColor       =   &H8000000A&
         Caption         =   "Price:"
         DataSource      =   "&H80000004&"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmShopSell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Green Effect RPG By Kevin Pfister
'~FrmShopSell~

Private Sub cmdsell_Click(Index As Integer)
    Call SellW(Index + 1)
End Sub

Private Sub CmdSellA_Click()
    Call SellA
End Sub

