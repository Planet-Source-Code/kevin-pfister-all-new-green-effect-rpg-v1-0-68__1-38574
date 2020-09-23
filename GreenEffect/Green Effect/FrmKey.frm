VERSION 5.00
Begin VB.Form FrmKey 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estate Agents"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Get Key Made"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton CmdMKey 
         Caption         =   "Make Key"
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Each key cost 200 Castras, Please enter the Key number, this will be given to you by the owner of the house"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4215
      End
   End
End
Attribute VB_Name = "FrmKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
