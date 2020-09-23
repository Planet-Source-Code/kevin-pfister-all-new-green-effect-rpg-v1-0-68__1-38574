VERSION 5.00
Begin VB.Form FrmPreview 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Map Preview(Maximise for Larger View)"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   ScaleHeight     =   268
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   301
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FrmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    GETEdit.DoPreview
End Sub
