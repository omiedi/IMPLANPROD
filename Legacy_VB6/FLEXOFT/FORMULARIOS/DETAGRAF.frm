VERSION 5.00
Begin VB.Form DETAGRAF 
   Caption         =   "Form1"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   225
      Left            =   1680
      TabIndex        =   2
      Top             =   3570
      Width           =   2115
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1590
      Left            =   3780
      TabIndex        =   1
      Top             =   1995
      Width           =   225
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      Height          =   1590
      Left            =   1680
      ScaleHeight     =   1530
      ScaleWidth      =   2055
      TabIndex        =   0
      Top             =   1995
      Visible         =   0   'False
      Width           =   2115
   End
End
Attribute VB_Name = "DETAGRAF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
