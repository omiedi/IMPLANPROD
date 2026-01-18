VERSION 5.00
Begin VB.Form MUESTRAPRIN 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.VScrollBar VScroll1 
      Height          =   8415
      Left            =   11700
      TabIndex        =   1
      Top             =   0
      Width           =   225
   End
   Begin VB.PictureBox Picture1 
      Height          =   18000
      Left            =   0
      ScaleHeight     =   17940
      ScaleWidth      =   11610
      TabIndex        =   0
      Top             =   1470
      Width           =   11670
      Begin VB.PictureBox Picture2 
         Height          =   1170
         Left            =   1365
         ScaleHeight     =   1110
         ScaleWidth      =   4470
         TabIndex        =   2
         Top             =   1890
         Width           =   4530
      End
   End
End
Attribute VB_Name = "MUESTRAPRIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub VScroll1_Change()
   Picture1.Top = (-1) * VScroll1.Value
End Sub
