VERSION 5.00
Begin VB.Form FILESEARCH 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FILESEARCH"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List8 
      Height          =   1230
      Left            =   1785
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.FileListBox File8 
      Height          =   1260
      Left            =   3360
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.DirListBox Dir8 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   330
      Left            =   1260
      TabIndex        =   3
      Top             =   1470
      Width           =   1905
   End
End
Attribute VB_Name = "FILESEARCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
