VERSION 5.00
Begin VB.Form FSCUECO07 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Interes Diario de Mora"
   ClientHeight    =   1545
   ClientLeft      =   5070
   ClientTop       =   4320
   ClientWidth     =   4035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      Height          =   5790
      Left            =   3165
      ScaleHeight     =   5730
      ScaleWidth      =   1845
      TabIndex        =   3
      Top             =   -1860
      Width           =   1905
      Begin VB.CommandButton Command3 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   90
         Picture         =   "FSCUECO07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1920
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   90
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   4
         Top             =   2760
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   5
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -315
         Picture         =   "FSCUECO07.frx":030A
         Top             =   3000
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   30
      TabIndex        =   0
      Top             =   120
      Width           =   3075
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   2040
         TabIndex        =   2
         Top             =   500
         Width           =   570
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2640
         TabIndex        =   7
         Top             =   560
         Width           =   405
      End
      Begin VB.Label Label1 
         Caption         =   "Interes por Mora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   560
         Width           =   1965
      End
   End
End
Attribute VB_Name = "FSCUECO07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
'Call GRACONTROL("FORIPRE", "LIVAVEN", Text1.TEXT)
End Sub

Private Sub Command4_Click()
'Call CIERRARCH("*.*")
   Hide
End Sub
