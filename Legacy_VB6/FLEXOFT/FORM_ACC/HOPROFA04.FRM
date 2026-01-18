VERSION 5.00
Object = "{CA8A9783-280D-11CF-A24D-444553540000}#1.3#0"; "pdf.ocx"
Begin VB.Form HOPROFA04 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Guia de Manufactura"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin PdfLib.Pdf Pdf1 
      Height          =   8330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11030
      _Version        =   327680
      _ExtentX        =   19456
      _ExtentY        =   14693
      _StockProps     =   0
      SRC             =   ""
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      Height          =   9000
      Left            =   11030
      ScaleHeight     =   8940
      ScaleWidth      =   1140
      TabIndex        =   1
      Top             =   0
      Width           =   1200
      Begin VB.CommandButton Command6 
         Height          =   540
         Left            =   105
         Picture         =   "HOPROFA04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   6510
         Width           =   645
      End
      Begin VB.CommandButton Command4 
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
         Left            =   105
         Picture         =   "HOPROFA04.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Ir al Final del Documento"
         Top             =   3780
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Left            =   105
         Picture         =   "HOPROFA04.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Página Siguiente"
         Top             =   3045
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Left            =   105
         Picture         =   "HOPROFA04.frx":0C7E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Página Anterior"
         Top             =   2310
         Width           =   650
      End
      Begin VB.CommandButton Command1 
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
         Left            =   105
         Picture         =   "HOPROFA04.frx":0F88
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Ir al Comienzo del Documento"
         Top             =   1575
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   90
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   3
         Top             =   7350
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command5 
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
         Left            =   105
         Picture         =   "HOPROFA04.frx":1292
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir"
         Top             =   105
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "HOPROFA04.frx":13DC
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   7710
         Width           =   1515
      End
   End
End
Attribute VB_Name = "HOPROFA04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Pdf1.gotoFirstPage
End Sub

Private Sub Command2_Click()
    Pdf1.gotoPreviousPage
End Sub

Private Sub Command3_Click()
    Pdf1.gotoNextPage
End Sub

Private Sub Command4_Click()
   Pdf1.gotoLastPage
End Sub

Private Sub Command5_Click()
   Unload Me
End Sub

Private Sub Command6_Click()
   Pdf1.printAll
End Sub

Private Sub Form_Resize()
   Picture1.Left = Width - 1000
   Picture1.Height = Height - 350
   Pdf1.Width = Width - 1000
   Pdf1.Height = Height - 350
End Sub

