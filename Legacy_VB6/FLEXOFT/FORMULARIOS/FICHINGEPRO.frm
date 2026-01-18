VERSION 5.00
Begin VB.Form FICHINGEPRO 
   BackColor       =   &H00E0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficha de Ingenieria de Producto"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   11910
   Begin VB.CommandButton Command13 
      Caption         =   "P.I.E."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   10920
      Picture         =   "FICHINGEPRO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "Edición de Plan de Inspección y Ensayo"
      Top             =   6195
      Width           =   855
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      Height          =   1590
      Left            =   2100
      ScaleHeight     =   1530
      ScaleWidth      =   2055
      TabIndex        =   38
      Top             =   6405
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   2
      Left            =   105
      MaxLength       =   128
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   5010
      Width           =   6735
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   1
      Left            =   105
      MaxLength       =   128
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   4240
      Width           =   6735
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   0
      Left            =   105
      MaxLength       =   128
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   3465
      Width           =   6735
   End
   Begin VB.CommandButton Command7 
      Height          =   260
      Left            =   105
      TabIndex        =   53
      Top             =   8295
      Width           =   1780
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Setup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10920
      Picture         =   "FICHINGEPRO.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2625
      Width           =   855
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   225
      Left            =   2100
      TabIndex        =   52
      Top             =   7245
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1170
      Left            =   2730
      TabIndex        =   51
      Top             =   6615
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0FFFF&
      Caption         =   "Fotografia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2650
      Index           =   2
      Left            =   7140
      TabIndex        =   49
      Top             =   5895
      Width           =   3630
      Begin VB.FileListBox File1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1845
         Index           =   2
         Left            =   315
         TabIndex        =   25
         Top             =   630
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.CommandButton Command3 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3255
         TabIndex        =   24
         Top             =   315
         Width           =   175
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   1590
         Index           =   2
         Left            =   735
         ScaleHeight     =   1530
         ScaleWidth      =   2265
         TabIndex        =   50
         Top             =   840
         Width           =   2325
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   315
         TabIndex        =   23
         Text            =   " "
         Top             =   315
         Width           =   3060
      End
      Begin VB.Image Image1 
         Height          =   19995
         Index           =   2
         Left            =   0
         Stretch         =   -1  'True
         Top             =   -1260
         Visible         =   0   'False
         Width           =   19995
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0FFFF&
      Caption         =   "Croquis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2650
      Index           =   1
      Left            =   7140
      TabIndex        =   47
      Top             =   3150
      Width           =   3630
      Begin VB.FileListBox File1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1845
         Index           =   1
         Left            =   315
         TabIndex        =   22
         Top             =   630
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.CommandButton Command3 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3255
         TabIndex        =   21
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   315
         TabIndex        =   20
         Text            =   " "
         Top             =   315
         Width           =   3060
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   1590
         Index           =   1
         Left            =   735
         ScaleHeight     =   1530
         ScaleWidth      =   2265
         TabIndex        =   48
         Top             =   840
         Width           =   2325
      End
      Begin VB.Image Image1 
         Height          =   20000
         Index           =   1
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   20000
      End
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3885
      TabIndex        =   4
      Top             =   1125
      Width           =   600
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3885
      TabIndex        =   2
      Top             =   735
      Width           =   600
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   960
      Left            =   1890
      TabIndex        =   43
      Top             =   2205
      Width           =   4950
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H80000009&
      Height          =   1240
      Left            =   4830
      ScaleHeight     =   1185
      ScaleWidth      =   1950
      TabIndex        =   41
      Top             =   210
      Width           =   2010
      Begin VB.TextBox ALIAS 
         BorderStyle     =   0  'None
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
         Index           =   3
         Left            =   0
         TabIndex        =   8
         Top             =   930
         Width           =   1905
      End
      Begin VB.TextBox ALIAS 
         BorderStyle     =   0  'None
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
         Index           =   2
         Left            =   0
         TabIndex        =   7
         Top             =   620
         Width           =   1905
      End
      Begin VB.TextBox ALIAS 
         BorderStyle     =   0  'None
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
         Index           =   1
         Left            =   0
         TabIndex        =   6
         Top             =   310
         Width           =   1905
      End
      Begin VB.TextBox ALIAS 
         BorderStyle     =   0  'None
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
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1905
      End
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1890
      TabIndex        =   3
      Top             =   1125
      Width           =   1905
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1890
      TabIndex        =   1
      Top             =   735
      Width           =   1905
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Selecc Todo"
      Height          =   260
      Left            =   5535
      TabIndex        =   16
      Top             =   8295
      Width           =   1275
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Pegar"
      Height          =   260
      Left            =   4245
      TabIndex        =   15
      Top             =   8295
      Width           =   1275
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Copiar"
      Height          =   260
      Left            =   3075
      TabIndex        =   14
      Top             =   8295
      Width           =   1170
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cortar"
      Height          =   260
      Left            =   1890
      TabIndex        =   13
      Top             =   8295
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0FFFF&
      Caption         =   "Plano"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2950
      Index           =   0
      Left            =   7140
      TabIndex        =   36
      Top             =   105
      Width           =   3630
      Begin VB.FileListBox File1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2040
         Index           =   0
         Left            =   315
         TabIndex        =   19
         Top             =   735
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.CommandButton Command3 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3255
         TabIndex        =   18
         Top             =   315
         Width           =   175
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   1590
         Index           =   0
         Left            =   735
         ScaleHeight     =   1530
         ScaleWidth      =   2265
         TabIndex        =   37
         Top             =   1155
         Width           =   2325
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   315
         TabIndex        =   17
         Text            =   " "
         Top             =   315
         Width           =   3060
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1930
         TabIndex        =   59
         Text            =   " "
         Top             =   735
         Width           =   1170
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimiento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   525
         TabIndex        =   58
         Top             =   765
         Width           =   1275
      End
      Begin VB.Image Image1 
         Height          =   10005
         Index           =   0
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   10005
      End
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   5800
      Width           =   6735
   End
   Begin VB.CommandButton Command21 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4305
      TabIndex        =   34
      Top             =   210
      Width           =   175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   10920
      Picture         =   "FICHINGEPRO.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Imprime Ficha de Ingenieria"
      Top             =   7035
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1890
      TabIndex        =   0
      Top             =   210
      Width           =   2430
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   540
      Left            =   1890
      MultiLine       =   -1  'True
      TabIndex        =   31
      Top             =   1575
      Width           =   4950
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   10920
      Picture         =   "FICHINGEPRO.frx":0DB6
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1785
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   10920
      Picture         =   "FICHINGEPRO.frx":10C0
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Sale de la Aplicación"
      Top             =   105
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   10920
      Picture         =   "FICHINGEPRO.frx":120A
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Graba las Modificaciones"
      Top             =   945
      Width           =   855
   End
   Begin VB.Line Line5 
      X1              =   1890
      X2              =   6825
      Y1              =   2120
      Y2              =   2120
   End
   Begin VB.Line Line4 
      X1              =   1890
      X2              =   6825
      Y1              =   3170
      Y2              =   3170
   End
   Begin VB.Line Line3 
      X1              =   6840
      X2              =   6840
      Y1              =   1575
      Y2              =   2100
   End
   Begin VB.Line Line2 
      X1              =   6840
      X2              =   6840
      Y1              =   3150
      Y2              =   2205
   End
   Begin VB.Line Line1 
      X1              =   6840
      X2              =   6840
      Y1              =   3230
      Y2              =   8550
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Referencia:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   105
      TabIndex        =   56
      Top             =   4770
      Width           =   6735
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   10360
      Picture         =   "FICHINGEPRO.frx":1514
      Top             =   8020
      Width           =   2010
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Alcance:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   1
      Left            =   105
      TabIndex        =   55
      Top             =   3990
      Width           =   6735
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Objeto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   0
      Left            =   105
      TabIndex        =   54
      Top             =   3225
      Width           =   6735
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3780
      TabIndex        =   46
      Top             =   1065
      Width           =   120
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3780
      TabIndex        =   45
      Top             =   690
      Width           =   120
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operaciones:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   44
      Top             =   2205
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "A L I A S"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   4620
      TabIndex        =   42
      Top             =   300
      Width           =   225
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nro de HE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   315
      TabIndex        =   40
      Top             =   1155
      Width           =   1485
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nro de Plano:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   315
      TabIndex        =   39
      Top             =   735
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Especificación:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   105
      TabIndex        =   35
      Top             =   5565
      Width           =   6735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   525
      TabIndex        =   33
      Top             =   315
      Width           =   1275
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion Corta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   315
      TabIndex        =   32
      Top             =   1575
      Width           =   1485
   End
End
Attribute VB_Name = "FICHINGEPRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CI%, CABAS, CODE$, CIE%
Dim HIDOCLICK As Integer
Dim TOPEINI(3), IZQUINI(3), ANCHINI(3), ALTUINI(3)
Dim IMACTI

Private Sub ALIAS_Change(Index As Integer)
   MODIFI% = 1
End Sub

'
Private Sub Command1_Click()
    Call GRABALIAS
    Call GRAESPECI
    Call GRABAPLANOS
    MODIFI% = 0
End Sub

Private Sub Command11_Click()
   COMIE = Text3.SelStart
   LOSEL = Text3.SelLength
   Text3.TEXT = Left$(Text3.TEXT, COMIE) + Clipboard.GetText + Mid$(Text3.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command12_Click()
   Text3.SelStart = 0
   Text3.SelLength = Len(Text3.TEXT)
   Text3.Refresh
End Sub

Private Sub Command13_Click()
   CIXI% = CODINT%(Text1.TEXT)
   If CIXI% > 0 Then
      If CIXI% <= NUMAS% Then
         PLAINSPEN.CONTROCAM.Caption = Timer
         PLAINSPEN.Show
      End If
   End If
End Sub


Private Sub Command26_Click()
   Command26.Enabled = False
   OPGRAFIC.Show 1
   Command26.Enabled = True
End Sub

Private Sub Command3_Click(Index As Integer)
   RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
   If RUTARGRA$ = "" Then RUTARGRA$ = LUDAT$
   File1(Index).Path = RUTARGRA$
   '
   TIARGRA$ = "*." + TRIM$(CONTROL$("", "TIARGRA"))
   If TIARGRA$ = "*." Then TIARGRA$ = "*.*"
   If TIARGRA$ = "*.*.*" Then TIARGRA$ = "*.*"
   File1(Index).Pattern = TIARGRA$
   File1(Index).Visible = True
End Sub

Private Sub Command4_Click()
   COMIE = Text3.SelStart
   LOSEL = Text3.SelLength
   Text3.TEXT = Left$(Text3.TEXT, COMIE) + Mid$(Text3.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command5_Click()
   CIIX% = CODINT%(Text1.TEXT)
   If Picture4.Visible = True Then
       Call PRIPLANO(CIIX%)
     Else
       Call PRIHOJESP(CIIX%)
   End If
End Sub
'

Private Sub Command8_Click()
   Clipboard.SetText Mid$(Text3.TEXT, Text3.SelStart + 1, Text3.SelLength)
End Sub

Private Sub File1_DblClick(Index As Integer)
   Text4(Index).TEXT = File1(Index).filename
   File1(Index).Visible = False
   Call MUESTRAPLANO(Index)
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    COMMA$ = Command$
    PPX% = InStr(COMMA$, "#" + JOBID$ + "#")
    If PPX% > 0 Then
        COMMA$ = Left$(COMMA$, PPX% - 1)
    End If
    If COMMA$ <> "" Then
        Text1.TEXT = COMMA$
    End If
    '
    For KK% = 0 To 2
      TOPEINI(KK%) = Picture1(KK%).Top
      IZQUINI(KK%) = Picture1(KK%).Left
      ANCHINI(KK%) = Picture1(KK%).Width
      ALTUINI(KK%) = Picture1(KK%).Height
    Next KK%
    Show
End Sub


Private Sub Form_Unload(Cancel As Integer)
    INI& = 1 + ULTREG&("ESTRUDOS")
    FIN& = ULTREG&("ESTRUNUE")
    For U& = INI& To FIN&
        RB$ = String$(256, 0)
        Mid$(RB$, 251, 4) = MKS$(0.0001)
        Call GRAREG("ESTRUDOS", RB$, U&)
    Next U&
    '
    Cancel = 0
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub HScroll1_Change()
     HCUADRO = Picture4.Width
     VCUADRO = Picture4.Height
     Picture4.PaintPicture Image1(IMACTI).Picture, (-1) * HScroll1.Value, (-1) * VScroll1.Value, 1 * HCUADRO, 1 * VCUADRO
End Sub

Private Sub List1_dblClick()
   List1.Enabled = False
   HOPROFABRI.Label7.Caption = Text1.TEXT
   HOPROFABRI.Label9.Caption = Text2.TEXT
   HOPROFABRI.Label8.Caption = TRIM$(Left$(List1.TEXT, 6))
   HOPROFABRI.Label10.Caption = TRIM$(Mid$(List1.TEXT, 7))
   HOPROFABRI.CONTROCAM.Caption = Timer
   HOPROFABRI.Show
   List1.Enabled = True
End Sub

Private Sub Picture1_Click(Index As Integer)
      
      Screen.MousePointer = 11
      '
      Picture4.Top = 500
      Picture4.Left = 500
      Picture4.Width = 10000
      Picture4.Height = 7000
      '
      'HIMAGEN = Picture1(Index).Picture.Width
      'VIMAGEN = Picture1(Index).Picture.Height
      HIMAGEN = Image1(Index).Picture.Width
      VIMAGEN = Image1(Index).Picture.Height
      If HIMAGEN < 1 Or VIMAGEN < 1 Then
         GoTo 99
      End If
      '
      HCUADRO = Picture4.Width
      VCUADRO = Picture4.Height
      '
      REDUH = HIMAGEN / HCUADRO
      REDUV = VIMAGEN / VCUADRO
      REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
      Picture4.Height = VCUADRO * REDUV / REDUC
      Picture4.Width = HCUADRO * REDUH / REDUC
      HCUADRO = Picture4.Width
      VCUADRO = Picture4.Height
      Picture4.Left = (12000 - HCUADRO) / 2
      '
      VScroll1.Top = Picture4.Top
      VScroll1.Left = Picture4.Left + Picture4.Width
      VScroll1.Height = Picture4.Height
      HScroll1.Top = Picture4.Top + Picture4.Height
      HScroll1.Width = Picture4.Width
      HScroll1.Left = Picture4.Left
      HScroll1.Max = HScroll1.Min   'HCUADRO / 2
      VScroll1.Max = VScroll1.Min   'VCUADRO / 2
      HScroll1.Value = 0
      VScroll1.Value = 0
      
      Picture4.PaintPicture Image1(Index).Picture, 0, 0, 1 * HCUADRO, 1 * VCUADRO
      Picture4.Visible = True
      HScroll1.Visible = True
      VScroll1.Visible = True
      IMACTI = Index
      '
99    Screen.MousePointer = 1
      '
End Sub

Private Sub Picture4_Click()
   Picture4.Visible = False
   HScroll1.Visible = False
   VScroll1.Visible = False
End Sub

Private Sub Text1_Change()
    If MODIFI% > 0 Then
        RESPU = MsgBox("Abandona sin Guardar los Cambios ?", 4371, "Operacion Critica")
        If RESPU <> 6 Then
            Command1.SetFocus
            Exit Sub
        End If
        MODIFI% = 0
    End If
    '
    MODIFI% = 0
    Text2.TEXT = ""
    Text3.TEXT = ""
    For KK% = 0 To 2
       Text4(KK%).TEXT = ""
    Next KK%
    MODIFI% = 0
End Sub

Private Sub Text1_DblClick()
    If MODIFI% > 0 Then
        RESPU = MsgBox("Abandona sin Guardar los Cambios ?", 4371, "Operacion Critica")
        If RESPU <> 6 Then
            Command1.SetFocus
            Exit Sub
        End If
        MODIFI% = 0
    End If
    '
    Text1.TEXT = ""
    Call SELECHO("MASTER")
End Sub
'
Private Sub Command21_Click()
    If MODIFI% > 0 Then
        RESPU = MsgBox("Abandona sin Guardar los Cambios ?", 4371, "Operacion Critica")
        If RESPU <> 6 Then
            Command1.SetFocus
            Exit Sub
        End If
        MODIFI% = 0
    End If
    '
    Text1.TEXT = ""
    Call SELECHO("MASTER")
    Text1.SetFocus
End Sub
'

Private Sub Text1_GotFocus()
   '
   If MODIFI% > 0 Then
       RESPU = MsgBox("Abandona sin Guardar los Cambios ?", 4371, "Operacion Critica")
       If RESPU <> 6 Then
           Command1.SetFocus
           Exit Sub
       End If
       MODIFI% = 0
   End If
   '
   Screen.MousePointer = 11
   Refresh
   If Text1.TEXT = "" Then
      Text1.TEXT = VALACT1$("MASTER")
   End If
   '
   CODX$ = TRIM$(Text1.TEXT)
   cix% = CODINT%(CODX$)
   If cix% > 0 Then
        COD$ = CODX$: CI% = cix%
        Text1.TEXT = CODX$
        Text1.Refresh
        Text2.TEXT = DESCRIT$(cix%)
        Text2.Refresh
        Call CARGALIAS
        Call CARHORUTA
        Call CARESPECI
        Call CARGAPLANOS
        MODIFI% = 0
      Else
        Text1.TEXT = COD$
   End If
   '
   Screen.MousePointer = 1
   '
End Sub
'

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CODX$ = TRIM$(Text1.TEXT)
        cix% = CODINT%(CODX$)
        If cix% > 0 Then
            COD$ = CODX$: CI% = cix%
            Text1.TEXT = CODX$
            COD$ = CODX$: CI% = cix%
            Text2.TEXT = DESCRIT$(CI%)
          Else
            Call MENSERR(24, "Código no Válido")
            Text1.TEXT = COD$
        End If
    End If
End Sub
'

Private Sub Text10_Change()
   MODIFI% = 1
End Sub

Private Sub Text2_Change()
   MODIFI% = 1
End Sub

Private Sub Text2_GotFocus()
    Text3.SetFocus
End Sub
'


Private Sub Command2_Click()
    Call HELPONLINE("FORMULAS")
End Sub
'

Private Sub Command9_Click()
    If MODIFI% > 0 Then
        RESPU = MsgBox("Abandona sin Guardar los Cambios ?", 4371, "Operacion Critica")
        If RESPU <> 6 Then
            Command1.SetFocus
            Exit Sub
        End If
        MODIFI% = 0
    End If
    '
    Call CIERRARCH("*.*")
    Call FINAL("")
    '
End Sub

Sub CARGALIAS()
   '
   COD$ = Text1.TEXT
   CI% = CODINT%(Text1.TEXT)
   For KK% = 0 To 3
      ALIAS(KK%).TEXT = ""
   Next KK%
   For KK% = 0 To 2
      Text5(KK%).TEXT = ""
   Next KK%
   '
   If CI% > 0 Then
      If CI% <= NUMAS% Then
         Screen.MousePointer = 11
         RNC$ = RECFILT$("ENCASPEC", 1, MKI$(CI%))
         URN& = ULTREG&("ENCASPEC")
         '
         If Len(RNC$) >= 4 Then
            RENOTA& = CVS(Left$(RNC$, 4))
            If RENOTA& > 0 Then
               If RENOTA& <= URN& Then
                  TTYY$ = REGLEIDO$("ENCASPEC", RENOTA&)
                  ALITOT$ = Mid$(TTYY$, 3, 64)
                  For KK% = 0 To 3
                     ALIAS(KK%).TEXT = TRIM$(Mid$(ALITOT$, 16 * KK% + 1, 16))
                     ALIAS(KK%).Refresh
                  Next KK%
                  For KK% = 0 To 2
                     Text5(KK%) = TRIM$(Mid$(TTYY$, 128 * KK% + 129, 128))
                     Text5(KK%).Refresh
                  Next KK%
               End If
            End If
         End If
      End If
   End If
   '
End Sub

Sub GRABALIAS()
   '
   COD$ = Text1.TEXT
   CI% = CODINT%(Text1.TEXT)
   If CI% > 0 Then
      If CI% <= NUMAS% Then
         Screen.MousePointer = 11
         URN& = ULTREG&("ENCASPEC")
         RNC$ = RECFILT$("ENCASPEC", 1, MKI$(CI%))
         '
         ALITOT$ = ""
         For KK% = 0 To 3
            ALITOT$ = ALITOT$ + AJUSTI$(ALIAS(KK%).TEXT, 16)
         Next KK%
         '
         ENCASPE$ = ""
         For KK% = 0 To 2
            ENCASPE$ = ENCASPE$ + AJUSTI$(Text5(KK%).TEXT, 128)
         Next KK%
         '
         RENOTA& = 0
         If Len(RNC$) >= 4 Then
            RENOTA& = CVS(Left$(RNC$, 4))
         End If
         '
         If RENOTA& > 0 Then
            If RENOTA& <= URN& Then
               TTXX$ = REGLEIDO$("ENCASPEC", RENOTA&)
               GoTo 10
            End If
         End If
         URN& = URN& + 1
         RENOTA& = URN&
         TTXX$ = REGBLAN$("ENCASPEC")
         Call REPLA(TTXX$, MKI$(CI%), 1, 2)
         '
10       Call REPLA(TTXX$, ALITOT$, 3, 64)
         Call REPLA(TTXX$, ENCASPE$, 129, 384)
         Call GRAREG("ENCASPEC", TTXX$, RENOTA&)
         Call CIERRARCH("ENCASPEC")
      End If
   End If
End Sub
'
Sub CARESPECI()
   '
   COD$ = Text1.TEXT
   CI% = CODINT%(Text1.TEXT)
   Text3.TEXT = ""
   If CI% > 0 Then
      If CI% <= NUMAS% Then
         Screen.MousePointer = 11
         RNC$ = RECFILT$("ESPECIFI", 1, MKI$(CI%))
         URN& = ULTREG&("ESPECIFI")
         '
         For U& = 1 To Len(RNC$) Step 4
            RENOTA& = CVS(Mid$(RNC$, U&, 4))
            If RENOTA& > 0 Then
               If RENOTA& <= URN& Then
                  TTYY$ = RTrim$(Mid$(REGLEIDO$("ESPECIFI", RENOTA&), 3))
                  If Text3.TEXT <> "" Then
                     Text3.TEXT = Text3.TEXT + Chr$(13) + Chr$(10)
                  End If
                  Text3.TEXT = Text3.TEXT + RTrim$(TTYY$)
               End If
            End If
         Next U&
         Screen.MousePointer = 1
      End If
   End If
   Text3.Refresh
   '
End Sub
'
Sub GRAESPECI()
   '
   COD$ = Text1.TEXT
   CI% = CODINT%(Text1.TEXT)
   If CI% > 0 Then
      If CI% <= NUMAS% Then
         Screen.MousePointer = 11
         URN& = ULTREG&("ESPECIFI")
         RNC$ = RECFILT$("ESPECIFI", 1, MKI$(CI%))
         KK& = (-3)
         '
         TEXTO$ = Text3.TEXT + Chr$(13) + Chr$(10)
         Call FRATEXTO(TEXTO$, 62)
         For LL% = 1 To CARETEX%
            TTYY$ = MKI$(CI%) + RETEX$(LL%)
            KK& = KK& + 4
            If KK& <= Len(RNC$) - 3 Then
                RENOTA& = CVS(Mid$(RNC$, KK&, 4))
              Else
                URN& = URN& + 1
                RENOTA& = URN&
            End If
            Call GRAREG("ESPECIFI", TTYY$, RENOTA&)
         Next LL%
         '
20       While KK& < Len(RNC$) - 3
            KK& = KK& + 4
            RENOTA& = CVS(Mid$(RNC$, KK&, 4))
            Call GRAREG("ESPECIFI", MKI$(0) + Space$(62), RENOTA&)
         Wend
         '
         Call CIERRARCH("ESPECIFI")
         Screen.MousePointer = 1
      End If
   End If

End Sub
'
Sub CARHORUTA()
   '
   COD$ = Text1.TEXT
   CI% = CODINT%(Text1.TEXT)
   List1.Clear
   If CI% > 0 Then
      If CI% <= NUMAS% Then
         Screen.MousePointer = 11
         RNC$ = RECFILT$("SECOPER", 1, MKI$(CI%))
         URN& = ULTREG&("SECOPER")
         '
         For U& = 1 To Len(RNC$) Step 4
            RENOTA& = CVS(Mid$(RNC$, U&, 4))
            If RENOTA& > 0 Then
               If RENOTA& <= URN& Then
                  TTYY$ = REGLEIDO$("SECOPER", RENOTA&)
                  TTXX$ = CSTRING$(Mid$(TTYY$, 3, 2), 1, 5, 0, 2) + "   " + Mid$(TTYY$, 5, 48)
                  List1.AddItem TTXX$
               End If
            End If
         Next U&
         Screen.MousePointer = 1
      End If
   End If
   Text3.Refresh
   '
End Sub
'
Sub CARGAPLANOS()
   '
   Text6.TEXT = ""
   Text7.TEXT = ""
   Text9.TEXT = ""
   Text8.TEXT = ""
   Text10.TEXT = ""
   For KK% = 0 To 2
      Text4(KK%).TEXT = ""
   Next KK%
   '
   COD$ = Text1.TEXT
   CI% = CODINT%(Text1.TEXT)
   If CI% > 0 Then
      If CI% <= NUMAS% Then
         Screen.MousePointer = 11
         RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CI%))
         If Len(RNC$) >= 4 Then
            REPLANO& = CVS(RNC$)
            If REPLANO& > 0 Then
               If REPLANO& <= ULTREG&("INDIPLAN") Then
                  RPLA$ = REGLEIDO$("INDIPLAN", REPLANO&)
                  Text7.TEXT = TRIM$(Mid$(RPLA$, 3, 16))
                  Text9.TEXT = TRIM$(Mid$(RPLA$, 19, 4))
                  Text8.TEXT = TRIM$(Mid$(RPLA$, 33, 16))
                  Text10.TEXT = TRIM$(Mid$(RPLA$, 49, 4))
                  Text6.TEXT = TRIM$(Mid$(RPLA$, 53, 12))
                  Text4(0).TEXT = TRIM$(Mid$(RPLA$, 65, 48))
                  Text4(1).TEXT = TRIM$(Mid$(RPLA$, 113, 48))
                  Text4(2).TEXT = TRIM$(Mid$(RPLA$, 161, 48))
               End If
            End If
         End If
         For KK% = 0 To 2
            Call MUESTRAPLANO(KK%)
         Next KK%
         Screen.MousePointer = 1
      End If
   End If
   '
End Sub
'
Sub GRABAPLANOS()
   '
   COD$ = Text1.TEXT
   CI% = CODINT%(Text1.TEXT)
   If CI% > 0 Then
      If CI% <= NUMAS% Then
         Screen.MousePointer = 11
         RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CI%))
         REPLANO& = 0
         If Len(RNC$) >= 4 Then REPLANO& = CVS(RNC$)
         '
         If REPLANO& < 1 Or REPLANO > ULTREG&("INDIPLAN") Then
            REPLANO& = 1 + ULTREG&("INDIPLAN")
         End If
         RPLA$ = REGBLAN$("INDIPLAN")
         Call REPLA(RPLA$, MKI$(CI%), 1, 2)
         Call REPLA(RPLA$, TRIM$(Text7.TEXT), 3, 16)
         Call REPLA(RPLA$, TRIM$(Text9.TEXT), 19, 4)
         Call REPLA(RPLA$, TRIM$(Text8.TEXT), 33, 16)
         Call REPLA(RPLA$, TRIM$(Text10.TEXT), 49, 4)
         Call REPLA(RPLA$, TRIM$(Text6.TEXT), 53, 12)
         Call REPLA(RPLA$, TRIM$(Text4(0).TEXT), 65, 48)
         Call REPLA(RPLA$, TRIM$(Text4(1).TEXT), 113, 48)
         Call REPLA(RPLA$, TRIM$(Text4(2).TEXT), 161, 48)
         Call GRAREG("INDIPLAN", RPLA$, REPLANO&)
         '
         Screen.MousePointer = 1
      End If
   End If
   '
End Sub

Private Sub Text3_Change()
   MODIFI% = 1
End Sub

Private Sub Text4_Change(Index As Integer)
   MODIFI% = 1
End Sub

Private Sub Text4_KeyPress(Index As Integer, KeyAscii As Integer)
   '
   If KeyAscii% = 13 Then
      Call MUESTRAPLANO(Index)
   End If
   '
End Sub
'
Sub MUESTRAPLANO(Index)
   '
   Screen.MousePointer = 11
   '
   RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
   If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
   If Right$(RUTARGRA$, 1) <> "\" Then
      RUTARGRA$ = RUTARGRA$ + "\"
   End If
   '
   Image1(Index).Picture = LoadPicture("")
   Picture1(Index).Cls
   COPLANO$ = RUTARGRA$ + TRIM$(Text4(Index).TEXT)
   If COPLANO$ <> "" Then
      If EXISTE%(COPLANO$) > 0 Then
         Picture1(Index).Visible = False
         Image1(Index).Picture = LoadPicture(COPLANO$)
         Picture1(Index).Top = TOPEINI(Index)
         Picture1(Index).Left = IZQUINI(Index)
         Picture1(Index).Width = ANCHINI(Index)
         Picture1(Index).Height = ALTUINI(Index)
         Picture1(Index).Refresh
         '
         HIMAGEN = Image1(Index).Picture.Width
         VIMAGEN = Image1(Index).Picture.Height
         
         HCUADRO = Picture1(Index).Width
         VCUADRO = Picture1(Index).Height
         '
         REDUH = HIMAGEN / HCUADRO
         REDUV = VIMAGEN / VCUADRO
         REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
         Picture1(Index).Height = VCUADRO * REDUV / REDUC
         Picture1(Index).Width = HCUADRO * REDUH / REDUC
         Picture1(Index).Refresh
         '
         HCUADRO = Picture1(Index).Width
         VCUADRO = Picture1(Index).Height
         Picture1(Index).Left = (Frame1(Index).Width - HCUADRO) / 2
         '
         Picture1(Index).PaintPicture Image1(Index).Picture, 0, 0, HCUADRO, VCUADRO
         Picture1(Index).Visible = True
         Picture1(Index).Refresh
       Else
         'Call MENSERR(24, "No Existe Archivo '" + COPLANO$ + "'.")
      End If
   End If
   '
   Screen.MousePointer = 1
   '
99 End Sub

Private Sub Text5_Change(Index As Integer)
   MODIFI% = 1
End Sub

Private Sub Text6_Change()
   MODIFI% = 1
End Sub

Private Sub Text7_Change()
   MODIFI% = 1
End Sub

Private Sub Text8_Change()
   MODIFI% = 1
End Sub

Private Sub Text9_Change()
   MODIFI% = 1
End Sub

Private Sub VScroll1_Change()
     HCUADRO = Picture4.Width
     VCUADRO = Picture4.Height
     Picture4.PaintPicture Image1(IMACTI).Picture, (-1) * HScroll1.Value, (-1) * VScroll1.Value, 1 * HCUADRO, 1 * VCUADRO
End Sub

