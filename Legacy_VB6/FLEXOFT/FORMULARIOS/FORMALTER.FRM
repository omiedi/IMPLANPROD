VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FORMALTER 
   BackColor       =   &H00EBE6DD&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Accesos Directos"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10245
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   10245
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   240
      Picture         =   "FORMALTER.frx":0000
      ScaleHeight     =   450
      ScaleWidth      =   480
      TabIndex        =   116
      Top             =   315
      Width           =   480
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00EBE6DD&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   240
      Picture         =   "FORMALTER.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   495
      TabIndex        =   115
      Top             =   315
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox LEYEN1 
      Height          =   435
      Index           =   0
      Left            =   5040
      ScaleHeight     =   375
      ScaleWidth      =   1425
      TabIndex        =   113
      Top             =   2940
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   8520
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   112
      Top             =   2520
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   7665
      ScaleHeight     =   540
      ScaleWidth      =   1275
      TabIndex        =   111
      ToolTipText     =   "Click con Boton Derecho para Cambiar la Imagen"
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EBE6DD&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   4200
      ScaleHeight     =   2295
      ScaleWidth      =   6015
      TabIndex        =   105
      Top             =   6600
      Visible         =   0   'False
      Width           =   6015
      Begin VB.Frame Frame2 
         BackColor       =   &H00EBE6DD&
         Caption         =   "Avance de Proceso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         TabIndex        =   106
         Top             =   120
         Width           =   3975
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Capturando Imagen . . . . . . . . . ."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   107
            Top             =   360
            Width           =   2775
         End
      End
      Begin VB.Label Label4 
         Height          =   255
         Left            =   0
         TabIndex        =   110
         Top             =   1800
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.Label Label3 
         Height          =   255
         Left            =   0
         TabIndex        =   109
         Top             =   1320
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   0
         TabIndex        =   108
         Top             =   960
         Visible         =   0   'False
         Width           =   3975
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   7680
   End
   Begin VB.ListBox SORTLIST 
      Height          =   255
      ItemData        =   "FORMALTER.frx":0614
      Left            =   8190
      List            =   "FORMALTER.frx":061B
      Sorted          =   -1  'True
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   1890
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.ListBox LISORT 
      Height          =   255
      ItemData        =   "FORMALTER.frx":0629
      Left            =   8190
      List            =   "FORMALTER.frx":0630
      TabIndex        =   71
      Top             =   1260
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   5145
      Pattern         =   "*.XYZ"
      TabIndex        =   70
      Top             =   945
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   105
      TabIndex        =   69
      Top             =   5040
      Visible         =   0   'False
      Width           =   9990
   End
   Begin VB.ListBox LINOSORT 
      Height          =   255
      ItemData        =   "FORMALTER.frx":063C
      Left            =   8190
      List            =   "FORMALTER.frx":0643
      TabIndex        =   68
      Top             =   630
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.ListBox PSPOOL 
      Height          =   1230
      Left            =   105
      TabIndex        =   67
      Top             =   3465
      Visible         =   0   'False
      Width           =   9990
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EBE6DD&
      BorderStyle     =   0  'None
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3060
      Left            =   1155
      TabIndex        =   33
      Top             =   105
      Width           =   3480
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1260
         TabIndex        =   0
         Top             =   210
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   0
         TabIndex        =   1
         Top             =   210
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   12
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   13
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   14
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   15
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   16
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   17
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   18
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   19
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   20
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   21
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   22
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   23
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   24
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   25
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   26
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   27
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   28
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   29
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   30
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   31
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.Label TEPRUEBA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   66
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   65
         Top             =   1050
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   0
         TabIndex        =   64
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   0
         TabIndex        =   63
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   0
         TabIndex        =   62
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   0
         TabIndex        =   61
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   0
         TabIndex        =   60
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   0
         TabIndex        =   59
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   0
         TabIndex        =   58
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   0
         TabIndex        =   57
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   0
         TabIndex        =   56
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   0
         TabIndex        =   55
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   0
         TabIndex        =   54
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   12
         Left            =   0
         TabIndex        =   53
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   13
         Left            =   0
         TabIndex        =   52
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   14
         Left            =   0
         TabIndex        =   51
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   15
         Left            =   0
         TabIndex        =   50
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   16
         Left            =   0
         TabIndex        =   49
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   17
         Left            =   0
         TabIndex        =   48
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   18
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   19
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   20
         Left            =   0
         TabIndex        =   45
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   21
         Left            =   0
         TabIndex        =   44
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   22
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   23
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   24
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   25
         Left            =   0
         TabIndex        =   40
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   26
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   27
         Left            =   0
         TabIndex        =   38
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   28
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   29
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   30
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   31
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   330
      Left            =   220
      Picture         =   "FORMALTER.frx":0651
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Imprime Mensaje por Impresora"
      Top             =   2730
      Visible         =   0   'False
      Width           =   435
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   6240
      OleObjectBlob   =   "FORMALTER.frx":0753
      TabIndex        =   114
      Top             =   1200
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   9840
      OleObjectBlob   =   "FORMALTER.frx":07CD
      Top             =   2400
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   31
      Left            =   0
      TabIndex        =   104
      Top             =   720
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   30
      Left            =   0
      TabIndex        =   103
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   29
      Left            =   0
      TabIndex        =   102
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   28
      Left            =   0
      TabIndex        =   101
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   27
      Left            =   0
      TabIndex        =   100
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   26
      Left            =   0
      TabIndex        =   99
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   25
      Left            =   0
      TabIndex        =   98
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   24
      Left            =   0
      TabIndex        =   97
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   23
      Left            =   0
      TabIndex        =   96
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   22
      Left            =   0
      TabIndex        =   95
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   21
      Left            =   0
      TabIndex        =   94
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   20
      Left            =   0
      TabIndex        =   93
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   19
      Left            =   0
      TabIndex        =   92
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   18
      Left            =   0
      TabIndex        =   91
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   17
      Left            =   0
      TabIndex        =   90
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   16
      Left            =   0
      TabIndex        =   89
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   15
      Left            =   0
      TabIndex        =   88
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   14
      Left            =   0
      TabIndex        =   87
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   13
      Left            =   0
      TabIndex        =   86
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   12
      Left            =   0
      TabIndex        =   85
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   11
      Left            =   0
      TabIndex        =   84
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   10
      Left            =   0
      TabIndex        =   83
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   9
      Left            =   0
      TabIndex        =   82
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   8
      Left            =   0
      TabIndex        =   81
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   7
      Left            =   0
      TabIndex        =   80
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   6
      Left            =   0
      TabIndex        =   79
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   5
      Left            =   0
      TabIndex        =   78
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   4
      Left            =   0
      TabIndex        =   77
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   3
      Left            =   0
      TabIndex        =   76
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   2
      Left            =   0
      TabIndex        =   75
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   1
      Left            =   0
      TabIndex        =   74
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label MODACDIR 
      Height          =   330
      Index           =   0
      Left            =   5250
      TabIndex        =   73
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Image Image3 
      Height          =   1050
      Left            =   240
      Stretch         =   -1  'True
      Top             =   7680
      Visible         =   0   'False
      Width           =   1320
   End
End
Attribute VB_Name = "FORMALTER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTALTER_Click(Index As Integer)
   '
   If MODALTER% = 1 Then
     Hide
     Call FRESHALL
     If FORMALTER.MODACDIR(Index).Caption <> "" Then
         Call CONECRUN(FORMALTER.MODACDIR(Index).Caption, "")
       Else
         Call FRESHALL
     End If
   ElseIf MODALTER% = 2 Then
     ALTERX% = Index% + 1
     On Error Resume Next
     FORMALTER.BOTALTER(0).SetFocus
     On Error GoTo 0
     Hide
     Call FRESHALL
   End If
   '
End Sub

Private Sub BOTALTER_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If Index = 0 Then
      If KeyCode = 34 Then
         Call BOTALTER_Click(Index)
      End If
   End If
End Sub

Private Sub Command1_Click()
   PRINTFORM
End Sub

Private Sub Form_Activate()
   Static YAPLI%
''''10 If YAPLI% < 2 Then
''''      Call APLICACIONSKINS(Me)
''''      YAPLI% = YAPLI% + 1
''''      TINI = Timer: While Timer < TINI + 1: Wend
''''      GoTo 10
''''   End If
''''   Refresh
   On Error Resume Next
   FORMALTER.BOTALTER(0).SetFocus
   On Error GoTo 0
End Sub

Private Sub Form_DblClick()
   PRINTFORM
End Sub

Sub FRESHALL()
    '
    If UCase$(App.EXEName) = "FAC2002" Then Exit Sub
    If UCase$(App.EXEName) = "FAC2001" Then GoTo 10
    If UCase$(App.EXEName) = "RFAC2001" Then Exit Sub
    DoEvents
Exit Sub
    If App.EXEName = "AMASTO01" Then Exit Sub
    '
10  For TTT& = 1 To Forms.Count
      If Forms(TTT& - 1).Visible = True Then
        'Forms(TTT& - 1).Refresh
        'NOMEFORM$ = UCase$(Forms(TTT& - 1).Name)
        'If InStr(NOMEFORM$, "FORVENTA") > 0 Or InStr(NOMEFORM$, "ORDEFAB") > 0 Then
            On Error Resume Next
            For U& = 0 To Forms(TTT& - 1).Controls.Count - 1
                Forms(TTT& - 1).Controls(U&).Refresh
            Next U&
            On Error GoTo 0
        'End If
      End If
    Next TTT&
    '
End Sub

Private Sub Form_Load()
   If Control("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
    End If

   CAOPS = 0
   ' Call APLICACIONSKINS(Me)
   
   For i = 0 To 30
        LEYEN(i).Font = "Arial"
        LEYEN(i).FontSize = 10
        LEYEN(i).FontBold = True
   Next i
   
   For i = 0 To 30
        BOTALTER(i).Font = "Arial"
        BOTALTER(i).FontSize = 8
        BOTALTER(i).FontBold = True
   Next i
   
End Sub

'
Private Sub Timer1_Timer()
     '
     RUTAFO$ = FORMALTER.Label2
     TINI = XVALO(FORMALTER.Label3)
     File1.Path = RUTAFO$
     File1.Pattern = "*.*"
     File1.Refresh
     '
     CANFIX& = File1.ListCount
     For KKII& = 1 To CANFIX&
        ARCHIX$ = File1.List(KKII& - 1)
        On Error Resume Next
        FHORA = CDbl(FileDateTime(RUTAFO$ + ARCHIX$))
        If FHORA >= TINI Then
          Label4 = RUTAFO$ + ARCHIX$
          Timer1.Enabled = False
          Hide
          Exit Sub
        End If
     Next KKII&
     '
End Sub
