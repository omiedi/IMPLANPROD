VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form OPCLAV04 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Resúmenes Mensuales de Ventas"
   ClientHeight    =   8190
   ClientLeft      =   1650
   ClientTop       =   600
   ClientWidth     =   8550
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   84
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   8520
      Left            =   7665
      ScaleHeight     =   8460
      ScaleWidth      =   1845
      TabIndex        =   67
      Top             =   0
      Width           =   1905
      Begin VB.PictureBox Picture16 
         Height          =   8200
         Left            =   0
         ScaleHeight     =   8145
         ScaleWidth      =   1020
         TabIndex        =   86
         Top             =   0
         Visible         =   0   'False
         Width           =   1080
         Begin VB.Image Image3 
            Height          =   8200
            Left            =   0
            Picture         =   "OPCLAV04.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   90
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   72
         Top             =   7245
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   73
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.CommandButton command1 
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
         Height          =   690
         Left            =   105
         Picture         =   "OPCLAV04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   180
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Left            =   105
         Picture         =   "OPCLAV04.frx":24E2
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   870
         Width           =   650
      End
      Begin VB.CommandButton Command7 
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
         Height          =   690
         Left            =   105
         Picture         =   "OPCLAV04.frx":27EC
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   1560
         Width           =   650
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Plan"
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
         Picture         =   "OPCLAV04.frx":2AF6
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Archivo Maestro de Cuentas Contables"
         Top             =   5985
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "OPCLAV04.frx":2C40
         Top             =   7590
         Width           =   1515
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Otros Movimientos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   210
      TabIndex        =   57
      Top             =   6860
      Width           =   7260
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   12
         Left            =   3150
         TabIndex        =   77
         Top             =   735
         Width           =   150
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
         Index           =   12
         Left            =   1995
         TabIndex        =   76
         Text            =   " "
         Top             =   735
         Width           =   1170
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
         Index           =   10
         Left            =   1995
         TabIndex        =   61
         Text            =   " "
         Top             =   420
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   10
         Left            =   3150
         TabIndex        =   60
         Top             =   420
         Width           =   150
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   825
         Left            =   210
         TabIndex        =   58
         Top             =   315
         Width           =   1905
         Begin VB.Label LABEL2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Gastos Embalaje"
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
            Index           =   12
            Left            =   -210
            TabIndex        =   79
            Top             =   495
            Width           =   1905
         End
         Begin VB.Label LABEL2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta IVA Ventas:"
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
            Index           =   10
            Left            =   -210
            TabIndex        =   59
            Top             =   195
            Width           =   1905
         End
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   12
         Left            =   3465
         TabIndex        =   78
         Top             =   735
         Width           =   3585
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   10
         Left            =   3465
         TabIndex        =   62
         Top             =   420
         Width           =   3585
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Clasificación de Débitos y Créditos Varios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   210
      TabIndex        =   47
      Top             =   5250
      Width           =   7260
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
         Index           =   11
         Left            =   1995
         TabIndex        =   65
         Text            =   " "
         Top             =   1050
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   3150
         TabIndex        =   64
         Top             =   1050
         Width           =   150
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   3150
         TabIndex        =   54
         Top             =   735
         Width           =   150
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
         Index           =   9
         Left            =   1995
         TabIndex        =   53
         Text            =   " "
         Top             =   735
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   3150
         TabIndex        =   52
         Top             =   420
         Width           =   150
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
         Index           =   8
         Left            =   1995
         TabIndex        =   51
         Text            =   " "
         Top             =   420
         Width           =   1170
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   950
         Left            =   25
         TabIndex        =   48
         Top             =   480
         Width           =   1905
         Begin VB.Label LABEL2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta IVA Ventas:"
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
            Index           =   11
            Left            =   0
            TabIndex        =   63
            Top             =   630
            Width           =   1905
         End
         Begin VB.Label LABEL2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta IVA Ventas:"
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
            Index           =   9
            Left            =   0
            TabIndex        =   50
            Top             =   315
            Width           =   1905
         End
         Begin VB.Label LABEL2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta IVA Ventas:"
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
            Index           =   8
            Left            =   0
            TabIndex        =   49
            Top             =   0
            Width           =   1905
         End
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   11
         Left            =   3465
         TabIndex        =   66
         Top             =   1050
         Width           =   3585
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   9
         Left            =   3465
         TabIndex        =   56
         Top             =   735
         Width           =   3585
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   8
         Left            =   3465
         TabIndex        =   55
         Top             =   420
         Width           =   3585
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cuentas Contables I.V.A."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1830
      Left            =   210
      TabIndex        =   1
      Top             =   105
      Width           =   7260
      Begin VB.CommandButton Command4 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   3150
         TabIndex        =   81
         Top             =   735
         Width           =   150
      End
      Begin VB.TextBox Text2 
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
         Index           =   3
         Left            =   1995
         TabIndex        =   80
         Text            =   " "
         Top             =   735
         Width           =   1170
      End
      Begin VB.TextBox Text100 
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
         Left            =   6510
         TabIndex        =   75
         Text            =   " "
         ToolTipText     =   "Código de Régimen de Percepción"
         Top             =   1365
         Width           =   540
      End
      Begin VB.TextBox Text2 
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
         Left            =   1995
         TabIndex        =   7
         Text            =   " "
         Top             =   420
         Width           =   1170
      End
      Begin VB.CommandButton Command4 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3150
         TabIndex        =   6
         Top             =   420
         Width           =   150
      End
      Begin VB.TextBox Text2 
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
         Left            =   1995
         TabIndex        =   5
         Text            =   " "
         Top             =   1050
         Width           =   1170
      End
      Begin VB.CommandButton Command4 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3150
         TabIndex        =   4
         Top             =   1050
         Width           =   150
      End
      Begin VB.TextBox Text2 
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
         Left            =   1995
         TabIndex        =   3
         Text            =   " "
         Top             =   1365
         Width           =   1170
      End
      Begin VB.CommandButton Command4 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3150
         TabIndex        =   2
         Top             =   1365
         Width           =   150
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta IVA T.Sec:"
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
         Index           =   4
         Left            =   0
         TabIndex        =   83
         Top             =   750
         Width           =   1905
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   3
         Left            =   3465
         TabIndex        =   82
         Top             =   735
         Width           =   3585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C.R.:"
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
         Index           =   3
         Left            =   4515
         TabIndex        =   74
         ToolTipText     =   "Código de Régimen de Percepción"
         Top             =   1365
         Width           =   1905
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3465
         TabIndex        =   38
         Top             =   1365
         Width           =   2430
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3465
         TabIndex        =   37
         Top             =   1050
         Width           =   3585
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3465
         TabIndex        =   36
         Top             =   420
         Width           =   3585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta IVA Ventas:"
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
         Index           =   0
         Left            =   0
         TabIndex        =   10
         Top             =   450
         Width           =   1905
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta IVA R.N.I:"
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
         Index           =   1
         Left            =   0
         TabIndex        =   9
         Top             =   1050
         Width           =   1905
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Percepción IVA:"
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
         Index           =   2
         Left            =   0
         TabIndex        =   8
         Top             =   1365
         Width           =   1905
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Clasificación de Ventas Según Posición de IVA Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   210
      TabIndex        =   0
      Top             =   2050
      Width           =   7260
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   2535
         Left            =   25
         TabIndex        =   27
         Top             =   480
         Width           =   1905
         Begin VB.Label LABEL2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta IVA Ventas:"
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
            Index           =   7
            Left            =   0
            TabIndex        =   35
            Top             =   2205
            Width           =   1905
         End
         Begin VB.Label LABEL2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta IVA Ventas:"
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
            Index           =   6
            Left            =   0
            TabIndex        =   34
            Top             =   1890
            Width           =   1905
         End
         Begin VB.Label LABEL2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta IVA Ventas:"
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
            Index           =   5
            Left            =   0
            TabIndex        =   33
            Top             =   1575
            Width           =   1905
         End
         Begin VB.Label LABEL2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta IVA Ventas:"
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
            Index           =   3
            Left            =   0
            TabIndex        =   32
            Top             =   945
            Width           =   1905
         End
         Begin VB.Label LABEL2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta IVA Ventas:"
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
            Index           =   2
            Left            =   0
            TabIndex        =   31
            Top             =   630
            Width           =   1905
         End
         Begin VB.Label LABEL2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta IVA Ventas:"
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
            Index           =   1
            Left            =   0
            TabIndex        =   30
            Top             =   315
            Width           =   1905
         End
         Begin VB.Label LABEL2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta IVA Ventas:"
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
            Index           =   0
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   1905
         End
         Begin VB.Label LABEL2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta IVA Ventas:"
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
            Index           =   4
            Left            =   0
            TabIndex        =   28
            Top             =   1260
            Width           =   1905
         End
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
         Index           =   7
         Left            =   1995
         TabIndex        =   26
         Text            =   " "
         Top             =   2625
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   3150
         TabIndex        =   25
         Top             =   2625
         Width           =   150
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
         Index           =   6
         Left            =   1995
         TabIndex        =   24
         Text            =   " "
         Top             =   2310
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   3150
         TabIndex        =   23
         Top             =   2310
         Width           =   150
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
         Index           =   5
         Left            =   1995
         TabIndex        =   22
         Text            =   " "
         Top             =   1995
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   3150
         TabIndex        =   21
         Top             =   1995
         Width           =   150
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
         Index           =   4
         Left            =   1995
         TabIndex        =   20
         Text            =   " "
         Top             =   1680
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   3150
         TabIndex        =   19
         Top             =   1680
         Width           =   150
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
         Left            =   1995
         TabIndex        =   18
         Text            =   " "
         Top             =   1050
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3150
         TabIndex        =   17
         Top             =   1050
         Width           =   150
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
         Left            =   1995
         TabIndex        =   16
         Text            =   " "
         Top             =   735
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3150
         TabIndex        =   15
         Top             =   735
         Width           =   150
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
         Left            =   1995
         TabIndex        =   14
         Text            =   " "
         Top             =   420
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3150
         TabIndex        =   13
         Top             =   420
         Width           =   150
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
         Index           =   3
         Left            =   1995
         TabIndex        =   12
         Text            =   " "
         Top             =   1365
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   3150
         TabIndex        =   11
         Top             =   1365
         Width           =   150
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   7
         Left            =   3465
         TabIndex        =   46
         Top             =   2625
         Width           =   3585
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   6
         Left            =   3465
         TabIndex        =   45
         Top             =   2310
         Width           =   3585
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   5
         Left            =   3465
         TabIndex        =   44
         Top             =   1995
         Width           =   3585
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   4
         Left            =   3465
         TabIndex        =   43
         Top             =   1680
         Width           =   3585
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3465
         TabIndex        =   42
         Top             =   1050
         Width           =   3585
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3465
         TabIndex        =   41
         Top             =   735
         Width           =   3585
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3465
         TabIndex        =   40
         Top             =   420
         Width           =   3585
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   3
         Left            =   3465
         TabIndex        =   39
         Top             =   1365
         Width           =   3585
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "OPCLAV04.frx":4B62
      TabIndex        =   85
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "OPCLAV04.frx":4BDC
      Top             =   1200
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu grabar 
         Caption         =   "Grabar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu accesodirecto 
      Caption         =   "Accesos Directos"
      Begin VB.Menu plandecuentas 
         Caption         =   "Plan de Cuentas"
      End
   End
End
Attribute VB_Name = "OPCLAV04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command8_Click(Index As Integer)
    Call SELECHO("ECUECON")
    If VALACT1$("ECUECON") <> "" Then
        Text4(Index).TEXT = VALACT1$("ECUECON")
        Label4(Index) = VALACT2$("ECUECON")
    End If
End Sub

Private Sub Form_Load()
   '
   If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      Me.Frame7.BackColor = &HFCD7B6
      Me.Frame8.BackColor = &HFCD7B6
      
      MARCOBARRA.Top = 7800
      '
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
    End If

   Call APLICACIONSKINS(Me)
   '
   CUEIVA$ = CONTROL$("SVD451S", "CUEIVA")
   Text2(0) = CUEIVA$
   Label3(0) = ECOARCH$("ECUECON", CUEIVA$)
   '
   CUEIVB$ = CONTROL$("SVD451S", "CUEIVB")
   Text2(3) = CUEIVB$
   Label3(3) = ECOARCH$("ECUECON", CUEIVB$)
   '
   CUEADU$ = CONTROL$("SVD451S", "CUEADU")
   Text2(1) = CUEADU$
   Label3(1) = ECOARCH$("ECUECON", CUEADU$)
   '
   CUERET$ = CONTROL$("SVD451S", "CUERET")
   Text2(2) = CUERET$
   Label3(2) = ECOARCH$("ECUECON", CUERET$)
   '
   Label2(0) = "Consumidor Final:"
   Label2(1) = "IVA Resp.Inscripto"
   Label2(2) = "Resp. No Inscripto"
   Label2(3) = "No Responsable IVA"
   Label2(4) = "IVA Exento"
   Label2(5) = "Comercio Exterior"
   Label2(6) = "Monotributo"
   Label2(7) = "Zona Promovida"
   '
   Label2(8) = "Cheques Rechazados"
   Label2(9) = "Diferencia de Cambio"
   Label2(10) = "Ajustes de Saldo"
   Label2(11) = "Otros Deb/Creditos"
   '
   For KKI% = 0 To 12
     Text4(KKI%).TEXT = CONTROL$("SVD451S", "CUEVE-" + TRIM$(Str$(KKI%)))
   Next KKI%
   '
   Screen.MousePointer = 1
   '
   '\\\OT-06-0158-WAR-30/05/06///
   Text100.TEXT = CONTROL$("COREPER", "COREPER")
   '\\\OT-06-0158-WAR-FIN///
   '
End Sub

Private Sub Command2_Click()
    '
    CUEIVA$ = AJUSTI$(Text2(0).TEXT, 12)
    CUEIVB$ = AJUSTI$(Text2(3).TEXT, 12)
    CUEADU$ = AJUSTI$(Text2(1).TEXT, 12)
    CUERET$ = AJUSTI$(Text2(2).TEXT, 12)
    '
    CUEIVA1% = 0: CUEIVB1% = 0: CUEADU1% = 0: CUERET1% = 0
    For U& = 1 To ULTREG&("ECUECON")
       X$ = REGLEIDO$("ECUECON", U&)
       If Left$(X$, 12) = CUEIVA$ Then CUEIVA1% = CVI(Mid$(X$, 41, 2)): DECIVA$ = Mid$(X$, 13, 28)
       If Left$(X$, 12) = CUEIVB$ Then CUEIVB1% = CVI(Mid$(X$, 41, 2)): DECIVA$ = Mid$(X$, 13, 28)
       If Left$(X$, 12) = CUEADU$ Then CUEADU1% = CVI(Mid$(X$, 41, 2)): DECADU$ = Mid$(X$, 13, 28)
       If Left$(X$, 12) = CUERET$ Then CUERET1% = CVI(Mid$(X$, 41, 2)): DECRET$ = Mid$(X$, 13, 28)
    Next U&
    '
    If CUEIVA1% < 1 Then
       Call MENSERR(24, "Falta Cuenta Contable I.V.A.")
       Text2(0).SetFocus
       Exit Sub
    End If
    '
    If CUEIVB1% < 1 Then CUEIVB$ = ""
    '
    If CUEADU1% < 1 Then
       Call MENSERR(24, "Falta Cuenta Contable I.V.A. R.N.I.")
       Text2(1).SetFocus
       Exit Sub
    End If
    '
    If CUERET1% < 1 Then
       Call MENSERR(24, "Falta Cuenta Contable Retenciones I.V.A.")
       Text2(2).SetFocus
       Exit Sub
    End If
    '
    YA1% = 0: YA2% = 0: YA3% = 0
    For U& = 1 To ULTREG&("CUECON")
       X$ = REGLEIDO$("CUECON", U&)
       If CVI(Left$(X$, 2)) = CUEIVA1% Then YA1% = 1
       If CVI(Left$(X$, 2)) = CUEADU1% Then YA2% = 1
       If CVI(Left$(X$, 2)) = CUERET1% Then YA3% = 1
    Next U&
    '
    If YA1% < 1 Then
       X$ = REGBLAN$("CUECON")
       Call REPLA(X$, MKI$(CUEIVA1%), 1, 2)
       Call REPLA(X$, DECIVA$, 3, 29)
       Call REPLA(X$, CUEIVA$, 33, 12)
       Call REPLA(X$, CUEIVA$, 45, 12)
       Call REGAPP("CUECON", X$)
    End If
    '
    If YA2% < 1 Then
       X$ = REGBLAN$("CUECON")
       Call REPLA(X$, MKI$(CUEADU1%), 1, 2)
       Call REPLA(X$, DECADU$, 3, 29)
       Call REPLA(X$, CUEADU$, 33, 12)
       Call REPLA(X$, CUEADU$, 45, 12)
       Call REGAPP("CUECON", X$)
    End If
    '
    If YA3% < 1 Then
       X$ = REGBLAN$("CUECON")
       Call REPLA(X$, MKI$(CUERET1%), 1, 2)
       Call REPLA(X$, DECRET$, 3, 29)
       Call REPLA(X$, CUERET$, 33, 12)
       Call REPLA(X$, CUERET$, 45, 12)
       Call REGAPP("CUECON", X$)
    End If
    '
    Call GRACONTROL("SVD451S", "CUEIVA", CUEIVA$)
    Call GRACONTROL("SVD451S", "CUEIVB", CUEIVB$)
    Call GRACONTROL("SVD451S", "CUEADU", CUEADU$)
    Call GRACONTROL("SVD451S", "CUERET", CUERET$)
    '
    MOCALVE$ = "PIVACLI"
    Call GRACONTROL("REMEVENT", "MODCALVE", MOCALVE$)
    '
    If MOCALVE$ = "PIVACLI" Then
      For KKI% = 0 To 12
        CUEVENTA$(KKI%) = AJUSTI$(Text4(KKI%).TEXT, 12)
      Next KKI%
      '
      For U& = 1 To ULTREG&("ECUECON")
        X$ = REGLEIDO$("ECUECON", U&)
        For KKI% = 0 To 12
          If Left$(X$, 12) = CUEVENTA$(KKI%) Then COCUEVEN%(KKI%) = CVI(Mid$(X$, 41, 2)): DECIVA$ = Mid$(X$, 13, 28)
        Next KKI%
      Next U&
      '
      For KKI% = 0 To 12
        If COCUEVEN%(KKI%) > 0 Or KKI% = 12 Then
            Call GRACONTROL("SVD451S", "CUEVE-" + TRIM$(Str$(KKI%)), CUEVENTA$(KKI%))
          Else
            Call MENSERR(24, "Faltan Cuentas Contables de Venta.")
            Text4(KKI%).SetFocus
            Exit Sub
        End If
      Next KKI%
    End If
    '\\\OT-06-0158-WAR-30/05/06///
      Call GRACONTROL("COREPER", "COREPER", Text100.TEXT)
    '\\\OT-06-0158-WAR-FIN///
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
    '
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Unload Me
    On Error Resume Next
    Hide
    On Error GoTo 0
End Sub

Private Sub grabar_Click()
    Call Command2_Click
End Sub

Private Sub plandecuentas_Click()
    Call Command11_Click
End Sub

Private Sub salir_Click()
    Call Command1_Click
End Sub

Private Sub Text2_Change(Index As Integer)
    DECUE$ = ECOARCH("ECUECON", TRIM$(Text2(Index).TEXT))
    If DECUE$ <> "" Then
        Label3(Index) = DECUE$
      Else
        Label3(Index) = ""
    End If
End Sub

Private Sub Text2_DblClick(Index As Integer)
    Call SELECHO("ECUECON")
    If VALACT1$("ECUECON") <> "" Then
        Text2(Index).TEXT = VALACT1$("ECUECON")
        Label3(Index) = VALACT2$("ECUECON")
    End If
End Sub
'
Private Sub Command4_Click(Index As Integer)
    Call SELECHO("ECUECON")
    If VALACT1$("ECUECON") <> "" Then
        Text2(Index).TEXT = VALACT1$("ECUECON")
        Label3(Index) = VALACT2$("ECUECON")
    End If
End Sub
'
Private Sub Text4_Change(Index As Integer)
    DECUE$ = ECOARCH("ECUECON", TRIM$(Text4(Index).TEXT))
    If DECUE$ <> "" Then
        Label4(Index) = DECUE$
      Else
        Label4(Index) = ""
    End If
End Sub

Private Sub Text4_DblClick(Index As Integer)
    Call SELECHO("ECUECON")
    If VALACT1$("ECUECON") <> "" Then
        Text4(Index).TEXT = VALACT1$("ECUECON")
        Label4(Index) = VALACT2$("ECUECON")
    End If
End Sub

Private Sub Command11_Click()
   '
   Command11.Enabled = False
   RETEJ% = EJER%
   EJER% = 1
   Call CARCUEINGAS
   EJER% = RETEJ%
   EJU$ = EJACT$
   Command11.Enabled = True
   '
End Sub


