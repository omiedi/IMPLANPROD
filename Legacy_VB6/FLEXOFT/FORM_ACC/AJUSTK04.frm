VERSION 5.00
Begin VB.Form AJUSTK04 
   BackColor       =   &H00EEEEB0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ajuste de Stocks a Fecha"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   9840
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture6 
      Height          =   6735
      Left            =   8925
      ScaleHeight     =   6675
      ScaleWidth      =   1035
      TabIndex        =   64
      Top             =   0
      Visible         =   0   'False
      Width           =   1100
      Begin VB.Image Image3 
         Height          =   6900
         Left            =   0
         Picture         =   "AJUSTK04.frx":0000
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      Height          =   6735
      Left            =   8925
      ScaleHeight     =   6675
      ScaleWidth      =   1950
      TabIndex        =   53
      Top             =   0
      Width           =   2010
      Begin VB.CommandButton Command25 
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
         Picture         =   "AJUSTK04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Importar Inventario de Stocks"
         Top             =   4515
         Width           =   650
      End
      Begin VB.CommandButton Command33 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   33.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         TabIndex        =   60
         ToolTipText     =   "Ajuste de Stocks a Cero"
         Top             =   3600
         Width           =   650
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Direct"
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
         Picture         =   "AJUSTK04.frx":24E2
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   1780
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "AJUSTK04.frx":27EC
         Style           =   1  'Graphical
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   2730
         Width           =   650
      End
      Begin VB.CommandButton Command1 
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
         Picture         =   "AJUSTK04.frx":2C2E
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   210
         Width           =   650
      End
      Begin VB.CommandButton Command29 
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
         Picture         =   "AJUSTK04.frx":2D78
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   900
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   54
         Top             =   5775
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   55
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "AJUSTK04.frx":3082
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   6090
         Width           =   1515
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00EEEEB0&
      Caption         =   "Consultas Complementarias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   4515
      TabIndex        =   43
      Top             =   105
      Width           =   4215
      Begin VB.PictureBox Picture12 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1260
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   44
         Top             =   1575
         Width           =   2535
         Begin VB.CommandButton Command13 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "AJUSTK04.frx":4FA4
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Consultas y Lista- dos  de Movimientos "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   630
            TabIndex        =   46
            Top             =   105
            Width           =   1830
         End
      End
      Begin VB.PictureBox Picture13 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   47
         Top             =   1110
         Width           =   2535
         Begin VB.CommandButton Command14 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   -20
            Picture         =   "AJUSTK04.frx":50EE
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Gestión de Inventario Continuo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   735
            TabIndex        =   49
            Top             =   0
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture14 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   50
         Top             =   630
         Width           =   2535
         Begin VB.CommandButton Command15 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "AJUSTK04.frx":53F8
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Ficha Kardex  de Movimientos"
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
            Left            =   840
            TabIndex        =   52
            Top             =   0
            Width           =   1410
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00EEEEB0&
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1850
      Left            =   210
      TabIndex        =   35
      Top             =   2835
      Width           =   4110
      Begin VB.Frame Frame7 
         BackColor       =   &H00EEEEB0&
         Caption         =   "Motivo - Referencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   140
         TabIndex        =   40
         Top             =   945
         Width           =   3820
         Begin VB.TextBox Text5 
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
            Left            =   210
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   315
            Width           =   3480
         End
      End
      Begin VB.TextBox Text4 
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
         Left            =   315
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   580
         Width           =   2220
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
         Height          =   285
         Left            =   2730
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   580
         Width           =   1065
      End
      Begin VB.CommandButton Command5 
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
         Left            =   3780
         TabIndex        =   36
         Top             =   580
         Width           =   175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   41
         Top             =   315
         Width           =   1065
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2625
         TabIndex        =   39
         Top             =   315
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEB0&
      Caption         =   "Consultas - Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3825
      Left            =   4515
      TabIndex        =   18
      Top             =   2730
      Width           =   4215
      Begin VB.Frame Frame6 
         BackColor       =   &H00EEEEB0&
         Caption         =   "Salida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   150
         TabIndex        =   34
         Top             =   2520
         Width           =   3060
         Begin VB.OptionButton Option4 
            BackColor       =   &H00EEEEB0&
            Caption         =   "Listado Impreso"
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
            Left            =   630
            TabIndex        =   30
            Top             =   735
            Width           =   1905
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00EEEEB0&
            Caption         =   "Por Pantalla"
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
            Left            =   630
            TabIndex        =   29
            Top             =   420
            Value           =   -1  'True
            Width           =   2115
         End
      End
      Begin VB.CommandButton Command7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   3360
         Picture         =   "AJUSTK04.frx":5702
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2625
         Width           =   645
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EEEEB0&
         Caption         =   "Rango Numérico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   150
         TabIndex        =   22
         Top             =   525
         Width           =   1800
         Begin VB.TextBox Text14 
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
            Left            =   525
            TabIndex        =   23
            Top             =   735
            Width           =   1065
         End
         Begin VB.TextBox Text13 
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
            Left            =   525
            TabIndex        =   24
            Top             =   1365
            Width           =   1065
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
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
            Left            =   0
            TabIndex        =   33
            Top             =   500
            Width           =   855
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
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
            Left            =   105
            TabIndex        =   32
            Top             =   1130
            Width           =   750
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EEEEB0&
         Caption         =   "Rango de Fechas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   2100
         TabIndex        =   19
         Top             =   525
         Width           =   1905
         Begin VB.TextBox Text12 
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
            Left            =   525
            TabIndex        =   27
            Text            =   " "
            Top             =   1365
            Width           =   1065
         End
         Begin VB.TextBox Text11 
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
            Left            =   525
            TabIndex        =   25
            Text            =   " "
            Top             =   735
            Width           =   1065
         End
         Begin VB.CommandButton BOTSEL 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1575
            Picture         =   "AJUSTK04.frx":5A0C
            TabIndex        =   26
            Top             =   735
            Width           =   175
         End
         Begin VB.CommandButton BOTSEL 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1575
            Picture         =   "AJUSTK04.frx":5D16
            TabIndex        =   28
            Top             =   1365
            Width           =   175
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
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
            Left            =   105
            TabIndex        =   21
            Top             =   1130
            Width           =   750
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
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
            Left            =   0
            TabIndex        =   20
            Top             =   500
            Width           =   855
         End
      End
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
      Left            =   2730
      TabIndex        =   1
      Top             =   210
      Width           =   175
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   600
      Width           =   4110
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3675
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   210
      Width           =   645
   End
   Begin VB.TextBox Text1 
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
      Left            =   1050
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   210
      Width           =   1695
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   210
      TabIndex        =   2
      Top             =   1470
      Width           =   4110
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EEEEB0&
      Caption         =   "Ingreso de Ajustes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1850
      Left            =   210
      TabIndex        =   12
      Top             =   4725
      Width           =   4110
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
         Height          =   285
         Left            =   1785
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   735
         Width           =   1170
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
         Height          =   285
         Left            =   1785
         TabIndex        =   4
         Top             =   1050
         Width           =   1170
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1785
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   420
         Width           =   1170
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
         Height          =   285
         Left            =   1785
         TabIndex        =   5
         Top             =   1365
         Width           =   1170
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
         Height          =   600
         Left            =   3255
         Picture         =   "AJUSTK04.frx":6020
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   420
         Width           =   645
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
         Height          =   600
         Left            =   3255
         Picture         =   "AJUSTK04.frx":632A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Entrada:"
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
         Left            =   105
         TabIndex        =   17
         Top             =   735
         Width           =   1590
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Salida:"
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
         Left            =   105
         TabIndex        =   16
         Top             =   1050
         Width           =   1590
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "StockTeórico:"
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
         Left            =   105
         TabIndex        =   15
         Top             =   420
         Width           =   1590
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Ajustado:"
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
         Left            =   105
         TabIndex        =   14
         Top             =   1365
         Width           =   1590
      End
   End
   Begin VB.Label Label17 
      Height          =   225
      Left            =   1575
      TabIndex        =   63
      Top             =   1260
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label16 
      Caption         =   "lote"
      Height          =   225
      Left            =   315
      TabIndex        =   62
      Top             =   1260
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9135
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "U-M:"
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
      Left            =   2730
      TabIndex        =   9
      Top             =   255
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
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
      Left            =   105
      TabIndex        =   8
      Top             =   255
      Width           =   855
   End
   Begin VB.Menu Archiv 
      Caption         =   "Archivos"
      Begin VB.Menu NUaju 
         Caption         =   "Nuevo Ajuste"
      End
      Begin VB.Menu spx2 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu Exit 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Complem 
      Caption         =   "Complementos"
      Begin VB.Menu Kardex 
         Caption         =   "Ficha Kardex"
      End
      Begin VB.Menu Invent 
         Caption         =   "Inventario Continuo"
      End
      Begin VB.Menu Reports 
         Caption         =   "Consultas Movimientos"
      End
      Begin VB.Menu S 
         Caption         =   "-"
      End
      Begin VB.Menu AJUSTEACERO 
         Caption         =   "Ajuste de Stocks a Cero"
      End
      Begin VB.Menu plxc 
         Caption         =   "-"
      End
      Begin VB.Menu IMPORTARINVENTARIOSTOCKS 
         Caption         =   "Importar Inventario de Stocks"
      End
      Begin VB.Menu S1 
         Caption         =   "-"
      End
      Begin VB.Menu Liaju 
         Caption         =   "Listado de Ajustes"
         Begin VB.Menu LiajuPan 
            Caption         =   "Por Pantalla"
         End
         Begin VB.Menu LiajuPrin 
            Caption         =   "Por Impresora"
         End
      End
   End
   Begin VB.Menu Acdir 
      Caption         =   "Accesos Directos"
      Begin VB.Menu Amasto 
         Caption         =   "Maestro de Items de Stock"
      End
   End
   Begin VB.Menu Setup 
      Caption         =   "Configuración"
   End
   Begin VB.Menu Help 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "AJUSTK04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MODENTRAJU%

Private Sub AJUSTEACERO_Click()
   Call Command33_Click
End Sub

Private Sub Amasto_Click()
    ACONEC$ = "ARCHIG04/AMASTO"
    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
End Sub

Private Sub BOTSEL_Click(Index As Integer)
    Call SELECHO("SELFECHA")
    If VALACT1$("SELFECHA") <> "" Then
        If Index = 0 Then
            Text11.TEXT = VALACT1$("SELFECHA")
          Else
            Text12.TEXT = VALACT1$("SELFECHA")
        End If
    End If
End Sub

Private Sub Command1_Click()
    PREPA% = 0
    Hide
End Sub

Private Sub Command13_Click()
    Command13.Enabled = False
    QUERSTK04.Show 1
    Command13.Enabled = True
End Sub

Private Sub Command14_Click()
    Command14.Enabled = False
    If App.EXEName = "INVENT04" Then
      Command14.Enabled = True
      Hide
      Exit Sub
    End If
    Call CONECRUN("INVENT04", "Inventarios Físicos y Valorizados")
    Command14.Enabled = True
End Sub

Private Sub Command15_Click()
    Command15.Enabled = False
    FIKARDEX.Text1 = Text1
    FIKARDEX.Show 1
    Command15.Enabled = True
End Sub

Private Sub Command2_Click()
    COD$ = ""
    CI% = 0
    Text1.TEXT = ""
    Text2.TEXT = ""
    Text3.TEXT = ""
    Text6.TEXT = ""
    Text7.TEXT = ""
    Text8.TEXT = ""
    Text9.TEXT = ""
    Text1.SetFocus
End Sub

Private Sub Command25_Click()
    '
    Command25.Enabled = False
    If DERACCE%("IMPINV", "Importación de Inventario de Stocks") >= 2 Then
       IMPSTO04.Show 1
    End If
    Command25.Enabled = True
    '
End Sub

Private Sub Command26_Click()
'   Call CONECRUN("FSETUP04/OPTRAJUST", "Configuración de Funcionamiento")
   OPTRAJU04.Show 1

   Call FINAL("")
End Sub

Private Sub Command27_Click()
    Call ACCDIR("AJUSTOCK")
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("AJUSTOCK")
End Sub

Private Sub Command3_Click()
    CIXI% = CODINT%(Text1)
    If CIXI% < 1 Then Exit Sub
    '
    If XVALO(Text8) > 0 Then    ' ES UN MOVIMIENTO DE SALIDA
       Call CARDEPOS
       CAITMO% = 1
       CIMOVIM%(CAITMO%) = CIXI%
       CAMOVIM(CAITMO) = (-1) * XVALO(Text8)
       ILISTA& = List1.ListIndex
       DEPO% = COSU%(ILISTA& + 1)
       If PUEDEMOVI%(DEPO%) < 1 Then
          Exit Sub
       End If
    End If
    '
    Label17 = ""
    If EXISTE%("TRAZAB04.EXE") > 0 Then
      'If XVALO(Text8) > 0 Then ' ES UN MOVIMIENTO DE SALIDA
       Call ARMALILOT(CIXI%)
       Call FRESHECHO("!LILOTSE")
       If ULTREG("!LILOTSE") < 1 Then
          Call MENSERR(24, "No Hay Lotes Disponibles\para el Ajuste Propuesto.")
          Exit Sub
       End If
       '
       Call SELECHO("!LILOTSE")
       VDEVU$ = VALACT1$("!LILOTSE")
       VDEV2$ = VALACT2$("!LILOTSE")
       If TRIM$(VDEVU$) = "" Then
          Call COMUNI("Ajuste Cancelado")
          Exit Sub
       End If
       CAMAXI = XVALO(Mid$(VDEV2$, 21, 12))
       If XVALO(Text8) > CAMAXI Then
          Call MENSERR(24, "Cantidad del Ajuste Supera\Saldo Fisico del Lote")
          Exit Sub
       End If
       '
       Label17 = TRIM$(VDEVU$)
      'End If
      '
    End If
    '
    Call AJUSTO3
    Call AJUSTO2
    Text6.TEXT = Text9.TEXT
    Text7.TEXT = ""
    Text8.TEXT = ""
    Text1.SetFocus
End Sub

Sub Command33_Click()
'\\\OT-05-0704-WAR-02/11/05///
  Command33.Enabled = False
   '
   Dim CATEO#(32767)
   '
   If DERACCE%("AJUCERO", "Ajuste de Stocks a Cero") < 2 Then
     GoTo 99
   End If
   '
10 Call SELECHO("SELFECHA")
   FA$ = VALACT1$("SELFECHA")
   FA1% = FECHANUM(FA$)
   If FA1% < 0 Then
     Call MENSERR(24, "Fecha de Ajuste no Válida")
     GoTo 10
   End If
   If FA1% = 0 Then
    GoTo 99
   End If
   '
20 Call SELECHO("TADEPOSI")
   DP1% = XVALO(VALACT1$("TADEPOSI"))
   If DP1% <= 0 Then
     GoTo 10
   End If
   '
   DPT1$ = FORMATNUM$(DP1%, "F3.0") + " " + ECOARCH$("TADEPOSI", Chr$(DP1%))
   If COMALTER%("Esta Seguro de Ajustar a cero\todos los Stocks del Depósito Nro. : " & DPT1$ & "\ al " & FA$ & " \\No, Cancelar\Si, Ajustar a Cero") = 2 Then
     '
     Call ABRESTOCKS
     SQLX$ = " SELECT * FROM MOVISTO "
     SQLX$ = SQLX$ + "WHERE CodiMovi = 'AJ'"
     SQLX$ = SQLX$ + "ORDER BY VAL(MID$(DoPriCor,4)) DESC"
     UNUM& = 0
     Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     On Error Resume Next
     MoviTemp.MoveFirst
     If Err < 1 Then
       UNUM& = XVALO(Mid$(MoviTemp!DoPriCor, 4))
     End If
     On Error GoTo 0
     NRO& = 1 + UNUM&
     MoviTemp.Close
     '
     For KI% = 1 To NUMAS%
       CATEO#(KI%) = 0
     Next KI%
     '
     Call ABRESTOCKS
     '
     CMOV$ = TRIM$(CONTROL$("", "CODAJUST"))
     If CMOV$ = "" Then CMOV$ = "AJ"
     DOSECO$ = "AJ." + Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Left$(HOS$, 2)
     DOPRICO$ = "AS." + TRIM$(Str$(NRO&))
     DOPRILA$ = DOPRICO$
     DOSELA$ = DOSECO$
     '
     Call CARDEPOS
     DEPOPRIN% = 0
     If DP1% = COSU%(1) Then DEPOPRIN% = 1
     '
     '\\\OT-06-0353-MH-29/08/06///
     OPXX% = ALTERNA%("Solo Articulos De Baja\Todos Los Articulos")
     '
    For CI% = 1 To NUMAS%
       '
       If OPXX% = 1 Then
         If STATITEM%(CI%) <> (-1) Then
           GoTo 39
         End If
       End If
     '\\\OT-06-0353-MH-29/08/06///
       '
       Screen.MousePointer = 11
       Call VESARRAS(CI%)
       '
       SQLX$ = " SELECT * FROM MOVISTO "
       SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
       SQLX$ = SQLX$ + "AND INT(CDBL(FechMov))<= " & CDbl(CVDAT(FA1%)) & " "
       If DEPOPRIN% = 1 Then
           SQLX$ = SQLX$ + "AND (DepoMovi = " & DP1% & " OR DepoMovi = 0)" & " "
         Else
           SQLX$ = SQLX$ + "AND DepoMovi = " & DP1% & " "
       End If
       '
       Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
       With MoviTemp
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
           Do While Not .EOF
             CANTI1# = !CANTINGRE - !CANTSALID
             CATEO#(CI%) = CATEO#(CI%) + CANTI1#
             .MoveNext
           Loop
         End If
       End With
       '
       If Abs(CATEO#(CI%)) < 0.001 Then GoTo 30
       'NUORIT% = 0
       NUORIT% = NUORIT% + 1
       VUNI# = 0
       MONEII% = 0
       NC% = 0
       DEPO% = DP1%
       CANTI = CATEO#(CI%) * (-1)
       REFE$ = "Ajuste de Stocks a Cero"
       '
       Call MOVISTOK(FA1%, CI%, LTX$, CMOV$, DOPRILA$, DOPRICO$, DOSELA$, DOSECO$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPO%, CANTI)
       '
30     IU& = XVALO(CI%)
       FIN& = XVALO(NUMAS%)
       Call TRACE(24, IU&, FIN&)
       'Screen.MousePointer = 1
39  Next CI%
    '
   Else
     Screen.MousePointer = 1
     GoTo 20
   End If
   '
   Call COMUNI("Proceso Completo")
   Screen.MousePointer = 1
99 Command33.Enabled = True
'\\\OT-05-0704-WAR-FIN///
End Sub

Private Sub Command4_Click()
    '
    Call SELECHO("MASTER")
    COD$ = VALACT1$("MASTER")
    Text1.TEXT = COD$
    CI% = CODINT%(COD$)
    Text3.TEXT = DESCRIT$(CI%)
    Text2.TEXT = UNIMED$(CI%)
    '
End Sub
'

Private Sub Command5_Click()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text10.TEXT = VDEV$
    End If
End Sub



Private Sub Command7_Click()
    '
    DESNUM& = 1: If XVALO(Text14.TEXT) > 0 Then DESNUM& = XVALO(Text14.TEXT)
    DESNUX$ = "AS." + TRIM$(Str$(DESNUM&))
    HASNUM& = 999999
    If TRIM$(Text13.TEXT) <> "" Then
        If XVALO(Text13.TEXT) >= DESNUM& Then HASNUM& = XVALO(Text13.TEXT)
    End If
    HASNUX$ = "AS." + TRIM$(Str$(HASNUM&))
    DESFE% = 0: If FECHANUM(TRIM$(Text11.TEXT)) > 0 Then DESFE% = FECHANUM(TRIM$(Text11.TEXT))
    HASFE% = 32767
    If TRIM$(Text12.TEXT) <> "" Then
        If FECHANUM(TRIM$(Text12.TEXT)) >= DESFE% Then HASFE% = FECHANUM(TRIM$(Text12.TEXT))
    End If
    '
    Screen.MousePointer = 11
    Call BLANARCH("LISTRAJ")
    '
    FEC1# = CDbl(CVDAT(DESFE%))
    FEC2# = CDbl(CVDAT(HASFE%))
    '
    Call ABRESTOCKS
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE CodiMovi =  'AJ' "
    SQLX$ = SQLX$ + " and DoPriCor >= '" & DESNUX$ & "' "
    SQLX$ = SQLX$ + " AND DoPriCor <= '" & HASNUX$ & "' "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) >= " & FEC1# & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & FEC2# & " "
    SQLX$ = SQLX$ + "ORDER BY VAL(MID$(DoPriCor,4,7)) ASC"
    '
    Set AJUSMOVTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With AJUSMOVTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No Existen Ajustes Registrados segun Filtro")
        On Error GoTo 0
        Screen.MousePointer = 1
        Exit Sub
      End If
      On Error GoTo 0
      Do While .EOF = False
      REFE$ = !ReferCor
      CI1% = !COD_INTE
      DP1% = !DEPOMOVI
      CAN# = !CANTINGRE
      FF% = CVINT(!FECHMOV)
      NRO& = Mid(!DoPriCor, 4)
      y$ = REGBLAN$("LISTRAJ")
      Call REPLA(y$, MKI$(FF%), 1, 2)
      Call REPLA(y$, REFE$, 3, 48)
      Call REPLA(y$, MKS$(NRO&), 51, 4)
      Call REPLA(y$, MKI$(CI1%), 55, 2)
      Call REPLA(y$, UNIMED$(CI1%), 57, 4)
      If CAN# > 0 Then
          Call REPLA(y$, MKD$(CAN#), 61, 8)
          Call REPLA(y$, MKD$(!CANTSALID), 69, 8)
        Else
          Call REPLA(y$, MKD$(CAN#), 61, 8)
          Call REPLA(y$, MKD$(!CANTSALID), 69, 8)
      End If
      Call REPLA(y$, Chr$(DP1%), 80, 1)
      Call REGAPP("LISTRAJ", y$)
      
      .MoveNext
9      Loop
    End With
    '
    OPCI% = 1
    If Option4.Value = True Then OPCI% = 2
    If OPCI% = 1 Then
        'Call OPNOIN("")
        Call COMUNI("Solo Disponible para Impresión. Genere el \ Listado y Exportelo a un Archivo de Texto")
        GoTo 10
      ElseIf OPCI% = 2 Then
10      Call CONECRUN("*LIAJUST", "")
    End If
    '
    Screen.MousePointer = 1
    '
End Sub

Private Sub Exit_Click()
   Call Command1_Click
End Sub

Private Sub Form_Load()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture6.Visible = True
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6
      Me.Frame7.BackColor = &HFCD7B6
      Me.Frame9.BackColor = &HFCD7B6
      Me.Option4.BackColor = &HFCD7B6
      Me.Option5.BackColor = &HFCD7B6
      MARCOBARRA.Top = 6950
      Set MARCOBARRA.Container = Me.Picture6

    End If

    Call TRANSLABELS(Name)
    Call AJUSTO1
End Sub
'

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   On Error Resume Next
   Hide
   On Error GoTo 0
End Sub

Private Sub Help_Click()
   Call Command29_Click
End Sub

Private Sub IMPORTARINVENTARIOSTOCKS_Click()
    Call Command25_Click
End Sub

Private Sub Invent_Click()
   Call Command14_Click
End Sub

Private Sub Kardex_Click()
   Call Command15_Click
End Sub

Private Sub LiajuPan_Click()
   Option5.Value = True
   Call Command7_Click
End Sub

Private Sub LiajuPrin_Click()
   Option4.Value = True
   Call Command7_Click
End Sub

Private Sub List1_Click()
    '
    ILISTA& = List1.ListIndex
    Text4.TEXT = Left$(List1.TEXT, 20)
    Text6.TEXT = TRIM$(Mid$(List1.TEXT, 21))
    '
    Text7.TEXT = ""
    Text8.TEXT = ""
    Text9.TEXT = TRIM$(Mid$(List1.TEXT, 21))
    If XVALO(Text9.TEXT) < 0 Then
        Text7.TEXT = TRIM$(CSTRING$(MKD$((-1) * XVALO(Text9.TEXT)), 4, 10, 1, 2))
        Text9.TEXT = "0.0"
    End If
    Text9.SetFocus
End Sub

Private Sub NUaju_Click()
   Call Command4_Click
End Sub

Private Sub Reports_Click()
   Call Command13_Click
End Sub

Private Sub Setup_Click()
   Call Command26_Click
End Sub

Private Sub Text1_Change()
    '
    Text2.TEXT = ""
    Text3.TEXT = ""
    If CODINT%(Text1.TEXT) > 0 Then
        COD$ = TRIM$(Text1.TEXT)
        CI% = CODINT%(COD$)
        Text2.TEXT = UNIMED$(CI%)
        Text3.TEXT = DESCRIT$(CI%)
        Call AJUSTO2               ' recalcular existencias a fecha
    End If
    '
End Sub
'

Private Sub Text1_DblClick()
    '
    Call SELECHO("MASTER")
    COD$ = VALACT1$("MASTER")
    Text1.TEXT = COD$
    CI% = CODINT%(COD$)
    Text3.TEXT = DESCRIT$(CI%)
    Text2.TEXT = UNIMED$(CI%)
    '
End Sub
'

Private Sub Text1_LostFocus()
    COD$ = TRIM$(Text1.TEXT)
    CI% = CODINT%(COD$)
    Text3.TEXT = DESCRIT$(CI%)
    Text2.TEXT = UNIMED$(CI%)
End Sub

Private Sub Text10_Change()
    If Visible = True Then
        If FECHANUM(Text10.TEXT) > 0 Then
            Call AJUSTO2
        End If
    End If
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text10.TEXT = VDEV$
    End If
End Sub

Private Sub Text11_DblClick()
    Call SELECHO("SELFECHA")
    If VALACT1$("SELFECHA") <> "" Then
        Text11.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text12_DblClick()
    Call SELECHO("SELFECHA")
    If VALACT1$("SELFECHA") <> "" Then
        Text12.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text2_GotFocus()
    Text7.SetFocus

End Sub

Private Sub Text3_GotFocus()
    Text7.SetFocus

End Sub

Private Sub Text4_GotFocus()
    Text7.SetFocus

End Sub

Private Sub Text6_GotFocus()
    COD$ = Text1.TEXT
    If CODINT%(COD$) < 1 Then
        Text1.SetFocus
        Exit Sub
    End If
    ILISTA& = List1.ListIndex
    If ILISTA& < 0 Then
        List1.SetFocus
        Exit Sub
    End If
    Text9.SetFocus
End Sub
'

Private Sub Text7_Change()
    If MODENTRAJU% = 7 Then
        Text8.TEXT = ""
        CAJUS# = XVALO(Text6.TEXT) + XVALO(Text7.TEXT) - XVALO(Text8.TEXT)
        Text9.TEXT = Str$(CAJUS#)
    End If
End Sub
'

Private Sub Text8_Change()
    If MODENTRAJU% = 8 Then
        Text7.TEXT = ""
        CAJUS# = XVALO(Text6.TEXT) + XVALO(Text7.TEXT) - XVALO(Text8.TEXT)
        Text9.TEXT = Str$(CAJUS#)
    End If
End Sub
'

Private Sub Text9_Change()
    If MODENTRAJU% = 9 Then
       DIFER# = XVALO(Text9.TEXT) - XVALO(Text6.TEXT)
       Text7.TEXT = ""
       Text8.TEXT = ""
       If DIFER# > 0 Then
           Text7.TEXT = Str$(DIFER#)
         ElseIf DIFER# < 0 Then
           Text8.TEXT = Str$(Abs(DIFER#))
       End If
    End If
End Sub
'
Private Sub Text7_GotFocus()
    COD$ = Text1.TEXT
    If CODINT%(COD$) < 1 Then
        Text1.SetFocus
        Exit Sub
    End If
    ILISTA& = List1.ListIndex
    If ILISTA& < 0 Then
        List1.SetFocus
        Exit Sub
    End If
    Text7.SelStart = 0
    Text7.SelLength = Len(Text7.TEXT)
    MODENTRAJU% = 7
End Sub

Private Sub Text8_GotFocus()
    COD$ = Text1.TEXT
    If CODINT%(COD$) < 1 Then
        Text1.SetFocus
        Exit Sub
    End If
    ILISTA& = List1.ListIndex
    If ILISTA& < 0 Then
        List1.SetFocus
        Exit Sub
    End If
    Text8.SelStart = 0
    Text8.SelLength = Len(Text8.TEXT)
    MODENTRAJU% = 8
End Sub
'

Private Sub Text9_GotFocus()
    COD$ = Text1.TEXT
    If CODINT%(COD$) < 1 Then
        Text1.SetFocus
        Exit Sub
    End If
    ILISTA& = List1.ListIndex
    If ILISTA& < 0 Then
        List1.SetFocus
        Exit Sub
    End If
    Text9.SelStart = 0
    Text9.SelLength = Len(Text9.TEXT)
    MODENTRAJU% = 9
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command3.SetFocus
    End If
End Sub

Sub AJUSTO1()
    '
    HOOI% = HOY(HOS$)
    '
    List1.Clear
    Call CARDEPOS
    '
111 If NUSU% < 1 Then
        Call MENSERR(24, "Falta Definir\Depósitos y Sucursales.")
        Hide
        Call FRESHALL
        PREPA% = 0
        Exit Sub
    End If
    '
115 Text10.TEXT = HOS$
    Text5.TEXT = "Diferencia de Inventario"
    '
End Sub
'

Sub AJUSTO2()
    '
    FF% = FECHANUM(Text10.TEXT)
    If FF% < 1 Then
        Text10.TEXT = HOS$
        Text10.SetFocus
        Exit Sub
    End If
    '
    CI% = CODINT%(Text1)
    If CI% > 0 Then
      '
      Screen.MousePointer = 11
      Call VESARRAS(CI%)
      For KI% = 1 To 64
        CATEO#(KI%) = 0
      Next KI%
      '
      Call ABRESTOCKS
      SQLX$ = " SELECT * FROM MOVISTO "
      SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
      SQLX$ = SQLX$ + "AND INT(CDBL(FechMov))<= " & CDbl(CVDAT(FF%)) & " "
      '
      Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
      With MoviTemp
        On Error Resume Next
        .MoveFirst
        If Err < 1 Then
          Do While Not .EOF
            CANTI# = !CANTINGRE - !CANTSALID
            CANTIS = !CANTINGRE - !CANTSALID
            dep% = !DEPOMOVI
            LOT$ = !IdenLote
            '
            COLUX% = 1
            For KI% = 1 To NUSU%
              If dep% = COSU%(KI%) Then COLUX% = KI%
            Next KI%
            '
            CATEO#(COLUX%) = CATEO#(COLUX%) + CANTIS
          .MoveNext
          Loop
        End If
      End With
    End If
    '
    List1.Clear
    For U& = 1 To NUSU%
       CAN = CATEO#(U&)
       List1.AddItem SUCU$(U&) + CSTRING$(MKS$(CAN), 4, 10, 1, 2)
    Next U&
    '
    List1.ListIndex = ILISTA&
    List1.Refresh
    Text4.TEXT = Left$(List1.TEXT, 20)
    Text6.TEXT = TRIM$(Mid$(List1.TEXT, 21))
    Text7.TEXT = ""
    Text8.TEXT = ""
    Text9.TEXT = TRIM$(Mid$(List1.TEXT, 21))
    If XVALO(Text9.TEXT) < 0 Then
        Text7.TEXT = TRIM$(CSTRING$(MKD$((-1) * XVALO(Text9.TEXT)), 4, 10, 1, 2))
        Text9.TEXT = "0.0"
    End If
    Text9.SetFocus
    '
    Screen.MousePointer = 1
    '
End Sub
'

Sub AJUSTO3()
    '
    Static YAFORI%, FORIPRE$   ', UREMOVI&
    '
'    If UREMOVI& < 1 Then
'       UREMOVI& = ULTREG&("MOVISTO")
'       Call CIERRARCH("MOVISTO")
'    End If
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORAJUST"))
    YAFORI% = 1
    If FORIPRE$ <> "" Then
      LU$ = LUDAT$
      If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
        RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
        Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Vale de Ajuste de Stocks:\  \'" + RMCC$ + "'.")
        YAFORI% = (-1)
      End If
    End If
    '
    HOII% = HOY(HOS$)
    '
    Call ABRESTOCKS
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + "WHERE CodiMovi = 'AJ'"
    '\\\OT-05-0303-WAR-19/05/05///
    SQLX$ = SQLX$ + "ORDER BY VAL(MID$(DoPriCor,4)) DESC"
    '\\\OT-05-0303-WAR-FIN///
    UNUM& = 0
    Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    On Error Resume Next
    MoviTemp.MoveFirst
    If Err < 1 Then
      UNUM& = XVALO(Mid$(MoviTemp!DoPriCor, 4))
    End If
    On Error GoTo 0
    NRO& = 1 + UNUM&
    MoviTemp.Close
    '
    CAVIE = XVALO(Text6.TEXT)
    CANUE = XVALO(Text9.TEXT)
    '
    FF% = FECHANUM(Text10.TEXT)
    If FF% < 1 Then
      Call MENSERR(24, "Fecha Incorrecta")
      Text10.SetFocus
      Exit Sub
    End If
    '
    CIII% = CODINT%(Text1.TEXT)
    If CIII% < 1 Then
      Call MENSERR(24, "Falta Código de Material")
      Text1.SetFocus
      Exit Sub
    End If
    '
    LOTE$ = Label17
    LTX$ = LOTE$
    If Left$(UCase$(LOTE$), 5) = "TODOS" Or TRIM$(LOTE$) = "" Then
      LTX$ = Space$(16)
    End If
    ILISTA& = List1.ListIndex
    DEPO% = COSU%(ILISTA& + 1)
    FECHA% = FF%
    CANTI = CANUE - CAVIE
    '
    If Abs(CANTI) >= 0.05 Then
        '
        CMOV$ = TRIM$(CONTROL$("", "CODAJUST"))
        If CMOV$ = "" Then CMOV$ = "AJ"
        DOSECO$ = "AJ." + Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Left$(HOS$, 2)
        DOPRICO$ = "AS." + TRIM$(Str$(NRO&))
        NUORIT% = 0
        VUNI# = 0
        MONEII% = 0
        NC% = 0
        REFE$ = Text5.TEXT
        If TRIM$(REFE$) = "" Then
            REFE$ = "Ajuste de Inventario"
        End If
        '
        Call MOVISTOK(FF%, CIII%, LTX$, CMOV$, DOPRICO$, DOPRICO$, DOSECO$, DOSECO$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPO%, CANTI)
        '
        If EXISTE%("TRAZAB04.EXE") > 0 Then
           If Abs(CANTI) >= 0.05 Then
             IDLOTXX$ = DOPRICO$
             If TRIM$(LTX$) <> "" Then IDLOTXX$ = LTX$
             'Call CREALOTIN(FF%, CIII%, IDLOTXX$, CDbl(CANTI))
           End If
           Call ACSALOTE(CIII%, IDLOTXX$, "NOMESS")
        End If
        '
237     If TRIM$(FORIPRE$) <> "" Then
          If YAFORI% > 0 Then
             LU$ = LUDAT$
             CODFOR$ = TRIM$(UCase$(FORIPRE$))
             '
             YAC = XVALO(ATRIFORM$(CODFOR$, "DETDESMM"))
             MARIZW = XVALO(ATRIFORM$(CODFOR$, "MARIZQMM"))
             MARSUW = XVALO(ATRIFORM$(CODFOR$, "MARSUPMM"))
             '
             INTERLI = XVALO(ATRIFORM$(CODFOR$, "INTERLIN"))
             If INTERLI < 1 Then
                Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
                INTERLI = 1
             End If
             '
             CACOPI% = XVALO(ATRIFORM$(CODFOR$, "CAN-COP"))
             If CACOPI% < 1 Or CACOPI% > 6 Then
                Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
                CACOPI% = 1
             End If
             '
             For COP% = 1 To CACOPI%
               '
               Call CARFORWIN
               '
               FEX = FECHA%
               Call PRINTFORWIN(FECHATEX$(FEX), "FECH-EMI", 0, 0)
               Call PRINTFORWIN(Time$, "HORA-EMI", 0, 0)
               Call PRINTFORWIN(USERNAME$, "USUARIO", 0, 0)
               Call PRINTFORWIN(Str$(NRO&), "NUME-COM", 0, 0)
               Call PRINTFORWIN(TRIM$(CODEXT$(CIII%)), "COD-MAT", 0, 0)
               Call PRINTFORWIN(DESCRIT$(CIII%), "DES-MAT", 0, 0)
               Call PRINTFORWIN(UNIMED$(CIII%), "UNIMED", 0, 0)
               Call PRINTFORWIN(ECOARCH$("ECODEP", Chr$(DEPO%)), "DEPOSITO", 0, 0)
               Call PRINTFORWIN(LTX$, "NRO-LOTE", 0, 0)
               CANS$ = CSTRING$(MKS$(CANTI), 3, 9, 1, 2)
               Call PRINTFORWIN(CANS$, "CANTIDAD", 0, 0)
               Call PRINTFORWIN(REFE$, "REF-MAT1", 0, 0)
               '
               'Printer.EndDoc
               Call SETSPOOL("ENDDOC", "")
               '
             Next COP%
             Call SAVESPOOL("Ajuste de Stocks" + Str$(NRO&), OKX%)
          End If
        End If
        '
        Call COMUNI("Ajuste Número" + Str$(NRO&) + "\Completo y Verificado")
        '
299 End If
    '
    Call CIERRARCH("*.*")
    '
End Sub
