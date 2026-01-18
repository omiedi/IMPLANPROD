VERSION 5.00
Begin VB.Form MODIPAMOD04 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modificación Parte Diario  M.O.D."
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Nuevos Valores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3690
      Left            =   210
      TabIndex        =   36
      Top             =   4515
      Width           =   7470
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
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
         Left            =   6300
         TabIndex        =   14
         Top             =   1995
         Width           =   960
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         Left            =   6300
         TabIndex        =   12
         Top             =   420
         Width           =   960
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "At.Múltiple"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1995
         TabIndex        =   17
         Top             =   1260
         Width           =   1275
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Re-Trabajo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   3780
         TabIndex        =   18
         Top             =   1260
         Width           =   1275
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
         Left            =   735
         TabIndex        =   6
         Text            =   " "
         Top             =   420
         Width           =   960
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
         Left            =   1680
         TabIndex        =   7
         Top             =   420
         Width           =   175
      End
      Begin VB.CommandButton Command4 
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
         Left            =   1575
         TabIndex        =   16
         Top             =   1335
         Width           =   175
      End
      Begin VB.TextBox Text1 
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
         Left            =   735
         TabIndex        =   15
         Top             =   1335
         Width           =   960
      End
      Begin VB.CommandButton Command7 
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
         Left            =   1680
         TabIndex        =   9
         Top             =   840
         Width           =   175
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
         Height          =   285
         Left            =   735
         TabIndex        =   8
         Top             =   840
         Width           =   960
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
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
         Left            =   2625
         TabIndex        =   10
         Text            =   " "
         Top             =   420
         Width           =   960
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
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
         Left            =   4305
         TabIndex        =   11
         Text            =   " "
         Top             =   420
         Width           =   960
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
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
         Left            =   6300
         TabIndex        =   13
         Top             =   1575
         Width           =   960
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Text            =   "MODIPAMOD04.frx":0000
         Top             =   2650
         Width           =   6945
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No Conf:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5460
         TabIndex        =   60
         Top             =   2070
         Width           =   810
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.OK:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4095
         TabIndex        =   59
         Top             =   1650
         Width           =   2175
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
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
         Left            =   6300
         TabIndex        =   58
         Top             =   840
         Width           =   960
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hs.Oper:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4305
         TabIndex        =   57
         Top             =   915
         Width           =   1965
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hs.Desc:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5355
         TabIndex        =   56
         Top             =   500
         Width           =   915
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         TabIndex        =   44
         Top             =   500
         Width           =   705
      End
      Begin VB.Label Label5 
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
         Height          =   555
         Left            =   735
         TabIndex        =   43
         ToolTipText     =   "Haga Doble-Click para Informe por Célula"
         Top             =   1680
         Width           =   4575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Legajo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   -210
         TabIndex        =   42
         Top             =   1335
         Width           =   915
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Equipo:"
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
         Left            =   -130
         TabIndex        =   41
         Top             =   920
         Width           =   855
      End
      Begin VB.Label Label10 
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
         Left            =   1995
         TabIndex        =   40
         Top             =   840
         Width           =   3270
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Com:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1995
         TabIndex        =   39
         Top             =   500
         Width           =   600
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fin:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3795
         TabIndex        =   38
         Top             =   500
         Width           =   495
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         TabIndex        =   37
         Top             =   2350
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Selección del Parte a Modificar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3480
      Left            =   210
      TabIndex        =   28
      Top             =   840
      Width           =   7470
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Left            =   105
         TabIndex        =   35
         Top             =   1995
         Width           =   7260
      End
      Begin VB.CommandButton Command3 
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
         Left            =   7035
         TabIndex        =   3
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text3 
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
         Left            =   1260
         TabIndex        =   4
         Top             =   1260
         Width           =   630
      End
      Begin VB.CommandButton Command6 
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
         Left            =   1890
         TabIndex        =   5
         Top             =   1260
         Width           =   175
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
         Left            =   1260
         TabIndex        =   1
         Top             =   420
         Width           =   2010
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
         Left            =   3255
         TabIndex        =   2
         Top             =   420
         Width           =   175
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Operario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5775
         TabIndex        =   53
         Top             =   1740
         Width           =   810
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4830
         TabIndex        =   52
         Top             =   1740
         Width           =   810
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Maquina"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3875
         TabIndex        =   51
         Top             =   1740
         Width           =   810
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "H. Fin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2625
         TabIndex        =   50
         Top             =   1740
         Width           =   810
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "H. Ini"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1545
         TabIndex        =   49
         Top             =   1740
         Width           =   810
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   840
         TabIndex        =   48
         Top             =   1740
         Width           =   600
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " OT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -315
         TabIndex        =   47
         Top             =   1740
         Width           =   915
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   6735
         TabIndex        =   46
         Top             =   1260
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N-Ord:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5355
         TabIndex        =   45
         Top             =   1335
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   1260
         TabIndex        =   34
         Top             =   840
         Width           =   5970
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5370
         TabIndex        =   33
         Top             =   420
         Width           =   1665
      End
      Begin VB.Label Label6 
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
         Height          =   300
         Left            =   -105
         TabIndex        =   32
         Top             =   495
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2205
         TabIndex        =   31
         Top             =   1260
         Width           =   3690
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Operación:"
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
         Left            =   -1155
         TabIndex        =   30
         Top             =   1320
         Width           =   2325
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "O/F Nro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3990
         TabIndex        =   29
         Top             =   495
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00004000&
      Height          =   8460
      Left            =   7875
      ScaleHeight     =   8400
      ScaleWidth      =   1740
      TabIndex        =   23
      Top             =   0
      Width           =   1800
      Begin VB.CommandButton Command9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "MODIPAMOD04.frx":0002
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Borra (Elimina) el Parte Activo"
         Top             =   4620
         Width           =   650
      End
      Begin VB.CommandButton Command16 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "MODIPAMOD04.frx":030C
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Listado de Informe de Produccion por Fecha"
         Top             =   1785
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
         Height          =   1350
         Left            =   105
         Picture         =   "MODIPAMOD04.frx":0616
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Aprueba y Graba los Nuevos Valores"
         Top             =   6800
         Width           =   650
      End
      Begin VB.CommandButton Command14 
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
         Height          =   640
         Left            =   105
         Picture         =   "MODIPAMOD04.frx":0920
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   3465
         Width           =   650
      End
      Begin VB.CommandButton Command19 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "MODIPAMOD04.frx":0D62
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "A-B-M Padrón de Operarios"
         Top             =   2820
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   24
         Top             =   5775
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   12000
         End
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
         Height          =   640
         Left            =   105
         Picture         =   "MODIPAMOD04.frx":106C
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Salir"
         Top             =   200
         Width           =   650
      End
      Begin VB.CommandButton Command29 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "MODIPAMOD04.frx":11B6
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   845
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "MODIPAMOD04.frx":14C0
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   6075
         Width           =   1515
      End
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "Corrección Parte Diario M.O.D."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   435
      Left            =   105
      TabIndex        =   26
      Top             =   210
      Width           =   5475
   End
End
Attribute VB_Name = "MODIPAMOD04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   STPPMOD04.Show 1
   Call COMUNI("Re-Inicie Ahora la Aplicación.")
   Call FINAL("")
   Command14.Enabled = True
End Sub

Private Sub Command16_Click()
   Command16.Enabled = False
   Call INFOFA04.Command26_Click
   Command16.Enabled = True
End Sub

Private Sub Command19_Click()
   Command19.Enabled = False
   Call INFOFA04.Command19_Click
   AppActivate Caption
   SendKeys Chr$(1)
   Command19.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   'VALIDACIONES
   HOII% = HOY(HOS$)
   If FECHANUM(Text4) < 33 Or FECHANUM(Text4) > HOII% Then
5     Call MENSERR(24, "Fecha no Válida")
      Text4.SetFocus
      GoTo 99
   End If
   '
   On Error Resume Next
   ABC12 = CDate(Text4)
   If Err Then
     On Error GoTo 0
     GoTo 5
   End If
   On Error GoTo 0
   '
   LEGAI% = XVALO(Text1)
   If LEGAI% > 0 Then
        LEGACELU$ = MKI$(LEGAI%)
      Else
        LEGACELU$ = ""
        TCELU$ = Label5 + "/"
8       PPXX% = InStr(TCELU$, "/")
        If PPXX% > 0 Then
          LEGAL& = XVALO(Left$(TCELU$, PPXX% - 1))
          TCELU$ = Mid$(TCELU$, PPXX% + 1)
          If LEGAL& > 0 Then
            If LEGAL& <= 32767 Then
              LEGAI% = LEGAL&
              If ECOARCH$("PADRON", MKI$(LEGAI%)) <> "" Then
                For KKII% = 1 To Len(LEGACELU$) Step 2
                  If CVI(Mid$(LEGACELU$, KKII%, 2)) = LEGAI% Then GoTo 8
                Next KKII%
                LEGACELU$ = LEGACELU$ + MKI$(LEGAI%)
                GoTo 8
              End If
            End If
          End If
          Call MENSERR(24, "Legajos No Válidos en Componentes de Célula")
        End If
   End If
   LEGAI% = CVI(Left$(LEGACELU$, 2))
   If LEGAI% < 1 Or ECOARCH$("PADRON", MKI$(LEGAI%)) = "" Or LEGACELU$ = "" Then
     Call MENSERR(24, "Operario / Célula Inexistente o no Válido")
     Text1.SetFocus
     GoTo 99
   End If
   '
   CID1% = CODINT%(Text10)
   If CID1% <= 0 Then
     Call MENSERR(24, "Código no Válido")
     Text10.SetFocus
     GoTo 99
   End If
   '
   If TRIM$(Label18) = "" Then
      Call MENSERR(24, "Debe Seleccionar Parte a Corregir")
      GoTo 99
   End If
   '
   If TRIM$(Label4) <> "" Then
     Call ABREORFAB
     OrdeFabr.FindFirst "IdSubOr = '" & TRIM$(Label4) & "' "
     If OrdeFabr.NoMatch = False Then
       If OrdeFabr!Cod_Inte = CID1% Then
         GoTo 19    ' VALIDA NUMERO DE o/f Y CODIGO
       End If
     End If
     Call MENSERR(24, "O/F no Corresponde a Código de Artículo")
     GoTo 99
   End If
   '
19 NOPXX$ = Text3
   PPXX% = InStr(NOPXX$, "/")
   If PPXX% < 1 Then PPXX% = 1 + Len(NOPXX$)
   NOPEX% = 0: ALTEX% = 0
   On Error Resume Next
   NOPEX% = XVALO(Left$(NOPXX$, PPXX% - 1))
   ALTEX% = XVALO(Mid$(NOPXX$, PPXX% + 1))
   On Error GoTo 0
   '
   If NOPEX% < 1 Or DESCRIOP$(CID1%, NOPEX%, ALTEX%) = "" Then
     Call MENSERR(24, "Operación no Válida")
     Text3.SetFocus
     GoTo 99
   End If
   '
   MAQUI1$ = Text5
   If ECOARCH$("PADMAQ", MAQUI1$) = "" Then
     Call MENSERR(24, "Máquina / Equipo no Válido")
     Text5.SetFocus
     GoTo 99
   End If
   '
   If XVALO(Text8.TEXT) < 0.005 Then
     Call MENSERR(24, "Cantidad no Válida")
     Text8.SetFocus
     GoTo 99
   End If
   '
   TDES% = HORANUM(Text6)
   THAS% = HORANUM(Text7)
   TFRI% = HORANUM(Text2)
   If TRIM$(Text6) = "" Or TRIM$(Text7) = "" Or THAS% < TDES% Then
     Call MENSERR(24, "Horarios de Comienzo y Fin no Válidos")
     Text6.SetFocus
     GoTo 99
   End If
   '
   'GRABAR INFORME
   NROOF$ = Label4
   CIXXI% = CODINT%(Text10)
   NOPE% = NOPEX%
   CAINFANT% = XVALO(Label18)
   '
   CAINF# = XVALO(Text8)
   CANOC# = XVALO(Text11)
   CAREAL# = CAINF#
   MAQUI$ = Text5
   TIEINF = 60 * (DIFHORA(TDES%, THAS%) - DIFHORA(0, TFRI%))
   If TIEINF < 0.01 Then
     Call MENSERR(24, "Horarios / Tiempos Informados no Válidos.")
     Text6.SetFocus
     GoTo 99
   End If
   '
   SQLX$ = " SELECT * FROM INFOFAB "
   SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
   SQLX$ = SQLX$ + " AND NumeOper = " & NOPE% & " "
   '
   Set INFOFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
   With INFOFABTMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
         If !NuOrInf <> CAINFANT% Then
           If IsNull(!ReTraba) Or !ReTraba < 1 Then
             CAREAL# = CAREAL# + !CantReal
           End If
         End If
       .MoveNext
       Loop
     End If
   End With
   '
   ABUS$ = "IdSubOr = '" & NROOF$ & "' "
   ABUS$ = ABUS$ + "AND Cod_Inte = " & CIXXI% & " "
   ABUS$ = ABUS$ + "AND NumeOper = " & NOPE% & " "
   ABUS$ = ABUS$ + "AND NuOrInf = " & CAINFANT% & " "
   
   'Call ABREORFAB
   With InfoFab
     .FindFirst ABUS$
     If .NoMatch = False Then
         .Edit
         !COD_EXTE = CODEXT$(CIXXI%)
         !Descrip = Left$(DESCRIT$(CIXXI%), 64)
           MARETRA% = 0: If Check1.Value = 1 Then MARETRA% = 1
         !ReTraba = MARETRA%
           ATENMUL% = 0: If Check2.Value = 1 Then ATENMUL% = 1
         !AtMulti = ATENMUL%
         !CantReal = CAINF#
         !CaNoConf = CANOC#
         !Maquina = Left$(MAQUI$, 6)
         !DescMaq = Left$(ECOARCH$("PADMAQ", MAQUI$), 64)
         !LegaOps = LEGACELU$
            FEX = FECHANUM(Text4)
            THAS1% = THAS%: If THAS1% > 2359 Then THAS1% = 0
            FEHOLI = CDate(FECHATEX$(FEX) + " " + HORATEXNOR$(Int(THAS1%)))
         !FeHorLib = FEHOLI
         !Hora_Des = TDES%
         !Hora_Has = THAS%
         !Hora_Fri = TFRI%
         !TiemEjec = TIEINF
            MODO% = CAINFANT%
         !TiemStan = TIESTAN(CIXXI%, NOPE%, CAINF#, MODO%, ALTEX%)  ' 1 o mas es para que no considere preparacion
              'COMENTE LA SIG. LINEA XQ LE FALTABA HACER REFERENCIA
              ' A LOS CAMPOS DEL REGISTRO. "!"
              'RENDI = 999.9: If !TiemStan > 0 Then RENDI = 100 * TiemEjec / TiemStan
            RENDI = 999.9: If !TiemStan > 0 Then RENDI = 100 * !TiemEjec / !TiemStan
            If RENDI > 999.9 Then RENDI = 999.9
         !RENDIM = RENDI
         !Observa = AJUSTI$(TRIM$(Text9), 30)
         !Referen = Text9
         .Update
       Else
         Call MENSERR(24, "No Fue Posible Actualizar el Informe.")
         GoTo 90
     End If
   End With
   '
   SQLX$ = " SELECT * FROM INFOFAB "
   SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
   SQLX$ = SQLX$ + " AND NumeOper = " & NOPE% & " "
   '
   CAREAL# = 0: TIEINF = 0
   Set INFOFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
   With INFOFABTMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
         If IsNull(!ReTraba) Or !ReTraba < 1 Then
           CAREAL# = CAREAL# + !CantReal
           TIEINF = TIEINF + !TiemEjec
         End If
       .MoveNext
       Loop
     End If
   End With
   '
   With OperFab
     .FindFirst "IdSubOr = '" & NROOF$ & "' AND NumeOper = " & NOPE%
     If .NoMatch = False Then
       .Edit
       !CantReal = CAREAL#
       !CantPend = !CantProg - CAREAL#
         If !CantPend < 0 Or STATUOP% = 3 Or STATUOP% = 8 Then !CantPend = 0
         CANTPE# = !CantPend
         MODO% = CAINFANT%
       !TiemEjec = TIEINF
       !TiemPend = TIEMAQUI(CIXXI%, NOPE%, CANTPE#, MODO%)  ' 1 es para que no considere preparacion
           If !TiemPend < 0 Or STATUOP% = 3 Or STATUOP% = 0 Then !TiemPend = 0
       !StatOper = STATUOP%
       .Update
     End If
   End With
   '
   If NOPE% = ULTIMOP%(CID1%) Then
        Call COMUNI("Ha Modificado el Parte de Fabricacion\Correspondiente a la Ultima Operacion.\  \Recuerde Cerrar Manualmente la Orden de Fabricacion\ Para Producir el Correcto Movimiento de Stock.")
   End If
         
   '
90 Text10 = ""
   Text1 = ""
   Label5 = ""
   Text2 = ""
   Text3 = ""
   Text4 = ""
   Text6 = ""
   Text7 = ""
   Text8 = ""
   Text9 = ""
   Text11 = ""
   Label18 = ""
   Label25 = ""
   '
   Text10.SetFocus
   '
99 Command2.Enabled = True
   '
End Sub



Private Sub Command29_Click()
  '
  Call OPNOIN("Ayuda FLEXOFT en Línea")
  '
End Sub

Private Sub Command3_Click()
   '
   NROOF$ = ""
'   Call INFOFA04.GENRESDAT
'   '
'   Call FRESHECHO("ECORDEP")
'   If ULTREG&("ECORDEP") < 1 Then
'     Call MENSERR(24, "No hay Ordenes de Fabricación Pendientes.")
'     Exit Sub
'   End If
'   '
   Call SELECHO("!ECORDEP/Ordenes de Trabajo en Curso")
   NROOF$ = AJUSTI$(VALACT1$("!ECORDEP"), 12)
   If TRIM$(NROOF$) = "" Then
     Label4 = ""
     Exit Sub
   End If
   '
   Label4 = NROOF$
   '
End Sub

Private Sub Command4_Click()
  '
  Call SELECHO("PADRON")
  VDEV$ = VALACT1$("PADRON")
  If VDEV$ <> "" Then
    Text1.TEXT = VDEV$
  End If
  '
End Sub

Private Sub Command5_Click()
  '
  Call SELECHO("SELFECHA")
  VDEV$ = VALACT1$("SELFECHA")
  If VDEV$ <> "" Then
    Text4.TEXT = VDEV$
  End If
  '
End Sub

Private Sub Command6_Click()
  '
  Call SELECHO("!OPERINF/Operaciones de Codigo: " + Text10)
  CODOP1$ = VALACT1$("!OPERINF")
  If CODOP1$ <> "" Then
    Y$ = FILTERGETRECORD$("!OPERINF", 1, MKI$(XVALO(CODOP1$)))
    CODOP1$ = Mid$(Y$, 3, 58)
    I% = InStr(1, CODOP1$, "-")
    CODOP2$ = TRIM$(Mid$(Y$, 1, I%))
    Text3.TEXT = CODOP2$
  End If
  '
End Sub

Private Sub Command7_Click()
   '
   Call SELECHO("PADMAQ")
   VDEVU$ = VALACT1$("PADMAQ")
   If VDEVU$ <> "" Then
      Text5.TEXT = VDEVU$
   End If
   '
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   Call SELECHO("MASTER")
   If VALACT1$("MASTER") <> "" Then
     Text10 = VALACT1$("MASTER")
   End If
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   '
   Command9.Enabled = False
   If DERACCE%("ANUPADIA", "Anulación Parte Diario de Mano de Obra") >= 2 Then
     '
     If TRIM$(Label18) = "" Then
        Call MENSERR(24, "Debe Seleccionar Parte a Eliminar")
        GoTo 99
     End If
     '
     If TRIM$(Label4) <> "" Then
       CIXXI% = CODINT%(Text10)
       Call ABREORFAB
       OrdeFabr.FindFirst "IdSubOr = '" & TRIM$(Label4) & "' "
       If OrdeFabr.NoMatch = False Then
         If OrdeFabr!Cod_Inte = CIXXI% Then
           GoTo 19    ' VALIDA NUMERO DE o/f Y CODIGO
         End If
       End If
       Call MENSERR(24, "O/F no Corresponde a Código de Artículo")
       GoTo 99
     End If
     '
19   NOPXX$ = Text3
     PPXX% = InStr(NOPXX$, "/")
     If PPXX% < 1 Then PPXX% = 1 + Len(NOPXX$)
     NOPEX% = 0: ALTEX% = 0
     On Error Resume Next
     NOPEX% = XVALO(Left$(NOPXX$, PPXX% - 1))
     ALTEX% = XVALO(Mid$(NOPXX$, PPXX% + 1))
     On Error GoTo 0
     '
     If NOPEX% < 1 Or DESCRIOP$(CIXXI%, NOPEX%, ALTEX%) = "" Then
       Call MENSERR(24, "Operación no Válida")
       Text3.SetFocus
       GoTo 99
     End If
     '
     If COMALTER%("¿ Está Seguro de ANULAR el Parte Activo ?\Esta Transacción NO Podrá Revertirse.\\No, Conservar\Si, Anular") <> 2 Then
       GoTo 99
     End If
     '
     'GRABAR INFORME
     NROOF$ = Label4
     CIXXI% = CODINT%(Text10)
     NOPE% = NOPEX%
     CAINFANT% = XVALO(Label18)
     '
     CAINF# = 0
     CAREAL# = CAINF#
     MAQUI$ = ""
     TIEINF = 0
     '
     If TRIM$(NROOF$) <> "" Then
        Call ABREORFAB
        ABUS$ = "IdSubOr = '" & NROOF$ & "' "
        ABUS$ = ABUS$ + "AND Cod_Inte = " & CIXXI% & " "
        OrdeFabr.FindFirst ABUS$
        If OrdeFabr.NoMatch = False Then
          If OrdeFabr!Statuorf < 3 Then                       ' no marca CERRADAS Y cumplidos"
            GoTo 20
          End If
        End If
        Call MENSERR(24, "Imposible Anular Este Parte.\La Orden de Fabricación Está Cerrada, Cumplida o Anulada.")
        GoTo 99
     End If
     '
20   ABUS$ = "IdSubOr = '" & NROOF$ & "' "
     ABUS$ = ABUS$ + "AND Cod_Inte = " & CIXXI% & " "
     ABUS$ = ABUS$ + "AND NumeOper = " & NOPE% & " "
     ABUS$ = ABUS$ + "AND NuOrInf = " & CAINFANT% & " "
     
     'Call ABREORFAB
     With InfoFab
       .FindFirst ABUS$
       If .NoMatch = False Then
           .Delete
         Else
           Call MENSERR(24, "No Fue Posible Anular el Informe.")
           GoTo 90
       End If
     End With
     '
     SQLX$ = " SELECT * FROM INFOFAB "
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
     SQLX$ = SQLX$ + " AND NumeOper = " & NOPE% & " "
     '
     CAREAL# = 0: TIEINF = 0
     Set INFOFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     With INFOFABTMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           If IsNull(!ReTraba) Or !ReTraba < 1 Then
             CAREAL# = CAREAL# + !CantReal
             TIEINF = TIEINF + !TiemEjec
           End If
         .MoveNext
         Loop
       End If
     End With
     '
     With OperFab
       .FindFirst "IdSubOr = '" & NROOF$ & "' AND NumeOper = " & NOPE%
       If .NoMatch = False Then
         .Edit
         !CantReal = CAREAL#
         !CantPend = !CantProg - CAREAL#
           If !CantPend < 0 Or STATUOP% = 3 Or STATUOP% = 8 Then !CantPend = 0
           CANTPE# = !CantPend
           MODO% = CAINFANT%
         !TiemEjec = TIEINF
         !TiemPend = TIEMAQUI(CIXXI%, NOPE%, CANTPE#, MODO%)  ' 1 es para que no considere preparacion
             If !TiemPend < 0 Or STATUOP% = 3 Or STATUOP% = 0 Then !TiemPend = 0
         !StatOper = STATUOP%
         .Update
       End If
     End With
     '
90   Text10 = ""
     Text1 = ""
     Text3 = ""
     Text4 = ""
     Text6 = ""
     Text7 = ""
     Text8 = ""
     Text9 = ""
     Label18 = ""
     '
     Text10.SetFocus
     '
   End If
   '
99 Command9.Enabled = True
   '
End Sub

Sub Form_Load()
  '
  Screen.MousePointer = 11
  Call CENTRAFORM(Me)
  '
  Call BLANARCH("!OPERINF")
  Call SETULTREG("!OPERINF", 0)
  Call FRESHECHO("!OPERINF")
  '
  If CONTROL$("PADIAMOD", "INFDESCA") <> "SI" Then
      Label26.Visible = False
      Text2.Visible = False
      Label27.Caption = "Horas Operación:"
      Label27.Top = 600
      Label27.FontUnderline = True
      Label27.Left = 5000
  End If
  '
  If CONTROL$("PADIAMOD", "PINOCONF") <> "SI" Then
      Label33.Visible = False
      Text11.Visible = False
      Text8.Top = Text11.Top
      Label32.Top = 1740
      Label32.Caption = "Pzas./ Cantidad:"
      Label32.Left = 4700
      Label32.FontUnderline = True
  End If
  '
  Screen.MousePointer = 1
  '
End Sub

Private Sub Label4_Change()
    '
    Call GENECODOPEOF(Label4)
    Call GENLISOP
    '
End Sub

Private Sub Label5_DblClick()
    '
    If TRIM$(Label18) = "" Then
      Call MENSERR(24, "Debe Seleccionar Parte a Corregir")
      Exit Sub
    End If
    '
    JJ& = 0
    REBLA$ = REGBLAN$("DCELULAS")
    TCELU$ = TRIM$(Label5)
    If XVALO(Text1) > 0 Then TCELU$ = TRIM$(Text1)
    TCELU$ = TCELU$ + "/"
8   PPXX% = InStr(TCELU$, "/")
    If PPXX% > 0 Then
       LEGAL& = XVALO(Left$(TCELU$, PPXX% - 1))
       TCELU$ = Mid$(TCELU$, PPXX% + 1)
       If LEGAL& > 0 Then
         If LEGAL& <= 32767 Then
           LEGAI% = LEGAL&
           If ECOARCH$("PADRON", MKI$(LEGAI%)) <> "" Then
             YY$ = REBLA$
             Call REPLA(YY$, MKI$(LEGAI%), 3, 2)
             JJ& = JJ& + 1
             Call GRAREG("!DCELULAS", YY$, JJ&)
           End If
         End If
       End If
       GoTo 8
    End If
    Call SETULTREG("!DCELULAS", JJ&)
    Call CIERRARCH("!DCELULAS")
    '
    VTB% = VENTABU%("DCELULAS/NC/TITUPAN=Procesos por Célula ")
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    TTXX2$ = ""
    For I& = 1 To ULTREG&("!DCELULAS")
       X$ = REGLEIDO$("!DCELULAS", I&)
       If CVI(Mid$(X$, 3, 2)) > 0 Then
         NLEGA% = CVI(Mid$(X$, 3, 2))
         For JJ& = 1 To Len(TTXX1$) Step 2
           If CVI(Mid$(TTXX1$, JJ&, 2)) = NLEGA% Then GoTo 19
         Next JJ&
         Call REPLA(X$, MKI$(ICELU%), 1, 2)
         J& = J& + 1
         Call GRAREG("!DCELULAS", X$, J&)
         If TTXX2$ <> "" Then TTXX2$ = TTXX2$ + " / "
         TTXX2$ = TTXX2$ + TRIM$(Str$(NLEGA%))
       End If
19  Next I&
    Call SETULTREG("!DCELULAS", J&)
    Call CIERRARCH("!DCELULAS")
    If ULTREG&("!DCELULAS") <= 1 Then
         Text1 = TRIM$(TTXX2$)
       Else
         Text1 = ""
         Label5 = TTXX2$
    End If
    '
End Sub

Private Sub List1_Click()
   Label18 = ""
   Text4.TEXT = ""
   Text6.TEXT = ""
   Text7.TEXT = ""
   Text5.TEXT = ""
   Text8.TEXT = ""
   Text11 = ""
   Text1.TEXT = ""
   Text2 = ""
   Label25 = ""
   Check1.Value = 0
   Check2.Value = 0
   '
   XX$ = List1.TEXT
   FE = Mid(XX$, 7, 8)
   HD = AJUSTD$(Mid(XX$, 17, 8), 8)
   HH = AJUSTD(Mid(XX$, 27, 8), 8)
   HF = AJUSTD$(Mid$(XX$, 97, 8), 8)
   MA = Mid(XX$, 37, 6)
   '
   If MA = "Ajuste" Then
      Call MENSERR(24, "Imposible Corregir Parte de Ajuste")
      Exit Sub
   End If
   '
   CA = Mid(XX$, 45, 8)
   CANC = Mid$(XX$, 105, 8)
   LG = AJUSTI$(Mid(XX$, 70, 8), 8)
   NORI% = XVALO(Mid(XX$, 1, 4))
   Label18 = TRIM$(Str$(NORI%))
   MARETRA% = XVALO(Mid$(XX$, 79, 1))
   ATENMUL% = XVALO(Mid$(XX$, 80, 1))
   OBSX$ = Mid$(XX$, 257)
   '
   Text4.TEXT = FE
   Text6.TEXT = HD
   Text7.TEXT = HH
   Text2.TEXT = HF
   Text5.TEXT = MA
   Text8.TEXT = CA
   Text11.TEXT = CANC
   Text1.TEXT = LG
   Text9 = OBSX$
   '
   If XVALO(LG) < 1 Then
     Label5 = TRIM$(Mid$(XX$, 129, 128))
   End If
   '
   Check1.Value = MARETRA%
   Check2.Value = ATENMUL%
End Sub

Private Sub Text1_Change()
  '
  If XVALO(Text1) < 32767 Then
      Label5.Caption = ECOARCH$("PADRON", MKI$(XVALO(Text1.TEXT)))
    Else
      Call MENSERR(24, "Legajo no Válido.")
      Text1 = ""
  End If
  '
End Sub
'
Private Sub Text1_DblClick()
  '
  Call Command4_Click
  '
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Text10.SetFocus
End If
End Sub

Private Sub Text10_Change()
   Label3 = DESCRIT0$(CODINT%(Text10))
   '
   CIXXI% = CODINT%(Text10)
   Label4 = "": CANOR% = 0
   Text3 = "": Text5 = ""
   Call SETULTREG("!ECORDEP", 0)
   If CIXXI% > 0 Then
      Call ABREORFAB
      SQLX$ = " SELECT * FROM ORDEFABR "
      SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CIXXI% & " "
      SQLX$ = SQLX$ + " ORDER BY IdSubOr ASC "
      '
      Set ORFASEL = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
      '
      On Error Resume Next
      ORFASEL.MoveLast
      If Err Then
         Exit Sub
      End If
      On Error GoTo 0
      Screen.MousePointer = 11
      CANOR% = ORFASEL.RecordCount
      JJ& = 0
      With ORFASEL
        .MoveFirst
        Do While (.EOF = False)
          If HAYOPINF%(!IdSubOr) > 0 Then
             If Label4 = "" Then Label4 = ORFASEL!IdSubOr
             Y$ = REGBLAN("ECORDEP")
             Call REPLA(Y$, !IdSubOr, 1, 12)
             Call REPLA(Y$, ECOARCH$("ECORDEF", !IdSubOr), 13, 52)
             '
             JJ& = JJ& + 1
             Call GRAREG("!ECORDEP", Y$, JJ&)
          End If
          '
        .MoveNext
        Loop
      End With
      Call SETULTREG("!ECORDEP", JJ&)
      Call CIERRARCH("!ECORDEP")
      Call FRESHECHO("!ECORDEP")
      Screen.MousePointer = 1
   End If
   Call GENLISOP
   '
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
  '
  If KeyAscii = 13 Then
     Text1.SetFocus
  End If
  '
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Text9.SetFocus
   End If
End Sub

Private Sub Text11_LostFocus()
  Text11.TEXT = TRIM$(NUMSTRI$(XVALO(Text11.TEXT), 12, 1, 0))
End Sub

Private Sub Text2_Change()
  '
  Call CALCUTIEMPO
  '
End Sub

Private Sub TEXT2_DBLCLICK()
  '
  'Call Command3_Click
  '
End Sub
Function CARPRODOF$(NROOF$)
  '
  CARPRODOF$ = ""
  Call ABREORFAB
  '
  SQLX$ = " SELECT * FROM OPERFAB "
  SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
  Set OPFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
  '
  With OPFABTMP
    '
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      Exit Function
    End If
    On Error GoTo 0
    '
    CI% = !Cod_Inte
    CARPRODOF$ = CODEXT$(CI%)
  End With
  '
End Function

Private Sub Text10_DblClick()
   Call Command8_Click
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Text3.SetFocus
End If


End Sub

Private Sub Text2_GotFocus()
  '
  Text2.SelStart = 0
  Text2.SelLength = Len(Text2)
  '
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  '
  If KeyAscii = 13 Then
     Text8.SetFocus
  End If
  '
End Sub

Private Sub Text2_LostFocus()
  '
   Text2 = HORATEX$(HORANUM(Text2))
   Exit Sub
'  Static MODORA%
'   If MODORA% < 1 Then
'      MODORA% = 1
'      If UCase$(TRIM$(CONTROL$("PADIAMOD", "CONFHORA"))) = "MINUCENT" Then
'         MODORA% = 2
'      End If
'   End If
'  '
'  If XVALO(Text2.TEXT) > 0 Then
'     Text2.TEXT = REPLACAR$(Text2.TEXT, ".", ":")
'     If InStr(1, Text2.TEXT, ".") = 0 Then
'      If InStr(1, Text2.TEXT, ":") = 0 Then
'         Text2.TEXT = Text2.TEXT + ":0"
'      End If
'     End If
'    '
'     If MODORA% = 2 Then
'        TTXX1$ = FORMATNUM$(REPLACAR$(Text2.TEXT, ":", "."), "F7.2")
'        Text2 = TRIM$(REPLACAR$(TTXX1$, ".", ","))
'        Exit Sub
'     End If
'    '
'     On Error Resume Next
'     DXX1 = CDate("0" + Text2.TEXT)
'     On Error GoTo 0
'     Text2.TEXT = Format$(DXX1, "hh:mm")
'  End If
  '
End Sub

Private Sub Text3_Change()
  '
  Label7 = ""
  Label18 = ""
  Check1.Value = 0
  Check2.Value = 0
  CID1% = CODINT%(Text10)
  NOPEX% = 0: ALTEX% = 0
  '
  NOPX$ = REPLACAR$(TRIM$(Text3), "-", "/")
  NOPX$ = REPLACAR$(NOPX$, ".", "/")
  PPXX% = InStr(NOPX$, "/")
  If PPXX% < 1 Then PPXX% = 1 + Len(NOPX$)
  On Error Resume Next
  NOPEX% = XVALO(Left$(NOPX$, PPXX% - 1))
  ALTEX% = XVALO(Mid$(NOPX$, PPXX% + 1))
  On Error GoTo 0
  Label7 = DESCRIOP$(CID1%, NOPEX%, ALTEX%)   ' XVALO(Text3.TEXT))
  '
  Call GENLISOP
  '
End Sub
'
Sub GENLISOP()
   List1.Clear
   Text1 = ""
   Text4 = ""
   Text5 = ""
   Text6 = ""
   Text7 = ""
   Text8 = ""
   Text9 = ""
   Check1.Value = 0
   Check2.Value = 0
   Label18 = ""
   '
   CIXI% = CODINT%(Text10)
   If CIXI% < 1 Then Exit Sub
   '
   NROOF$ = TRIM$(Label4)
   '
   NOPX$ = REPLACAR$(TRIM$(Text3), "-", "/")
   NOPX$ = REPLACAR$(NOPX$, ".", "/")
   PPXX% = InStr(NOPX$, "/")
   If PPXX% < 1 Then PPXX% = 1 + Len(NOPX$)
   NOPE% = 0: ALTEX% = 0
   On Error Resume Next
   NOPE% = XVALO(Left$(NOPX$, PPXX% - 1))
   ALTEX% = XVALO(Mid$(NOPX$, PPXX% + 1))
   On Error GoTo 0
   If NOPE% < 1 Then Exit Sub
     '
     Call ABREORFAB
     SQLX$ = " SELECT * FROM INFOFAB "
     SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CIXI% & " "
     SQLX$ = SQLX$ + "AND IdSubOr = '" & NROOF$ & "' "
     SQLX$ = SQLX$ + "AND NumeOper = " & NOPE% & " "
     SQLX$ = SQLX$ + "ORDER BY NuOrInf ASC"
     '
     Set FABRTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With FABRTMP
       On Error Resume Next
       .MoveFirst
       If Err Then
         On Error GoTo 0
         Exit Sub
       End If
       '
       On Error GoTo 0
       Do While (.EOF = False)
'\\\OT-06-0309-MH-28/08/06///
          ALTIN% = 0
          On Error Resume Next
          ALTIN% = !AltOper
          On Error GoTo 0
          '
          If ALTIN% = ALTEX% Then
             NORI% = !NuOrInf      'numero de INFORME
             'AGREGADAS POR MI
             FEC% = CVINT(!FeHorLib)
             HDES% = !Hora_Des
             HHAS% = !Hora_Has
             HFRI% = 0
             On Error Resume Next
             HFRI% = !Hora_Fri
             On Error GoTo 0
             MAQ$ = !Maquina
             CAN# = !CantReal
             CANOCO# = 0
             On Error Resume Next
             CANOCO# = !CaNoConf
             On Error GoTo 0
             LEGAX$ = !LegaOps
             If Len(LEGAX$) > 2 Then
                  LEG% = 0
                  DES$ = "CÉLULA"
                Else
                  LEG% = CVI(LEGAX$)
                  DES$ = ECOARCH$("PADRON", MKI$(XVALO(LEG)))
             End If
             LEGAY$ = ""
             For KKI& = 1 To Len(LEGAX$) Step 2
                NLEGA% = CVI(Mid$(LEGAX$, KKI&, 2))
                If NLEGA% > 0 Then
                   If LEGAY$ <> "" Then LEGAY$ = LEGAY$ + " / "
                   LEGAY$ = LEGAY$ + TRIM$(Str$(NLEGA%))
                End If
             Next KKI&
             '
             MARETRA$ = " ": ATEMUL$ = " ": OBSX$ = ""
             On Error Resume Next
             MARETRA$ = TRIM$(Str$(!ReTraba))
             ATEMUL$ = TRIM$(Str$(!AtMulti))
             OBSX$ = !Referen
             On Error GoTo 0
             'Label5.Caption = ECOARCH$("PADRON", MKI$(XVALO(Text1.TEXT)))
             YY$ = Space$(256)
             Call REPLA(YY$, AJUSTD$(Str$(NORI%), 4), 1, 4)
             'AGREGADAS POR MI
             Call REPLA(YY$, FECHATEX$(FEC%), 7, 8)
             Call REPLA(YY$, HORATEX(HDES), 17, 8)
             Call REPLA(YY$, HORATEX$(HHAS), 27, 8)
             Call REPLA(YY$, HORATEX$(HFRI%), 97, 8)
             Call REPLA(YY$, MAQ, 37, 6)
             Call REPLA(YY$, FORMATNUM(CAN#, "F8.1"), 45, 8)
             Call REPLA(YY$, FORMATNUM(CANOCO#, "F8.1"), 105, 8)
             Call REPLA(YY$, DES$, 55, 15)
             Call REPLA(YY$, FORMATNUM(LEG%, "F8"), 70, 8)
             Call REPLA(YY$, MARETRA$, 79, 1)
             Call REPLA(YY$, ATEMUL$, 80, 1)
             Call REPLA(YY$, LEGAY$, 129, 128)
             YY$ = YY$ + OBSX$
             '
             List1.AddItem YY$
          End If
          '
       .MoveNext
       Loop
     End With
'\\\OT-06-0309-MH-28/08/06///
End Sub
'
Sub GENECODOPEOF(NROOF$)
     '
     If TRIM$(NROOF$) = "" Then
       CIXXI% = CODINT%(Text10)
       If CIXXI% > 0 Then Call GENECOPER(CIXXI%)
       Exit Sub
     End If
     '
     Call BLANARCH("!OPERINF")
     Call SETULTREG("!OPERINF", 0)
     Call FRESHECHO("!OPERINF") 'nuevo
     'ACA OBTENGO LOS REGISTROS DE OPERFAB, CUYO IdSubOr = NROOF$
     'Luego cambio las variable por los campos correspondientes
     '
     Call ABREORFAB
     SQLX$ = " SELECT * FROM INFOFAB "
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
     SQLX$ = SQLX$ + "ORDER BY NumeOper ASC, AltOper ASC"
     '
     Set FABRTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With FABRTMP
       On Error Resume Next
       .MoveFirst
       If Err Then
         Call MENSERR(24, "No hay Operaciones Informadas\para la Orden de Fabricación Seleccionada.")
         On Error GoTo 0
         Exit Sub
       End If
       '
       On Error GoTo 0
       JJ& = 0: NOPANT% = 0
       YY$ = REGBLAN("OPERINF")
       Do While (.EOF = False)
          NOPE% = !NumeOper      'numero de operación
          ALTOP% = 0
          On Error Resume Next
          ALTOP% = !AltOper
          On Error GoTo 0
          If NOPE% <> NOPANT% Or ALTOP% <> ALTAN% Then
            NOPANT% = NOPE%
            ALTAN% = ALTOP%
            DOPE$ = !DescOper      'Descripcion de la operacion
            MAQUI$ = !Maquina
            '
            JJ& = JJ& + 1: JJI% = JJ&
            YY$ = MKI$(JJI%) + Space$(58) + String(4, 0)
            NOPX$ = FORMATNUM$(NOPE%, "I3")
              If ALTOP% > 0 Then NOPX$ = NOPX$ + "/" + TRIM$(Str$(ALTOP%))
            NOPX$ = AJUSTI$(NOPX$, 6) + "- "
            NOPX$ = NOPX$ + AJUSTI$(DOPE$, 41) + " - "
            NOPX$ = NOPX$ + AJUSTI$(MAQUI$, 6)
            Call REPLA(YY$, NOPX$, 3, 58)
            Call GRAREG("!OPERINF", YY$, JJ&)
            '
          End If
       .MoveNext
       Loop
     End With
        
     Call SETULTREG("!OPERINF", JJ&)
     Call CIERRARCH("!OPERINF")
     '
     Screen.MousePointer = 1
     '
End Sub
'
Sub GENECOPER(CIJK%)
     '
     Call BLANARCH("!OPERINF")
     Call SETULTREG("!OPERINF", 0)
     Call FRESHECHO("!OPERINF") 'nuevo
     'ACA OBTENGO LOS REGISTROS DE OPERFAB, CUYO IdSubOr = NROOF$
     'Luego cambio las variable por los campos correspondientes
     '
     Call ABREMASTER
     Set SecopTemp = Articulos.OpenRecordset("SELECT * FROM SecOper WHERE CodIConj=" & CIJK% & " AND StatOper = 1 ORDER BY NumeOper ASC, AltOper ASC", dbOpenSnapshot)
     On Error Resume Next
     SecopTemp.MoveFirst
     If Err < 1 Then
       JJ& = 0
       YY$ = REGBLAN("OPERINF")
       With SecopTemp
         Do While Not .EOF
           CAPROD# = 0      'Cantidad Programada
           CAREA# = 0       'Cantidad realizada
           SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
           '
           NOPE% = !NumeOper      'numero de operación
           ALTE% = !AltOper
           NO3$ = FORMATNUM(NOPE%, "F3.0")
           NOPX$ = AJUSTD$(Str$(NOPE%), 3)
             If ALTE% > 0 Then NOPX$ = NOPX$ + "/" + TRIM$(Str$(ALTE%))
           NOPX$ = Left$(NOPX$ + "      ", 6) + "- "
           DOPE$ = !DescOper      'Descripcion de la operacion
           MAQUI$ = !CodMaq0      'Máquina
           '
           JJ& = JJ& + 1: JJI% = JJ&
           YY$ = MKI$(JJI%) + Space$(58) + String(4, 0)
           NOPX$ = NOPX$ + AJUSTI$(DOPE$, 29) + " - "
           NOPX$ = NOPX$ + AJUSTI$(MAQUI$, 6)
           NOPX$ = NOPX$ + FORMATNUM$(SALDO#, "F12.1")
           Call REPLA(YY$, NOPX$, 3, 58)
              CAPROD1 = CAPROD#
           Call REPLA(YY$, MKS$(CAPROD1), 61, 4)
           Call GRAREG("!OPERINF", YY$, JJ&)
           '
         .MoveNext
         Loop
       End With
     End If
     '
     Call SETULTREG("!OPERINF", JJ&)
     Call CIERRARCH("!OPERINF")
     '
     Screen.MousePointer = 1
     '
End Sub

Function CAPENOP$(NROOF$, NOPE1%, CANPROGRA)
  '
  CAPENOP$ = ""
  Call ABREORFAB
  SQLX$ = " SELECT * FROM OPERFAB "
  SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
  SQLX$ = SQLX$ + "AND NumeOper = " & NOPE1% & " "
  SQLX$ = SQLX$ + " ORDER BY NumeOper "
  '
  Set FABRTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
  '
  With FABRTMP
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      Exit Function
    End If
    On Error GoTo 0
    CAPROD# = !CantProg      'Cantidad Programada
    CAREA# = !CantReal       'Cantidad realizada
    SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
    CAPENOP$ = NUMSTRI$(SALDO#, 12, 1, 0)
    CANPROGRA = CAPROD#
  End With
  '
End Function
'
Function MAQDEFIOP$(NROOF$, NOPE1%)
  '
  MAQDEFIOP$ = ""
  Call ABREORFAB
  SQLX$ = " SELECT * FROM OPERFAB "
  SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
  SQLX$ = SQLX$ + "AND NumeOper = " & NOPE1% & " "
  SQLX$ = SQLX$ + " ORDER BY NumeOper "
  '
  Set FABRTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
  '
  With FABRTMP
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      GoTo 20
    End If
    On Error GoTo 0
    MAQDEFIOP$ = !Maquina
  End With
  '
20 If MAQDEFIOP$ = "" Then
     CIXI% = CODINT%(Text10)
     MAQDEFIOP$ = MAQUISTAN$(CIXI%, NOPE1%)
   End If
  '
End Function

Private Sub Text3_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   Text5.SetFocus
 End If
End Sub

Private Sub Text3_LostFocus()
  NOPX$ = REPLACAR$(TRIM$(Text3), "-", "/")
  NOPX$ = REPLACAR$(NOPX$, ".", "/")
  Text3 = NOPX$
End Sub

Private Sub Text4_DblClick()
   Call Command5_Click
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
      Text1.SetFocus
   End If
End Sub

  '
  '
Private Sub Text5_dblclick()
  '
  Call Command7_Click
  '
End Sub
Private Sub text5_change()
  '
  Label10.Caption = TRIM(ECOARCH$("PADMAQ", Text5.TEXT))
  '
  
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Text6.SetFocus
End If
End Sub

Private Sub Text6_Change()
   '
   Call CALCUTIEMPO
   '
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
    Text7.SetFocus
   End If
End Sub

Private Sub Text6_lostfocus()
  '
  Text6 = HORATEX$(HORANUM(Text6))
  Exit Sub
'   Static MODORA%
'   If MODORA% < 1 Then
'      MODORA% = 1
'      If UCase$(TRIM$(CONTROL$("PADIAMOD", "CONFHORA"))) = "MINUCENT" Then
'         MODORA% = 2
'      End If
'   End If
'   '
'  If XVALO(Text6.TEXT) > 0 Then
'    Text6.TEXT = REPLACAR$(Text6.TEXT, ".", ":")
'    If InStr(1, Text6.TEXT, ".") = 0 Then
'      If InStr(1, Text6.TEXT, ":") = 0 Then
'        Text6.TEXT = Text6.TEXT + ":0"
'      End If
'    End If
'    '
'    If MODORA% = 2 Then
'       TTXX1$ = FORMATNUM$(REPLACAR$(Text6.TEXT, ":", "."), "F7.2")
'       Text6 = TRIM$(REPLACAR$(TTXX1$, ".", ","))
'       Exit Sub
'    End If
'    '
'    On Error Resume Next
'    DXX1 = CDate("0" + Text6.TEXT)
'    On Error GoTo 0
'    Text6.TEXT = Format$(DXX1, "hh:mm")
'  End If
  '
End Sub

Private Sub Text7_Change()
   '
   Call CALCUTIEMPO
   '
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Text2.Visible = True Then
         Text2.SetFocus
      Else
         Text8.SetFocus
      End If
   End If
End Sub

Private Sub Text7_lostfocus()
  '
  Text7 = HORATEX$(HORANUM(Text7))
  Exit Sub
'  Static MODORA%
'  If MODORA% < 1 Then
'     MODORA% = 1
'     If UCase$(TRIM$(CONTROL$("PADIAMOD", "CONFHORA"))) = "MINUCENT" Then
'        MODORA% = 2
'     End If
'  End If
'  '
'  If XVALO(Text7.TEXT) > 0 Then
'    Text7.TEXT = REPLACAR$(Text7.TEXT, ".", ":")
'    If InStr(1, Text7.TEXT, ".") = 0 Then
'      If InStr(1, Text7.TEXT, ":") = 0 Then
'        Text7.TEXT = Text7.TEXT + ":0"
'      End If
'    End If
'    '
'    If MODORA% = 2 Then
'       TTXX1$ = FORMATNUM$(REPLACAR$(Text7.TEXT, ":", "."), "F7.2")
'       Text7 = TRIM$(REPLACAR$(TTXX1$, ".", ","))
'       Exit Sub
'    End If
'    '
'    On Error Resume Next
'    DXX1 = CDate("0" + Text7.TEXT)
'    On Error GoTo 0
'    Text7.TEXT = Format$(DXX1, "hh:mm")
'  End If
  '
End Sub

Private Sub Text8_GotFocus()
  Text8.SelStart = 0
  Text8.SelLength = Len(Text8)
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If Text11.Visible = True Then
        Text11.SetFocus
     Else
        Text9.SetFocus
     End If
  End If
End Sub

Private Sub Text8_LOSTFOCUS()
  '
  Text8.TEXT = TRIM$(NUMSTRI$(XVALO(Text8.TEXT), 12, 1, 0))
  '
End Sub

Function HAYOPINF%(NROOF$)
   '
   HAYOPINF% = 0
   Call ABREORFAB
   SQLX$ = " SELECT * FROM INFOFAB "
   SQLX$ = SQLX$ + "WHERE IdSubOr = '" & NROOF$ & "' "
   Set FABRTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   With FABRTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
       Exit Function
     End If
     HAYOPINF% = 1
   End With
   '
End Function

Sub CALCUTIEMPO()
   '
   HCOM% = HORANUM(Text6)
   HFIN% = HORANUM(Text7)
   HDESCAN% = HORANUM(Text2)
   DIFDES = DIFHORA(0, HDESCAN%)
   '
   Label25 = ""
   If HFIN% >= HCOM% Then
     TDIF = (DIFHORA(HCOM%, HFIN%) - DIFDES) * 3600
     If TDIF >= 0 Then
        HMINX$ = HORMINSE$(TDIF)
        HMINY$ = Left$(HMINX$, Len(HMINX$) - 3)
        Label25.Caption = HMINY$
     End If
   End If
   '
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   '
   If KeyAscii = 13 Then
     If TRIM$(Text9) = "" Then
       Command2.SetFocus
     End If
   End If
   '
End Sub
