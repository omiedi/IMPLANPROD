VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form DESCRIREP 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descripción de Reparación"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11835
   Icon            =   "DESCRIREP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   11835
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame9 
      BackColor       =   &H00CDDADA&
      Caption         =   "DESCRIPCION DE LA REPARACION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   19
      Top             =   2040
      Width           =   10695
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
         Height          =   650
         Left            =   9840
         Picture         =   "DESCRIREP.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   135
         ToolTipText     =   "Graba e Imprime Registro"
         Top             =   3000
         Width           =   690
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   5160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   2760
         Width           =   4605
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   5160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   1560
         Width           =   4605
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   5160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   480
         Width           =   4605
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         IntegralHeight  =   0   'False
         Left            =   240
         Style           =   1  'Checkbox
         TabIndex        =   21
         Top             =   2280
         Width           =   4725
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1290
         IntegralHeight  =   0   'False
         Left            =   240
         Style           =   1  'Checkbox
         TabIndex        =   20
         Top             =   600
         Width           =   4725
      End
      Begin VB.Label Label8 
         BackColor       =   &H00CDDADA&
         Caption         =   "REACONDICIONAMIENTO REALIZADO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   29
         Top             =   2520
         Width           =   3615
      End
      Begin VB.Label Label7 
         BackColor       =   &H00CDDADA&
         Caption         =   "TRABAJOS REALIZADOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   27
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Label Label6 
         BackColor       =   &H00CDDADA&
         Caption         =   "OBSERVACIONES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   24
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label5 
         BackColor       =   &H00CDDADA&
         Caption         =   "ANALISIS DE CAUSAS DE FALLA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   2040
         Width           =   3615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00CDDADA&
         Caption         =   "FALLAS DETECTADAS POR DPTO TECNICO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00CDDADA&
      Caption         =   "CONTROL FINAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   18
      Top             =   5880
      Width           =   10695
      Begin VB.CommandButton Command8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   7050
         Picture         =   "DESCRIREP.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   138
         ToolTipText     =   "Quitar Datos del Ultimo Control"
         Top             =   3360
         Width           =   650
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
         Height          =   650
         Left            =   7750
         Picture         =   "DESCRIREP.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   137
         ToolTipText     =   "Documentos Externos"
         Top             =   3360
         Width           =   650
      End
      Begin VB.CommandButton Command6 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   8470
         Picture         =   "DESCRIREP.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   125
         Top             =   3360
         Width           =   650
      End
      Begin VB.CommandButton Command23 
         Height          =   650
         Left            =   9190
         Picture         =   "DESCRIREP.frx":0D72
         Style           =   1  'Graphical
         TabIndex        =   124
         ToolTipText     =   "Imprime Registro de Control Final"
         Top             =   3360
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
         Height          =   650
         Left            =   9910
         Picture         =   "DESCRIREP.frx":13DC
         Style           =   1  'Graphical
         TabIndex        =   123
         ToolTipText     =   "Graba e Imprime Registro"
         Top             =   3360
         Width           =   650
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   9600
         TabIndex        =   114
         Text            =   " "
         Top             =   720
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   9600
         TabIndex        =   113
         Text            =   " "
         Top             =   1080
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   8
         Left            =   9600
         TabIndex        =   112
         Text            =   " "
         Top             =   1440
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   9
         Left            =   9600
         TabIndex        =   111
         Text            =   " "
         Top             =   1800
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   10
         Left            =   9600
         TabIndex        =   110
         Text            =   " "
         Top             =   2160
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   11
         Left            =   9600
         TabIndex        =   109
         Text            =   " "
         Top             =   2520
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   7800
         TabIndex        =   99
         Top             =   720
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   7800
         TabIndex        =   98
         Text            =   " "
         Top             =   1080
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   7800
         TabIndex        =   97
         Text            =   " "
         Top             =   1440
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   7800
         TabIndex        =   96
         Text            =   " "
         Top             =   1800
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   7800
         TabIndex        =   95
         Text            =   " "
         Top             =   2160
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   7800
         TabIndex        =   94
         Text            =   " "
         Top             =   2520
         Width           =   945
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1725
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   2160
         Width           =   3330
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   1920
         TabIndex        =   31
         Text            =   " "
         Top             =   1560
         Width           =   1305
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
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
         Left            =   120
         TabIndex        =   30
         Text            =   " "
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label Label53 
         BackStyle       =   0  'Transparent
         Caption         =   "Técnico Reparación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6840
         TabIndex        =   136
         Top             =   2880
         Width           =   945
      End
      Begin VB.Label Label117 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2520
         TabIndex        =   134
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label83 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7850
         TabIndex        =   133
         Top             =   3000
         Width           =   2685
      End
      Begin VB.Label Label88 
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
         Left            =   120
         TabIndex        =   132
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label Label115 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   131
         Top             =   960
         Width           =   1440
      End
      Begin VB.Label Label58 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº O/F."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         TabIndex        =   130
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label54 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   129
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label112 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   128
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mediciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7200
         TabIndex        =   127
         Top             =   240
         Width           =   3405
      End
      Begin VB.Label Label51 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1320
         TabIndex        =   126
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   2640
         Picture         =   "DESCRIREP.frx":16E6
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   2760
         Picture         =   "DESCRIREP.frx":19F0
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label85 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   840
         TabIndex        =   122
         ToolTipText     =   "Número de Revisión"
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label86 
         BackStyle       =   0  'Transparent
         Caption         =   "Revisión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   121
         Top             =   480
         Width           =   795
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CH07:"
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
         Left            =   9000
         TabIndex        =   120
         Top             =   795
         Width           =   540
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CH08:"
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
         Left            =   9000
         TabIndex        =   119
         Top             =   1155
         Width           =   540
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CH09:"
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
         Left            =   9000
         TabIndex        =   118
         Top             =   1515
         Width           =   540
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CH10:"
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
         Left            =   9000
         TabIndex        =   117
         Top             =   1875
         Width           =   540
      End
      Begin VB.Label Label56 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CH11:"
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
         Left            =   9000
         TabIndex        =   116
         Top             =   2235
         Width           =   540
      End
      Begin VB.Label Label57 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CH12:"
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
         Left            =   9000
         TabIndex        =   115
         Top             =   2595
         Width           =   540
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "11"
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
         Left            =   5280
         TabIndex        =   108
         Top             =   765
         Width           =   300
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   10
         Left            =   5640
         TabIndex        =   107
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   10
         Left            =   6120
         TabIndex        =   106
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CH01:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7200
         TabIndex        =   105
         Top             =   795
         Width           =   540
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CH02:"
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
         Left            =   7200
         TabIndex        =   104
         Top             =   1155
         Width           =   540
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CH03:"
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
         Left            =   7200
         TabIndex        =   103
         Top             =   1515
         Width           =   540
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CH04:"
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
         Left            =   7200
         TabIndex        =   102
         Top             =   1875
         Width           =   540
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CH05:"
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
         Left            =   7200
         TabIndex        =   101
         Top             =   2235
         Width           =   540
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CH06:"
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
         Left            =   7200
         TabIndex        =   100
         Top             =   2595
         Width           =   540
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "19"
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
         Left            =   5280
         TabIndex        =   93
         Top             =   3360
         Width           =   300
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "12"
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
         Left            =   5280
         TabIndex        =   92
         Top             =   1095
         Width           =   300
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   11
         Left            =   5640
         TabIndex        =   91
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   11
         Left            =   6120
         TabIndex        =   90
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "13"
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
         Left            =   5280
         TabIndex        =   89
         Top             =   1455
         Width           =   300
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   12
         Left            =   5640
         TabIndex        =   88
         Top             =   1395
         Width           =   360
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   12
         Left            =   6120
         TabIndex        =   87
         Top             =   1395
         Width           =   360
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "14"
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
         Left            =   5280
         TabIndex        =   86
         Top             =   1770
         Width           =   300
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   13
         Left            =   5640
         TabIndex        =   85
         Top             =   1725
         Width           =   360
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   13
         Left            =   6120
         TabIndex        =   84
         Top             =   1725
         Width           =   360
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "15"
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
         Left            =   5280
         TabIndex        =   83
         Top             =   2070
         Width           =   300
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   14
         Left            =   5640
         TabIndex        =   82
         Top             =   2040
         Width           =   360
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   14
         Left            =   6120
         TabIndex        =   81
         Top             =   2040
         Width           =   360
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   19
         Left            =   6120
         TabIndex        =   80
         Top             =   3645
         Width           =   360
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   19
         Left            =   5640
         TabIndex        =   79
         Top             =   3645
         Width           =   360
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   18
         Left            =   6120
         TabIndex        =   78
         Top             =   3330
         Width           =   360
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   18
         Left            =   5640
         TabIndex        =   77
         Top             =   3330
         Width           =   360
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   17
         Left            =   6120
         TabIndex        =   76
         Top             =   3015
         Width           =   360
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   17
         Left            =   5640
         TabIndex        =   75
         Top             =   3015
         Width           =   360
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   16
         Left            =   6120
         TabIndex        =   74
         Top             =   2685
         Width           =   360
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   16
         Left            =   5640
         TabIndex        =   73
         Top             =   2685
         Width           =   360
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   15
         Left            =   6120
         TabIndex        =   72
         Top             =   2370
         Width           =   360
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   15
         Left            =   5640
         TabIndex        =   71
         Top             =   2370
         Width           =   360
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "20"
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
         Left            =   5280
         TabIndex        =   70
         Top             =   3645
         Width           =   300
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "18"
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
         Left            =   5280
         TabIndex        =   69
         Top             =   3030
         Width           =   300
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "17"
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
         Left            =   5280
         TabIndex        =   68
         Top             =   2670
         Width           =   300
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "16"
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
         Left            =   5280
         TabIndex        =   67
         Top             =   2370
         Width           =   300
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "10"
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
         Left            =   3840
         TabIndex        =   66
         Top             =   3660
         Width           =   300
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   9
         Left            =   4200
         TabIndex        =   65
         Top             =   3600
         Width           =   360
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   9
         Left            =   4680
         TabIndex        =   64
         Top             =   3600
         Width           =   360
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "8"
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
         Left            =   3840
         TabIndex        =   63
         Top             =   3000
         Width           =   300
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "9"
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
         Left            =   3840
         TabIndex        =   62
         Top             =   3300
         Width           =   300
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1"
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
         Left            =   3840
         TabIndex        =   61
         Top             =   750
         Width           =   300
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   4200
         TabIndex        =   60
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   4680
         TabIndex        =   59
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "2"
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
         Left            =   3840
         TabIndex        =   58
         Top             =   1065
         Width           =   300
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   4200
         TabIndex        =   57
         Top             =   1035
         Width           =   360
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   4680
         TabIndex        =   56
         Top             =   1035
         Width           =   360
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "3"
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
         Left            =   3840
         TabIndex        =   55
         Top             =   1410
         Width           =   300
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   4200
         TabIndex        =   54
         Top             =   1365
         Width           =   360
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   4680
         TabIndex        =   53
         Top             =   1365
         Width           =   360
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "4"
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
         Left            =   3840
         TabIndex        =   52
         Top             =   1710
         Width           =   300
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   4200
         TabIndex        =   51
         Top             =   1680
         Width           =   360
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   4680
         TabIndex        =   50
         Top             =   1680
         Width           =   360
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "5"
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
         Left            =   3840
         TabIndex        =   49
         Top             =   2025
         Width           =   300
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   4200
         TabIndex        =   48
         Top             =   1995
         Width           =   360
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   4680
         TabIndex        =   47
         Top             =   1995
         Width           =   360
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "6"
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
         Left            =   3840
         TabIndex        =   46
         Top             =   2340
         Width           =   300
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   4200
         TabIndex        =   45
         Top             =   2325
         Width           =   360
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   4680
         TabIndex        =   44
         Top             =   2325
         Width           =   360
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "7"
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
         Left            =   3840
         TabIndex        =   43
         Top             =   2700
         Width           =   300
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   4200
         TabIndex        =   42
         Top             =   2640
         Width           =   360
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   4680
         TabIndex        =   41
         Top             =   2640
         Width           =   360
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   4200
         TabIndex        =   40
         Top             =   2955
         Width           =   360
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   4680
         TabIndex        =   39
         Top             =   2955
         Width           =   360
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   4200
         TabIndex        =   38
         Top             =   3285
         Width           =   360
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   4680
         TabIndex        =   37
         Top             =   3285
         Width           =   360
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nº Control"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3960
         TabIndex        =   36
         Top             =   240
         Width           =   2565
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -120
         TabIndex        =   35
         Top             =   1920
         Width           =   1410
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Tensión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   34
         Top             =   1320
         Width           =   810
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Dictámen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         TabIndex        =   33
         Top             =   1320
         Width           =   1065
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00CDDADA&
      Caption         =   "FIRMA ELECTRONICA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   780
      Left            =   4920
      TabIndex        =   13
      Top             =   8760
      Width           =   5525
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1410
         TabIndex        =   15
         Top             =   315
         Width           =   4005
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   315
         TabIndex        =   14
         Top             =   400
         Width           =   1005
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00CDDADA&
      Caption         =   "CODIGO - MODELO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1995
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   10695
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   885
         Left            =   1005
         TabIndex        =   12
         Top             =   990
         Width           =   4005
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Equipo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   -90
         TabIndex        =   11
         Top             =   680
         Width           =   1005
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1005
         TabIndex        =   10
         Top             =   600
         Width           =   4005
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Defectos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   -90
         TabIndex        =   9
         Top             =   1020
         Width           =   1005
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Talón Nro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   8350
         TabIndex        =   8
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   9465
         TabIndex        =   7
         Top             =   525
         Width           =   930
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6555
         TabIndex        =   6
         Top             =   525
         Width           =   2700
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Serie"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   5910
         TabIndex        =   5
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   885
         Left            =   5460
         TabIndex        =   4
         Top             =   990
         Width           =   4950
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Accesorios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   5040
         TabIndex        =   3
         Top             =   720
         Width           =   1245
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   11745
      Left            =   10920
      ScaleHeight     =   11715
      ScaleWidth      =   2115
      TabIndex        =   1
      Top             =   0
      Width           =   2145
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   3000
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
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
         Left            =   120
         Picture         =   "DESCRIREP.frx":1CFA
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Muestra el Historial del Equipo Elegido"
         Top             =   1800
         Width           =   690
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
         Left            =   120
         Picture         =   "DESCRIREP.frx":243C
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Acceso a Tablas de Fallas y Causas"
         Top             =   1080
         Width           =   690
      End
      Begin VB.CommandButton Command2 
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
         Height          =   650
         Left            =   120
         Picture         =   "DESCRIREP.frx":2586
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Cierra y Abondona la Aplicación"
         Top             =   315
         Width           =   690
      End
   End
End
Attribute VB_Name = "DESCRIREP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
Dim CantidadControlesArticulo As Integer
Private Sub Command1_Click()
    Command1.Enabled = False
    '
    NUREP& = XVALO(Label25)
    CONDI$ = "NUMREPA = " & NUREP& & ""  '
    If List2.SelCount < 1 Then
       Call COMUNI("Debe Ingresar una o más Causas de Falla")
       GoTo 99
    End If
    '
    If List1.SelCount < 1 Then
       Call COMUNI("Debe Ingresar uno o más Tipos de Falla")
       GoTo 99
    End If
    '
    Screen.MousePointer = 11
    
''    Call ACTUREGISTRO("REPARA", "REPARACION", "NUMREPA = " & NUREP& & "", Text3, REGAF&, "NOMESS")
''    Call ACTUREGISTRO("REPARA", "OBSERVA", "NUMREPA = " & NUREP& & "", Text2, REGAF&, "NOMESS")
''    Call ACTUREGISTRO("REPARA", "REACONDI", "NUMREPA = " & NUREP& & "", Text1, REGAF&, "NOMESS")
''    Call ACTUREGISTRO("REPARA", "TEC_REPAR", "NUMREPA = " & NUREP& & "", Label2, REGAF&, "NOMESS")
    SQLX$ = "SELECT REPARACION,OBSERVA, REACONDI,TEC_REPARA,NUMREPA  FROM REPARA "
    SQLX$ = SQLX$ + " WHERE NUMREPA = " & NUREP&
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
25  On Error Resume Next
    RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
    End If
    On Error GoTo 0
    '
    With RS
      Do While Not .EOF
         !NumRepa = NUREP&
         !REPARACION = Text3
         !Observa = Text2
         !REACONDI = Text1
         !TEC_REPARA = Left$(Label2, 64)
        .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing
    '
    ' DESCRIPCION DE LOS FALLOS
''    Screen.MousePointer = 11
''    DELARFA$ = ""
''    CONDI1$ = "NUMREPA = " & NUREP& & " AND ORIGEN = 2 "  '
''    Call BORRAREGISTROS("TIPOFALL", CONDI1$, RAFE&, "NOMESS")
''    For KKI% = 1 To List1.ListCount
''      If List1.Selected(KKI% - 1) = True Then
''         If DELARFA$ <> "" Then DELARFA$ = DELARFA$ + vbCrLf
''         DELARFA$ = DELARFA$ + TRIM$(List1.List(KKI% - 1))
''         '
''28       RNDX& = XVALO(RANDSTRING$(-6))
''         CONDI1$ = "NUMREPA = " & RNDX&
''         If CantRegistros("TIPOFALL", CONDI1$, "NOMESS") > 0 Then GoTo 28
''         Call CreaRegistro("TIPOFALL", "NUMREPA", RNDX&, "NOMESS")
''         Call ACTUREGISTRO("TIPOFALL", "ORIGEN", CONDI1$,  2, RAFE&, "NOMESS")  ' 1 CLIENTE; 2 TECNICO
''         Call ACTUREGISTRO("TIPOFALL", "CODIFALL", CONDI1$, XVALO(Left$(List1.List(KKI% - 1), 4)), RAFE&, "NOMESS")
''         Call ACTUREGISTRO("TIPOFALL", "DESCFALL", CONDI1$, TRIM$(Mid$(List1.List(KKI% - 1), 6)), RAFE&, "NOMESS")
''         Call ACTUREGISTRO("TIPOFALL", "NUMREPA", CONDI1$, NUREP&, RAFE&, "NOMESS")
''      End If
''    Next KKI%
    '
    CONDI1$ = "NUMREPA = " & NUREP& & " AND ORIGEN = 2 "  '
    Call BORRAREGISTROS("TIPOFALL", CONDI1$, RAFE&, "NOMESS")
    
    Screen.MousePointer = 11
    DELARFA$ = ""

    SQLX$ = "SELECT NUMREPA,ORIGEN, CODIFALL,DESCFALL FROM TIPOFALL "
    SQLX$ = SQLX$ + " WHERE NUMREPA = " & NUREP&
    Dim RS1 As ADODB.Recordset
    Set RS1 = New ADODB.Recordset
26  On Error Resume Next
    RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 26
    End If
    On Error GoTo 0
    '
    With RS1
        For KKI% = 1 To List1.ListCount
          If List1.Selected(KKI% - 1) = True Then
                If DELARFA$ <> "" Then DELARFA$ = DELARFA$ + vbCrLf
                DELARFA$ = DELARFA$ + TRIM$(List1.List(KKI% - 1))
                .AddNew
                !NumRepa = NUREP&
                !ORIGEN = 2
                !CODIFALL = XVALO(Left$(List1.List(KKI% - 1), 4))
                !DESCFALL = TRIM$(Mid$(List1.List(KKI% - 1), 6))
                .Update
          End If
        Next KKI%
    End With
    RS1.Close
    Set RS1 = Nothing
    
    Call ACTUREGISTRO("REPARA", "DETIPFALLA", "NUMREPA =" & NUREP&, DELARFA$, REGAF&, "NOMESS")
    '
    ' DESCRIPCION DE CAUSA DE FALLOS
    Screen.MousePointer = 11
    DECAUFA$ = ""
    Call BORRAREGISTROS("CAUSFALL", CONDI$, RAFE&, "NOMESS")
    
''    For KKI% = 1 To List2.ListCount
''      If List2.Selected(KKI% - 1) = True Then
''         If DECAUFA$ <> "" Then DECAUFA$ = DECAUFA$ + vbCrLf
''         DECAUFA$ = DECAUFA$ + TRIM$(List2.List(KKI% - 1))
''         '
''38       RNDX& = XVALO(RANDSTRING$(-6))
''         CONDI1$ = "NUMREPA = " & RNDX&
''         If CantRegistros("CAUSFALL", CONDI1$, "NOMESS") > 0 Then GoTo 38
''         Call CreaRegistro("CAUSFALL", "NUMREPA", RNDX&, "NOMESS")
''         Call ACTUREGISTRO("CAUSFALL", "CODIFALL", CONDI1$, XVALO(Left$(List2.List(KKI% - 1), 4)), RAFE&, "NOMESS")
''         Call ACTUREGISTRO("CAUSFALL", "DESCFALL", CONDI1$, TRIM$(Mid$(List2.List(KKI% - 1), 6)), RAFE&, "NOMESS")
''         Call ACTUREGISTRO("CAUSFALL", "NUMREPA", CONDI1$, NUREP&, RAFE&, "NOMESS")
''      End If
''    Next KKI%
    SQLX$ = "SELECT NUMREPA,CODIFALL, CODIFALL,DESCFALL FROM CAUSFALL "
    SQLX$ = SQLX$ + " WHERE NUMREPA = " & NUREP&
    '
    Dim RS2 As ADODB.Recordset
    Set RS2 = New ADODB.Recordset
27  On Error Resume Next
    RS2.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 27
    End If
    On Error GoTo 0
    '
    With RS2
        For KKI% = 1 To List2.ListCount
          If List2.Selected(KKI% - 1) = True Then
                If DECAUFA$ <> "" Then DECAUFA$ = DECAUFA$ + vbCrLf
                DECAUFA$ = DECAUFA$ + TRIM$(List2.List(KKI% - 1))
                .AddNew
                !NumRepa = NUREP&
                !CODIFALL = XVALO(Left$(List2.List(KKI% - 1), 4))
                !DESCFALL = TRIM$(Mid$(List2.List(KKI% - 1), 6))
                .Update
          End If
        Next KKI%
    End With
    RS2.Close
    Set RS2 = Nothing

    Call ACTUREGISTRO("REPARA", "DECAUFALLA", "NUMREPA =" & NUREP&, DECAUFA$, REGAF&, "NOMESS")
    '
    Screen.MousePointer = 11
    NUMESER$ = TRIM$(Label26)
    RACLI$ = VALOBADA("REPARA", "RASO_CLI", "NUMREPA = " & NUREP& & "", "NOMESS")
    CUERPO$ = "N-Rep: " & NUREP& & " - Cliente: " & RACLI$ + vbCrLf + vbCrLf
    '
    Screen.MousePointer = 11
    If DELARFA$ <> "" Then
       CUERPO$ = CUERPO$ + "Fallas Halladas Depto Tecnico: " + vbCrLf + DELARFA$ + vbCrLf + vbCrLf
    End If
    '
    If DECAUFA$ <> "" Then
       CUERPO$ = CUERPO$ + "Causas de Fallas: " + vbCrLf + DECAUFA$ + vbCrLf + vbCrLf
    End If
    '
    If TRIM$(Text3) <> "" Then
       CUERPO$ = CUERPO$ + "Trabajos Realizados: " + vbCrLf + Text3 + vbCrLf + vbCrLf
    End If
    '
    If TRIM$(Text1) <> "" Then
       CUERPO$ = CUERPO$ + "Reacondicionamiento: " + vbCrLf + Text1 + vbCrLf + vbCrLf
    End If
    '
    Screen.MousePointer = 11
    Call ANOTANUSE(NUMESER$, CUERPO$, , "NOMESS")
    'Call Command2_Click
    '
99  Screen.MousePointer = 1
    Command1.Enabled = True
    
End Sub


Private Sub Command15_Click()
    '
    command15.Enabled = False
    Call SELECHO("TACAUFA/Causa de Fallas")
    COD$ = VALACT1$("TACAUFA")
    CODI& = XVALO(COD$)
    If CODI& > 0 Then
      Text2 = COD$
    End If
    command15.Enabled = True
    '
End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Unload Me
 End Sub

Private Sub Command23_Click()

    NUSERIE$ = TRIM$(Label51)
    CANTRE& = CantRegistros&("TANUMSE", "NumeroSerie ='" & NUSERIE$ & "'")
    If CANTRE& < 1 Then
       Call MENSERR(24, "Número de Serie Incorrecto")
       Exit Sub
    End If
    Call PRINTMEDI(0) '  0 IMPRESION SIMPLE --- UNO GRABACION
    
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    '
    'VALIDACION POR FECHA
    
    VDEVV$ = TRIM$(Label112)
    If FECHANUM(VDEVV$) < 1 Then
      Call MENSERR(24, "Fecha Inválida")
      GoTo 99
    End If
    '
    'VALIDACION DE NUMERO DE SERIE
    '
    NUSERIE$ = TRIM$(Label51)
    CANTRE& = CantRegistros&("TANUMSE", "NumeroSerie ='" & NUSERIE$ & "'")
    If CANTRE& < 1 Then
       Call MENSERR(24, "Número de Serie Incorrecto")
       GoTo 99
    End If
    '
    'VALIDACION DE DICTAMEN
    If TRIM$(Text5) = "" Then
      Call MENSERR(24, "Falta Informar Dictamen")
      GoTo 99
    End If
    '
    If TRIM$(Text8) = "" Then
      Call MENSERR(24, "Falta Tensión del Equipo")
      GoTo 99
    End If
    '
    If TRIM$(Label83) = "" Then
      Call MENSERR(24, "Falta Elegir Técnico para el Control Final")
      GoTo 99
    End If
    '
    Screen.MousePointer = 11
    '
    SQLX$ = "NumeroSerie = '" & NUSERIE$ & "'"
    '
    Call ACTUREGISTRO("TANUMSE", "Dic_Ensa", SQLX$, Left$(TRIM$(Text5), 4), REGAF&)
    Call ACTUREGISTRO("TANUMSE", "Obse_Ensa", SQLX$, Text4, REGAF&)
    Call ACTUREGISTRO("TANUMSE", "Fech_Ensa", SQLX$, TRIM$(Label112), REGAF&)
    Call ACTUREGISTRO("REPARA", "UsRealizaRepa", "NUMREPA = " & Label25, Left$(TRIM$(Label83), 32), REGAF&)
    '
    'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
    CUERPO$ = "Registro de Control Final de Reparaciones" & vbCrLf
    CUERPO$ = CUERPO$ + "Técnico: " & UCase$(TRIM$(Label83))
    CUERPO$ = CUERPO$ + " - Resultado: " & UCase$(TRIM$(Text5))
    Call ANOTANUSE(NUSERIE$, CUERPO$)
    '
    REVIS% = 0
    '
    CONDI$ = " NumeroSerie ='" & NUSERIE$ & "'  and ORIGEN = 'REP' "
    ULTREVIS% = XVALO(VALOBADA("CONTROLFIN", "MAX(REVISION)", CONDI$, "NOMESS"))
    Label85 = ULTREVIS%
    '
    SQLX$ = "SELECT * FROM CONTROLFIN "
    SQLX$ = SQLX$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
    SQLX$ = SQLX$ + " AND REVISION =" & ULTREVIS% & ""
    SQLX$ = SQLX$ + " AND ORIGEN = 'REP' "
    '
    Dim CONTRTMP As ADODB.Recordset
    Set CONTRTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With CONTRTMP
     If Not .EOF Then
       REVIS% = 1 + ULTREVIS%
     Else
       REVIS% = 0 'SI ES EL PRIMERO EL VALOR VA A SER 0
     End If
    End With
    CONTRTMP.Close
    Set CONTRTMP = Nothing
    '
    'GRABO LOS DATOS COMO NUEVOS EN CONTROLFIN
    SQLX$ = "SELECT * FROM CONTROLFIN"
    SQLX$ = SQLX$ + " WHERE NumeroSerie = '" & NUSERIE$ & "'"
    Dim CTRLFINTMP As ADODB.Recordset
    Set CTRLFINTMP = New ADODB.Recordset
    CTRLFINTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    For U& = 0 To Lis1.UBound
        With CTRLFINTMP
          .AddNew   ' YA CONVERTIDO
          On Error Resume Next
          !CodiEmpr = CodiEmp%
          On Error GoTo 0
          !NUMEROSERIE = NUSERIE$
          !COD_INTE = CODINT%(Label117)
          !NumeroOrden = U& + 1
          RESTDO$ = ""
          If Lis1(U&) <> "" Then RESTDO$ = "OK"
          If Lis2(U&) <> "" Then RESTDO$ = "NC"
          !Resultado = RESTDO$
          '!MARBORRA = 0
          !Fech_Ensa = TRIM$(Label112)
          !Revision = REVIS%
          !ORIGEN = "REP"
          .Update
        End With
    Next U&
    CTRLFINTMP.Close
    '
    '
    'GRABO LOS DATOS COMO NUEVOS EN MEDIFIN
    SQLX$ = "SELECT * FROM MEDIFIN"
    SQLX$ = SQLX$ + " WHERE NumeroSerie = '" & NUSERIE$ & "'"
    Dim MEDIFINTMP As ADODB.Recordset
    Set MEDIFINTMP = New ADODB.Recordset
    MEDIFINTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    For i& = 0 To 11
        With MEDIFINTMP
          .AddNew   ' YA CONVERTIDO
          On Error Resume Next
          !CodiEmpr = CodiEmp%
          On Error GoTo 0
          !NUMEROSERIE = NUSERIE$
          !COD_INTE = CODINT%(Label117)
          TTT$ = TRIM$(Str(i& + 1))
          While Len(TTT$) < 2: TTT$ = "0" + TTT$: Wend
          !CHN = "CH" + TTT$
          !MEDA = XVALO(TextA(i&))
          !Fech_Ensa = TRIM$(Label112)
          !Revision = REVIS%
          !ORIGEN = "REP"
          .Update
        End With
    Next i&
    MEDIFINTMP.Close
    '
    Dictamen$ = TRIM$(Text5)
    If InStr(1, Dictamen$, "A") > 0 Or InStr(1, Dictamen$, "APRO") > 0 Then
        CIERREP09.Label25 = Label25
        CIERREP09.Command1_Click
        CIERREP09.Command2_Click
    End If
    '
    CANCELPRINT% = 1
    Call PRINTMEDI(1)
    CANCELPRINT% = 0
    '
    FILTX$ = TRIM$(Label25) & ";" & Label85
    Call STDFORM("ORCF-MDT", FILTX$, "PRINT")
    '
    Label51.Caption = ""
    Label51.Caption = NUSERIE$
    '
99  Screen.MousePointer = 1
    '
    Command3.Enabled = True

End Sub

Sub PRINTMEDI(N%)
    '
    REVIS% = XVALO(Label85)
    If N% > 0 Then
        REVIS% = 0
        NUSERIE$ = TRIM$(Label51)
        SQLZ$ = "SELECT MAX(REVISION)AS MAXREVI FROM CONTROLFIN"
        SQLZ$ = SQLZ$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
        Dim MAXREVTMP As ADODB.Recordset
        Set MAXREVTMP = FLEXCONN.Execute(FILTSQL$(SQLZ$))
        ULTREVIS% = XVALO(MAXREVTMP!MAXREVI)
        MAXREVTMP.Close
        Set MAXREVTMP = Nothing
        REVIS% = ULTREVIS%
    End If
    TIPODOC$ = "Registro de Reparacion" ' guardar el documento immpreso
    NUMEDOC$ = TRIM$(Label51) & "-" & REVIS%
    '
    CODPARAM$(1) = "FECH-EMI": DOCPARAM$(1) = TRIM$(Label112) 'FECHA
    CODPARAM$(2) = "NRO-AENT": DOCPARAM$(2) = Label115 'o.f
    CODPARAM$(3) = "TALLE": DOCPARAM$(3) = TRIM$(Label83) 'TECNICO
    'CODPARAM$(4) = "COLOR": DOCPARAM$(4) = TRIM$(Text7) 'A.EXTERNO
    CODPARAM$(5) = "REF-MAT1":  DOCPARAM$(5) = TRIM$(Label51) 'NRO SERIE
    CODPARAM$(6) = "DES-MAT": DOCPARAM$(6) = Label22.Caption 'DESCRIPCION DE MATERIAL
    CODPARAM$(7) = "REF-MAT2":  DOCPARAM$(7) = Text5.TEXT  'DICTAMEN
    CODPARAM$(8) = "OBSERVA":  DOCPARAM$(8) = Text4.TEXT  'OBSERVACIONES
    CODPARAM$(9) = "N-HPROF":  DOCPARAM$(9) = XVALO((Text8.TEXT)) '\\\OT-08-1074-OD-24/11/08-MEDITEA-TENSION
    CAPARDOC% = 10
    '
    CODTABU1$(1) = "MEDIDA-1"
    CODTABU1$(2) = "MEDIDA-2"
    CODTABU1$(3) = "MEDIDA-3"
    CODTABU1$(4) = "N-PLANO"
    CODTABU1$(5) = "R-PLANO"
    CODTABU1$(6) = "F-PLANO"
    CODTABU1$(7) = "Z-PLANO"
    CACOLTAB1% = 7
    '
    U& = 0
    J& = 0
    CAITAB1% = 0
    For U& = 0 To 11
        CAITAB1% = CAITAB1% + 1
        TETABU1$(1, CAITAB1%) = TRIM$(TextA(U&))     'MEDICIONES A
    Next U&
    CAITAB1% = 0
    For J& = 0 To Lis1.UBound
        CAITAB1% = CAITAB1% + 1
        TETABU1$(4, CAITAB1%) = TRIM$(Lis1(J&))   ' NUMERO
        TETABU1$(5, CAITAB1%) = TRIM$(Lis2(J&))    'CONTROL
    Next J&
    '
    FORIPRE$ = TRIM$(Control("", "RECONRE"))
    Call PRINTDOC("@" + FORIPRE$)
End Sub

Sub BLANFORMU()
    i& = 0
    For i& = 0 To 14
        Lis1(i&) = ""
        Lis2(i&) = ""
    Next i&
    i& = 0
    For i& = 0 To 11
        TextA(i&) = ""
    Next i&
    Label115 = ""
    Label117 = ""
    'Label13 = ""
    Text5 = ""
    'Label8 = ""
    Text4 = ""
    Label112 = FECHATEX$(HOY(HO$))
    Label85 = ""
    Text8.Visible = False
End Sub
Private Sub Command4_Click()
    '
    Command4.Enabled = False
    If DERACCE%("ACTAFALL", "Actualizacion de Tablas de Fallas") >= 2 Then
10     OPX% = COMALTER%("Opciones de Tablas a Actualizar:\\Tipos de Falla\Causas de Falla")
       If OPX% = 1 Then
            Call CONFREPA09.Option3_DblClick
            GoTo 10
          ElseIf OPX% = 2 Then
            Call CONFREPA09.Option1_DblClick
            GoTo 10
       End If
       Call Form_Load
    End If
    '
99  Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    If TRIM$(Label26) <> "" Then
        CONUSE07.Label7 = Label26
        CONUSE07.Show 1
    End If
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()

   FILTX$ = TRIM$(Label25) & ";" & Label85
   Call STDFORM("ORCF-MDT", FILTX$, "PRINT")

   'If XVALO(VALOBADA("TANUMSE", "NumeroSerie", "NumeroSerie ='" & TRIM$(Label51) & "'")) > 0 Then
   '  Call VERIMPENSAYOS
   'End If
   
End Sub

Private Sub Command7_Click()

    Opcion% = COMALTER%("Documentos Externos\\Visualizar\Agregar Nuevos")
    
    If Opcion% = 1 Then
        TituloLista$ = "DOCUMENTOS EXTERNOS"
        
        CAMPOS$ = "Nombre/A48;Ruta/A0"
        Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
        FRMZELECHO.Caption = TituloLista$
        
        IDLIST& = ObtenerIdTablaAuxiliar("DOCEXCONFIN")
        SQLX$ = "SELECT CAMPO4, CAMPO5 FROM DATASQL WITH(NOLOCK) WHERE ID_LISTA=" & CStr(IDLIST&) & " AND CAMPO2='" & Label117 & "'"
        Set RS = READSET(SQLX$)
        With RS
            If Not (.EOF And .BOF) Then
                FIN& = RS.RecordCount
                FRMZELECHO.msf2.Rows = FIN& + 1 ' ESTABLECE DE ANTEMANO LA CANTIDAD DE FILAS A UTILIZAR
            End If
            Do While Not .EOF
                NombreDocumento$ = !CAMPO4
                RutaDocumento$ = !CAMPO5
                
                J& = J& + 1
                FRMZELECHO.msf2.TextMatrix(J&, 1) = NombreDocumento$
                FRMZELECHO.msf2.TextMatrix(J&, 2) = RutaDocumento$
                
                .MoveNext
            Loop
        End With
        On Error Resume Next
        RS.Close
        Set RS = Nothing
        '
        On Error GoTo 0
        
        Screen.MousePointer = 1
        
        FRMZELECHO.Width = 8000
50      FRMZELECHO.Show 1
    
        RutaDocumento$ = TRIM$(RESP_ZELE_VECT(2))
        If RutaDocumento$ <> "" Then
            Call ShellExecute(Me.hwnd, "Open", RutaDocumento$, "", "", 1)
        Else
            GoTo 99
        End If
            
        GoTo 50
    
    ElseIf Opcion% = 2 Then
        
        CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
        CommonDialog1.DialogTitle = "Archivo Origen de Datos"
        '     cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
        CommonDialog1.Filter = "Todos los Archivos(*.*)|*.*|"
        CommonDialog1.FilterIndex = 1
        On Error GoTo 99
        CommonDialog1.ShowOpen
        On Error GoTo 0
        '
        DoEvents
        RutaDocumento$ = TRIM$(UCase$(CommonDialog1.FileName))
        
        If RutaDocumento$ <> "" Then
            NombreDocumento$ = TRIM$(InputBox("Nombre", "Ingrese Nombre para el Documento"))
            
            If NombreDocumento$ <> "" Then
                CodigoExterno$ = Label117
                DescripcionProducto$ = Label22
                
                IDLIST& = ObtenerIdTablaAuxiliar("DOCEXCONFIN")
                
                ConsultaSql$ = "INSERT INTO DATASQL (ID_LISTA, CAMPO1, CAMPO2, CAMPO3, CAMPO4, CAMPO5) VALUES ('" & CStr(IDLIST&) & "','','" & CodigoExterno$ & "','" & DescripcionProducto$ & "','" & NombreDocumento$ & "','" & RutaDocumento$ & "')"
                FLEXCONN.Execute ConsultaSql$
            End If
        End If

    End If
    
99  Exit Sub
End Sub



Private Sub Command8_Click()

    RESPUESTA% = COMALTER%("¿Esta Seguro que Desea Quitar los Datos del Ultimo Control?\\No, Continuar\Si, Quitar")
    
    If RESPUESTA% = 2 Then
        For i% = 0 To 19
            Lis1(i%) = ""
            Lis2(i%) = ""
        Next i%
        
        For i% = 0 To 11
            TextA(i%) = ""
        Next i%
        
        Label83 = USERNAME
        
        Text5 = ""
    End If
    
End Sub

Private Sub Form_Load()
    '
    Label112 = FECHATEX$(HOY(HO$))
    '
    List1.Clear
    For i& = 1 To ULTREG&("TATIPFA")
       X$ = REGLEIDO$("TATIPFA", i&)
       If CVS(Left$(X$, 4)) > 0 Then
           TTXX$ = FORMATNUM$(CVS(Left$(X$, 4)), "F4") + " - " + Mid$(X$, 5, 44)
           List1.AddItem TTXX$
       End If
    Next i&
    Call CIERRARCH("TATIPFA")
    '
    List2.Clear
    For i& = 1 To ULTREG&("TACAUFA")
       X$ = REGLEIDO$("TACAUFA", i&)
       If CVS(Left$(X$, 4)) > 0 Then
           TTXX$ = FORMATNUM$(CVS(Left$(X$, 4)), "F4") + " - " + Mid$(X$, 5, 44)
           List2.AddItem TTXX$
       End If
    Next i&
    Call CIERRARCH("TACAUFA")
    '
End Sub

Sub VERIMPENSAYOS()
   Command4.Enabled = False
   '
   '
   Call COPYSTRU("INDIENSA", "INDIENRE")
   Call SETULTREG("!INDIENRE", 0)

   SQLX$ = "SELECT DISTINCT NumeroSerie, REVISION, FECH_ENSA FROM CONTROLFIN "
   SQLX$ = SQLX$ + " WHERE NumeroSerie = '" & Label51 & "'"
   SQLX$ = SQLX$ + " AND ORIGEN = 'REP'"
   Dim CTRLFINTMP As ADODB.Recordset
   Set CTRLFINTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   With CTRLFINTMP
     U& = 0
     Do While Not .EOF
       U& = U& + 1
       NUSRIE$ = SAFETEXT$(!NUMEROSERIE)
       RVSON% = XVALO(!Revision)
       FENSA% = CVINT(!Fech_Ensa)
       NUREV$ = "Registro de Reparacion " & NUSRIE$ & "-" & RVSON%
       Call REPLA(TX$, MKI$(FENSA%), 1, 2)
       Call REPLA(TX$, NUREV$, 3, 62)
       Call GRAREG("!INDIENRE", TX$, U&)
       .MoveNext
     Loop
   End With
   CTRLFINTMP.Close
   Set CTRLFINTMP = Nothing
   '
   '
   Call SETULTREG("!INDIENRE", U&)
   Call FRESHECHO("!INDIENRE")
   '
10 If ULTREG&("!INDIENRE") > 0 Then
     Call SELECHO("!INDIENRE/Indice de Registro de Reparaciones")
   Else
     GoTo 99
   End If
   '
   NUDOCU$ = VALACT2$("!INDIENRE") ' MUESTRO EL FORMULARIO IMPRESO  SELECCIONAD0
   If NUDOCU$ = "" Then GoTo 99 ' por si no elige ninguno y cierra el indice
    Screen.MousePointer = 11
    '
    POS% = InStr(NUDOCU$, TRIM$(Label51)) ' ASIGNO A POS% LA POSICION DESDE DONDE EMPIEZA EL NUMERO DE SEREIE
    TIDOCUM$ = TRIM$(Mid$(NUDOCU$, 1, POS% - 1)) ' ASIGNO A TIDOCUM SOLO EL TIPO DE ENSAYO QUE VA A SER DESDE 1 HASTA EL VALOR DE POS% - 1
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
         XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
         If InStr(XX$, TIDOCUM$) > 4 Then
           PPXX% = InStr(XX$, "Nro.")
           If PPXX% > 0 Then
             XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
           End If
           If InStr(XX$, NUDOCU$) > 4 Then
             DOCUNU& = CVS(Left$(XX$, 4))
             GoTo 5
           End If
         End If
    Next UI&
       '
    Screen.MousePointer = 1
    Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Ensayo.\ ")
    GoTo 10
       '
5   Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
   
   GoTo 10
99 Command4.Enabled = True
End Sub
Private Sub Image1_Click()
   If DERACCE("MODENS70", "Modificacion de Informe de Ensayo") > 1 Then
        Image3.Visible = True
        Image1.Visible = False
        Command3.Enabled = False
        Command23.Enabled = True ' HABILITA LA IMPRESORA
        ULTREVIS% = XVALO(Label85)
        'SI YA TIENE ALGUNA REVISION MUESTRA LA FECHA DE LA REVISION
        If ULTREVIS% > 0 Then
          FEC = VALOBADA("TANUMSE", "FECH_ENSA", "NumeroSerie ='" & Label51 & "'")
          FE% = CVINT(FEC) ' SI VUELVE A ESTAR EL  CANDADO VISIBLE PONGO
          Label112 = FECHATEX$(FE%) ' LA FECHA QUE TIENE EL ENSAYO GRABADA
        Else 'SI ES LA PRIMERA REVISION MUESTRA LA DEL DIA
          HOII% = HOY(HOS$)
          Label112 = FECHATEX$(HOII%)
        End If
   End If
End Sub

Private Sub Image3_Click()
   If DERACCE("MODCONFI", "Modificacion  de Registro de Control Final") > 1 Then
        Image1.Visible = True
        Image3.Visible = False
        Command3.Enabled = True
        Command23.Enabled = False
        HOII% = HOY(HOS$)
        Label112 = HOS$ ' PONGO LA FECHA ACTUAL
   End If
End Sub



Private Sub Label25_Change()
    '
    NUREP& = XVALO(Label25)
    NUMEREPA$ = "OR-" & TRIM$(Str$(NUREP&))
    '
    CONDI$ = "NUMREPA = " & NUREP& & ""
    Label22 = VALOBADA("REPARA", "DESCRIPCION", CONDI$, "NOMESS")
    Label26 = VALOBADA("REPARA", "NUMEROSERIE", CONDI$, "NOMESS")
    Label20 = VALOBADA("REPARA", "DEFEC_INFOR", CONDI$, "NOMESS")
    Label28 = VALOBADA("REPARA", "ACCESORIOS", CONDI$, "NOMESS")
    
    Text1 = VALOBADA("REPARA", "REACONDI", CONDI$, "NOMESS")
    
    REPAX$ = VALOBADA("REPARA", "REPARACION", CONDI$, "NOMESS")
    If TRIM$(REPAX$) = "" Then REPAX$ = VALOBADA("REPARA", "TRAB_EFECT", CONDI$, "NOMESS")
    Text3 = REPAX$
    
    Text2 = VALOBADA("REPARA", "OBSERVA", CONDI$, "NOMESS")
    
    Label2 = VALOBADA("REPARA", "TEC_REPAR", CONDI$, "NOMESS")
    If TRIM$(Label2) = "" Then Label2 = USERNAME
    '
    ' tipo de falla
    For KKI% = 1 To List1.ListCount
       List1.Selected(KKI% - 1) = False
    Next KKI%
    '
    SQLX$ = "SELECT * FROM TIPOFALL WHERE NUMREPA = " & NUREP& & " AND ORIGEN = 2"
    Dim FASEL As ADODB.Recordset
    Set FASEL = New ADODB.Recordset
    Set FASEL = FLEXCONN.Execute(SQLX$)
    With FASEL
       Do While Not .EOF
          NUFAL& = XVALO(!CODIFALL)
          DEFAL$ = SAFETEXT$(!DESCFALL)
          For KKI% = 1 To List1.ListCount
             If XVALO(Left$(List1.List(KKI% - 1), 4)) = NUFAL Then
                List1.Selected(KKI% - 1) = True
                GoTo 14
             End If
          Next KKI%
14     .MoveNext
       Loop
    End With
    FASEL.Close
    Set FASEL = Nothing
    '
    ' CAUSA DE FALLA
    SQLX$ = "SELECT * FROM CAUSFALL WHERE NUMREPA = " & NUREP&
    Set FASEL = New ADODB.Recordset
    Set FASEL = FLEXCONN.Execute(SQLX$)
    With FASEL
       Do While Not .EOF
          NUFAL& = XVALO(!CODIFALL)
          DEFAL$ = SAFETEXT$(!DESCFALL)
          For KKI% = 1 To List2.ListCount
             If XVALO(Left$(List2.List(KKI% - 1), 4)) = NUFAL Then
                List2.Selected(KKI% - 1) = True
                GoTo 24
             End If
          Next KKI%
24     .MoveNext
       Loop
    End With
    FASEL.Close
    Set FASEL = Nothing
    '
End Sub

Sub LIMPIAR()
     Label22 = ""
     Label25 = ""
     Label26 = ""
     Label28 = ""
End Sub

Private Sub Label51_Change()

   Image3.Visible = False
   Image1.Enabled = False
   Image1.Visible = True
   Command23.Enabled = True
   Call BLANFORMU
   NUSERIE$ = TRIM$(Label51)
   Label83 = USERNAME
   
   'Call LIMPIAINFORMACION
   If InStr(NUSERIE$, "'") < 1 And NUSERIE$ <> "" Then

       SQLX$ = "SELECT * FROM TANUMSE"
       SQLX$ = SQLX$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
       Dim TNMSETMP As ADODB.Recordset
       Set TNMSETMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
       With TNMSETMP
            If Not .EOF Then
                
                CI% = XVALO(!COD_INTE)
                Label115 = SAFETEXT$(!IDSUBOR)
                Label117 = CODEXT$(CI%)
                'Label13 = DESCRIT0$(CI%)
                Text4 = SAFETEXT$(!Obse_Ensa)
                Text5 = SAFETEXT$(!Dic_Ensa)
                Label88 = SAFETEXT$(!TEN_ENSA)
                Text8 = XVALO(!TEN_ENSA)
                fecccha$ = SAFETEXT$(!Fech_Ensa) ' si no tiene fecha pone la de hoy
                If fecccha$ <> "" Then fff% = CVINT(fecccha$): Label112 = FECHATEX(fff%)
                If TRIM$(fecccha$) = "" Then FF% = HOY(HO$): Label112 = HO$
    
                GoSub CargaRestDatos
            End If
      End With
CargaRestDatos:
       '
       CONDI$ = " NumeroSerie ='" & NUSERIE$ & "'  and ORIGEN = 'REP' "
       ULTREVIS% = XVALO(VALOBADA("CONTROLFIN", "MAX(REVISION)", CONDI$, "NOMESS"))
       Label85 = ULTREVIS%
       UsuarioRealizoControl$ = TRIM$(VALOBADA("REPARA", "UsRealizaRepa", "NUMREPA = " & Label25))
       If UsuarioRealizoControl$ <> "" Then Label83 = UsuarioRealizoControl$
       
       'AGREGAR VALIDACION SI ES LA PRIMERA REVISION QUE PRESENTE FECHA DE HOY
       If ULTREVIS% < 1 Then
          HOII% = HOY(HOS$)
          Label112 = FECHATEX$(HOII%)
       End If
       '
'       If ULTREVIS% > 0 Then ' SI TIENE REVISION DE REPARACION MUESTRA SI NO NO
       SQLX$ = "SELECT * FROM CONTROLFIN WITH(NOLOCK)"
       SQLX$ = SQLX$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
       SQLX$ = SQLX$ + " AND REVISION =" & ULTREVIS% & ""
       SQLX$ = SQLX$ + " AND ORIGEN = 'REP' "
       Dim CTRLFINTMP As ADODB.Recordset
       Set CTRLFINTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
       With CTRLFINTMP
        If Not .EOF Then
         Call bloquear
         Do While Not .EOF
           IND% = XVALO(!NumeroOrden)
           If !Resultado = "OK" Then Lis1(IND% - 1) = SAFETEXT$(!Resultado)
           If !Resultado = "NC" Then Lis2(IND% - 1) = SAFETEXT$(!Resultado)
           If !Resultado = "" Then Lis2(IND% - 1) = ""
         .MoveNext
         Loop
        Else
         Command23.Enabled = False
        End If
       End With
       '
       J& = 0
       SQLX$ = "SELECT * FROM MEDIFIN WITH(NOLOCK)"
       SQLX$ = SQLX$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
       SQLX$ = SQLX$ + " AND REVISION =" & ULTREVIS% & ""
       SQLX$ = SQLX$ + " AND ORIGEN = 'REP' "
       SQLX$ = SQLX$ + " ORDER BY CHN "

       Dim MEDIFINTMP As ADODB.Recordset
       Set MEDIFINTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
       With MEDIFINTMP
        If Not .EOF Then
         Call bloquear
       
         Do While Not .EOF
           If J& > 11 Then Exit Do
           TextA(J&) = FORMATNUM$(XVALO(!MEDA), "F7.2")  '
           J& = J& + 1
           .MoveNext
         Loop
        End If
       End With
       
       IDLIST& = ObtenerIdTablaAuxiliar("CANCONART")
       Codigo$ = TRIM$(Label117)
       CantidadControlesArticulo = XVALO(VALOBADA("DATASQL", "CAMPO3", "ID_LISTA=" & CStr(IDLIST&) & " AND CAMPO1='" & Codigo$ & "'"))
       
   End If

End Sub

Sub bloquear()
    Image3.Visible = True
    Image1.Visible = False
    Command3.Enabled = False
    Image1.Enabled = True
End Sub

Private Sub Label83_DblClick()

    CAMPOS$ = "Código/f6;Nombre/A32"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
    FRMZELECHO.Caption = "TABLA DE TÉCNICOS"
    
    Dim obj As New RECORXET
    Set RS1 = obj.RS_TECNICOS_MDT()
    Call Carga_MSF_Recordset(RS1, CAMPOS$, FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 8000
50  FRMZELECHO.Show 1

    Label83 = TRIM(RESP_ZELE_VECT(2))

End Sub

Private Sub Lis1_DblClick(Index As Integer)
  If Image3.Visible = True Then Text5.SetFocus: Exit Sub

   '\\\OT-08-0933-OD-06/10/08///
    If Lis1(Index) = "" Then
       Lis1(Index) = "OK"
       Lis2(Index) = ""
     Else
       Lis1(Index) = ""
    End If
   '\\\OT-08-0933-OD-FIN///
   Call VeredGeneral 'MEDITEA-VALIDAR OK Y NC PARA ANOTAR EL DICTAMEN
   
End Sub

Sub VeredGeneral()
  OK% = 0
  NC% = 0
  For i& = 0 To Lis2.UBound
    If Lis2(i&) = "NC" Then
      Text5 = "R"
      Exit Sub
    End If
  Next i&

  For i& = 0 To Lis1.UBound
    If Lis1(i&) = "OK" Then
      OK% = OK% + 1
    End If
  Next i&
  
  If OK% = Lis1.UBound + 1 Or (OK% >= CantidadControlesArticulo And CantidadControlesArticulo > 0) Then
    Text5 = "A"
  Else
    Text5 = ""
  End If
End Sub

Private Sub Lis2_DblClick(Index As Integer)
   '\\\OT-08-0933-OD-06/10/08///
    If Image3.Visible = True Then Text5.SetFocus: Exit Sub

    If Lis2(Index) = "" Then
       Lis2(Index) = "NC"
       Lis1(Index) = ""
     Else
       Lis2(Index) = ""
    End If
   
    Call VeredGeneral 'VALIDAR OK Y NC PARA ANOTAR EL DICTAMEN
End Sub


Private Sub Text2_Change()
'   If XVALO(Text2) > 0 Then
'     Label5 = ECOARCH$("TACAUFA", MKS$(XVALO(Text2)))
'   Else
'     Label5 = ""
'   End If
End Sub

Private Sub Text2_DblClick()
   'Call Command15_Click
End Sub

Private Sub Text5_Change()
   If XVALO(Text5) > 0 Then
      Label2 = ECOARCH$("TATIPFA", MKS$(XVALO(Text5)))
   Else
      Label2 = ""
   End If

End Sub



Private Sub TextA_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Index = 11 Then
            TextA(0).SetFocus
        Else
            TextA(Index + 1).SetFocus
        End If
    End If
    
End Sub

Private Sub TextA_LostFocus(Index As Integer)

  VALOR1 = ROUND#(XVALO(TextA(Index)), 2)

  If XVALO(VALOR1) < 10000 Then
   TextA(Index) = FORMATNUM(TextA(Index), "F7.2")
  Else
    Call MENSERR(24, "Sólo Se Aceptan Valores Menores a 10000.\Debe Corregir Para Poder Continuar.")
    TextA(Index).SetFocus
  End If
  
End Sub


