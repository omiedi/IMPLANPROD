VERSION 5.00
Begin VB.Form FISECOM09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seguimiento Comercial FLEXOFT"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11865
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FISECOM09.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11865
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame8 
      Caption         =   "Contactos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1575
      Left            =   120
      TabIndex        =   105
      Top             =   4350
      Width           =   10620
      Begin VB.Frame Frame14 
         BorderStyle     =   0  'None
         Caption         =   "Frame14"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   45
         TabIndex        =   106
         Top             =   510
         Width           =   10485
         Begin VB.Frame Frame15 
            BorderStyle     =   0  'None
            Caption         =   "Frame15"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2800
            Left            =   0
            TabIndex        =   120
            Top             =   0
            Width           =   10350
            Begin VB.PictureBox Picture5 
               BorderStyle     =   0  'None
               Height          =   2700
               Index           =   4
               Left            =   9000
               ScaleHeight     =   2700
               ScaleWidth      =   4230
               TabIndex        =   121
               Top             =   0
               Width           =   4230
               Begin VB.TextBox Text30 
                  Height          =   285
                  Index           =   8
                  Left            =   105
                  TabIndex        =   130
                  Top             =   2400
                  Width           =   4100
               End
               Begin VB.TextBox Text30 
                  Height          =   285
                  Index           =   7
                  Left            =   105
                  TabIndex        =   129
                  Top             =   2100
                  Width           =   4100
               End
               Begin VB.TextBox Text30 
                  Height          =   285
                  Index           =   6
                  Left            =   105
                  TabIndex        =   128
                  Top             =   1800
                  Width           =   4100
               End
               Begin VB.TextBox Text30 
                  Height          =   285
                  Index           =   5
                  Left            =   105
                  TabIndex        =   127
                  Top             =   1500
                  Width           =   4100
               End
               Begin VB.TextBox Text30 
                  Height          =   285
                  Index           =   4
                  Left            =   105
                  TabIndex        =   126
                  Top             =   1200
                  Width           =   4100
               End
               Begin VB.TextBox Text30 
                  Height          =   285
                  Index           =   3
                  Left            =   105
                  TabIndex        =   125
                  Top             =   900
                  Width           =   4100
               End
               Begin VB.TextBox Text30 
                  Height          =   285
                  Index           =   0
                  Left            =   105
                  TabIndex        =   124
                  Top             =   0
                  Width           =   4100
               End
               Begin VB.TextBox Text30 
                  Height          =   285
                  Index           =   1
                  Left            =   105
                  TabIndex        =   123
                  Top             =   300
                  Width           =   4100
               End
               Begin VB.TextBox Text30 
                  Height          =   285
                  Index           =   2
                  Left            =   105
                  TabIndex        =   122
                  Top             =   600
                  Width           =   4100
               End
            End
            Begin VB.PictureBox Picture5 
               BorderStyle     =   0  'None
               Height          =   2700
               Index           =   3
               Left            =   7600
               ScaleHeight     =   2700
               ScaleWidth      =   4230
               TabIndex        =   131
               Top             =   0
               Width           =   4230
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   8
                  Left            =   105
                  TabIndex        =   140
                  Top             =   2400
                  Width           =   4100
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   7
                  Left            =   105
                  TabIndex        =   139
                  Top             =   2100
                  Width           =   4100
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   6
                  Left            =   105
                  TabIndex        =   138
                  Top             =   1800
                  Width           =   4100
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   5
                  Left            =   105
                  TabIndex        =   137
                  Top             =   1500
                  Width           =   4100
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   4
                  Left            =   105
                  TabIndex        =   136
                  Top             =   1200
                  Width           =   4100
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   3
                  Left            =   105
                  TabIndex        =   135
                  Top             =   900
                  Width           =   4100
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   0
                  Left            =   105
                  TabIndex        =   134
                  Top             =   0
                  Width           =   4100
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   1
                  Left            =   105
                  TabIndex        =   133
                  Top             =   300
                  Width           =   4100
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   2
                  Left            =   105
                  TabIndex        =   132
                  Top             =   600
                  Width           =   4100
               End
            End
            Begin VB.PictureBox Picture5 
               BorderStyle     =   0  'None
               Height          =   2700
               Index           =   2
               Left            =   5500
               ScaleHeight     =   2700
               ScaleWidth      =   5175
               TabIndex        =   141
               Top             =   0
               Width           =   5175
               Begin VB.TextBox Text8 
                  Height          =   285
                  Index           =   8
                  Left            =   105
                  TabIndex        =   150
                  Top             =   2400
                  Width           =   5050
               End
               Begin VB.TextBox Text8 
                  Height          =   285
                  Index           =   7
                  Left            =   105
                  TabIndex        =   149
                  Top             =   2100
                  Width           =   5050
               End
               Begin VB.TextBox Text8 
                  Height          =   285
                  Index           =   6
                  Left            =   105
                  TabIndex        =   148
                  Top             =   1800
                  Width           =   5050
               End
               Begin VB.TextBox Text8 
                  Height          =   285
                  Index           =   5
                  Left            =   105
                  TabIndex        =   147
                  Top             =   1500
                  Width           =   5050
               End
               Begin VB.TextBox Text8 
                  Height          =   285
                  Index           =   4
                  Left            =   105
                  TabIndex        =   146
                  Top             =   1200
                  Width           =   5050
               End
               Begin VB.TextBox Text8 
                  Height          =   285
                  Index           =   3
                  Left            =   105
                  TabIndex        =   145
                  Top             =   900
                  Width           =   5050
               End
               Begin VB.TextBox Text8 
                  Height          =   285
                  Index           =   0
                  Left            =   120
                  TabIndex        =   144
                  ToolTipText     =   "Doble-Click para Enviar un Mail"
                  Top             =   0
                  Width           =   5050
               End
               Begin VB.TextBox Text8 
                  Height          =   285
                  Index           =   1
                  Left            =   105
                  TabIndex        =   143
                  Top             =   300
                  Width           =   5050
               End
               Begin VB.TextBox Text8 
                  Height          =   285
                  Index           =   2
                  Left            =   105
                  TabIndex        =   142
                  Top             =   600
                  Width           =   5050
               End
            End
            Begin VB.PictureBox Picture5 
               BorderStyle     =   0  'None
               Height          =   2700
               Index           =   1
               Left            =   3000
               ScaleHeight     =   2700
               ScaleWidth      =   5175
               TabIndex        =   151
               Top             =   0
               Width           =   5175
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   2
                  Left            =   105
                  TabIndex        =   160
                  Top             =   600
                  Width           =   5000
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   1
                  Left            =   105
                  TabIndex        =   159
                  Top             =   300
                  Width           =   5000
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   0
                  Left            =   105
                  TabIndex        =   158
                  Top             =   0
                  Width           =   5000
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   3
                  Left            =   105
                  TabIndex        =   157
                  Top             =   900
                  Width           =   5000
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   4
                  Left            =   105
                  TabIndex        =   156
                  Top             =   1200
                  Width           =   5000
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   5
                  Left            =   105
                  TabIndex        =   155
                  Top             =   1500
                  Width           =   5000
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   6
                  Left            =   105
                  TabIndex        =   154
                  Top             =   1800
                  Width           =   5000
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   7
                  Left            =   105
                  TabIndex        =   153
                  Top             =   2100
                  Width           =   5000
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   8
                  Left            =   105
                  TabIndex        =   152
                  Top             =   2400
                  Width           =   5000
               End
            End
            Begin VB.PictureBox Picture5 
               BorderStyle     =   0  'None
               Height          =   2700
               Index           =   0
               Left            =   250
               ScaleHeight     =   2700
               ScaleWidth      =   6120
               TabIndex        =   161
               Top             =   0
               Width           =   6120
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   2
                  Left            =   0
                  TabIndex        =   170
                  Top             =   600
                  Width           =   6000
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   1
                  Left            =   0
                  TabIndex        =   169
                  Top             =   300
                  Width           =   6000
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   0
                  Left            =   0
                  TabIndex        =   168
                  ToolTipText     =   "Doble-Click para Enviar un Mail"
                  Top             =   0
                  Width           =   6000
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   3
                  Left            =   0
                  TabIndex        =   167
                  Top             =   900
                  Width           =   6000
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   4
                  Left            =   0
                  TabIndex        =   166
                  Top             =   1200
                  Width           =   6000
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   5
                  Left            =   0
                  TabIndex        =   165
                  Top             =   1500
                  Width           =   6000
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   6
                  Left            =   0
                  TabIndex        =   164
                  Top             =   1800
                  Width           =   6000
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   7
                  Left            =   0
                  TabIndex        =   163
                  Top             =   2100
                  Width           =   6000
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   8
                  Left            =   0
                  TabIndex        =   162
                  Top             =   2400
                  Width           =   6000
               End
            End
            Begin VB.Label LEFE 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "3)"
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
               Index           =   2
               Left            =   -1365
               TabIndex        =   179
               Top             =   630
               Width           =   1590
            End
            Begin VB.Label LEFE 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "2)"
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
               Index           =   1
               Left            =   -1365
               TabIndex        =   178
               Top             =   330
               Width           =   1590
            End
            Begin VB.Label LEFE 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "1)"
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
               Index           =   0
               Left            =   -1365
               TabIndex        =   177
               Top             =   45
               Width           =   1590
            End
            Begin VB.Label LEFE 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "6)"
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
               Index           =   6
               Left            =   -1365
               TabIndex        =   176
               Top             =   1530
               Width           =   1590
            End
            Begin VB.Label LEFE 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "5)"
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
               Index           =   7
               Left            =   -1365
               TabIndex        =   175
               Top             =   1230
               Width           =   1590
            End
            Begin VB.Label LEFE 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "4)"
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
               Index           =   8
               Left            =   -1365
               TabIndex        =   174
               Top             =   930
               Width           =   1590
            End
            Begin VB.Label LEFE 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "9)"
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
               Index           =   9
               Left            =   -1365
               TabIndex        =   173
               Top             =   2430
               Width           =   1590
            End
            Begin VB.Label LEFE 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "8)"
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
               Index           =   10
               Left            =   -1365
               TabIndex        =   172
               Top             =   2130
               Width           =   1590
            End
            Begin VB.Label LEFE 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "7)"
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
               Index           =   11
               Left            =   -1365
               TabIndex        =   171
               Top             =   1830
               Width           =   1590
            End
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   960
            LargeChange     =   7
            Left            =   10360
            TabIndex        =   107
            Top             =   0
            Width           =   120
         End
      End
      Begin VB.Label LabeL9 
         Caption         =   "Celular:"
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
         Index           =   4
         Left            =   9135
         TabIndex        =   108
         Top             =   270
         Width           =   6000
      End
      Begin VB.Label LabeL9 
         Caption         =   "Telefono:"
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
         Index           =   3
         Left            =   7770
         TabIndex        =   180
         Top             =   270
         Width           =   6000
      End
      Begin VB.Label LabeL9 
         Caption         =   "E-Mail:"
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
         Index           =   2
         Left            =   5670
         TabIndex        =   109
         Top             =   270
         Width           =   6000
      End
      Begin VB.Label LabeL9 
         Caption         =   "Cargo / Función - Sector:"
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
         Index           =   1
         Left            =   3150
         TabIndex        =   110
         Top             =   270
         Width           =   6000
      End
      Begin VB.Label LabeL9 
         Caption         =   "Persona de Contacto:"
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
         Index           =   0
         Left            =   270
         TabIndex        =   111
         Top             =   270
         Width           =   6000
      End
   End
   Begin VB.CommandButton Command30 
      Caption         =   "x"
      Height          =   345
      Left            =   10200
      TabIndex        =   104
      Top             =   0
      Width           =   495
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8355
      Left            =   10800
      ScaleHeight     =   8295
      ScaleWidth      =   1035
      TabIndex        =   86
      Top             =   0
      Width           =   1095
      Begin VB.CommandButton Command32 
         Height          =   600
         Left            =   160
         Picture         =   "FISECOM09.frx":1272
         Style           =   1  'Graphical
         TabIndex        =   181
         ToolTipText     =   "Generar un Presupuesto"
         Top             =   3120
         Width           =   690
      End
      Begin VB.CommandButton AGenRep 
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
         Left            =   160
         Picture         =   "FISECOM09.frx":13BC
         Style           =   1  'Graphical
         TabIndex        =   99
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   3720
         Width           =   690
      End
      Begin VB.CommandButton Command1 
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
         Height          =   600
         Left            =   165
         Picture         =   "FISECOM09.frx":16C6
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   1320
         Width           =   690
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H80000012&
         Caption         =   "Varios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   2895
         Left            =   40
         TabIndex        =   94
         Top             =   4410
         Width           =   920
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
            Height          =   600
            Left            =   120
            Picture         =   "FISECOM09.frx":19D0
            Style           =   1  'Graphical
            TabIndex        =   118
            ToolTipText     =   "Contactos Programados"
            Top             =   315
            Width           =   690
         End
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
            Left            =   120
            Picture         =   "FISECOM09.frx":2112
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Configuración de Funcionamiento"
            Top             =   910
            Width           =   690
         End
         Begin VB.CommandButton Command18 
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
            Left            =   120
            Picture         =   "FISECOM09.frx":2554
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Listados por Impresora"
            Top             =   1520
            Width           =   690
         End
         Begin VB.CommandButton Command20 
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
            Left            =   120
            Picture         =   "FISECOM09.frx":285E
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Administrar Mail Circular"
            Top             =   2120
            Width           =   690
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000012&
         Caption         =   "Exp/Imp"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1065
         Left            =   15
         TabIndex        =   89
         Top             =   3255
         Visible         =   0   'False
         Width           =   990
         Begin VB.CommandButton Command7 
            Caption         =   "B"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   105
            TabIndex        =   93
            ToolTipText     =   "Confecciona Copia de Seguridad (Zip)"
            Top             =   315
            Width           =   330
         End
         Begin VB.CommandButton Command8 
            Caption         =   "R"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   525
            TabIndex        =   92
            ToolTipText     =   "Restaura Copia de Seguridad (Zip)"
            Top             =   315
            Width           =   330
         End
         Begin VB.CommandButton Command10 
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   525
            TabIndex        =   91
            ToolTipText     =   "Exporta Direcciones para Outlook Express"
            Top             =   650
            Width           =   330
         End
         Begin VB.CommandButton Command12 
            Caption         =   "I"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   105
            TabIndex        =   90
            ToolTipText     =   "Importa Libreta de Direcciones (de C:\FLEXOFT)"
            Top             =   650
            Width           =   330
         End
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Sort"
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
         Left            =   160
         Picture         =   "FISECOM09.frx":2B68
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Ordenamiento Alfabético"
         Top             =   2520
         Width           =   690
      End
      Begin VB.CommandButton Command3 
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
         Height          =   600
         Left            =   160
         Picture         =   "FISECOM09.frx":2E72
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   120
         Width           =   690
      End
      Begin VB.CommandButton Command21 
         Caption         =   "New"
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
         Left            =   160
         Picture         =   "FISECOM09.frx":2FBC
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Alta de Nueva Ficha de Seguimiento"
         Top             =   720
         Width           =   690
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Erase"
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
         Left            =   160
         Picture         =   "FISECOM09.frx":3106
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Elimina Contacto Activo de la Base"
         Top             =   1920
         Width           =   690
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   160
         ScaleHeight     =   75
         ScaleWidth      =   675
         TabIndex        =   87
         Top             =   7440
         Width           =   740
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   0
            TabIndex        =   88
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -240
         Picture         =   "FISECOM09.frx":3410
         Top             =   7680
         Width           =   1515
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Misc."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   2145
      Left            =   9660
      TabIndex        =   68
      Top             =   5970
      Width           =   1065
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
         Height          =   600
         Left            =   105
         Picture         =   "FISECOM09.frx":5332
         Style           =   1  'Graphical
         TabIndex        =   117
         ToolTipText     =   "Imprime Ficha del Contacto"
         Top             =   1420
         Width           =   850
      End
      Begin VB.CommandButton Command17 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   105
         Picture         =   "FISECOM09.frx":599C
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Documentos Conexos"
         Top             =   870
         Width           =   850
      End
      Begin VB.CommandButton Command11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   105
         Picture         =   "FISECOM09.frx":5AE6
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Abrir Anotador de Novedades"
         Top             =   320
         Width           =   850
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   4245
      Left            =   105
      TabIndex        =   31
      Top             =   75
      Width           =   10620
      Begin VB.TextBox Text22 
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
         Left            =   3675
         TabIndex        =   5
         Text            =   " "
         Top             =   1665
         Width           =   3375
      End
      Begin VB.CommandButton Command33 
         BackColor       =   &H00C0C0C0&
         Height          =   580
         Left            =   240
         Picture         =   "FISECOM09.frx":5DF0
         Style           =   1  'Graphical
         TabIndex        =   100
         ToolTipText     =   "Oportunidades para este Prospecto"
         Top             =   3540
         Width           =   700
      End
      Begin VB.CommandButton Command35 
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
         ToolTipText     =   "Selecciona Contacto - Lista Completa"
         Top             =   2040
         Width           =   175
      End
      Begin VB.TextBox Text27 
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
         Left            =   1155
         TabIndex        =   6
         Top             =   2040
         Width           =   555
      End
      Begin VB.Frame Frame7 
         Caption         =   "Representante"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   4920
         TabIndex        =   84
         Top             =   1935
         Width           =   2445
         Begin VB.CommandButton Command31 
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
            Height          =   235
            Left            =   2190
            TabIndex        =   11
            ToolTipText     =   "Selecciona Contacto - Lista Completa"
            Top             =   165
            Width           =   150
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
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
            Height          =   225
            Left            =   1680
            TabIndex        =   98
            Top             =   180
            Width           =   525
         End
         Begin VB.Label LABEL40 
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
            Height          =   225
            Left            =   120
            TabIndex        =   85
            Top             =   480
            Width           =   2205
         End
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
         Height          =   280
         Left            =   7500
         Picture         =   "FISECOM09.frx":6232
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Marca / Quita Prioridad"
         Top             =   330
         Width           =   285
      End
      Begin VB.TextBox Text15 
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
         Left            =   7245
         TabIndex        =   83
         Text            =   " "
         Top             =   330
         Width           =   330
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   9345
         TabIndex        =   82
         Text            =   " "
         Top             =   1050
         Width           =   1065
      End
      Begin VB.CommandButton Command28 
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
         Left            =   8925
         TabIndex        =   78
         Top             =   1050
         Width           =   150
      End
      Begin VB.TextBox Text13 
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
         Left            =   1155
         TabIndex        =   2
         Text            =   " "
         ToolTipText     =   "Doble-Click para Enviar un Mail"
         Top             =   960
         Width           =   6630
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   8820
         TabIndex        =   74
         Top             =   3885
         Width           =   790
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   9620
         TabIndex        =   73
         Top             =   3885
         Width           =   790
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8820
         Sorted          =   -1  'True
         TabIndex        =   18
         ToolTipText     =   "Doble Click para Actualizar Tabla"
         Top             =   3180
         Width           =   1590
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   1155
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   3180
         Width           =   7260
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
         Height          =   280
         Left            =   9660
         TabIndex        =   16
         Text            =   " "
         Top             =   2820
         Width           =   750
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
         Height          =   280
         Left            =   1155
         TabIndex        =   9
         Text            =   " "
         Top             =   2820
         Width           =   2115
      End
      Begin VB.TextBox Text2 
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
         Left            =   4830
         TabIndex        =   10
         Text            =   " "
         Top             =   2820
         Width           =   3585
      End
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
         Height          =   280
         Left            =   1155
         TabIndex        =   1
         Text            =   " "
         ToolTipText     =   "Doble-Click para Enviar un Mail"
         Top             =   630
         Width           =   6630
      End
      Begin VB.Frame Frame5 
         Caption         =   "Nro. Prospecto / Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   850
         Left            =   7875
         TabIndex        =   50
         Top             =   150
         Width           =   2720
         Begin VB.CommandButton Command2 
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   2130
            Picture         =   "FISECOM09.frx":6674
            Style           =   1  'Graphical
            TabIndex        =   103
            ToolTipText     =   "Búsqueda Secuencial por Texto"
            Top             =   240
            Width           =   480
         End
         Begin VB.CommandButton Command13 
            Caption         =   "!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   260
            Left            =   1920
            TabIndex        =   101
            ToolTipText     =   "Contactos Programados"
            Top             =   240
            Width           =   160
         End
         Begin VB.CommandButton COMMAND16 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   260
            Left            =   1770
            TabIndex        =   96
            ToolTipText     =   "Lista Completa de Contactos"
            Top             =   240
            Width           =   150
         End
         Begin VB.CommandButton Command27 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   260
            Left            =   1920
            Picture         =   "FISECOM09.frx":697E
            TabIndex        =   76
            ToolTipText     =   "Contacto Siguiente"
            Top             =   525
            Width           =   160
         End
         Begin VB.CommandButton Command26 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   260
            Left            =   1770
            Picture         =   "FISECOM09.frx":6C88
            TabIndex        =   75
            ToolTipText     =   "Contacto Anterior"
            Top             =   525
            Width           =   160
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            Caption         =   "Cte:"
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
            Left            =   100
            TabIndex        =   102
            Top             =   550
            Width           =   560
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
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
            Height          =   255
            Left            =   700
            TabIndex        =   97
            Top             =   525
            Width           =   990
         End
         Begin VB.Label Label26 
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
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   1575
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
         Height          =   280
         Left            =   1155
         TabIndex        =   0
         Text            =   " "
         Top             =   330
         Width           =   6420
      End
      Begin VB.TextBox Text26 
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
         Left            =   7875
         TabIndex        =   13
         Text            =   " "
         Top             =   1770
         Width           =   2150
      End
      Begin VB.TextBox Text24 
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
         Left            =   7875
         TabIndex        =   15
         Text            =   " "
         Top             =   2400
         Width           =   2150
      End
      Begin VB.TextBox Text23 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1155
         TabIndex        =   8
         Text            =   " "
         Top             =   2400
         Width           =   3645
      End
      Begin VB.TextBox Text21 
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
         Left            =   1155
         TabIndex        =   4
         Top             =   1680
         Width           =   1485
      End
      Begin VB.TextBox Text20 
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
         Left            =   1155
         TabIndex        =   3
         Text            =   " "
         Top             =   1350
         Width           =   5895
      End
      Begin VB.TextBox Text18 
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
         Left            =   7875
         TabIndex        =   12
         Text            =   " "
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox Text17 
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
         Left            =   7875
         TabIndex        =   14
         Text            =   " "
         Top             =   2085
         Width           =   2535
      End
      Begin VB.Label LABEL52 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   95
         Top             =   2040
         Width           =   2865
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "M:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9135
         TabIndex        =   81
         Top             =   1080
         Width           =   1380
      End
      Begin VB.Label Label38 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8250
         TabIndex        =   80
         Top             =   1050
         Width           =   680
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Baja:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7875
         TabIndex        =   79
         ToolTipText     =   "Doble-Click para Re-Activar"
         Top             =   1080
         Width           =   1380
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         Caption         =   "Web-Site:"
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
         TabIndex        =   77
         Top             =   1065
         Width           =   960
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Gr:"
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
         Left            =   7845
         TabIndex        =   72
         ToolTipText     =   "Doble Click para Actualizar Tabla"
         Top             =   3180
         Width           =   960
      End
      Begin VB.Line Line1 
         X1              =   1155
         X2              =   8400
         Y1              =   4125
         Y2              =   4125
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Observac:"
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
         TabIndex        =   56
         Top             =   3180
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Empleados:"
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
         Left            =   8460
         TabIndex        =   55
         Top             =   2880
         Width           =   1065
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Ramo:"
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
         TabIndex        =   54
         Top             =   2925
         Width           =   960
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Especialidad:"
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
         Left            =   2970
         TabIndex        =   53
         Top             =   2880
         Width           =   1800
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Dir-elec:"
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
         TabIndex        =   52
         ToolTipText     =   "Doble Click para Generar Direcciones Automáticas a Partir del Web Site"
         Top             =   750
         Width           =   960
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Empresa:"
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
         TabIndex        =   49
         Top             =   435
         Width           =   960
      End
      Begin VB.Image Image5 
         Height          =   300
         Left            =   10080
         Picture         =   "FISECOM09.frx":6F92
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   315
      End
      Begin VB.Image Image4 
         Height          =   300
         Left            =   10080
         Picture         =   "FISECOM09.frx":729C
         Stretch         =   -1  'True
         Top             =   1770
         Width           =   315
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Calle y Nº:"
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
         TabIndex        =   39
         Top             =   1455
         Width           =   960
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Ciudad:"
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
         Left            =   2625
         TabIndex        =   38
         Top             =   1770
         Width           =   960
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "País:"
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
         TabIndex        =   37
         Top             =   2505
         Width           =   960
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Tel:"
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
         Left            =   6825
         TabIndex        =   36
         Top             =   1875
         Width           =   960
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "C.Post:"
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
         TabIndex        =   35
         Top             =   1770
         Width           =   960
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Fax:"
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
         Left            =   6825
         TabIndex        =   34
         Top             =   2505
         Width           =   960
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Provincia:"
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
         Left            =   120
         TabIndex        =   33
         Top             =   2160
         Width           =   960
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "N.Fant:"
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
         Left            =   6825
         TabIndex        =   32
         Top             =   1545
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Novedades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   2160
      Left            =   120
      TabIndex        =   57
      Top             =   5970
      Width           =   9460
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   30
         ScaleHeight     =   420
         ScaleWidth      =   9405
         TabIndex        =   112
         Top             =   1680
         Width           =   9400
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   495
            TabIndex        =   114
            Text            =   "Combo1"
            Top             =   60
            Width           =   2415
         End
         Begin VB.TextBox Text88 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3780
            TabIndex        =   113
            Top             =   60
            Width           =   5505
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Con:"
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
            Left            =   60
            TabIndex        =   116
            Top             =   120
            Width           =   420
         End
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Asunto:"
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
            Left            =   2745
            TabIndex        =   115
            Top             =   120
            Width           =   960
         End
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   600
         Width           =   7155
      End
      Begin VB.CommandButton Command15 
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
         Height          =   280
         Left            =   8295
         TabIndex        =   66
         Top             =   1350
         Width           =   175
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
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
         Left            =   7350
         TabIndex        =   67
         Top             =   1350
         Width           =   1000
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         ScaleHeight     =   240
         ScaleWidth      =   9360
         TabIndex        =   58
         Top             =   315
         Width           =   9420
         Begin VB.Label Label14 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Detalle                                                    Usuario"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   105
            TabIndex        =   60
            Top             =   20
            Width           =   9360
         End
         Begin VB.Label Label13 
            Caption         =   "Label5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   15
            Left            =   0
            TabIndex        =   59
            Top             =   420
            Width           =   1380
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   1155
            Y1              =   420
            Y2              =   420
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "O.K"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   745
         Left            =   8600
         Picture         =   "FISECOM09.frx":75A6
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   960
         Width           =   750
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   9500
         Y1              =   2130
         Y2              =   2130
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Label35"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7245
         TabIndex        =   65
         Top             =   720
         Width           =   4005
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   330
         Left            =   80
         TabIndex        =   64
         Top             =   800
         Width           =   1065
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Próx.Contacto:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7245
         TabIndex        =   63
         Top             =   1095
         Width           =   1260
      End
      Begin VB.Line Line4 
         X1              =   7140
         X2              =   9450
         Y1              =   950
         Y2              =   950
      End
      Begin VB.Line Line7 
         X1              =   9450
         X2              =   9450
         Y1              =   315
         Y2              =   2210
      End
      Begin VB.Label Label15 
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
         Height          =   225
         Left            =   8400
         TabIndex        =   61
         Top             =   840
         Width           =   4000
      End
      Begin VB.Line Line3 
         X1              =   -30
         X2              =   7120
         Y1              =   1650
         Y2              =   1650
      End
      Begin VB.Label Label32 
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
         Height          =   1170
         Left            =   7140
         TabIndex        =   62
         Top             =   600
         Width           =   2310
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   10680
      Sorted          =   -1  'True
      TabIndex        =   48
      Top             =   735
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00D0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   2835
      ScaleHeight     =   2475
      ScaleWidth      =   4155
      TabIndex        =   40
      Top             =   1890
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton Command6 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3885
         TabIndex        =   47
         Top             =   -10
         Width           =   270
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   1260
         TabIndex        =   46
         Top             =   2100
         Width           =   1485
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Levante el Receptor !!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   330
         Left            =   420
         TabIndex        =   45
         Top             =   1680
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Left            =   210
         TabIndex        =   44
         Top             =   1365
         Width           =   3795
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "45241284"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   420
         TabIndex        =   43
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Llamando a:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   315
         TabIndex        =   42
         Top             =   420
         Width           =   1800
      End
      Begin VB.Label Label25 
         BackColor       =   &H00800000&
         Caption         =   "Monitor de Estado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   4110
      End
   End
   Begin VB.Label ORIMAI 
      Height          =   330
      Left            =   1470
      TabIndex        =   119
      Top             =   0
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Apel:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   30
      Top             =   1785
      Width           =   960
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Apel:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   29
      Top             =   1260
      Width           =   960
   End
End
Attribute VB_Name = "FISECOM09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cancelflag
Dim IMPORTA$
Dim NOVALID%
Dim ARRABOR%
'
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'
Private Sub AGenRep_Click()
    'se comento por que daba un error 5 en el procedimiento
    ' text4.gostfocus ***Queda para ver***
    AGenRep.Enabled = False '
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Sub Command11_Click()
  '\\\OT-05-0966-OD-27/10/08///
   If TRIM$(Label26) = "" Then
     COMMAND16.SetFocus
     Exit Sub
   End If
   '
   'VALIDACION DE QUE LA FICHA ESTA GRABADA
   COCLI$ = TRIM$(Label26)
   'NUFICH& = XVALO(Mid$((VALOBADA("BSEGCOM", "CODI_CLI", "CODI_CLI ='" & COCLI$ & "'")), 4))
   'If NUFICH& < 1 Then Command16.SetFocus: Exit Sub
   '
   Command11.Enabled = False
   '
   If TRIM$(Label26) <> "" Then
     ANOTASEG09.Label26 = Label26
     ANOTASEG09.Label3 = Text4
     ANOTASEG09.Text12 = Text12
     '
     ANOTASEG09.List1.Clear
     Call ABRECLIEN
     ANOTACION$ = VALOBADA("BSEGCOM", "HIS_CONTAC", "CODI_CLI = '" & COCLI$ & "'")
     XX$ = ANOTACION$: LONGI& = Len(XX$)
     INICIO& = 1
     Do While INICIO& < LONGI&
         FIN& = InStr(INICIO&, XX$, Chr$(10))
         FINB& = InStr(INICIO&, XX$, CRLF$)
         If FINB& > 0 And FINB& < FIN& Then FIN& = FINB&
         If FIN& > 0 Then
            TTXX$ = Mid$(XX$, INICIO&, FIN& - INICIO&)
            ANOTASEG09.List1.AddItem TTXX$
            INICIO& = FIN& + 2
         Else
            TTXX$ = Mid$(XX$, INICIO&)
            ANOTASEG09.List1.AddItem TTXX$
            Exit Do
         End If
     Loop
     '
     TPI = ANOTASEG09.List1.ListCount - 12
     If TPI < 0 Then TPI = 0
     ANOTASEG09.List1.TopIndex = TPI
     '
     ANOTASEG09.Show 1
   End If
   '
   Command11.Enabled = True
   '
End Sub

Private Sub Command15_Click()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text12.TEXT = VDEV$
    End If
End Sub

Sub Command16_Click()
   If TRIM$(FORSELEC.Text1) = "" Then
     FORSELEC.ABUSQUE = TRIM$(Label26)
   End If
   ABUSX$ = "INDIPROS" ' : If USNBR% < 100 Then ABUSX$ = "BSEGALE"
21 Call SELECHO(ABUSX$ & "/Indice Prospect")
   VDEVU$ = VALACT1$(ABUSX$)
   If VDEVU$ <> "" Then
      Label26 = TRIM$(VDEVU$)
      If Visible = False Then
         Hide
         Show 1
         GoTo 21
      End If
   End If
End Sub

Private Sub Command17_Click()
   If TRIM$(Label26) = "" Then
     COMMAND16.SetFocus
     Exit Sub
   End If
   '
   Command17.Enabled = False
   '
   Call CREACAMPO("BADACLIEN", "DOCADCRM", "Codi_Cli", 10, 12, 1)
   Call CREACAMPO("BADACLIEN", "DOCADCRM", "NomAsignado", 10, 256)
   Call CREACAMPO("BADACLIEN", "DOCADCRM", "FILE_NAME", 10, 256)
   Call CREACAMPO("BADACLIEN", "DOCADCRM", "FILE_DATA", 12, 0)
   '
   DOCONEC09.Label26 = Label26
   DOCONEC09.Show 1
   '
   Command17.Enabled = True
   '
End Sub

Private Sub Command18_Click()
   '
   Command18.Enabled = False
   
   OPX% = ALTERNA%("Base de Datos Completa\Alarma de Prospect\Por Clasificacion")
   Select Case OPX%
     Case 1
        Screen.MousePointer = 11
        Call ActualizaArchivo(OPX%)
        Screen.MousePointer = 1
        Call FINAL("*BSEGCOM7")
     Case 2
       Call ALARMAPROSPECT
     Case 3
       Call LIPORCLASI
   End Select
   Command18.Enabled = True
   '
End Sub

Sub LIPORCLASI()
   '
   SELDESTI09.Show 1
   OPENFORMS = DoEvents
   If SELDESTI09.Label1 <> "OK" Then Exit Sub
   '
   ENVIOMASIVO% = 1
   TODOS% = 0: SELSTRING$ = String$(SELDESTI09.List5.ListCount, "0")
   If SELDESTI09.Check1.Value = 1 Then
       SELSTRING$ = String$(SELDESTI09.List5.ListCount, "1")
     Else
       For KKI% = 1 To SELDESTI09.List5.ListCount
         If SELDESTI09.List5.Selected(KKI% - 1) = True Then Mid$(SELSTRING$, KKI%) = "1"
       Next KKI%
   End If
   '
   Screen.MousePointer = 11
   Call CREACAMPO("", "BSEGCOM", "MARULIMP", 3, 2, 1)
   Call ACTUREGISTRO("BSEGCOM", "MARULIMP", "MARULIMP > 0", 0, RAFE&)
   '
   FIN& = CantRegistros("BSEGCOM")
   Dim CTACSEL As ADODB.Recordset
   Set CTASEL = New ADODB.Recordset
   SQLX$ = "SELECT * FROM BSEGCOM WITH(NOLOCK) ORDER BY CODI_CLI ASC"
   Set CTASEL = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   JJ& = 0
   With CTASEL
      Do While Not .EOF
        JJ& = JJ& + 1
        Call TRACE(20, JJ&, FIN&)
        PERLIST$ = SAFETEXT$(!GRUPOS)
        For KKI% = 1 To Len(PERLIST$)
           PERLI% = Asc(Mid$(PERLIST$, KKI%, 1))
           If PERLI% > 0 Then
             If Mid$(SELSTRING$, PERLI%, 1) = "1" Then
               CODCLISE$ = SAFETEXT$(!CODI_CLI)
               Call ACTUREGISTRO("BSEGCOM", "MARULIMP", "CODI_CLI = '" & CODCLISE & "'", 1, RAFE&)
             End If
           End If
        Next KKI%
        '
      .MoveNext
      Loop
   End With
   CTASEL.Close
   Set CTASEL = Nothing
   Screen.MousePointer = 1
   '
   Call CONECRUN("?LIPROSEL", "Listado Prospectos Seleccionados")
   '
End Sub
'
Private Sub Command19_Click()
   RCONTA$ = AJUSTI$(Text15, 3)
   If RCONTA$ = "!!!" Then
       RCONTA$ = "   "
     Else
       RCONTA$ = "!!!"
   End If
   Text15 = TRIM$(RCONTA$)
End Sub

Sub Command2_Click()
   '
   Static ARGUME$
   Command2.Enabled = False
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
10 ARGUME$ = TRIM$(UCase$(InputBox$("Buscar", "Búsqueda Cliente por Texto", ARGUME$)))
   '
   If ARGUME$ <> "" Then
     Screen.MousePointer = 11
     JJ& = 0
     SQLX$ = "SELECT * FROM BSEGCOM"
     Dim BSEGTMP As ADODB.Recordset
     Set BSEGTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
     With BSEGTMP
       Do While Not .EOF
         For KKI% = 1 To .Fields.Count
             On Error Resume Next
             VATEX$ = ""
             VATEX$ = .Fields(KKI% - 1).Value
             On Error GoTo 0
             If InStr(VATEX$, ARGUME$) > 0 Then
               Call REPLA(EL1$, !CODI_CLI, 1, 12)
               Call REPLA(EL1$, AJUSTI$(SAFETEXT$(!Raso_Cli), 50), 13, 50)   ' empresa
               JJ& = JJ& + 1
               Call GRAREG("!SECONTA7", EL1$, JJ&)
               Exit For
             End If
         Next KKI%

       
'        ABUSX$ = "BSEGCOM": If USNBR% < 100 Then ABUSX$ = "BSEGALE"
'        For U& = 1 To ULTREG&(ABUSX$)
'          XXX$ = REGLEIDO$(ABUSX$, U&)
'          If InStr(UCase$(Mid$(XXX$, 3)), ARGUME$) > 0 Then
'            JJ& = JJ& + 1
'            Call GRAREG("!SECONTAC", XXX$, JJ&)
'          End If
'        Next U&
        .MoveNext
       Loop
     End With
     '
     Call SETULTREG("!SECONTA7", JJ&)
     Call CIERRARCH("!SECONTA7")
     Call BAJALDISCO
     Call FRESHECHO("!SECONTA7")
     '
20   If JJ& > 0 Then
21     Call SELECHO("!SECONTA7")
       UCONTA$ = VALACT1$("!SECONTA7")
     
        If UCONTA$ <> "" Then
          Label26 = UCONTA$
          If Visible = False Then
            Hide
            Show 1
            GoTo 21
          End If
        End If
     End If
     '
   End If
   Screen.MousePointer = 1
   Command2.Enabled = True
End Sub

Sub Command20_Click()
   Command20.Enabled = False
   SELDESTI09.Show 1
   '
   If SELDESTI09.Label1 <> "OK" Then GoTo 99
   '
   Top = Screen.Height
   ARMANEW10.Show 1
   Call CENTRAFORM(Me)
   '
99 Command20.Enabled = True

End Sub

Private Sub Command22_Click()
   If List2.ListIndex >= 0 Then
     List2.RemoveItem (List2.ListIndex)
   End If
End Sub

Private Sub Command23_Click()
   If TRIM$(Label26) = "" Then
     COMMAND16.SetFocus
     Exit Sub
   End If
   '
   XX$ = REGSEL$("TAGRUCON")
   If Len(XX$) > 4 Then
     RAGRE% = Val(CVI(Left$(XX$, 2)))
     TAGRE$ = AJUSTI$(Mid$(XX$, 51, 24), 48) + AJUSTI$(Mid$(XX$, 3, 48), 48) + Str$(RAGRE%)
     For U& = 1 To List2.ListCount
       If Val(Mid$(List2.List(U& - 1), 97, 6)) = RAGRE% Then Exit Sub
     Next U&
     List2.AddItem TAGRE$
   End If
End Sub

Private Sub Command24_Click()
   Command24.Enabled = False
   If COMALTER%("Realmente Desea Eliminar PERMANENTEMENTE\la Ficha Activa ?\\No, Cancelar\Si, Eliminar") = 2 Then
5    Screen.MousePointer = 11
     COCLI$ = TRIM$(Label26)
     'borra de la tabla
     '
     Call ABRECLIEN
     SQLX$ = "DELETE FROM BSEGCOM WHERE CODI_CLI = '" & COCLI$ & "'"
      FLEXCONN.Execute (FILTSQL$(SQLX$))
     '
     ABASE$ = "INDIPROS": GoSub BORRABASE
     '
     Label26 = ""
   End If
   '
   Screen.MousePointer = 1
   Command24.Enabled = True
   Exit Sub
   '
   '
BORRABASE:
   JJ& = 0
   FIN& = ULTREG&(ABASE$)
   For U& = 1 To ULTREG&(ABASE$)
     Call TRACE(24, U&, FIN&)
     XX$ = REGLEIDO$(ABASE$, U&)
     IFICHB$ = Left$(XX$, 12)
     If TRIM$(IFICHB$) <> COCLI$ Then
       'If IFICHB% > IFICHA% Then IFICHB% = IFICHB% - 1
       Call REPLA(XX$, COCLI$, 1, 2)
       JJ& = JJ& + 1
       Call GRAREG(ABASE$, XX$, JJ&)
     End If
   Next U&
   Call SETULTREG(ABASE$, JJ&)
   Call CIERRARCH(ABASE$)
   Call FRESHECHO(ABASE$)

Return

End Sub

Private Sub Command26_Click()
   ABUSX$ = "INDIPROS" ': If USNBR% < 100 Then ABUSX$ = "BSEGALE"
   UKU& = REGNUM&(ABUSX$)
   If UKU& < 1 Then
      XX$ = FILTERGETRECORD$("INDIPROS", 1, Label26)
      UKU& = REGNUM&(ABUSX$)
   End If
   If UKU& > 1 Then
     XX$ = REGLEIDO$(ABUSX$, UKU& - 1)
     Label26 = TRIM$(Left$(XX$, 15))
   End If
End Sub

Private Sub Command27_Click()
   ABUSX$ = "INDIPROS"
   UKU& = REGNUM&(ABUSX$)
   If UKU& < 1 Then
      XX$ = FILTERGETRECORD$("INDIPROS", 1, Label26)
      UKU& = REGNUM&(ABUSX$)
   End If
   If UKU& > 0 Then
      XX$ = REGLEIDO$(ABUSX$, UKU& + 1)
      Label26 = TRIM$(Left$(XX$, 12))
   End If

End Sub

Private Sub Command28_Click()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Label38 = VDEV$
    End If
End Sub


Private Sub Command29_Click()
    Call CONACT09.Command1_Click
    CONACT09.Show 1
End Sub

Private Sub Command30_Click()
    '
    NOVALID% = 1
'Exit Sub
    Call ABRECONEXION
    SQLX$ = "SELECT * FROM BSEGCOM WITH(NOLOCK)"
    Dim SECOM As ADODB.Recordset
    Set SECOM = New ADODB.Recordset
    Set SECOM = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With SECOM
      Do While Not .EOF
        COCLI$ = SAFETEXT$(!CODI_CLI)
        MAICLI$ = SAFETEXT$(!Mail_Cli)
        PPXX% = InStr(MAICLI$, "/")
        If PPXX% > 0 Then
           Label26 = COCLI$
           Text5 = Left$(MAICLI$, PPXX% - 1)
           MAICLI$ = Mid$(MAICLI$, PPXX% + 1)
           Do While TRIM$(MAICLI$) <> ""
             PPXX% = InStr(MAICLI$, "/")
             If PPXX% < 1 Then PPXX% = 1 + Len(MAICLI$)
             EMA$ = Left$(MAICLI$, PPXX% - 1)
             For JJ& = 0 To Text8.Count - 1
               If TRIM$(Text8(JJ&).TEXT) = "" Then
                  Text8(JJ&) = EMA$
                  GoTo 19
               End If
             Next JJ&
             '
19           MAICLI$ = Mid$(MAICLI$, PPXX% + 1)
           Loop
           OPENFORMS = DoEvents
           Exit Sub
           '
        End If
      .MoveNext
      Loop
    End With
    SECOM.Close
    Set SECOM = Nothing
    NOVALID% = 0
End Sub

Private Sub Command31_Click()

  If DERACCE%("MODIVEN", "Modificación de Vendedor") > 1 Then ' derecho de acceso para modificar vendedor
    '
    Call SELECHO("REPRCOME")
    VDEVU$ = VALACT1$("REPRCOME")
    If TRIM$(VDEVU$) <> "" Then
      Label42 = VDEVU$
    End If
    '
   End If
End Sub

Private Sub Command32_Click()
   '
   CONEXI$ = "PRESUP07" ' CONECTA CON PRESUPUESTO Y LE PASA EL NUMERO DE PROSEPCTO
   If TRIM$(Label26) <> "" Then CONEXI$ = CONEXI$ + "/" + TRIM$(Label26)
   Call CONECRUN(CONEXI$, "Presupuestos de Clientes")
   '
End Sub

Sub ALARMAPROSPECT()
   '
   '
   Call SELECHO("VENDEDOR")
   VEND% = XVALO(VALACT1$("VENDEDOR"))
   If VEND% < 1 Then
     Exit Sub
   End If
   '
   HOII% = HOY(HOS$)
   FECHOY = FETEXCOR1(HOII%)
   '
   SQLX$ = "SELECT * FROM BSEGCOM WITH(NOLOCK) WHERE "
   CONDI$ = " VEND_CLI =" & VEND%
   CONDI$ = CONDI$ + " AND FECHPROXCTO <= '" & FECHOY & " '"
   CONDI$ = CONDI$ + " AND (FECHBAJA >= '" & FECHOY & " ' OR FECHBAJA IS NULL OR FECHBAJA = '')"
   FIN& = CantRegistros("BSEGCOM", CONDI$)
   If FIN& < 1 Then
      Call COMUNI("No Se Encontraron Contactos Vencidos Para Este Vendedor")
      Exit Sub
   End If
   SQLX$ = SQLX$ + CONDI$
   '
   JJ& = 0
   JX& = 0
   TX$ = REGBLAN$("ALARPROS")
   Call SETULTREG("ALARPROS", 0)
   Dim BSEGTMP As ADODB.Recordset
   Set BSEGTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With BSEGTMP
    Do While Not .EOF
        Call TRACE(24, JX&, FIN&)
        JX& = JX& + 1
        COCLI$ = SAFETEXT$(!CODI_CLI)
        CONDI$ = "CODI_CLI =  '" & COCLI$ & "'"
        NC% = XVALO(VALOBADA("DEUDOR12", "NUME_CLI", CONDI$, "NOMESS"))
        If NC% > 0 Then GoTo 50 ' VALIDACION QUE NO ESTE DADO DE ALTA COMO CLIENTE
        Call REPLA(TX$, COCLI$, 1, 10)
        Call REPLA(TX$, SAFETEXT$(!Raso_Cli), 13, 64)
        TELE1$ = SAFETEXT$(!Tele_Cli)
        TELE2$ = SAFETEXT$(!TELE_Cli2)
        TELE$ = ""
        If TRIM$(TELE1$) <> "" Then TELE$ = TELE1$
        If TRIM$(TELE2$) <> "" Then TELE$ = TELE$ + "/" + TELE2$
        Call REPLA(TX$, TELE$, 77, 24)
          FEPROCTO% = CVINT(!FechProxCto)
        Call REPLA(TX$, MKI$(FEPROCTO%), 101, 2)
        notax$ = SAFETEXT$(!His_Contac)
        i& = 0
        CONTA% = 0
        While InStr(notax$, Chr$(10)) > 0
           CONTA% = CONTA% + 1
           POS2% = InStr(notax$, Chr$(10))
           NOTAX1$ = Mid$(notax$, 1, POS2% - 1)
           NOTAX1$ = REPLACAR$(NOTAX1$, Chr$(61), "") 'QUITA EL =
           NOTAX1$ = REPLACAR$(NOTAX1$, Chr$(45), "") ' QUITA EL -
           Call REPLA(TX$, SAFETEXT$(NOTAX1$), 103, 128)
           
           JJ& = JJ& + 1
           Call GRAREG("ALARPROS", TX$, JJ&)
           TX$ = REGBLAN$("ALARPROS")
           notax$ = Mid$(notax$, POS2% + 2)
        Wend
        If CONTA% < 1 Then
           JJ& = JJ& + 1
           Call REPLA(TX$, SAFETEXT$(NOTAX1$), 103, 128)
           Call GRAREG("ALARPROS", TX$, JJ&)
           TX$ = REGBLAN$("ALARPROS")
        End If
        
50  .MoveNext
    Loop
   End With
   BSEGTMP.Close
   Set BSEGTMP = Nothing
   If JJ& < 1 Then
      Call COMUNI("No Se Encontraron Contactos Vencidos Para Este Vendedor")
      Exit Sub
   End If

   Call SETULTREG("ALARPROS", JJ&)
   Call FINAL("*ALARPROS")

End Sub

Private Sub Command33_Click()
   Command33.Enabled = False
   COCLI$ = TRIM$(Label26)
   If COCLI$ = "" Then
21    Call SELECHO("INDIPROS" & "/Indice Prospect")
      COCLI$ = VALACT1$("INDIPROS")
      Label26 = COCLI$
   End If
   '
   If COCLI$ <> "" Then
      CONDI$ = "CODI_CLI = '" & COCLI$ & "'"
      CANOP% = CantRegistros("OPORTUNI", CONDI$)
      If CANOP% < 1 Then GoTo 30              ' NUEVA OPORTUNIDAD
      '
      OPX% = COMALTER%("Oportunidades para " + TRIM$(Text4) + "\\Nueva Oportunidad\Abrir Existente")
      If OPX% = 1 Then
30         Call INICRM09.NUEVAOPORTUN(COCLI$)
        ElseIf OPX% = 2 Then
           CONDI$ = "CODI_CLI = '" & COCLI$ & "' AND CodOport <>'' ORDER BY CodOport"
           Call CARGALISEL("!INDOPORT", "OPORTUNI", "CodOport/A12;RefeOport/A48", CONDI$, "NOMESS")
           Call SELECHO("!INDOPORT/INDICE DE OPORTUNIDADES")
           CODIOPO$ = TRIM$(VALACT1$("!INDOPORT"))
           If CODIOPO$ <> "" Then
              Call STDFORM("NUOPOR", CODIOPO$)
           End If
      End If
   End If
   '
   Command33.Enabled = True
End Sub

Private Sub Command35_Click()
   Call SELECHO("PROVINC")
   VDEVU$ = VALACT1$("PROVINC")
   If TRIM$(VDEVU$) <> "" Then
      Text27 = VDEVU$
   End If

End Sub

Private Sub Form_Load()
   '
   Call CENTRAFORM(Me)
   Call TRANSLABELS(Name)
   DoEvents
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
     Caption = Caption + "  -  " + EPAC$
   End If
   Label35 = USERNAM$(USNBR%)
   HOII% = HOY(HOS$)
   Label34 = HOS$
   Text12 = FECHATEX$(DIPOST(HOII%))
   '
   VScroll1.Max = (Frame15.Height - VScroll1.Height) / 100
   VScroll1.Value = 0
   '
   For KKI% = 1 To Text6.Count
      Text6(KKI% - 1).FontSize = 8
      Text7(KKI% - 1).FontSize = 8
      Text8(KKI% - 1).FontSize = 8
      Text10(KKI% - 1).FontSize = 8
      Text30(KKI% - 1).FontSize = 8
   Next KKI%
   '
   ' Al establecer 0 en InputLen se indica a MSComm que
   ' lea todo el contenido del búfer de entrada cuando
   ' se use la propiedad Input.
   'MSComm1.InputLen = 0
   '
   Call ABRECLIEN
   '
   If EXISTE%("PRESUP07.EXE") > 0 Then Command32.Visible = True
   '
   Call CREACAMPO("", "CONTACSCOM", "CODIEMPR", 3, 0)
   Call CREACAMPO("", "CONTACSCOM", "CODICLI", 10, 12)
   Call CREACAMPO("", "CONTACSCOM", "ORDI_CTAC", 3, 0)
   Call CREACAMPO("", "CONTACSCOM", "NOME_CTAC", 10, 64)
   Call CREACAMPO("", "CONTACSCOM", "CARG_CTAC", 10, 64)
   Call CREACAMPO("", "CONTACSCOM", "MAIL_CTAC", 10, 128)
   Call CREACAMPO("", "CONTACSCOM", "TELE_CTAC", 10, 64)
   Call CREACAMPO("", "CONTACSCOM", "MARBORRA", 3, 0)
   '
   For KKI% = 1 To Text6.Count - 1
      Text6(Index).ToolTipText = Text6(0).ToolTipText
   Next KKI%
   '
   For KKI% = 1 To Text8.Count - 1
      Text8(Index).ToolTipText = Text8(0).ToolTipText
   Next KKI%
   '
End Sub

Sub Command1_Click()
   '
   If TRIM$(Label26) = "" Then
     COMMAND16.SetFocus
     Exit Sub
   End If
   '
   Command1.Enabled = False
   If TRIM$(Label26) <> "" Then
     Screen.MousePointer = 11
      Call GRAFICON
     Screen.MousePointer = 1
    End If
   Command1.Enabled = True
   '
   If NOVALID% > 0 Then Exit Sub
   '
   IMPORTANDO = False
   On Error Resume Next
   IMPORTANDO = IMPORCRM.Visible
   PROCESANDO = PRINMAI09.Visible
   On Error GoTo 0
   If IMPORTANDO = True Then Exit Sub
   If PROCESANDO = True Then Exit Sub
   '
   If IMPORTA$ <> "SI" Then
     If TRIM$(Text20) = "" Then
        Call COMUNI("Falta Ingresar Domicilio")
        On Error Resume Next
        Text20.SetFocus
        On Error GoTo 0
        Exit Sub
     End If
   End If
   '
   If IMPORTA$ <> "SI" Then
     If TRIM$(Text22) = "" Then
       Call COMUNI("Falta Ingresar Ciudad")
       On Error Resume Next
       Text22.SetFocus
       On Error GoTo 0
       Exit Sub
     End If
   End If
   '
   If IMPORTA$ <> "SI" Then
     If ECOARCH("provinc", Text27) = "" Then
       Call COMUNI("Falta Ingresar Provincia o Provincia No Valida")
       On Error Resume Next
       Text27.SetFocus
       On Error GoTo 0
       Exit Sub
     End If
   End If
   
   
End Sub

Private Sub Command10_Click()
   '
MsgBox "No Disponible"
Exit Sub
   '
   Command10.Enabled = False
   Call EXPORAGE
   Command10.Enabled = True
End Sub

Private Sub Command12_Click()
   '
   If USNBR% <> 101 Then Exit Sub
   Command12.Enabled = False
   IMPORCONTAC.Show 1
   'Call IMPORAGE
   Command12.Enabled = True
End Sub

Sub Command13_Click()
   '
   Static FIDATX1#, FIDATX2#, UCONTA$
   '
   Command13.Enabled = False
   ABUSX$ = "BSEGCOM7": If USNBR% < 100 Then ABUSX$ = "BSEGALE"
   If FIDATIM#(ABUSX$) = FIDAT1X# Then
     If FIDATIM#("!SECONTA7") = FIDATX2# Then
       JJ& = ULTREG&("!SECONTA7")
       GoTo 20
     End If
   End If
   '
   Call SETULTREG("SECONTA7", 0)

   FIDATX1# = FIDATIM#(ABUSX$)
   JJ& = 0
   HOII% = HOY(HOS$)
   HOOOY = FETEXCOR1(HOII%)
   FINIPRO% = FECHANUM("30/09/10")  ' PARA EXCLUIR CONTACTOS VIEJOS
   FIPRO = FETEXCOR1(FINIPRO%)
   SQLX$ = "SELECT * FROM BSEGCOM"
   SQLX$ = SQLX$ + " WHERE FechProxCto >= '" & FIPRO & "'"
   SQLX$ = SQLX$ + " AND FechProxCto <= '" & HOOOY & "'"
   SQLX$ = SQLX$ + " ORDER BY CODI_CLI"
   Dim BSEGTMP As ADODB.Recordset
   Set BSEGTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With BSEGTMP
    Do While Not .EOF
        Call REPLA(EL1$, !CODI_CLI, 1, 12)
        Call REPLA(EL1$, AJUSTI$(SAFETEXT$(!Raso_Cli), 50), 13, 50)   ' empresa
        JJ& = JJ& + 1
        Call GRAREG("!SECONTA7", EL1$, JJ&)
    .MoveNext
    Loop
   End With
   Call SETULTREG("!SECONTA7", JJ&)
   
'   For U& = 1 To ULTREG&(ABUSX$)
'     XXX$ = REGLEIDO$(ABUSX$, U&)
'     If CVI(Mid$(XXX$, 1023, 2)) <= HOII% Then
'       If CVI(Mid$(XXX$, 1023, 2)) > FINIPRO% Then
'         JJ& = JJ& + 1
'         Call GRAREG("!SECONTAC", XXX$, JJ&)
'       End If
'     End If
'   Next U&
   Call SETULTREG("!SECONTA7", JJ&)
   Call CIERRARCH("!SECONTA7")
   Call BAJALDISCO
   FIDATX2# = FIDATIM#("!SECONTA7")
   Call FRESHECHO("!SECONTA7")
   '
20 If JJ& > 0 Then
     FORSELEC.ABUSQUE = UCONTA$
21   Call SELECHO("!SECONTA7")
     UCONTA$ = VALACT1$("!SECONTA7")
     
     If UCONTA$ <> "" Then
       Label26 = UCONTA$
       Call Command11_Click
       If Visible = False Then
          Hide
          Show 1
          GoTo 21
       End If
     End If
   End If
   '
   Command13.Enabled = True
   '
End Sub

Private Sub Command14_Click()
   If USNBR% <> 101 Then GoTo 20
   Command14.Enabled = False
   OPX% = COMALTER%("Opciones de Configuración:\\Conf.General\Particular\Clasificaciones")
   If OPX% = 1 Then
        Call CONECRUN("FSETUP04", "")
      ElseIf OPX% = 2 Then
20      STPCRM09.Show 1
      ElseIf OPX% = 3 Then
        Call ACTAGRUP
   End If
   Command14.Enabled = True
End Sub
'
Sub Command21_Click()
   '
   Call ABRECLIEN
10 NCLI$ = NUEVOPROESP$
   '
   VDEF$ = String$(128, 0)
   Call REPLA(VDEF$, NCLI$, 1, 12)
   Call REPLA(VDEF$, Space$(62), 13, 62)
      If PRINMAI09.Visible = True Then
         Call REPLA(VDEF$, UCase$(PRINMAI09.Label14.Caption), 13, 62)
      End If
   OBX$ = OBJPLA$("ALTACLIENPROSP", VDEF$)
   If OBX$ = "" Then Label26 = "": Exit Sub
   VDEF$ = OBX$
   
   NCLI$ = TRIM$(Mid$(OBX$, 1, 12))
   If NCLI$ = "" Then
      Call MENSERR(24, "Número de Cuenta no Válido.")
      GoTo 10
   End If
   '
   CONDI$ = "Codi_cli='" & NCLI$ & "'"
   '
   NXLI$ = VALOBADA("BSEGCOM", "CODI_CLI", CONDI$, "NOMESS")
   If NXLI$ <> "" Then
      Call MENSERR(24, "Ingreso no Válido.\Código de Cuenta Pre-existente.")
      GoTo 10
   End If
   '
   RAZCL$ = TRIM$(Mid$(OBX$, 13, 62))
   If Len(RAZCL$) < 1 Then
      Call MENSERR(24, "Ingreso no Válido.\Falta Razón Social del Cliente")
      GoTo 10
   End If
   '
   SQLX$ = "SELECT * FROM BSEGCOM"
   SQLX$ = SQLX$ + " WHERE Codi_Cli = '" & NCLI$ & "'"
   Dim BSEGTMP As ADODB.Recordset
   Set BSEGTMP = New ADODB.Recordset
   BSEGTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With BSEGTMP
       .AddNew   ' YA CONVERTIDO
       On Error Resume Next
       !CodiEmpr = CodiEmp%
       On Error GoTo 0
       !CODI_CLI = NCLI$
       !Raso_Cli = RAZCL$
       .Update
   End With
   BSEGTMP.Close
   Set BSEGTMP = Nothing
   '
   HOII% = HOY(HOS$)
   Label34 = HOS$
   Text12 = HOS$
   Label26 = NCLI$
   Text4.TEXT = RAZCL$
   '
   If USNBR% < 100 Then Label42 = TRIM$(Str$(USNBR%))
   '
   USNU% = USNBR% 'CARGA EL VENDEDOR SEGUN EL USUARIO
   If USNBR% > 100 Then '
     USNU% = USNBR% Mod 100
   End If
   NVEND% = 0
   For i& = 1 To ULTREG&("DATUSERS")
       X$ = REGLEIDO$("DATUSERS", i&)
       If Asc(Mid$(X$, 1, 1)) = USNU% Then
          NVEND% = Asc(Mid$(X$, 33, 2))
          Exit For
       End If
   Next i&
   Label42 = NVEND%
   '
   Call GENEFICHA
   If PRINMAI09.Visible = True Then
      Exit Sub                       ' PARA QUE NO LO MUESTRE ANTES DE ENTRAR
   End If
   '
   If Visible = False Then
      Show 1
   End If
   '
End Sub

Function NUEVOPROESP$()
   '
   Static COTERMI$
   '
   If COTERMI$ = "" Then
      ITERMI$ = IDENTER$
10    COTERMI$ = CONTROL$(ITERMI$, "COTERMI")
      If Len(COTERMI$) <> 2 Then
        Call MENSERR(24, "Falta Identificador de Estacion de Trabajo")
        OPCOTEL.Show 1
        GoTo 10
      End If
   End If
   '
   Call ABRECLIEN
   OTX$ = "CX-______-" & COTERMI$
   CAMPOSE$ = "MAX(CAST(SUBSTRING(CODI_CLI,4,6) AS Integer))"
   CONDI$ = "CODI_CLI LIKE '" & OTX$ & "'"
   NUCONSU& = 1 + XVALO(VALOBADA("BSEGCOM", CAMPOSE$, CONDI$))
'   If NUCONSU& < 100000 Then
'      NUCONSU& = 100001
'   End If
   '
20 NEXTNU$ = TRIM$(Str$(NUCONSU&))
   While Len(NEXTNU$) < 6: NEXTNU$ = "0" + NEXTNU$: Wend
   NUEVOPROESP$ = "CX-" + NEXTNU$ + "-" + COTERMI$
   '
End Function

Private Sub Command25_Click()
   ANCHAC = Width
   VISIAC = Picture1.Visible
   Image2.Visible = False
   Picture1.Visible = False
   Width = 10800
   On Error Resume Next
   PRINTFORM
   If Err Then
      On Error GoTo 0
      Call MENSERR(24, "Error de Impresión de Ficha")
   End If
   On Error GoTo 0
   Image2.Visible = True
   Picture1.Visible = VISIAC
   Width = ANCHAC
End Sub

Sub Command3_Click()
   Call CIERRARCH("*.*")
   On Error Resume Next
   Hide
   On Error GoTo 0
End Sub

Private Sub Command4_Click()
   cancelflag = 1
End Sub

Private Sub Command5_Click()
   '
   Command5.Enabled = False
   COCLI$ = TRIM$(Label26)
   CODOP$ = ""
   If COCLI$ <> "" Then
      If FISECOM09.VALIPROCON%(Text12, Text88) > 0 Then
         PERSCON$ = Combo1.TEXT
         REFACT$ = TRIM$(Text88)
         FEPROG% = FECHANUM(Text12)
         Call FISECOM09.PROGRACTI(COCLI$, CODOP$, PERSCON$, REFACT$, FEPROG%)
         '
         If TRIM$(Text11) <> "" Then
            Call ANOTAHISTORIAL(Text11, COCLI$)
            Text11 = ""
         End If
      End If
   End If
   Command5.Enabled = True
   '
'   COCLI$ = TRIM$(Label26)
'   If COCLI$ <> "" Then
'     FEC% = FECHANUM(Text12)
'     If FEC% <= HOY(HO$) Then FEC% = HOY(HO$)
'     FEPRO = FETEXCOR1$(FEC%)
'     CONDI$ = "CODI_CLI = '" & COCLI$ & "'"
'     Call ACTUREGISTRO("BSEGCOM", "FechProxCto", CONDI$, FEPRO, RAFE&, "NOMESS")
'     '
'     If TRIM$(Text11) <> "" Then
'       Call FRATEXTO(Text11, 56)
'       TTXX1$ = ""
'       For KKI% = 1 To CARETEX%
'         TTXX$ = Space$(76)
'         Call REPLA(TTXX$, RETEX$(KKI%), 11, 56)
'         If KKI% = 1 Then
'           Call REPLA(TTXX$, HOS$, 1, 8)
'           Call REPLA(TTXX$, USERTER$, 69, 24)
'         End If
'         TTXX1$ = TTXX1$ + TTXX$ + Chr$(10) + Chr$(13)
'       Next KKI%
'       TTXX1$ = TTXX1$ + String$(88, "=") + Chr$(10) + Chr$(13)
'       '
'       CONDI$ = "CODI_CLI = '" & COCLI$ & "'"
'       ANOTANTE$ = VALOBADA("BSEGCOM", "HIS_CONTAC", CONDI$)
'       Call ACTUREGISTRO("BSEGCOM", "HIS_CONTAC", CONDI$, ANOTANTE$ + TTXX1$, RAFE&)
'       '
'       Text11 = ""
'       '
'     End If
'   End If
'   '
'   Command5.Enabled = True
'   '
End Sub

Private Sub Command6_Click()
   cancelflag = 1
End Sub

Private Sub Command7_Click()
   '
MsgBox "No Disponible"
Exit Sub
   '
   Command7.Enabled = False
   Call CIERRARCH("BSEGCOM7")
   Call CIERRARCH("NOTASEG")
   Call CIERRARCH("*.*")
   '
   If TRIM$(UCase$(CurDir)) <> "C:\AGETEL" Then
     TTXX$ = "Su Agenda Telefónica FLEXOFT está Instalada"
     TTXX$ = TTXX$ + "\en la Carpeta " + REPLACAR$(CurDir, "\", "/") + "."
     TTXX$ = TTXX$ + "\  \Las Opciones de Respaldo y Recuperación"
     TTXX$ = TTXX$ + "\Pueden no Funcionar Correctamente en esta carpeta."
     TTXX$ = TTXX$ + "\  \Conviene Re-Instalar (Copiar) Toda la Agenda"
     TTXX$ = TTXX$ + "\a la Carpeta 'C:/AGETEL'."
     Call MENSERR(24, TTXX$)
   End If
   '
   If COMALTER%("Coloque Diskette de Respaldo\\Cancelar\Continuar") <> 2 Then
     GoTo 99
   End If
   '
   Screen.MousePointer = 11
   LDT$ = TRIM$(UCase$(LUDAT$))
   NX% = FILEOPEN%(LUCOM$ + "BKAGENTE.BAT", 2, 0)
   Print #NX%, LUBAS$ + "PKZIP " + "A:BCONTACT " + LDT$ + "BCONTACT.DAT " + LDT$ + "BCONTACT.X01 " + LDT$ + "NOTACONT.DAT " + LDT$ + "NOTACONT.X01 "
   Print #NX%, "EXIT"
   Close #NX%: NX% = 0
   Call BAJALDISCO
   '
   Shell LUCOM$ + "BKAGENTE.PIF", 2
   TINI = Timer
   While Timer < TINI + 5: Wend
   '
99 Command7.Enabled = True
   Screen.MousePointer = 1
   '
End Sub

Private Sub Command8_Click()
   '
MsgBox "No Disponible"
Exit Sub
   '
   Command8.Enabled = False
   Call CIERRARCH("BCONTACT")
   Call CIERRARCH("*.*")
   '
   If TRIM$(UCase$(CurDir)) <> "C:\AGETEL" Then
     TTXX$ = "Su Agenda Telefónica FLEXOFT está Instalada"
     TTXX$ = TTXX$ + "\en la Carpeta " + REPLACAR$(CurDir, "\", "/") + "."
     TTXX$ = TTXX$ + "\  \Las Opciones de Respaldo y Recuperación"
     TTXX$ = TTXX$ + "\Pueden no Funcionar Correctamente en esta carpeta."
     TTXX$ = TTXX$ + "\  \Conviene Re-Instalar (Copiar) Toda la Agenda"
     TTXX$ = TTXX$ + "\a la Carpeta 'C:/AGETEL'."
     Call MENSERR(24, TTXX$)
   End If
   '
   If COMALTER%("Coloque Diskette de Respaldo\\Cancelar\Continuar") <> 2 Then
     GoTo 99
   End If
   '
   If EXISTE%("A:BCONTACT.ZIP") <> 1 Then
     Call COMUNI("Recuperación Cancelada")
     GoTo 99
   End If
   '
   Screen.MousePointer = 11
   LDT$ = TRIM$(UCase$(LUDAT$))
   NX% = FILEOPEN%(LUCOM$ + "BKAGENTE.BAT", 2, 0)
   Print #NX%, LUBAS$ + "PKUNZIP -O " + "A:BCONTACT " + LDT$
   Print #NX%, "EXIT"
   Close #NX%: NX% = 0
   Call BAJALDISCO
   '
   Shell LUCOM$ + "BKAGENTE.PIF", 2
   TINI = Timer
   While Timer < TINI + 5: Wend
   '
99 Command8.Enabled = True
   Screen.MousePointer = 1
   '
End Sub

Private Sub Command9_Click()
   '
   If USNBR% <> 101 Then Exit Sub
   Command9.Enabled = False
   Call SORCONTAC
   Command9.Enabled = True
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
End Sub

Private Sub Image4_Click()
   Image4.Enabled = False
   'Call Dial(TRIM$(Text26.TEXT), Name)
   Image4.Enabled = True
End Sub

Private Sub Image5_Click()
   Image5.Enabled = False
   'Call Dial(TRIM$(Text24.TEXT), Name)
   Image5.Enabled = True
End Sub

Private Sub Label16_DblClick()
   Call ACTAGRUP
End Sub

Private Sub Label26_Change()
   '
   If IMPORTA$ = "SI" Then
     Call CARFICHA_DAT(Label26)
   Else
     Call CARFICHA
   End If
   '
End Sub

Private Sub Label37_DblClick()
   Label38 = ""
   Text14 = ""
End Sub

Private Sub Label7_DblClick()
   '
   If DERACCE%("GENDIMAI", "Generar Direcciones de Mail del Web-Site") >= 2 Then
      Screen.MousePointer = 11
      Call ABRECONEXION
      Screen.MousePointer = 11
      CONDI$ = "MAIL_CLI = '' OR MAIL_CLI IS NULL "
      CONDI$ = CONDI$ + " AND PAGWEB LIKE 'www.%'"
      Dim CLISE As ADODB.Recordset
      Set CLISE = FLEXCONN.Execute("SELEcT * FROM BSEGCOM WHERE " & CONDI$)
      '
      With CLISE
        Do While Not .EOF
           CCX$ = TRIM$(SAFETEXT$(!CODI_CLI))
           PWW$ = TRIM$(SAFETEXT$(!PAGWEB))
           If PWW$ <> "" Then
              PPXX% = InStr(UCase$(PWW$), "WWW.")
              If PPXX% > 0 Then
                 MCX$ = "info@" + Mid$(PWW$, PPXX% + 4)
                 PPXX% = InStr(MCX$, "/")
                 If PPXX% > 1 Then MCX$ = Left$(MCX$, PPXX% - 1)
                 '
                 CONDI$ = "CODI_CLI = '" & CCX$ & "'"
                 Call ACTUREGISTRO("BSEGCOM", "MAIL_CLI", CONDI$, MCX$, RAFE&)
              End If
           End If
        .MoveNext
        Loop
      End With
      '
      Screen.MousePointer = 1
      Call COMUNI("Completado")
      '
   End If
   '
End Sub

Private Sub List2_Click()
   List2.ToolTipText = TRIM$(Mid$(List2.TEXT, 49, 48))
End Sub

Private Sub List2_DblClick()
   Call ACTAGRUP
End Sub

Private Sub Text1_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text10_GotFocus(Index As Integer)
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text11_DblClick()
   Text11.Enabled = False
   Call Dial(TRIM$(Text11.TEXT), Name)
   Text11.Enabled = True
End Sub

Private Sub Text11_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text12_DblClick()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text12.TEXT = VDEV$
    End If
End Sub

Private Sub Text12_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text13_DblClick()
   WSITE$ = Text13
   Call ShellExecute(Me.hwnd, "Open", WSITE$, "", "", 1)
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Text20.SetFocus
End Sub

Private Sub Text15_GotFocus()
   Text4.SetFocus
End Sub

Private Sub label42_Change()
    VEND% = XVALO(Label42)
    LABEL40 = TRIM$(ECOARCH$("REPRCOME", MKI$(VEND%)))
End Sub

Private Sub label42_DblClick()
'
Call Command31_Click
'
End Sub


Private Sub label42_GotFocus()
   If USNBR% < 100 Then
      Text9.SetFocus
   End If
End Sub

Private Sub Text17_DblClick()
   Text17.Enabled = False
   'Call Dial(TRIM$(Text17.TEXT), Name)
   Text17.Enabled = True
End Sub

Private Sub Text17_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text18_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Text21.SetFocus
End Sub

Private Sub Text27_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text2_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text20_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text21_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text22_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text23_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text24_DblClick()
   Text24.Enabled = False
   'Call Dial(TRIM$(Text24.TEXT), Name)
   Text24.Enabled = True
End Sub

Private Sub Text24_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text26_DblClick()
   Text26.Enabled = False
   'Call Dial(TRIM$(Text26.TEXT), Name)
   Text26.Enabled = True
End Sub

'--------------------------------------------------------
'   DIALER.FRM
'   Copyright (c) 1994 Crescent Software, Inc.
'   por Carl Franklin
'
'   Actualizado por Anton de Jong
'
'   Muestra cómo marcar números telefónicos con un módem.
'
'   Para que este programa funcione, el teléfono y el
'   módem deben estar conectados a la misma línea.
'--------------------------------------------------------
'Option Explicit

'DefInt A-Z

' Este indicador se establece cuando el usuario elige Cancelar.

Sub Dial(NULLAMA$, FORNAME$)
   '
   Dim DialString$, FromModem$, dummy
    ' AT es el comando compatible Hayes ATTENTION, y es necesario para enviar comandos al módem.
    ' DT significa "marcar con tonos". El comando Dial utiliza tonos, en lugar de pulsos (DP =
    ' marcar con pulsos). Number$ es el número de teléfono a marcar.
    ' El punto y coma indica al módem que vuelva al modo de comando después de marcar
    ' (importante). Es necesario un retorno de carro, vbCr, al enviar comandos al módem.
    '
   '
   For KI% = 1 To Forms.Count
     If Forms(KI% - 1).Name = FORNAME$ Then GoTo 5
   Next KI%
   Exit Sub
   '
5  Number$ = ""
   INTERN$ = ""
   TOMALIN$ = TRIM$(CONTROL("AGETELEF", "TOMALIN"))
   '
   PPXX% = InStr(NULLAMA$, "/")
   If PPXX% < 1 Then PPXX% = 1 + Len(NULLAMA$)
   INTERN$ = TRIM$(Mid$(NULLAMA$, PPXX% + 1))
   NULLAMA$ = TRIM$(Left$(NULLAMA$, PPXX% - 1))
   '
   Number$ = TOMALIN$
   For KKI% = 1 To Len(NULLAMA$)
     CARA% = Asc(Mid$(NULLAMA$, KKI%, 1))
     If CARA% = 44 Or CARA% >= 48 Then
       If CARA% <= 57 Then
         Number$ = Number$ + Chr$(CARA%)
       End If
     End If
   Next KKI%
   If Number$ = "" Then Exit Sub
   '
10 Label28 = Number$
   Label29 = ""
   Picture1.Visible = True
   Picture1.Refresh
   cancelflag = 0
   '
   DialString$ = "ATDT" + Number$ + ";" + vbCr
    '
    ' Configuración del puerto de comunicaciones.
    ' Suponiendo que hay conectado un mouse a COM1, se establece 2 en CommPort
   SETUPSTRI$ = CONTROL("AGETELEF", "SETUPST")
   If SETUPSTRI$ = "" Then
     SETUPSTRI$ = "2     9600  N     8     1     "
   End If
   COPORT = Val(Left$(SETUPSTRI$, 6))
   On Error Resume Next
   MSComm1.CommPort = COPORT
   On Error GoTo 0
   '
   SETTI$ = ""
   For KKI% = 1 To 4
     SETTI$ = SETTI$ + TRIM$(Mid$(SETUPSTRI$, 6 * KKI% + 1, 6))
     If KKI% < 4 Then SETTI$ = SETTI$ + ","
   Next KKI%
   On Error Resume Next
   MSComm1.Settings = SETTI$
   If Err Then
      Call MENSERR(24, "Configuración de Puerto no Válida.\Corrija o Aplique Valores Predeterminados por Omisión.")
      Picture1.Visible = False
      SETUPTEL.Show 1
      Exit Sub
   End If
    '
    ' Abre el puerto de comunicaciones.
   On Error Resume Next
   MSComm1.PortOpen = True
   If Err Then
      Call MENSERR(24, "COM" + TRIM$(Str$(COPORT)) + " no está disponible.\Cambie la propiedad CommPort a otro puerto.")
      Picture1.Visible = False
      Exit Sub
   End If
    '
    ' Vuelca el búfer de entrada.
   MSComm1.InBufferCount = 0
    '
    ' Marca el número.
   MSComm1.Output = DialString$
   TFIN = Timer + 10
   Label30.Visible = False
    '
    ' Espera a que venga el "OK" del módem.
   Do
     dummy = DoEvents()
     CEVEN = MSComm1.CommEvent
     Select Case CEVEN
        Case comEvReceive
          Label29 = "Llamando ..."
          Label29.Refresh
          Label30.Visible = True
        Case comEvSend
          Label29 = "Marcando ..."
        Case Else
          'MsgBox "Error de Comunicacion"
          'exit do
     End Select
      '
      ' Si hay datos en el búfer, los lee.
     If MSComm1.InBufferCount Then
       FromModem$ = FromModem$ + MSComm1.Input
        ' Comprueba si hay "OK".
       If InStr(FromModem$, "OK") Then
             Label29 = "Llamando ..."
             Label29.Refresh
             Label30.Visible = True
             TFIN = Timer + 3
             While Timer < TFIN
               dummy = DoEvents()
             Wend
             Exit Do
          End If
       End If
        
       ' ¿El usuario ha elegido Cancelar?
       If cancelflag Then
          cancelflag = False
          Exit Do
       End If
       '
       If Timer > TFIN Then
         Picture1.Visible = False
         MsgBox "Imposible Establecer Comunicación"
         Exit Do
       End If
       '
    Loop
    '
    ' Desconecta el módem.
    MSComm1.Output = "ATH" + vbCr
    
    ' Cierra el puerto.
    MSComm1.PortOpen = False
    Picture1.Visible = False
    Call FRESHALL
    '
    If INTERN$ <> "" Then
      TFIN = Timer + 3
      While Timer < TFIN: Wend
      Number$ = INTERN$
      INTERN$ = ""
      VUELTA% = 2
      GoTo 10
    End If
    '
End Sub

Sub GRAFICON()
   '
   '\\\OT-05-0966-OD-27/10/08///

    CODICLIX$ = TRIM$(Label26)
    '
    'VALIDACION DEL TEXT EMPLEADO
    If XVALO(Text3) > 32766 Then
      Call MENSERR(24, "Números de Empleados Mayor al Admitido")
      GoTo 99
    End If
    Call ABRECLIEN
    SQLX$ = "SELECT * FROM BSEGCOM"
    SQLX$ = SQLX$ + " WHERE Codi_Cli = '" & CODICLIX$ & "'"
    Dim FichaTmp As ADODB.Recordset
    Set FichaTmp = New ADODB.Recordset
    FichaTmp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With FichaTmp
     If .EOF Then
       .AddNew   ' YA CONVERTIDO
       On Error Resume Next
       !CodiEmpr = CodiEmp%
       On Error GoTo 0
       !CODI_CLI = CODICLIX$
     End If
        !Nume_Cli = 0
        !Raso_Cli = AJUSTI$(Text4, 59) + AJUSTI$(Text15, 3) ' empresa
        !Mail_Cli = Left$(TRIM$(Text5), 64) 'DIRECCION ELECTRONICA
        !Domi_Cli = Left$(TRIM$(Text20), 64) ' CALLE Y NUMERO
        !Cpos_Cli = Left$(TRIM$(Text21), 16) ' C.POSTAL
        !Loca_Cli = Left$(TRIM$(Text22), 48) ' LOCALIDAD
        !Prov_Cli = Left$(TRIM$(Text27), 2) ' PROVINCIA  **********Verificar
        !Pais_Cli = Left$(TRIM$(Text23), 48)    ' PAIS
        !Origen = Left$(TRIM$(Text18), 64)  ' ORIGEN
        !Vend_Cli = XVALO(Label42)
        !RamoActi = Left$(TRIM$(Text1), 64) ' RAMO
        !ProduPrin = Left$(TRIM$(Text2), 128) ' PRINC.PRODUCTO
        !Empleados = XVALO(Text3) '****************Ver
        !Tele_Cli = Left$(TRIM$(Text26), 32)  ' TELEFONO
        !TELE_Cli2 = Left$(TRIM$(Text17), 32) ' telef.partic.
        !Fax_Cli = Left$(TRIM$(Text24), 24) ' telef2.partic
        !Ctac_Cli = Left$(TRIM$(Text6(0)), 48)
        !Cargo1 = Left$(TRIM$(Text7(0)), 32)
        !SECTOR1 = Left$(TRIM$(Text8(0)), 64)
        !Interno1 = Left$(TRIM$(Text10(0)), 32)
        !TCelu1 = Left$(TRIM$(Text30(0)), 32)
        !Contacto2 = Left$(TRIM$(Text6(1)), 48)
        !Cargo2 = Left$(TRIM$(Text7(1)), 32)
        !SECTOR2 = Left$(TRIM$(Text8(1)), 64)
        !Interno2 = Left$(TRIM$(Text10(1)), 32)
        !TCelu2 = Left$(TRIM$(Text30(1)), 32)
        !Contacto3 = Left$(TRIM$(Text6(2)), 48)
        !Cargo3 = Left$(TRIM$(Text7(2)), 32)
        !SECTOR3 = Left$(TRIM$(Text8(2)), 64)
        !Interno3 = Left$(TRIM$(Text10(2)), 32)
        !TCelu3 = Left$(TRIM$(Text30(2)), 32)
        !PAGWEB = Left$(TRIM$(Text13), 48)
        !fechbaja = Label38
        !Mot_Baja = Left$(TRIM$(Text14), 64)
        '!FechUltEnvio =
        
        PUNGRU$ = ""
        For U& = 1 To List2.ListCount
          PGR% = XVALO(Mid$(List2.List(U& - 1), 97, 6))
          If PGR% < 256 Then PUNGRU$ = PUNGRU$ + Chr$(PGR%)
        Next U&
        !GRUPOS = Left$(PUNGRU$ + String$(30, 0), 30)
        '
        FEC% = FECHANUM(Text12)
        !FechProxCto = CVDAT(FEC%)
        !Nota_Cli = Text9
        .Update
   
    End With
    FichaTmp.Close
    Set FichaTmp = Nothing
    '
    CODICLIX$ = TRIM$(Label26)
    CONDI$ = "CodiCli = '" & CODICLIX$ & "'"
    Call ACTUREGISTRO("CONTACSCOM", "Marborra", CONDI$, 1, RAFE&)
    '
    SQLX$ = "SELECT * FROM CONTACSCOM WHERE " + CONDI$
    Dim Fichanew As ADODB.Recordset
    Set Fichanew = New ADODB.Recordset
    Fichanew.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    For IKI% = 0 To Text6.Count - 1
        TEXTOX$ = Text6(IKI%).TEXT + Text7(IKI%).TEXT + Text8(IKI%).TEXT + Text10(IKI%).TEXT + Text30(IKI%)
        If TRIM$(TEXTOX$) <> "" Then
            With Fichanew
                .AddNew   ' YA CONVERTIDO
                On Error Resume Next
                !CodiEmpr = CodiEmp%
                On Error GoTo 0
                !CODICLI = CODICLIX$
                !ORDI_CTAC = i
                !NOME_CTAC = Text6(IKI%)
                !CARG_CTAC = Text7(IKI%)
                !MAIL_CTAC = Text8(IKI%)
                !TELE_CTAC = Text10(IKI%)
                !CELU_CTAC = Text30(IKI%)
                !MarBorra = 0
                .Update
            End With
        End If
    Next IKI%
    Fichanew.Close
    Set Fichanew = Nothing
    '
    CONDI1$ = "MARBORRA = 1"
    Call BORRAREGISTROS("CONTACSCOM", CONDI1$, RAFE&)
    '
    Call GENEFICHA 'GENERA EL INDICE DE FICHAS
99  Screen.MousePointer = 1
    '\\\OT-05-0966-OD-FIN///
    '


            '   Call REPLA(RELE$, AJUSTI$(Text4, 59) + AJUSTI$(Text15, 3), 3, 62)  ' empresa
            '   Call REPLA(RELE$, Text5, 65, 64)     ' DIRECCION ELECTRONICA
               ''''   Call REPLA(RELE$, Text20, 129, 48)   ' CALLE Y NUMERO
            '   Call REPLA(RELE$, Text21, 177, 8)    ' C.POSTAL
            '   Call REPLA(RELE$, Text22, 185, 32)   ' LOCALIDAD
            '
            '   '
            '   Call REPLA(RELE$, Text3, 369, 8) ' provincia particular
            '   Call REPLA(RELE$, Text26, 377, 24)   ' TELEFONO
            '   Call REPLA(RELE$, Text17, 401, 24) ' telef.partic.
            '   Call REPLA(RELE$, Text24, 425, 24) ' telef2.partic
            '   '
            '   For KKI% = 0 To 2
            '     TTXX$ = Space$(96)
            '     Call REPLA(TTXX$, Text6(KKI%), 1, 32)    ' CONTACTO PERSONA 1 / 2 / 3
            '     Call REPLA(TTXX$, Text7(KKI%), 33, 24)
            '     Call REPLA(TTXX$, Text8(KKI%), 57, 24)
            '     Call REPLA(TTXX$, Text10(KKI%), 81, 16)
            '     Call REPLA(RELE$, TTXX$, 449 + 96 * KKI%, 96)
            '   Next KKI%
            '   '
            '   Call REPLA(RELE$, Text13, 737, 64) ' WEB-SITE
            '
            '   PUNGRU$ = ""
            '   For U& = 1 To List2.ListCount
            '     PGR% = Val(Mid$(List2.List(U& - 1), 73, 6))
            '     If PGR% < 256 Then PUNGRU$ = PUNGRU$ + Chr$(PGR%)
            '   Next U&
            '   Call REPLA(RELE$, PUNGRU$ + String$(30, 0), 993, 30)
            '   '
            '   IVENDE% = XVALO(Text16)
            '   Call REPLA(RELE$, MKI$(IVENDE%), 971, 2)
            '   Call REPLA(RELE$, MKI$(FECHANUM(Label38)), 973, 2)   ' FECHA DE BAJA
            '   Call REPLA(RELE$, TRIM$(Text14), 975, 16)             ' MOTIVO DE BAJA
            '   Call REPLA(RELE$, MKI$(FECHANUM(Text12)), 1023, 2)   ' FECHA PROXIMO CONTACTO
            '''''   'HOII% = HOY(HOS$)
            '''''   'If CVI(Mid$(RELE$, 1023, 2)) < HOII% Then
            '''''   '  Call REPLA(RELE$, MKI$(HOII%), 1023, 2)
            '''''   'End If
            '   '
            '   Call GRAREG("BSEGCOM", RELE$, REG&)
   '
'   Call CIERRARCH("BSEGCOM")
'   Call FRESHECHO("BSEGCOM")
'   '
'   Call COPYSTRU("BSEGCOM", "BSEGALE")
'   JJ& = 0
'   FIN& = ULTREG&("BSEGCOM")
'   For U& = 1 To FIN&
'      XX$ = REGLEIDO$("BSEGCOM", U&)
'      If CVI(Mid$(XX$, 971, 2)) = 3 Then
'         JJ& = JJ& + 1
'         Call GRAREG("BSEGALE", XX$, JJ&)
'      End If
'   Next U&
'   Call SETULTREG("BSEGALE", JJ&)
'   Call FRESHECHO("BSEGALE")
'   '
'   Screen.MousePointer = 11
'   NCON% = Val(Label26)
'   URN& = ULTREG&("NOTASEG")
'   RNC$ = RECFILT$("NOTASEG", 1, MKI$(NCON%))
'   KK& = (-3)
'   '
'   TTYZ$ = Text9.TEXT
'   LOCADE& = Len(TTYZ$)
'   For U& = LOCADE& To 1 Step -1
'     If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 10
'   Next U&
'   U& = 0
'   '
'10 TEXTO$ = Left$(TTYZ$, U&)
'   While Len(TEXTO$) > 0
'      TTYY$ = MKI$(NCON%) + TEXTO$ + Space$(62)
'      KK& = KK& + 4
'      If KK& <= Len(RNC$) - 3 Then
'           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
'         Else
'           URN& = URN& + 1
'           RENOTA& = URN&
'      End If
'      Call GRAREG("NOTASEG", TTYY$, RENOTA&)
'      TEXTO$ = Mid$(TEXTO$, 63)
'   Wend
'   '
'20 While KK& < Len(RNC$) - 3
'      KK& = KK& + 4
'      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
'      Call GRAREG("NOTASEG", MKI$(0) + Space$(62), RENOTA&)
'   Wend
'   '
'   Call CIERRARCH("NOTASEG")
'
End Sub

Sub CARFICHA()
   '
   Screen.MousePointer = 1
   
5  COCLI$ = TRIM$(Label26)
'   REG& = REGBUS&("BSEGCOM", 1, MKI$(IFICHA%))
'   If REG& < 1 Or REG& > ULTREG&("BSEGCOM") Then
   Call ABRECLIEN
   
   SQLX$ = "SELECT * FROM BSEGCOM"
   SQLX$ = SQLX$ + " WHERE CODI_CLI = '" & COCLI$ & "'"
   Dim FICHTMP As ADODB.Recordset
   Set FICHTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With FICHTMP
    If .EOF Then
       Text1 = ""
       Text2 = ""
       Text3 = ""
       Text4 = ""
       Text5 = ""
       Text9 = ""
       Text11 = ""
       Label42 = ""
       Text15 = ""
       Text13 = ""
       Text12 = ""
       Text25 = ""
       Text14 = ""
       Text15 = ""
       Text11 = ""
       Text20 = ""
       Text21 = ""
       Text22 = ""
       Text27 = ""
       Text23 = ""
       Text26 = ""
       Label42 = "": LABEL40 = ""
       Text17 = ""
       Text24 = ""
       Text18 = ""
       '
       For KKI% = 0 To Text6.Count - 1
         Text6(KKI%) = ""
         Text7(KKI%) = ""
         Text8(KKI%) = ""
         Text10(KKI%) = ""
         Text30(KKI%) = ""
       Next KKI%
       '
       HOII% = HOY(HOS$)
       Label34 = HOS$
       Text12 = FECHATEX$(DIPOST(HOII%))
       '
       List2.Clear
       '
     Else
       'RELE$ = REGLEIDO$("BSEGCOM", REG&)
       Text4 = TRIM$(Mid$(SAFETEXT$(!Raso_Cli), 1, 59))
       Text15 = TRIM$(Mid$(SAFETEXT$(!Raso_Cli), 60, 3))
       Text5 = SAFETEXT$(!Mail_Cli)
       Text20 = SAFETEXT$(!Domi_Cli)
       Text21 = SAFETEXT$(!Cpos_Cli)
       '
       Text22 = SAFETEXT$(!Loca_Cli)
       Text27 = SAFETEXT$(!Prov_Cli)
       Text23 = SAFETEXT$(!Pais_Cli)
       Text18 = SAFETEXT$(!Origen)
       Text1 = SAFETEXT$(!RamoActi)
       Text2 = SAFETEXT$(!ProduPrin)
       Text3 = XVALO(!Empleados)
       '
       Text26 = SAFETEXT$(!Tele_Cli)
       Text17 = SAFETEXT$(!TELE_Cli2)
       Text24 = SAFETEXT$(!Fax_Cli)
       Text11 = ""
       '
       For KKI% = 0 To Text6.Count - 1
          Text6(KKI%) = ""
          Text7(KKI%) = ""
          Text8(KKI%) = ""
          Text10(KKI%) = ""
          Text30(KKI%) = ""
       Next KKI%
       '
       JJ1% = 0
       Text6(0) = TRIM$(SAFETEXT$(!Ctac_Cli))
       Text7(0) = TRIM$(SAFETEXT$(!Cargo1))
       Text8(0) = TRIM$(SAFETEXT$(!SECTOR1))
       Text10(0) = TRIM$(SAFETEXT$(!Interno1))
       Text30(0) = TRIM$(SAFETEXT$(!TCelu1))
       If Text6(JJ1%) + Text7(JJ1%) + Text8(JJ1%) + Text10(JJ1%) + Text30(JJ1%) <> "" Then JJ1% = JJ1% + 1
       '
       Text6(JJ1%) = TRIM$(SAFETEXT$(!Contacto2))
       Text7(JJ1%) = TRIM$(SAFETEXT$(!Cargo2))
       Text8(JJ1%) = TRIM$(SAFETEXT$(!SECTOR2))
       Text10(JJ1%) = TRIM$(SAFETEXT$(!Interno2))
       Text30(JJ1%) = TRIM$(SAFETEXT$(!TCelu2))
       If Text6(JJ1%) + Text7(JJ1%) + Text8(JJ1%) + Text10(JJ1%) + Text30(JJ1%) <> "" Then JJ1% = JJ1% + 1
       '
       Text6(JJ1%) = TRIM$(SAFETEXT$(!Contacto3))
       Text7(JJ1%) = TRIM$(SAFETEXT$(!Cargo3))
       Text8(JJ1%) = TRIM$(SAFETEXT$(!SECTOR3))
       Text10(JJ1%) = TRIM$(SAFETEXT$(!Interno3))
       Text30(JJ1%) = TRIM$(SAFETEXT$(!TCelu3))
       If Text6(JJ1%) + Text7(JJ1%) + Text8(JJ1%) + Text10(JJ1%) + Text30(JJ1%) <> "" Then JJ1% = JJ1% + 1
       '
       Text13 = SAFETEXT$(!PAGWEB)
       '
       IVENDE% = XVALO(!Vend_Cli)
       If IVENDE% = 8224 Then IVENDE% = 0
       Label42 = TRIM$(FORMATNUM$(IVENDE%, "I4"))
       '
       List2.Clear
       '
       PUNGRU$ = SAFETEXT$(!GRUPOS) 'Mid$(RELE$, 993, 30)
       UTACON% = ULTREG&("TAGRUCON")
       For UKI% = 1 To Len(PUNGRU$)
         RETA& = Asc(Mid$(PUNGRU$, UKI%, 1))
         If RETA& > 0 Then
           If RETA& <= UTACON% Then
             XX$ = REGLEIDO$("TAGRUCON", RETA&)
             RAGRE% = Val(CVI(Left$(XX$, 2)))
             TAGRE$ = AJUSTI$(Mid$(XX$, 51, 24), 48) + AJUSTI$(Mid$(XX$, 3, 48), 48) + Str$(RAGRE%)
             List2.AddItem TAGRE$
           End If
         End If
       Next UKI%
       '
       FEBAJA$ = SAFETEXT$(!fechbaja)         ' FECHA DE BAJA
       Label38 = FEBAJA$
       '
       Text14 = SAFETEXT$(!Mot_Baja)                 ' MOTIVO DE BAJA
       '
       FE% = CVINT(SAFETEXT$(!FechProxCto))
       Text12 = FECHATEX(FE%)
       '
       Text9 = "" ': TTYZ$ = ""
       Text9 = SAFETEXT$(!Nota_Cli)
       '
       ' BUSCA Y PRESENTA, SI EXISTE, EL NUMERO DE CLIENTE ASOCIADO AL PROSPECTO
       NCLI% = XVALO(VALOBADA("DEUDOR12", "Nume_Cli", "Codi_Cli='" & TRIM$(Label26) & "'"))
       Label41 = NCLI%
       '
    End If
   End With
   '
   FICHTMP.Close
   Set FICHTMP = Nothing
   '
   On Error Resume Next
   If IMPORCRM.Visible = True Then GoTo 99
   On Error GoTo 0
   '
   SQLX$ = "SELECT * FROM CONTACSCOM"
   SQLX$ = SQLX$ + " WHERE CODICLI = '" & COCLI$ & "'"
   Dim FICHCON As ADODB.Recordset
   Set FICHCON = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With FICHCON
      Do While Not .EOF
        NOME$ = TRIM$(SAFETEXT$(!NOME_CTAC))
        CARG$ = TRIM$(SAFETEXT$(!CARG_CTAC))
        MAIL$ = TRIM$(SAFETEXT$(!MAIL_CTAC))
        TELE$ = TRIM$(SAFETEXT$(!TELE_CTAC))
        CELU$ = TRIM$(SAFETEXT$(!CELU_CTAC))
        '
        For JJ2% = 1 To JJ1%
           If UCase$(Text6(JJ2% - 1)) = UCase$(NOME$) Then
              If Text7(JJ2% - 1) = "" Then Text7(JJ2% - 1) = CARG$
              If Text8(JJ2% - 1) = "" Then Text8(JJ2% - 1) = MAIL$
              If Text10(JJ2% - 1) = "" Then Text10(JJ2% - 1) = TELE$
              If Text30(JJ2% - 1) = "" Then Text30(JJ2% - 1) = CELU$
              GoTo 19
           End If
        Next JJ2%
        '
15      If JJ1% >= Text6.Count Then
           Call AGREGARENGLONES
           GoTo 15
        End If
        '
        Text6(JJ1%) = NOME$
        Text7(JJ1%) = CARG$
        Text8(JJ1%) = MAIL$
        Text10(JJ1%) = TELE$
        Text30(JJ1%) = CELU$
        If Text6(JJ1%) + Text7(JJ1%) + Text8(JJ1%) + Text10(JJ1%) + Text30(JJ1%) <> "" Then
           JJ1% = JJ1% + 1
16         If JJ1% > Text6.Count Then
              Call AGREGARENGLONES
              GoTo 16
           End If
        End If
        '
19    .MoveNext
      Loop
   End With
   FICHCON.Close
   Set FICHCON = Nothing
   '
   ' carga el combo de contactos
   Call CARLICONTAC(Combo1, Label26)
   '
   ' carga proxima actividad proyectada
   If COCLI$ <> "" Then
      CONDI$ = "Codi_cli = '" & COCLI$ & "'"
      CONDI$ = CONDI$ + " AND STAT_ACTI < 3 "
      CONDI$ = CONDI$ + " AND (FHOR_REAL < '" & FETEXCOR1$(FECHANUM("01/01/10")) & "' OR FHOR_REAL IS NULL) "
      CONDI$ = CONDI$ + " ORDER BY FHOR_PROG DESC"
      '
      Combo1.TEXT = VALOBADA("ACTIPROY", "CONTACTO", CONDI$, "NOMESS")
      Text88 = VALOBADA("ACTIPROY", "REFE_ACTI", CONDI$, "NOMESS")
      Text12 = FECHATEX$(CVINT(VALOBADA("ACTIPROY", "FHOR_PROG", CONDI$)))
   End If
   '
99 On Error GoTo 0
   Screen.MousePointer = 1
   '
End Sub

Sub CARLICONTAC(NOMBREDECOMBO, CODCLIX$, Optional NOCONTA$)
   '
   On Error Resume Next
   NOMBREDECOMBO.Clear
   If Err Then
      On Error GoTo 0
      Exit Sub
   End If
   '
   CONDI$ = "CODI_CLI = '" & CODCLIX$ & "'"
   NOCON1$ = TRIM$(VALOBADA("BSEGCOM", "Ctac_Cli", CONDI$)) ', "NOMESS"))
   If NOCON1$ <> "" Then NOMBREDECOMBO.AddItem NOCON1$
   '
   NOCON1$ = TRIM$(VALOBADA("BSEGCOM", "CONTACTO2", CONDI$, "NOMESS"))
   If NOCON1$ <> "" Then NOMBREDECOMBO.AddItem NOCON1$
   '
   NOCON1$ = TRIM$(VALOBADA("BSEGCOM", "CONTACTO3", CONDI$, "NOMESS"))
   If NOCON1$ <> "" Then NOMBREDECOMBO.AddItem NOCON1$
   '
   SQLX$ = "SELECT * FROM CONTACSCOM"
   SQLX$ = SQLX$ + " WHERE CODICLI = '" & CODCLIX$ & "'"
   Dim FICHCON As ADODB.Recordset
   Set FICHCON = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With FICHCON
      Do While Not .EOF
        NOCON1$ = TRIM$(SAFETEXT$(!NOME_CTAC))
        '
        For JJ2% = 1 To NOMBREDECOMBO.ListCount
           If UCase$(NOMBREDECOMBO.List(JJ2% - 1)) = UCase$(NOCON1$) Then GoTo 19
        Next JJ2%
        '
        NOMBREDECOMBO.AddItem NOCON1$
19    .MoveNext
      Loop
   End With
   FICHCON.Close
   Set FICHCON = Nothing
   '
   On Error Resume Next
   If NOCONTA$ <> "" Then
        NOMBREDECOMBO.TEXT = NOCONTA$
     Else
        If NOMBREDECOMBO.ListCount > 0 Then NOMBREDECOMBO.TEXT = NOMBREDECOMBO.List(0)
   End If
   On Error GoTo 0
   '
End Sub
















Sub SORCONTAC()
   '
   ' SELECCIONA MODO DE ORDENAMIENTO
   OPXX% = COMALTER%("Elija si Criterio de Ordenamiento:\\Por Número\Alfabético\Cancelar")
   '
   SQLX$ = "SELECT * FROM BSEGCOM"
   If OPXX% = 1 Then
     SQLX$ = SQLX$ + " ORDER BY CAST(SUBSTRING(CODI_CLI,4,9)AS INTEGER) ASC"
   Else
     SQLX$ = SQLX$ + " ORDER BY RASO_CLI ASC"
   End If
   Dim BSGCTMP As ADODB.Recordset
   Set BSGCTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   FIN& = ULTREG&("INDIPROS")
   Call SETULTREG("INDIPROS", 0)
   ' GENERA LISTA DE SELECCION DE PROSPECTOS SEGUN ORDENAMIENTO ELEGIDO
   With BSGCTMP
    Do While Not .EOF
      Call TRACE(24, JJ&, FIN&)
      EL1$ = REBLA$
      Call REPLA(EL1$, !CODI_CLI, 1, 15)
      Call REPLA(EL1$, AJUSTI$(SAFETEXT$(!Raso_Cli), 50), 16, 50)   ' empresa
      JJ& = JJ& + 1
      Call GRAREG("INDIPROS", EL1$, JJ&)
    .MoveNext
    Loop
   End With
   ' CIERRA ARCHIVO Y REFRESCO EL ARCHIVO ECO
   Call SETULTREG("INDIPROS", JJ&)
   Call CIERRARCH("INDIPROS")
   Call FRESHECHO("INDIPROS")
   Exit Sub
   '
   List1.Clear
   Screen.MousePointer = 11
   '
   For U& = 1 To ULTREG&("BSEGCOM7")
     RELE$ = REGLEIDO$("BSEGCOM7", U&)
     TX2$ = UCase$(TRIM$(Mid$(RELE$, 3, 62)))
     TTXX$ = AJUSTI$(TX2$, 64) + Str$(U&)
     List1.AddItem TTXX$
   Next U&
   '
   JJ& = 0
   Call SETULTREG("!BSEGCOM7", 0)
   For U& = 1 To List1.ListCount
     TTXX$ = List1.List(U& - 1)
     REG& = Val(Mid$(TTXX$, 65))
     X$ = REGLEIDO$("BSEGCOM7", REG&)
     JJ& = JJ& + 1
     Call GRAREG("!BSEGCOM7", X$, JJ&)
   Next U&
   Call SETULTREG("!BSEGCOM7", JJ&)
   Call CIERRARCH("*.*")
   '
   Call TRACENTRAL("BSEGCOM7", "")
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
End Sub

Sub EXPORAGE()
   '
   On Error GoTo 90
   N1% = FreeFile
   Open "A:EXPORLIB.CSV" For Output As #N1%
   Print #N1%, "Nombre,Dirección electrónica"
   On Error GoTo 0
   '
   Screen.MousePointer = 11
   For U& = 1 To ULTREG&("BCONTACT")
     RELE$ = REGLEIDO$("BCONTACT", U&)
     TX2$ = TRIM$(Mid$(RELE$, 89, 56))
     TX4$ = TRIM$(Mid$(RELE$, 145, 64))
     TTXX$ = TRIM$(TX2$)
     If TRIM$(TX2$) <> "" And TRIM$(TX4$) <> "" Then
         TTXX$ = TTXX$ + " - "
       Else
         TTXX$ = TTXX$ + " "
     End If
     '
     TTYY$ = TRIM$(TTXX$ + TRIM$(TX4$))
     TTZZ$ = REPLACAR$(TTYY$, Chr$(34), "")
     TTXX$ = REPLACAR$(TTZZ$, ",", " ")
     '
     Call REPLA(RELE$, TTXX$, 3, 78)
     Call GRAREG("BCONTACT", RELE$, U&)
     '
     EMA$ = TRIM$(Mid$(RELE$, 217, 64))
     If EMA$ <> "" Then
       If Asc(Left$(EMA$, 1)) <> 39 Then
         AX$ = TTXX$ + "," + EMA$
         Print #N1%, AX$
         JJ& = JJ& + 1
       End If
     End If
   Next U&
   '
   Close #N1%
   Call CIERRARCH("*.*")
   MsgBox Str$(JJ&) + " Direcciones Exportadas"
   Screen.MousePointer = 1
   Exit Sub
   '
90 Resume 91
91 On Error GoTo 0
   '
End Sub

Sub IMPORAGE()
   '
   Screen.MousePointer = 11
   Dim RAEMP$(4096)
   For U& = 1 To ULTREG&("BCONTACT")
     X$ = REGLEIDO$("BCONTACT", U&)
     RAEMP$(U&) = UCase$(TRIM$(Mid$(X$, 145, 64)))
     CAEMP& = U&
   Next U&
   '
   FIN& = ULTREG&("!BCONTACT")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("!BCONTACT", U&)
     RASOE$ = UCase$(TRIM$(Mid$(X$, 145, 64)))
     EMAI$ = LCase$(TRIM$(Mid$(X$, 217, 64)))
     For JJ& = 1 To CAEMP&
       If RAEMP$(JJ&) = RASOE$ Then
         Y$ = REGLEIDO$("BCONTACT", JJ&)
         If InStr(Y$, EMAI$) > 0 Then GoTo 19 ' YA ESTA
         If TRIM$(Mid$(Y$, 217, 64)) = "" Then
             Call REPLA(Y$, EMAI$, 217, 64)
           ElseIf TRIM$(Mid$(Y$, 281, 64)) = "" Then
             Call REPLA(Y$, EMAI$, 281, 64)
           ElseIf TRIM$(Mid$(Y$, 345, 64)) = "" Then
             Call REPLA(Y$, EMAI$, 345, 64)
         End If
         Call GRAREG("BCONTACT", Y$, JJ&)
         GoTo 19
       End If
    Next JJ&
    '
    CAEMP& = CAEMP& + 1
    RAEMP$(CAEMP&) = UCase$(TRIM$(Mid$(X$, 145, 64)))
    KKI% = CAEMP&
    Call REPLA(X$, MKI$(KKI%), 1, 2)
    Call REGAPP("BCONTACT", X$)
19 Next U&
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
End Sub

Sub IMPOROUT()
   '
   On Error GoTo 90
   N1% = FreeFile
   Open "A:\OUTLOOK.CSV" For Input As #N1%
   Line Input #N1%, A$
   On Error GoTo 0
   '
   REG& = 0
   While Not EOF(N1%)
     REG& = REG& + 1
     REGI% = REG&
     RELE$ = MKI$(REGI%) + Space$(1022)
     '
     Line Input #N1%, TTXX$
     PPXX% = InStr(TTXX$, ",")
     If PPXX% > 0 Then
        EMPRE$ = Left$(TTXX$, PPXX% - 1)
        TTXX$ = Mid$(TTXX$, PPXX% + 1)
        PPXX% = InStr(TTXX$, ",")
        If PPXX% > 0 Then
          EMAI$ = Left$(TTXX$, PPXX% - 1)
          '
          Call REPLA(RELE$, EMPRE$, 3, 78)
          Call REPLA(RELE$, "Sr.", 81, 8)
          Call REPLA(RELE$, EMPRE$, 145, 64)
          Call REPLA(RELE$, EMAI$, 217, 64)
          Call GRAREG("BCONTACT", RELE$, REG&)
        End If
     End If
     '
   Wend
   '
   Call CIERRARCH("BCONTACT")
   Call FRESHECHO("BCONTACT")
   Exit Sub
   '
90 Resume 91
91 On Error GoTo 0
   '
End Sub

Private Sub Text26_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub TEXT27_Change()
    LABEL52 = TRIM$(ECOARCH$("provinc", Text27))
End Sub

Private Sub Text3_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text4_GotFocus()
   If TRIM$(Label26) = "" Then
      On Error Resume Next ' evitar el error 5 al clickear en boton visual data
      COMMAND16.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Text5.SetFocus
End Sub

Public Sub Text5_DblClick()
   '
   COCLI$ = TRIM$(Label26)
   If COCLI$ <> "" Then
     '
     Call CREAMAIL(COCLI$)
     '
     ATENTI% = 1
     MAILMESS07.Label26 = Label26
     MAILMESS07.Show 1
   End If
   '
End Sub

Private Sub Text5_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Text13.SetFocus
End Sub

Sub Text6_DblClick(Index As Integer)
   '
   COCLI$ = TRIM$(Label26)
   If COCLI$ <> "" Then
       DESTI$ = TRIM$(Text6(Index))
       If DESTI$ = "" Then
          Call MENSERR(24, "Falta Destinatario")
       End If
       DIMAI$ = TRIM$(Text8(Index))
       If DIMAI$ = "" Or InStr(DIMAI$, "@") < 2 Then
          Call MENSERR(24, "Falta Direccion de E-mail")
       End If
       '
       Call CREAMAIL(COCLI$, DESTI$, DIMAI$, ASUNTO$)
       '
       ATENTI% = 1
       MAILMESS07.Label26 = Label26
       MAILMESS07.Show 1
   End If
   '
End Sub

Private Sub Text6_GotFocus(Index As Integer)
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text6_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   KCD% = KeyCode
   If KCD% = 38 Then
        If TRIM$(Text6(Index)) <> "" Then
           If Index > 0 Then
              If TRIM$(Text6(Index - 1)) = "" Then
                   Text6(Index - 1) = Text6(Index)
                   Text6(Index) = ""
                   Text6(Index - 1).SetFocus
                 Else
                   Call INTERCAM(Index)
              End If
           End If
        End If
     ElseIf KCD% = 40 Then
        If TRIM$(Text6(Index)) <> "" Then
           If Index < Text6.Count - 1 Then
              If TRIM$(Text6(Index + 1)) = "" Then
                    Text6(Index + 1) = Text6(Index)
                    Text6(Index) = ""
                    Text6(Index + 1).SetFocus
                 Else
                    Call INTERCAM(Index + 1)
              End If
           End If
        End If
   End If
End Sub

Sub INTERCAM(IKI%)
   '
   T6$ = Text6(IKI% - 1)
   T7$ = Text7(IKI% - 1)
   T8$ = Text8(IKI% - 1)
   T10$ = Text10(IKI% - 1)
   T30$ = Text30(IKI% - 1)
   Text6(IKI% - 1) = Text6(IKI%)
   Text7(IKI% - 1) = Text7(IKI%)
   Text8(IKI% - 1) = Text8(IKI%)
   Text10(IKI% - 1) = Text10(IKI%)
   Text30(IKI% - 1) = Text30(IKI%)
   Text6(IKI%) = T6$
   Text7(IKI%) = T7$
   Text8(IKI%) = T8$
   Text10(IKI%) = T10$
   Text30(IKI%) = T30$
   '
End Sub

Private Sub Text7_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   KCD% = KeyCode
   If KCD% = 38 Then
        If TRIM$(Text7(Index)) <> "" Then
           If Index > 0 Then
              If TRIM$(Text7(Index - 1)) = "" Then
                   Text7(Index - 1) = Text7(Index)
                   Text7(Index) = ""
                   Text7(Index - 1).SetFocus
                 Else
                   Call INTERCAM(Index)
              End If
           End If
        End If
     ElseIf KCD% = 40 Then
        If TRIM$(Text7(Index)) <> "" Then
           If Index < Text6.Count - 1 Then
              If TRIM$(Text7(Index + 1)) = "" Then
                 Text7(Index + 1) = Text7(Index)
                 Text7(Index) = ""
                 Text7(Index + 1).SetFocus
              End If
           End If
        End If
     ElseIf KCD% = 46 Then
        Text7(Index) = ""
   End If
End Sub

Private Sub Text8_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   KCD% = KeyCode
   If KCD% = 38 Then
        If TRIM$(Text8(Index)) <> "" Then
           If Index > 0 Then
              If TRIM$(Text8(Index - 1)) = "" Then
                   Text8(Index - 1) = Text8(Index)
                   Text8(Index) = ""
                   Text8(Index - 1).SetFocus
                 Else
                   Call INTERCAM(Index)
              End If
           End If
        End If
     ElseIf KCD% = 40 Then
        If TRIM$(Text8(Index)) <> "" Then
           If Index < Text6.Count - 1 Then
              If TRIM$(Text8(Index + 1)) = "" Then
                 Text8(Index + 1) = Text8(Index)
                 Text8(Index) = ""
                 Text8(Index + 1).SetFocus
              End If
           End If
        End If
   End If
End Sub

Private Sub Text10_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   KCD% = KeyCode
   If KCD% = 38 Then
        If TRIM$(Text10(Index)) <> "" Then
           If Index > 0 Then
              If TRIM$(Text10(Index - 1)) = "" Then
                   Text10(Index - 1) = Text10(Index)
                   Text10(Index) = ""
                   Text10(Index - 1).SetFocus
                 Else
                   Call INTERCAM(Index)
              End If
           End If
        End If
     ElseIf KCD% = 40 Then
        If TRIM$(Text10(Index)) <> "" Then
           If Index < Text6.Count - 1 Then
              If TRIM$(Text10(Index + 1)) = "" Then
                 Text10(Index + 1) = Text10(Index)
                 Text10(Index) = ""
                 Text10(Index + 1).SetFocus
              End If
           End If
        End If
   End If
End Sub

Private Sub Text30_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   KCD% = KeyCode
   If KCD% = 38 Then
        If TRIM$(Text30(Index)) <> "" Then
           If Index > 0 Then
              If TRIM$(Text30(Index - 1)) = "" Then
                   Text30(Index - 1) = Text30(Index)
                   Text30(Index) = ""
                   Text30(Index - 1).SetFocus
                 Else
                   Call INTERCAM(Index)
              End If
           End If
        End If
     ElseIf KCD% = 40 Then
        If TRIM$(Text30(Index)) <> "" Then
           If Index < Text6.Count - 1 Then
              If TRIM$(Text30(Index + 1)) = "" Then
                 Text30(Index + 1) = Text30(Index)
                 Text30(Index) = ""
                 Text30(Index + 1).SetFocus
              End If
           End If
        End If
   End If
End Sub

Private Sub Text7_GotFocus(Index As Integer)
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text8_DblClick(Index As Integer)
   Call Text6_DblClick(Index)
End Sub

Private Sub Text8_GotFocus(Index As Integer)
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text9_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Sub ACTAGRUP()
   Call TRALOCAL("TAGRUCON", "")
   VTB% = VENTABU%("TAGRUCON")
   If VTB% >= 0 Then
     For U& = 1 To ULTREG&("!TAGRUCON")
       XX$ = REGLEIDO$("!TAGRUCON", U&)
       UKI% = U&
       Call REPLA(XX$, MKI$(UKI%), 1, 2)
       Call GRAREG("!TAGRUCON", XX$, U&)
     Next U&
     Call TRACENTRAL("TAGRUCON", "")
     If ULTREG&("TAGRUCON") > 255 Then
       Call SETULTREG("TAGRUCON", 255)
     End If
   End If
End Sub

Sub SEGUIOPER()
   Screen.MousePointer = 11

   Dim NCON&(), PROCON%(), PRIMOV&()
   ' 1.- SELECCION CONTACTOS MARCADOS
   FIN& = ULTREG&("BSEGCOM7")
   CANSEL& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     RELE$ = REGLEIDO$("BSEGCOM7", U&)
     'If Mid$(RELE$, 62, 3) = "!!!" Then
       If TRIM$(Mid$(RELE$, 3, 59)) <> "" Then
         CANSEL& = CANSEL& + 1
         ReDim Preserve NCON&(CANSEL&), PROCON%(CANSEL&), PRIMOV&(CANSEL&)
         NCON&(CANSEL&) = XVALO(Mid$(RELE$, 4, 9))
         PROCON%(CANSEL&) = CVI(Mid$(RELE$, 1033, 2))
         PRIMOV&(CANSEL&) = 0
       End If
     'End If
    Next U&
   '
   ' 2.- ORDENAR POR FECHA PRÓXIMO CONTACTO
   For U1& = 1 To CANSEL&
     Call TRACE(20, U1&, CANSEL&)
     For U2& = U1& + 1 To CANSEL&
       If PROCON%(U2&) < PROCON%(U1&) Then
         NCONI% = NCON&(U2&)
         NCON&(U2&) = NCON&(U1&)
         NCON&(U1&) = NCONI%
         PROCONI% = PROCON%(U2&)
         PROCON%(U2&) = PROCON%(U1&)
         PROCON%(U1&) = PROCONI%
       End If
     Next U2&
   Next U1&
   '
   ' 3.- IDENTIFICAR ULTINA ANOTACIÓN
   For ISELL& = 1 To CANSEL&
     Call TRACE(20, ISELL&, CANSEL&)
     NPX& = NCON&(ISELL&)
     RFF$ = RECFILT$("ANOTASEG", 1, MKS$(NPX&))
     '
     For U& = Len(RFF$) To 1 Step -4
       RENOTA& = CVS(Mid$(RFF$, U& - 3, 4))
       XX$ = REGLEIDO$("ANOTASEG", RENOTA&)
       If TRIM$(Mid$(XX$, 5, 8)) <> "" Then
         If TRIM$(Mid$(XX$, 5, 8)) <> "========" Then
           PRIMOV&(ISELL&) = RENOTA&
           Exit For
         End If
       End If
     Next U&
   Next ISELL&
   '
   ' 4.- GENERAR LISTADO
   REBLA$ = REGBLAN$("SEGOPEN7")
   JJ& = 0
   For ISELL& = 1 To CANSEL&
     Call TRACE(20, ISELL&, CANSEL&)
     NPX& = NCON&(ISELL&)
     YY$ = FILTERGETRECORD$("BSEGCOM7", 1, MKS$(NCON&(ISELL&)))
     EMPRA$ = Mid$(YY$, 3, 59)
     NOMCON$ = Mid$(YY$, 449, 32)
     TELEF$ = Mid$(YY$, 377, 24)
     '
     RFF$ = RECFILT$("ANOTASEG", 1, MKS$(NPX&))
     KL& = 0
     '
     For U& = 1 To Len(RFF$) Step 4
       RENOTA& = CVS(Mid$(RFF$, U&, 4))
       If RENOTA& >= PRIMOV&(ISELL&) Then
         XX$ = REGLEIDO$("ANOTASEG", RENOTA&)
         ANOTA$ = Mid$(XX$, 5, 68)
         '
         ' REGISTROS DESCARTADOS
         ANOTB$ = TRIM$(ANOTA$)
         If ULTIA$ = "" Then
           If ANOTB$ = "" Then
             GoTo 29
           End If
         End If
         If Left$(ANOTB$, 10) = "==========" Then GoTo 29
         If ANOTB$ = "==" Then GoTo 29
         If ANOTB$ = "Viviana Sorrentino" Then GoTo 29
         If ANOTB$ = "Viviana" Then GoTo 29
         If ANOTB$ = "Sorrentino" Then GoTo 29
         If ANOTB$ = "Departamento" Then GoTo 29
         If ANOTB$ = "Comercial" Then GoTo 29
         If ANOTB$ = "FLEXOFT Soluciones Informáticas Integrales" Then GoTo 29
         If ANOTB$ = "Miembro de la Cámara de Empresas de Software y Servicios" Then GoTo 29
         If ANOTB$ = "Informáticos" Then GoTo 29
         If ANOTB$ = "Bauness 2170 - 5to 'D' - 1431 Buenos Aires - Argentina" Then GoTo 29
         If ANOTB$ = "Tel (00-54-11) 4524-1284 / 4541-1515 - e-mail:" Then GoTo 29
         If ANOTB$ = "info@flexoft.com.ar" Then GoTo 29
         If ANOTB$ = "Flavio Grasso." Then GoTo 29
         If ANOTB$ = "Miembro de la Cámara de Empresas de Software" Then GoTo 29
         If ANOTB$ = "y Servicios Informáticos" Then GoTo 29
         If ANOTB$ = "Tel (00-54-11) 4524-1284 / 4541-1515 / 4522-2048" Then GoTo 29
         If ANOTB$ = "e-mail: info@flexoft.com.ar" Then GoTo 29
         '
         ZZ$ = REBLA$
         If KL& = 0 Then
           Call REPLA(ZZ$, EMPRA$, 1, 32)
           Call REPLA(ZZ$, MKI$(PROCON%(ISELL&)), 137, 2)
           Call REPLA(ZZ$, NOMCON$, 139, 24)
           Call REPLA(ZZ$, TELEF$, 163, 24)
         End If
         Call REPLA(ZZ$, ANOTA$, 69, 68)
         KL& = KL& + 1
         JJ& = JJ& + 1
         Call GRAREG("SEGOPEN7", ZZ$, JJ&)
         ULTIA$ = TRIM$(ANOTA$)
       End If
29   Next U&
     JJ& = JJ& + 1
     Call GRAREG("SEGOPEN7", String$(256, "-"), JJ&)
   Next ISELL&
   Screen.MousePointer = 1

   Call SETULTREG("SEGOPEN7", JJ&)
   Call CIERRARCH("SEGOPEN7")
   Call FINAL("*SEGOPEN7")
   '
End Sub
Sub GENEFICHA()
     'REGENERA EL INDICE DE FICHAS
     '
     Call ABRECLIEN
     CONDI$ = "CODI_CLI <> '' ORDER BY Raso_Cli ASC"
     '
     
     Call CARGALISEL("INDIPROS", "BSEGCOM", "Codi_Cli/A12;Raso_Cli/A40", CONDI$, "NOMESS")
     Call FRESHECHO("INDIPROS")
     '
'     IDFICHA$ = TRIM$(Label26)
'     RASO2$ = TRIM$(Text4)
'     REGI$ = RECFILT$("INDIPROS", 1, TRIM$(IDFICHA$))
'     REG& = CVS(REGI$)
'     If REG& > 0 Then
'       XX$ = REGLEIDO$("INDIPROS", REG&)
'       RASO1$ = TRIM$(Mid$(XX$, 13, 50))
'       If RASO1$ <> RASO2$ Then
'          Call REPLA(XX$, AJUSTI$(RASO2$, 50), 13, 50)
'          Call GRAREG("INDIPROS", XX$, REG&)
'          Call CIERRARCH("INDIPROS")
'          Call FRESHECHO("INDIPROS")
'       End If
'     Else
'       XX$ = REGBLAN$("INDIPROS")
'       Call REPLA(XX$, IDFICHA$, 1, 12)
'       Call REPLA(XX$, AJUSTI$(RASO2$, 50), 13, 50)
'       Call REGAPP("INDIPROS", XX$)
'       Call CIERRARCH("INDIPROS")
'       Call FRESHECHO("INDIPROS")
'     End If
     '
End Sub

Sub ActualizaArchivo(OPCION%)
    Call ABRECLIEN
    '
    SQLX$ = "SELECT * FROM BSEGCOM "
    If OPCION% = 3 Then
      SQLX$ = SQLX$ + " WHERE SUBSTRING(RASO_CLI,60,3)= '!!!'"
      FIN& = CantRegistros&("BSEGCOM", "SUBSTRING(RASO_CLI,60,3)= '!!!'")
    End If
    
    SQLX$ = SQLX$ + " ORDER BY RASO_CLI"
    If OPCION% < 3 Then
      FIN& = CantRegistros&("BSEGCOM")
    End If
    Dim BSEGTMP As ADODB.Recordset
    Set BSEGTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    Call SETULTREG("BSEGCOM7", 0)
    
    REG& = 0
    With BSEGTMP
     Do While Not .EOF
        REG& = REG& + 1
        Call TRACE(20, REG&, FIN&)
        REGI$ = SAFETEXT$(!CODI_CLI)
        RELE$ = REGI$ + Space$(1022)
        Call REPLA(RELE$, AJUSTI$(!Raso_Cli, 62), 3, 62)   ' empresa
        Call REPLA(RELE$, SAFETEXT(!Mail_Cli), 65, 64)     ' DIRECCION ELECTRONICA
        Call REPLA(RELE$, SAFETEXT(!Domi_Cli), 129, 48)   ' CALLE Y NUMERO
        Call REPLA(RELE$, SAFETEXT(!Cpos_Cli), 177, 8)    ' C.POSTAL
        Call REPLA(RELE$, SAFETEXT(!Loca_Cli), 185, 32)   ' LOCALIDAD
         PROVI$ = TRIM$(SAFETEXT(!Prov_Cli))
        Call REPLA(RELE$, ECOARCH$("PROVINC", PROVI$), 217, 32)  ' PROVINCIA
        Call REPLA(RELE$, SAFETEXT(!Pais_Cli), 249, 24)   ' PAIS
        '
        Call REPLA(RELE$, SAFETEXT(!Origen), 273, 24)   ' ORIGEN
        Call REPLA(RELE$, SAFETEXT(!RamoActi), 297, 24)    ' RAMO
        Call REPLA(RELE$, SAFETEXT(!ProduPrin), 321, 48)    ' PRINC.PRODUCTO
'        Call REPLA(RELE$, XVALO(!Empleados), 361, 8) ' provincia particular
'        Call REPLA(RELE$, SAFETEXT(!Tele_Cli), 369, 24)   ' TELEFONO
'        Call REPLA(RELE$, SAFETEXT(!Tele_Cli2), 393, 24) ' telef.partic.
'        Call REPLA(RELE$, SAFETEXT(!Origen), 417, 24) ' telef2.partic
'        '
'        Call REPLA(RELE$, SAFETEXT(!Ctac_Cli), 441, 32)
'        Call REPLA(RELE$, SAFETEXT(!Cargo1), 473, 24)
'        Call REPLA(RELE$, SAFETEXT(!Sector1), 497, 24)
'        Call REPLA(RELE$, SAFETEXT(!Interno1), 521, 16)
'        '
'        Call REPLA(RELE$, SAFETEXT(!Contacto2), 537, 32)
'        Call REPLA(RELE$, SAFETEXT(!Cargo2), 569, 24)
'        Call REPLA(RELE$, SAFETEXT(!Sector2), 593, 24)
'        Call REPLA(RELE$, SAFETEXT(!Interno2), 617, 16)
'        '
'        Call REPLA(RELE$, SAFETEXT(!Contacto3), 633, 32)
'        Call REPLA(RELE$, SAFETEXT(!Cargo3), 665, 24)
'        Call REPLA(RELE$, SAFETEXT(!Sector3), 679, 24)
'        Call REPLA(RELE$, SAFETEXT(!Interno3), 713, 16)

        TTYZ$ = SAFETEXT$(!Nota_Cli)
        LOCADE& = Len(TTYZ$)
        For U& = LOCADE& To 1 Step -1
          If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 10
        Next U&
        U& = 0
        '
10      TEXTOXX$ = Left$(TTYZ$, U&)
        While Len(TEXTOXX$) > 0
           TTYY$ = MKI$(NCON%) + TEXTOXX$ + Space$(62)
           KK& = KK& + 4
           If KK& <= Len(RNC$) - 3 Then
                RENOTA& = CVS(Mid$(RNC$, KK&, 4))
              Else
                URN& = URN& + 1
                RENOTA& = URN&
           End If
           Call GRAREG("NOTASEG", TTYY$, RENOTA&)
           TEXTOXX$ = Mid$(TEXTOXX$, 63)
        Wend
        '
        Call REPLA(RELE$, SAFETEXT(!PAGWEB), 733, 64) ' WEB-SITE
        '
        IVENDE% = XVALO(!Vend_Cli)
        Call REPLA(RELE$, MKI$(IVENDE%), 900, 2)
        Call REPLA(RELE$, MKI$(FECHANUM(SAFETEXT$(!fechbaja))), 973, 2)   ' FECHA DE BAJA
        Call REPLA(RELE$, SAFETEXT$(!Mot_Baja), 975, 16)             ' MOTIVO DE BAJA
        Call REPLA(RELE$, MKI$(FECHANUM(SAFETEXT$(!FechProxCto))), 1023, 2)   ' FECHA PROXIMO CONTACTO
        HOII% = HOY(HOS$)
        If CVI(Mid$(RELE$, 1013, 2)) < HOII% Then
        Call REPLA(RELE$, MKI$(FECHANUM(SAFETEXT$(!FechUltEnvio))), 1025, 2)
        End If
        '
        Call GRAREG("BSEGCOM7", RELE$, REG&)
     .MoveNext
     Loop
    End With
    Call SETULTREG("BSEGCOM7", REG&)
    Call CIERRARCH("BSEGCOM7")
End Sub
'
Sub CREAMAIL(CODICLI$, Optional DESTI$, Optional DIMAI$, Optional ASUNTO$, Optional ORIMESS$)
   '
   KOKLI$ = TRIM$(CODICLI$)
   CONDICION$ = "CODI_CLI = '" & KOKLI$ & "'"
   '
   EMPRE$ = TRIM$(VALOBADA("BSEGCOM", "RASO_CLI", CONDICION$)) 'TRIM$(Mid$(RELE$, 3, 59))
   CALNUM$ = TRIM$(VALOBADA("BSEGCOM", "DOMI_CLI", CONDICION$)) 'TRIM$(Mid$(RELE$, 129, 48))
   CPOST$ = TRIM$(VALOBADA("BSEGCOM", "Cpos_Cli", CONDICION$)) 'TRIM$(Mid$(RELE$, 177, 8))
   If CPOST$ <> "" Then CPOST$ = CPOST$ + " "
   LOCAX$ = CPOST$ + TRIM$(VALOBADA("BSEGCOM", "Loca_Cli", CONDICION$)) ' CPOST$ + TRIM$(Mid$(RELE$, 185, 32))
   PROPA$ = TRIM$(VALOBADA("BSEGCOM", "Prov_Cli", CONDICION$)) 'TRIM$(Mid$(RELE$, 217, 32))
   PRO$ = TRIM$(ECOARCH$("provinc", PROPA$))
   If PRO$ <> "" Then PRO$ = PRO$ + " - "
   
   PRO$ = PRO$ + TRIM$(VALOBADA("BSEGCOM", "Pais_Cli", CONDICION$)) 'TRIM$(PROPA$ + TRIM$(Mid$(RELE$, 249, 24)))
   '
   CADEST% = 0
'   TTXX$ = TRIM$(Mid$(RELE$, 449, 32))
'   If TTXX$ <> "" Then
'      CADEST% = CADEST% + 1
'      DESTIX$ = TTXX$
'   End If
   '
   TTXX$ = ""
   If EMPRE$ <> "" Then TTXX$ = TTXX$ + EMPRE$ + Chr$(13) + Chr$(10)
   If CALNUM$ <> "" Then TTXX$ = TTXX$ + CALNUM$ + Chr$(13) + Chr$(10)
   If LOCAX$ <> "" Then TTXX$ = TTXX$ + LOCAX$ + Chr$(13) + Chr$(10)
   If PRO$ <> "" Then TTXX$ = TTXX$ + PRO$ + Chr$(13) + Chr$(10)
   '
   DESTIX$ = DESTI$
   MMT4$ = TRIM$(DESTI$)
   If MMT4$ <> "" Then
      If InStr(MMT4$, "@") > 0 Then
         CONDI44$ = "CODICLI = '" & KOKLI$ & "' AND MAIL_CTAC = '" & DESTI$ & "'"
         MMT44$ = VALOBADA("CONTACSCOM", "NOME_CTAC", CONDI44$, "NOMESS")
         If MMT44$ <> "" Then
           MMT4$ = MMT44$
           DESTIX$ = MMT4$
         End If
      End If
      MMT4$ = MMT4$ + " - "
      CADEST% = 1
   End If
   MAILMESS07.Text4 = TRIM$(MMT4$ + EMPRE$)
   MAILMESS07.Label26 = KOKLI$
   MAILMESS07.Combo1 = DESTIX$
   '
   MAIDES$ = REPLACAR$(TRIM$(FISECOM09.Text5), "/", ";")
   MAILMESS07.Text5 = MAIDES$: MAILMESS07.Text3 = ""
      If DIMAI$ <> "" Then
         If DIMAI$ <> MAIDES$ Then
            MAILMESS07.Text3 = MAILMESS07.Text5
            MAILMESS07.Text5 = DIMAI$
         End If
      End If
      '
   ASUNTX$ = TRIM$(ASUNTO$)
   If ASUNTX$ = "" Then ASUNTX$ = CONTROL$("SEGUICOM", "ASUNTO")
   If ASUNTX$ = "" Then ASUNTX$ = "FLEXOFT"
   If ASUNTX$ = "FLEXOFT" Then
      If TRIM$(DESTIX$) <> "" Then
         ASUNTX$ = ASUNTX$ + " para " + DESTIX$
      End If
   End If
   '
   MAILMESS07.Text20 = ASUNTX$
   '
   HOII% = HOY(HOS$)
   MAILMESS07.Label5 = "Buenos Aires, " + FETEXTO$(HOII%) + ".-"
   MARIZ$ = ""
   While MAILMESS07.Label5.Width < MAILMESS07.Text2.Width - 1500
     MAILMESS07.Label5 = " " + MAILMESS07.Label5
     MARIZ$ = " " + MARIZ$
   Wend
   '
   TTYY$ = MAILMESS07.Label5 + Chr$(13) + Chr$(10)
   If CADEST% < 1 Then
       TTYY$ = TTYY$ + "Sres." + Chr$(13) + Chr$(10)
    Else
       TTYY$ = TTYY$ + "Sr./Sra." + Chr$(13) + Chr$(10)
       TTYY$ = TTYY$ + DESTIX$ + Chr$(13) + Chr$(10)
   End If
   '
   TTYY$ = TTYY$ + TTXX$
   If DESTIX$ <> "" Then
      TTYY$ = TTYY$ + CRLF$ + MARIZ$ + "At: " + DESTIX$
   End If
   For KKII% = 1 To 2: TTYY$ = TTYY$ + CRLF: Next KKII%
   '
   ' RE-ENVIO A OTRO DESTINATARIO
   If ORIMAI = "SHOWMAIL09" Then
       MAILMESS07.List3.Clear
       On Error Resume Next
       MkDir LUCOM$ + "ARADJUN"
       On Error GoTo 0
       RUTAR$ = LUCOM$ + "ARADJUN\"
       '
       NUMAI& = XVALO(SHOWMAIL09.Label13)
       For KKII% = 1 To SHOWMAIL09.List3.ListCount
          NORMAI& = XVALO(Left$(SHOWMAIL09.List3.List(KKII% - 1), 3))
          CONDI$ = "NUME_MAIL = " & NUMAI& & " AND NORD_MAIL = " & NORMAI&
          FINAM$ = VALOBADA("ADJUNMAI", "FILE_NAME", CONDI$, "NOMESS")
          XXXX$ = VALOBADA("ADJUNMAI", "FILE_DATA", CONDI$, "NOMESS")
          If FINAM$ <> "" Then
             If XXXX$ <> "" Then
                Call SAVEFILE(RUTAR$ + FINAM$, XXXX$)
                DOCUACT$ = RUTAR$ + FINAM$
                MAILMESS07.List3.AddItem AJUSTI$(Mid$(SHOWMAIL09.List3.List(KKII% - 1), 4), 128) + DOCUACT$
             End If
          End If
       Next KKII%
       '
       TTZZ$ = SHOWMAIL09.Text2.TEXT
       PPXX% = InStr(UCase$(TTZZ$), "DE NUESTRA MAYOR CONSIDE")
       If PPXX% < 1 Then
          PPXX% = InStr(UCase$(TTZZ$), "ESTIMADO")
          If PPXX% > 128 Then PPXX% = 0
          If PPXX% < 1 Then
             Call MENSERR(24, "Imposible Pre-Armar Texto del Mail.\   \Debe Configurarse Manualmente.")
             PPXX% = 1
             GoTo 15
          End If
       End If
15     RESTODELMAIL$ = Mid$(TTZZ$, PPXX%)
       TTYY$ = TTYY$ + CRLF + String$(72, "*") + CRLF + RESTODELMAIL$
       ORIMAI = ""
       GoTo 40
   End If
   '
   '
   ' RE-ENVIO A OTRO DESTINATARIO
   If ORIMAI = "SHINMAIL09" Then
       MAILMESS07.List3.Clear
       On Error Resume Next
       MkDir LUCOM$ + "ARADJUN"
       On Error GoTo 0
       RUTAR$ = LUCOM$ + "ARADJUN\"
       '
       NUMAI& = XVALO(SHINMAIL09.Label13)
       For KKII% = 1 To SHINMAIL09.List3.ListCount
          NORMAI& = XVALO(Left$(SHINMAIL09.List3.List(KKII% - 1), 3))
          CONDI$ = "NUME_MAIL = " & NUMAI& & " AND NORD_MAIL = " & NORMAI&
          FINAM$ = VALOBADA("ADJUENTRA", "FILE_NAME", CONDI$, "NOMESS")
          XXXX$ = VALOBADA("ADJUENTRA", "FILE_DATA", CONDI$, "NOMESS")
          If FINAM$ <> "" Then
             If XXXX$ <> "" Then
                Call SAVEFILE(RUTAR$ + FINAM$, XXXX$)
                DOCUACT$ = RUTAR$ + FINAM$
                MAILMESS07.List3.AddItem AJUSTI$(Mid$(SHINMAIL09.List3.List(KKII% - 1), 4), 128) + DOCUACT$
             End If
          End If
       Next KKII%
       '
       RESTODELMAIL$ = SHINMAIL09.Text1.TEXT
       TTYY$ = TTYY$ + CRLF + "Re-envío correo recibido de " + SHINMAIL09.Label14 + " <" + SHINMAIL09.Label16 + ">" + CRLF + CRLF + CRLF + String$(72, "*") + CRLF + RESTODELMAIL$
       ORIMAI = ""
       GoTo 40
   End If
   '
   '
   TTYY$ = TTYY$ + "De nuestra mayor consideracion:" + CRLF + CRLF
   SST = Len(TTYY$)
   '
   If ORIMAI = "SHOWMESS09" Then
       TTYY$ = TTYY$ + SHOWMESS09.Text2
       ORIMAI = ""
   End If
   '
   For KKII% = 1 To 2: TTYY$ = TTYY$ + Chr$(13) + Chr$(10): Next KKII%
   TTYY$ = TTYY$ + "Quedamos a la espera de sus comentarios con un cordial saludo," + CRLF + CRLF
      PIEDMA$ = CONTROL_SQL$("SEGUICOM", "PIEDMAI-" & TRIM$(TRIM$(USNBR%)))
   TTYY$ = TTYY$ + PIEDMA$ + CRLF
   If InStr(UCase$(PIEDMA$), "FLEXOFT") > 0 Then
      TTYY$ = TTYY$ + "=========================================================" + Chr$(13) + Chr$(10)
      TTYY$ = TTYY$ + "FLEXOFT S.A. - Soluciones Informáticas Integrales" + Chr$(13) + Chr$(10)
      TTYY$ = TTYY$ + "Bauness 2170 - 5to. 'D' - 1431 Buenos Aires - Argentina" + Chr$(13) + Chr$(10)
      TTYY$ = TTYY$ + "Telefax: (00-54-11) 4524-1284 / 4522-2048 / 4524-0796 " + Chr$(13) + Chr$(10)
      TTYY$ = TTYY$ + "website: www.flexoft.com.ar - email: info@flexoft.com.ar" + Chr$(13) + Chr$(10)
      TTYY$ = TTYY$ + "=========================================================" + Chr$(13) + Chr$(10)
   End If
   '
   If ORIMESS$ <> "" Then
      TTYY$ = TTYY$ + CRLF + CRLF + String$(73, "*") + CRLF + ORIMESS$
   End If
   '
40 MAILMESS07.Text2 = TTYY$ ' = TTYY$ + Text1
   MAILMESS07.Text2.SelStart = SST
   MAILMESS07.Text2.SelLength = 0
   '
   MAILMESS07.List1.Clear
   Call ABRECLIEN
   ANOTACION$ = SAFETEXT$(VALOBADA("BSEGCOM", "HIS_CONTAC", CONDICION$))
   XX$ = ANOTACION$: LONGI& = Len(XX$)
   INICIO& = 1
   Do While INICIO& < LONGI&
        FIN& = InStr(INICIO&, XX$, Chr$(10))
        FINB& = InStr(INICIO&, XX$, CRLF$)
        If FINB& > 0 And FINB& < FIN& Then FIN& = FINB&
        If FIN& > 0 Then
           TTXX$ = Mid$(XX$, INICIO&, FIN& - INICIO&)
           MAILMESS07.List1.AddItem TTXX$
           INICIO& = FIN& + 2
        Else
           TTXX$ = Mid$(XX$, INICIO&)
           MAILMESS07.List1.AddItem TTXX$
           Exit Do
        End If
   Loop
   TPU& = MAILMESS07.List1.ListCount - 7
   If TPU& < 0 Then TPU& = 0
   MAILMESS07.List1.TopIndex = TPU&
   '
99 End Sub

Sub CARFICHA_DAT(NUME$)
   '
5  IFICHA% = Val(Label26)
   REG& = REGBUS&("BSEGCOM", 1, MKI$(IFICHA%))
   If REG& < 1 Or REG& > ULTREG&("BSEGCOM") Then
       Text1 = ""
       Text2 = ""
       Text3 = ""
       Text4 = ""
       Text5 = ""
       Text9 = ""
       Text16 = ""
       Text15 = ""
       Text13 = ""
       Text12 = ""
       Text25 = ""
       Text14 = ""
       Text15 = ""
       Text11 = ""
       Text20 = ""
       Text21 = ""
       Text22 = ""
       Text19 = ""
       Text23 = ""
       Text26 = ""
       Text16 = "": LABEL40 = ""
       Text17 = ""
       Text24 = ""
       Text18 = ""
       '
       For KKI% = 0 To 2
         Text6(KKI%) = ""
         Text7(KKI%) = ""
         Text8(KKI%) = ""
         Text10(KKI%) = ""
         Text30(KKI%) = ""
       Next KKI%
       '
       HOII% = HOY(HOS$)
       Label34 = HOS$
       Text12 = FECHATEX$(DIPOST(HOII%))
       '
       List2.Clear
       '
     Else
       RELE$ = REGLEIDO$("BSEGCOM", REG&)
       Text4 = TRIM$(Mid$(RELE$, 3, 59))
       Text15 = TRIM$(Mid$(RELE$, 62, 3))
       Text5 = TRIM$(Mid$(RELE$, 65, 64))
       Text20 = TRIM$(Mid$(RELE$, 129, 48))
       Text21 = TRIM$(Mid$(RELE$, 177, 8))
       '
       Text22 = TRIM$(Mid$(RELE$, 185, 32))
       Text19 = TRIM$(Mid$(RELE$, 217, 32))
       Text23 = TRIM$(Mid$(RELE$, 249, 24))
       Text18 = TRIM$(Mid$(RELE$, 273, 24))
       Text1 = TRIM$(Mid$(RELE$, 297, 24))
       Text2 = TRIM$(Mid$(RELE$, 321, 48))
       Text3 = TRIM$(Mid$(RELE$, 369, 8))
       '
       Text26 = TRIM$(Mid$(RELE$, 377, 24))
       Text17 = TRIM$(Mid$(RELE$, 401, 24))
       Text24 = TRIM$(Mid$(RELE$, 425, 24))
       '
       For KKI% = 0 To 2
         TTXX$ = Mid$(RELE$, 449 + KKI% * 96, 96)
         Text6(KKI%) = Mid$(TTXX$, 1, 32)
         Text7(KKI%) = Mid$(TTXX$, 33, 24)
         Text8(KKI%) = Mid$(TTXX$, 57, 24)
         Text10(KKI%) = Mid$(TTXX$, 81, 16)
       Next KKI%
       '
       Text13 = TRIM$(Mid$(RELE$, 737, 64))
       '
       IVENDE% = CVI(Mid$(RELE$, 971, 2))
       If IVENDE% = 8224 Then IVENDE% = 0
       Text16 = TRIM$(FORMATNUM$(IVENDE%, "I4"))
       '
       List2.Clear
       PUNGRU$ = Mid$(RELE$, 993, 30)
       UTACON% = ULTREG&("TAGRUCON")
       For UKI% = 1 To Len(PUNGRU$)
         RETA& = Asc(Mid$(PUNGRU$, UKI%, 1))
         If RETA& > 0 Then
           If RETA& <= UTACON% Then
             XX$ = REGLEIDO$("TAGRUCON", RETA&)
             RAGRE% = Val(CVI(Left$(XX$, 2)))
             TAGRE$ = AJUSTI$(Mid$(XX$, 51, 24), 48) + AJUSTI$(Mid$(XX$, 3, 48), 48) + Str$(RAGRE%)
             List2.AddItem TAGRE$
           End If
         End If
       Next UKI%
       '
       FEBAJA% = CVI(Mid$(RELE$, 973, 2))                ' FECHA DE BAJA
       Label38 = ""
       If FEBAJA% > 8224 Then Label38 = FECHATEX$(FEBAJA%)
       Text14 = Mid$(RELE$, 975, 16)                      ' MOTIVO DE BAJA
       '
       Text12 = FECHATEX$(CVI(Mid$(RELE$, 1023, 2)))
       '
       Text9 = "": TTYZ$ = ""
       If IFICHA% > 0 Then
         Screen.MousePointer = 11
         RNC$ = RECFILT$("NOTASEG", 1, MKI$(IFICHA%))
         URN& = ULTREG&("NOTASEG")
         '
         For U& = 1 To Len(RNC$) Step 4
           RENOTA& = CVS(Mid$(RNC$, U&, 4))
           If RENOTA& > 0 Then
             If RENOTA& <= URN& Then
               TTYZ$ = TTYZ$ + Mid$(REGLEIDO$("NOTASEG", RENOTA&), 3)
             End If
           End If
         Next U&
         Screen.MousePointer = 1
       End If
       '
       LOCADE& = Len(TTYZ$)
       For U& = LOCADE& To 1 Step -1
         If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 30
       Next U&
       U& = 0
30     Text9 = Left$(TTYZ$, U&)
       Text9.Refresh
       '
       CODICLIX$ = TRIM$(Label26)
       QTR& = CantRegistros("BSEGCOM", "Codi_Cli = '" & CODICLIX$ & "'")
       If QTR& < 1 Then
          Call Command1_Click    ' GUARDA SI ES INEXISTENTE
       End If
       '
       '
       NPX& = XVALO(CODICLIX$)
       RFF$ = RECFILT$("ANOTASEG", 1, MKS$(NPX&))
       '
       TTXX1$ = ""
       For U& = 1 To Len(RFF$) Step 4
         RENOTA& = CVS(Mid$(RFF$, U&, 4))
         XX$ = REGLEIDO$("ANOTASEG", RENOTA&)
         For KU% = 5 To 92
           If Asc(Mid$(XX$, KU%)) < 32 Then
             Mid$(XX$, KU%, 1) = " "
           End If
         Next KU%
         TTXX1$ = TTXX1$ + Mid$(XX$, 5, 92) + Chr$(10) + Chr$(13)
       Next U&
       '
       CONDI$ = "CODI_CLI = '" & CODICLIX$ & "'"
       ANOTANTE$ = TTXX1$ + VALOBADA("BSEGCOM", "HIS_CONTAC", CONDI$)
       If TRIM$(ANOTANTE$) <> "" Then
          Call ACTUREGISTRO("BSEGCOM", "HIS_CONTAC", CONDI$, ANOTANTE$, RAFE&)
       End If
       '
   End If
   '
End Sub

Private Sub VScroll1_Change()
   Frame15.Top = (-1) * VScroll1.Value * 100
End Sub

Private Sub VScroll1_Scroll()
   Frame15.Top = (-1) * VScroll1.Value * 100
End Sub

Sub ANOTAHISTORIAL(TTTT$, COCLI$, Optional CODOP$, Optional FANOTA%)
    '
    HOII% = HOY(HO$)
    Call FRATEXTO(TTTT$, 56)
    TTXX1$ = ""
    For KKI% = 1 To CARETEX%
        TTXX$ = Space$(76)
        Call REPLA(TTXX$, RETEX$(KKI%), 11, 56)
        If KKI% = 1 Then
           Call REPLA(TTXX$, HOS$, 1, 8)
           If FANOTA% > 0 Then
              Call REPLA(TTXX$, FECHATEX$(FANOTA%), 1, 8)
           End If
           Call REPLA(TTXX$, USERTER$, 69, 24)
        End If
        TTXX1$ = TTXX1$ + TTXX$ + CRLF$
    Next KKI%
    TTXX1$ = TTXX1$ + String$(88, "=") + CRLF$
    '
    CONDI$ = "CODI_CLI = '" & COCLI$ & "'"
    ANOTANTE$ = VALOBADA("BSEGCOM", "HIS_CONTAC", CONDI$)
    Call ACTUREGISTRO("BSEGCOM", "HIS_CONTAC", CONDI$, ANOTANTE$ + TTXX1$, RAFE&)
    '
    If CODOP$ <> "" Then
        CONDI$ = "CODI_CLI = '" & COCLI$ & "' and CodOport = '" & CODOP$ & "'"
        ANOTANTE$ = VALOBADA("OPORTUNI", "HIS_OPORTU", CONDI$)
        Call ACTUREGISTRO("OPORTUNI", "HIS_OPORTU", CONDI$, ANOTANTE$ + TTXX1$, RAFE&)
    End If
    '
    Call ANOTASEG09.CARGANOTA
    '
End Sub

Sub PROGRACTI(COCLI$, COPOR$, PERSCON$, REFACT$, FEPROG%)
    '
    PROACT$ = INICRM09.PROXIACTI$
    '
    ' VERIFICAR SI NO ESTA YA GRABADA
    CONDI$ = "Codi_Cli = '" & COCLI$ & "' "                              ' CODIGO DE PROSPECTO
    CONDI$ = CONDI$ & " AND CODOPORT = '" & COPOR$ & "' "                ' CODIGO DE OPORTUNIDAD
    CONDI$ = CONDI$ & " AND CONTACTO = '" & PERSCON$ & "' "              ' PERSONA DE CONTACTO
    CONDI$ = CONDI$ & " AND REFE_ACTI = '" & REFACT$ & "' "              ' REFERENCIA DE ACTIVIDAD
    CONDI$ = CONDI$ & " AND FHOR_PROG = '" & FETEXCOR1$(FEPROG%) & "' "  ' FECHA Y HORA PROGRAMADO
    If CantRegistros("ACTIPROY", CONDI$) > 0 Then GoTo 99                ' YA ESTABA REGISTRADA
    '
    Call CreaRegistro("ACTIPROY", "NUME_ACTI", PROACT$)
    '
    CONDI$ = "NUME_ACTI = '" & PROACT$ & "'"
    Call ACTUREGISTRO("ACTIPROY", "Codi_Cli", CONDI$, COCLI$, RAFE&, "NOMESS")               ' CODIGO DE PROSPECTO
    Call ACTUREGISTRO("ACTIPROY", "CODOPORT", CONDI$, COPOR$, RAFE&, "NOMESS")               ' CODIGO DE OPORTUNIDAD
    Call ACTUREGISTRO("ACTIPROY", "CONTACTO", CONDI$, PERSCON$, RAFE&, "NOMESS")             ' PERSONA DE CONTACTO
    Call ACTUREGISTRO("ACTIPROY", "REFE_ACTI", CONDI$, REFACT$, RAFE&, "NOMESS")             ' REFERENCIA DE ACTIVIDAD
    Call ACTUREGISTRO("ACTIPROY", "DESC_ACTI", CONDI$, "", RAFE&, "NOMESS")                  ' DESCRIPCION LARGA DE ACTIVIDAD
    Call ACTUREGISTRO("ACTIPROY", "USUA_PROG", CONDI$, USNBR% Mod 100, RAFE&, "NOMESS")      ' USUARIO PROGRAMADO
    Call ACTUREGISTRO("ACTIPROY", "USUA_ASIG", CONDI$, USNBR% Mod 100, RAFE&, "NOMESS")      ' USUARIO ASIGNADO
    Call ACTUREGISTRO("ACTIPROY", "FHOR_PROG", CONDI$, FETEXCOR1$(FEPROG%), RAFE&, "NOMESS") ' FECHA Y HORA PROGRAMADO
    Call ACTUREGISTRO("ACTIPROY", "USUA_REAL", CONDI$, 0, RAFE&, "NOMESS")                   ' USUARIO QUE REALIZO
    Call ACTUREGISTRO("ACTIPROY", "FHOR_REAL", CONDI$, 0, RAFE&, "NOMESS")                   ' FECHA DE REALIZACION
    Call ACTUREGISTRO("ACTIPROY", "COMENTAR", CONDI$, "", RAFE&, "NOMESS")                   ' COMENTARIO DE REALIZACION
    Call ACTUREGISTRO("ACTIPROY", "USUA_CTRL", CONDI$, 0, RAFE&, "NOMESS")                   ' USUARIO DE CONTROL Y APROBACION
    Call ACTUREGISTRO("ACTIPROY", "FHOR_CTRL", CONDI$, 0, RAFE&, "NOMESS")                   ' FECHA DE CONTROL
    Call ACTUREGISTRO("ACTIPROY", "STAT_ACTI", CONDI$, 0, RAFE&, "NOMESS")                   ' STATUS (0 = ACTIVA ; 3 = CERRADA)
    '
99  Screen.MousePointer = SCMA
    '
End Sub

Function VALIPROCON%(FEPROCO$, REFEACT$)
    '
    VPC% = 0
    FEPROG% = FECHANUM(FEPROCO$)
    If FEPROG% < HOY(HOS$) Then
       Call MENSERR(24, "Fecha Proximo Contacto No Valida.\Debe ser Igual o Posterior a la Actual.")
       GoTo 99
    End If
    '
    If TRIM$(REFEACT$) = "" Then
       Call MENSERR(24, "Mandatorio Indicar Asunto Proximo Contacto")
       GoTo 99
    End If
    VPC% = 1
    '
99  VALIPROCON% = VPC%
    '
End Function

Private Sub Picture5_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 1
End Sub

Private Sub PICTURE5_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Index > 0 Then
      Picture5(Index).MousePointer = 7
      If ARRABOR% = 1 Then
        If X > 0 Then
          If Index < 4 Then
              BORDELIM = Picture5(Index + 1).Left
            Else
              BORDELIM = Frame15.Width
          End If
          If Picture5(Index).Left > BORDELIM - 500 Then Exit Sub
        End If
        If X < 0 Then
          If Picture5(Index).Left < Picture5(Index - 1).Left + 500 Then Exit Sub
        End If
        Picture5(Index).Left = Picture5(Index).Left + X
        Label9(Index).Left = Picture5(Index).Left + 160
      End If
   End If
End Sub

Private Sub PICTURE5_MouseUP(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 0
   Picture5(Index).MousePointer = 1
End Sub

Sub AGREGARENGLONES()
   '
   INTERVER = 25
   NUPIC% = Text6.Count
   Load Text6(NUPIC%)
   Text6(NUPIC%) = ""
   Text6(NUPIC%).Left = Text6(NUPIC% - 1).Left
   Text6(NUPIC%).Top = Text6(NUPIC% - 1).Top + Text6(NUPIC% - 1).Height + INTERVER
   Text6(NUPIC%).Visible = True
   '
   Load Text7(NUPIC%)
   Text7(NUPIC%) = ""
   Text7(NUPIC%).Left = Text7(NUPIC% - 1).Left
   Text7(NUPIC%).Top = Text7(NUPIC% - 1).Top + Text7(NUPIC% - 1).Height + INTERVER
   Text7(NUPIC%).Visible = True
   '
   Load Text8(NUPIC%)
   Text8(NUPIC%) = ""
   Text8(NUPIC%).Left = Text8(NUPIC% - 1).Left
   Text8(NUPIC%).Top = Text8(NUPIC% - 1).Top + Text8(NUPIC% - 1).Height + INTERVER
   Text8(NUPIC%).Visible = True
   '
   Load Text10(NUPIC%)
   Text10(NUPIC%) = ""
   Text10(NUPIC%).Left = Text10(NUPIC% - 1).Left
   Text10(NUPIC%).Top = Text10(NUPIC% - 1).Top + Text10(NUPIC% - 1).Height + INTERVER
   Text10(NUPIC%).Visible = True
   '
   Load Text30(NUPIC%)
   Text30(NUPIC%) = ""
   Text30(NUPIC%).Left = Text30(NUPIC% - 1).Left
   Text30(NUPIC%).Top = Text30(NUPIC% - 1).Top + Text30(NUPIC% - 1).Height + INTERVER
   Text30(NUPIC%).Visible = True
   '
   For KKI% = 0 To Picture5.Count - 1
      If Picture5(KKI%).Height < Text6(NUPIC%).Top + Text6(NUPIC%).Height Then
        Picture5(KKI%).Height = Text6(NUPIC%).Top + Text6(NUPIC%).Height
      End If
   Next KKI%
   If Frame15.Height < Picture5(0).Top + Picture5(0).Height Then
        Frame15.Height = Picture5(0).Top + Picture5(0).Height
   End If
   '
   VScroll1.Max = (Frame15.Height - VScroll1.Height) / 100
   '
End Sub


