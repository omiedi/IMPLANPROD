VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form AJUSTK07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ajuste de Stocks a Fecha"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   9960
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   6720
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   67
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   7560
      OleObjectBlob   =   "AJUSTK07.frx":0000
      TabIndex        =   66
      Top             =   -120
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4725
      OleObjectBlob   =   "AJUSTK07.frx":007A
      Top             =   2835
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   7050
      Left            =   9120
      ScaleHeight     =   7020
      ScaleWidth      =   1980
      TabIndex        =   53
      Top             =   0
      Width           =   2010
      Begin VB.CommandButton AGenRep 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "AJUSTK07.frx":02AE
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   3240
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   68
         Top             =   6120
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   69
            Top             =   0
            Width           =   12000
         End
      End
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
         Height          =   650
         Left            =   -240
         Picture         =   "AJUSTK07.frx":05B8
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Ajuste de Lote a Cero"
         Top             =   5040
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         Enabled         =   0   'False
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
         Left            =   360
         Picture         =   "AJUSTK07.frx":0702
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Ajuste de Saldo por Lotes"
         Top             =   4320
         Visible         =   0   'False
         Width           =   650
      End
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
         Height          =   650
         Left            =   360
         Picture         =   "AJUSTK07.frx":084C
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Importar Inventario de Stocks"
         Top             =   5040
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command33 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   -240
         TabIndex        =   58
         ToolTipText     =   "Ajuste de Stocks a Cero"
         Top             =   4320
         Width           =   650
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Direct"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "AJUSTK07.frx":0996
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   1680
         Width           =   650
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "AJUSTK07.frx":0CA0
         Style           =   1  'Graphical
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   2400
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "AJUSTK07.frx":10E2
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   210
         Width           =   650
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "AJUSTK07.frx":122C
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   900
         Width           =   650
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00CDDADA&
      Caption         =   "CONSULTAS COMPLEMENTARIAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   4755
      TabIndex        =   43
      Top             =   105
      Width           =   4215
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3480
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox Picture12 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   1260
         ScaleHeight     =   615
         ScaleWidth      =   2505
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
            Height          =   645
            Left            =   0
            Picture         =   "AJUSTK07.frx":1536
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Consultas y Listados de Movimientos "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   630
            TabIndex        =   46
            Top             =   70
            Width           =   1830
         End
      End
      Begin VB.PictureBox Picture13 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   840
         ScaleHeight     =   615
         ScaleWidth      =   2505
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
            Height          =   645
            Left            =   -20
            Picture         =   "AJUSTK07.frx":1680
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Gestión de Inventario Continuo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   735
            TabIndex        =   49
            Top             =   40
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture14 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   420
         ScaleHeight     =   615
         ScaleWidth      =   2505
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
            Height          =   645
            Left            =   0
            Picture         =   "AJUSTK07.frx":198A
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Ficha Kardex de Movimientos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   52
            Top             =   40
            Width           =   1410
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00CDDADA&
      Caption         =   "IDENTIFICACIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      Left            =   90
      TabIndex        =   35
      Top             =   2640
      Width           =   4470
      Begin VB.CommandButton Command10 
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
         Height          =   285
         Left            =   4280
         TabIndex        =   71
         ToolTipText     =   "Solo Con Saldo"
         Top             =   940
         Width           =   175
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
         Left            =   4095
         TabIndex        =   61
         Top             =   940
         Width           =   175
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00CDDADA&
         Caption         =   "MOTIVO - REFERENCIA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   140
         TabIndex        =   40
         Top             =   1260
         Width           =   4185
         Begin VB.TextBox Text5 
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
            Height          =   285
            Left            =   210
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   315
            Width           =   3840
         End
      End
      Begin VB.TextBox Text4 
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
         Height          =   285
         Left            =   195
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   580
         Width           =   2700
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   3090
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
         Left            =   4140
         TabIndex        =   36
         Top             =   580
         Width           =   175
      End
      Begin VB.Label Label17 
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
         Left            =   2205
         TabIndex        =   63
         Top             =   945
         Width           =   1950
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lote / Partida"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   420
         TabIndex        =   62
         Top             =   1020
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   41
         Top             =   350
         Width           =   1065
      End
      Begin VB.Label Label4 
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
         Height          =   225
         Left            =   3120
         TabIndex        =   39
         Top             =   350
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "CONSULTAS - LISTADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3930
      Left            =   4755
      TabIndex        =   18
      Top             =   2650
      Width           =   4215
      Begin VB.Frame Frame6 
         BackColor       =   &H00CDDADA&
         Caption         =   "SALIDA"
         BeginProperty Font 
            Name            =   "Arial"
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
         Top             =   2625
         Width           =   3060
         Begin VB.OptionButton Option4 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Listado Impreso"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   630
            TabIndex        =   30
            Top             =   735
            Width           =   1905
         End
         Begin VB.OptionButton Option5 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Por Pantalla"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
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
         Picture         =   "AJUSTK07.frx":1C94
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2730
         Width           =   645
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00CDDADA&
         Caption         =   "RANGO NUMÉRICO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   120
         TabIndex        =   22
         Top             =   525
         Width           =   1800
         Begin VB.TextBox Text14 
            Alignment       =   2  'Center
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
            Height          =   285
            Left            =   360
            TabIndex        =   23
            Top             =   720
            Width           =   1065
         End
         Begin VB.TextBox Text13 
            Alignment       =   2  'Center
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
            Height          =   285
            Left            =   360
            TabIndex        =   24
            Top             =   1365
            Width           =   1065
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   360
            TabIndex        =   33
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   360
            TabIndex        =   32
            Top             =   1125
            Width           =   510
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00CDDADA&
         Caption         =   "RANGO DE FECHAS"
         BeginProperty Font 
            Name            =   "Arial"
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
            Alignment       =   2  'Center
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
            Height          =   285
            Left            =   360
            TabIndex        =   27
            Text            =   " "
            Top             =   1365
            Width           =   1065
         End
         Begin VB.TextBox Text11 
            Alignment       =   2  'Center
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
            Height          =   285
            Left            =   360
            TabIndex        =   25
            Text            =   " "
            Top             =   720
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
            Left            =   1440
            Picture         =   "AJUSTK07.frx":1F9E
            TabIndex        =   26
            Top             =   720
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
            Left            =   1440
            Picture         =   "AJUSTK07.frx":22A8
            TabIndex        =   28
            Top             =   1365
            Width           =   175
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   360
            TabIndex        =   21
            Top             =   1125
            Width           =   750
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   360
            TabIndex        =   20
            Top             =   495
            Width           =   735
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
      Left            =   3000
      TabIndex        =   1
      Top             =   210
      Width           =   175
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   620
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   600
      Width           =   4470
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3915
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   210
      Width           =   645
   End
   Begin VB.TextBox Text1 
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
      Height          =   285
      Left            =   690
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   210
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   90
      TabIndex        =   2
      Top             =   1365
      Width           =   4470
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "INGRESO DE AJUSTE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   90
      TabIndex        =   12
      Top             =   4800
      Width           =   4470
      Begin VB.TextBox Text7 
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
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   735
         Width           =   1650
      End
      Begin VB.TextBox Text8 
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
         Height          =   285
         Left            =   1785
         TabIndex        =   4
         Top             =   1050
         Width           =   1650
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
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
         Width           =   1650
      End
      Begin VB.TextBox Text9 
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
         Height          =   285
         Left            =   1785
         TabIndex        =   5
         Top             =   1365
         Width           =   1650
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
         Left            =   3615
         Picture         =   "AJUSTK07.frx":25B2
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
         Left            =   3615
         Picture         =   "AJUSTK07.frx":28BC
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Entrada"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   17
         Top             =   810
         Width           =   1590
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Salida"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   16
         Top             =   1140
         Width           =   1590
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "StockTeórico"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   15
         Top             =   510
         Width           =   1590
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Ajustado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   14
         Top             =   1440
         Width           =   1590
      End
   End
   Begin VB.Label Label16 
      Caption         =   "lote"
      Height          =   225
      Left            =   315
      TabIndex        =   60
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
      Caption         =   "U-M"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3360
      TabIndex        =   9
      Top             =   285
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   -240
      TabIndex        =   8
      Top             =   285
      Width           =   855
   End
   Begin VB.Menu Archiv 
      Caption         =   "Movimientos"
      Begin VB.Menu NUaju 
         Caption         =   "Nuevo Ajuste de Stock"
      End
      Begin VB.Menu mnuAjusteStockLecturaEtiqueta 
         Caption         =   "Ajuste de Stock por Lectura de Etiqueta (QR)"
      End
      Begin VB.Menu AnulaAjuste 
         Caption         =   "Anula Ajuste de Stock"
      End
      Begin VB.Menu spcl1 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuAjusteCeroLecturaEtiqueta 
         Caption         =   "Ajuste de Stock a Cero por Lectura de Etiqueta (QR)"
      End
      Begin VB.Menu AjusteACero 
         Caption         =   "Ajuste Masivo de Stock a Cero Por Depósito (No Trazables y Trazables)"
      End
      Begin VB.Menu AjusteACeroPorArtDep 
         Caption         =   "Ajuste de Stock A Cero Por Artículo y Depósito"
         Begin VB.Menu AjusteACeroArtDepManual 
            Caption         =   "Ajuste a Cero Manual"
         End
         Begin VB.Menu AjusteACeroArtDepExcel 
            Caption         =   "Ajuste a Cero vía Excel (Codigo-Deposito-Fecha)"
         End
      End
      Begin VB.Menu AjustePorLote 
         Caption         =   "Ajuste de Saldo por Lote (Solo Trazables)"
      End
      Begin VB.Menu spd3 
         Caption         =   "-"
      End
      Begin VB.Menu ImportaInventario 
         Caption         =   "Importador de Inventario"
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
      Caption         =   "Consultas y Listados"
      Begin VB.Menu Reports 
         Caption         =   "Consultas Movimientos"
      End
      Begin VB.Menu plxc 
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
      Begin VB.Menu Kardex 
         Caption         =   "Ficha Kardex"
      End
      Begin VB.Menu Invent 
         Caption         =   "Inventario Continuo"
      End
   End
   Begin VB.Menu mnuConfiguracion 
      Caption         =   "Configuración"
      Begin VB.Menu mnuTablaTaras 
         Caption         =   "Tabla de Taras"
      End
      Begin VB.Menu line100 
         Caption         =   "-"
      End
      Begin VB.Menu Setup 
         Caption         =   "Configuración General"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "AJUSTK07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MODENTRAJU%

Sub AJUSTE_APRO(MODO%)
   '
   'Call COPYSTRU("LILOTSE", "AJUSTEAP")
   Call TRALOCAL("AJUSTEAP", "")
   FIN& = ULTREG("AJUPEND")
   Call SETULTREG("AJUSTEAP", 0)
   Call TRALOCAL("AJUSTEAP", "")
   
   'CARGO LA LISTA DE SELECCION DESDE EL ARCHIVO AJUPEND
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TXTX$ = REGLEIDO$("AJUPEND", U&)
     
     CIII% = CVI(Mid$(TXTX$, 3, 2))
     If CIII% > 0 Then
        DOPRI$ = TRIM$(Mid$(TXTX$, 23, 10))
        DESCRI$ = AJUSTI$(DESCRIT0$(CIII%), 16)
        CANTINGRE = CVD(Mid$(TXTX$, 88, 8))
        CANTSALI = CVD(Mid$(TXTX$, 96, 8))
        NUMUSU% = Asc(Mid$(TXTX$, 126, 1))
        AJUST = CANTINGRE - CANTSALI
        Call REPLA(TTXX$, TRIM$(DOPRI$), 1, 16)
        Call REPLA(TTXX$, CODEXT$(CIII%), 18, 16)
        Call REPLA(TTXX$, AJUSTI$(DESCRIT0$(CIII%), 24), 36, 24)
        Call REPLA(TTXX$, FORMATNUM$(AJUST, "F10.1"), 62, 10)
        Call REPLA(TTXX$, USERNAM$(NUMUSU%), 74, 16)
        Call REGAPP("!AJUSTEAP", TTXX$)
     End If
   Next U&
   Call TRACENTRAL("AJUSTEAP", "")
   If U& < 2 Then Call COMUNI("No Hay Ajustes Para Autorizar"): Exit Sub
   '
   'PRESENTO LISTA DE SELECCION
50 Call SELECHO("AJUSTEAP/Cod.Movimiento          Código Artículo             Descripción                                                    Cantidad     Usuario")
   '
   If MODO% = 2 Then Exit Sub
   '
   DOPRI$ = TRIM$(VALACT1$("AJUSTEAP"))
   If DOPRI$ = "" Then Exit Sub
   '
   'VALIDACION DE USUARIO CASO GABAL
   User% = USNBR% Mod 100
   If User% <> 2 And User% <> 13 Then
      Call COMUNI("Usuario No Habilitado Para Autorizar Ajustes")
      Exit Sub
   End If
   '
   'LEO LA INFORMACION DEL REGISTRO SELECCIONADO
   TX1$ = FILTERGETRECORD$("AJUPEND", 6, DOPRI$)
   REFE$ = TRIM$(Mid$(TX1$, 43, 30))
   FF% = CVI(Mid$(TX1$, 1, 2))  'TOMA LA FECHA EN QUE SE GENERO, EL AJUSTE, NO LA DE LA APROBACION
   CIII% = CVI(Mid$(TX1$, 3, 2))
   LTX$ = TRIM$(Mid$(TX1$, 5, 16))
   DEPO% = Asc(Mid$(TX1$, 83, 1))
   CANTIENT1 = CVD(Mid$(TX1$, 88, 8))
   CANTISAL1 = CVD(Mid$(TX1$, 96, 8))
   CANTI = CANTIENT1 - CANTISAL1
   '
   If ESTRAZABLE%(CIXI%) > 0 Then
    If LTX$ = "" Then
         OPX% = COMALTER%("Material Trazable - Falta Lote / Partida a Ajustar\\Cancelar\Ingresar Lote")
         If OPX% > 1 Then
             LTX$ = UCase$(TRIM$(InputBox$("Lote", "Ingreso de Lote", LTX$)))
             If LTX$ = "" Then GoTo 50
         Else
             GoTo 50
         End If
    End If
   End If
   '
   If CANTI < 0.05 Then  ' es un movimiento de salida
          IDLO$ = LTX$
          CANSALI = CANTI
          If CANSALI > SaldoLote(CIII%, LTX$, DEPO%) + 0.05 Then
             Call MENSERR(24, "Cantidad del Ajuste Supera\Saldo Total del Lote\en el Depósito: " & DEPO%)
             GoTo 50
          End If
   End If
   '
   'CONSULTA DE CONFIRMACION DE AJUSTE
   OPX% = COMALTER%("Realmente Desea Aprobar el Ajuste \Para el Código " & CODEXT$(CIII%) & " Por Cantidad: " & CANTI & "\\Cancelar\Aprobar")
   If OPX% < 2 Then
      Call COMUNI("No se Ha Aprobado el Ajuste")
      'SI ELIGIO CONCELAR O MARCO X VUELVE ATRAS
      GoTo 50
   End If
   'FIN DE CONSULTA DE CONFIRMACION
   '
   'LEVANTO LOS DATOS PARA REALIZAR EL MOVIMIENTO DE STOCKS
   CMOV$ = TRIM$(Control$("", "CODAJUST"))
   If CMOV$ = "" Then CMOV$ = "AJ"
   '
   NVM$ = TRIM$(ProNroComprobante("AJ"))
   While Len(NVM$) < 6: NVM$ = "0" + NVM$: Wend
   NRO& = XVALO(NVM$)
   '
   DOSECO$ = "AJ." + Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Left$(HOS$, 2)
   DOPRICO$ = CMOV$ + "." + NVM$
   NUORIT% = 0
   VUNI# = 0
   MONEII% = 0
   NC% = 0
   '
   If TRIM$(REFE$) = "" Then
      REFE$ = "Ajuste de Inventario Autorizado"
   End If
   '
   'GENERA EL MOVIMIENTO
   Call MOVISTOK(FF%, CIII%, LTX$, CMOV$, DOPRICO$, DOPRICO$, DOSECO$, DOSECO$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPO%, CANTI)
   '
   'BORRO EL REGISTRO
   Call FILTERDELETE("AJUPEND", 6, AJUSTI$(DOPRI$, 10))
   '
   'SI EL CODIGO ESTA EN PANTALLA LO REFRESCA
   If TRIM$(Text1) = CODEXT$(CIII%) Then
      Call AJUSTO2
   End If
   '
   If ESTRAZABLE(CIII%) > 0 Then
       If Abs(CANTI) >= 0.05 Then
          IDLOTXX$ = LTX$
       End If
       Call ACSALOTE(CIII%, LTX$, "NOMESS")
   End If
   '
    
   FORIPRE$ = TRIM$(Control$("", "FORAJUST"))
    YAFORI% = 1
    If FORIPRE$ <> "" Then
      LU$ = LUDAT$
      If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
        RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
        Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Vale de Ajuste de Stocks:\  \'" + RMCC$ + "'.")
        YAFORI% = (-1)
      End If
   End If

   If TRIM$(FORIPRE$) <> "" Then
      If YAFORI% > 0 Then
         'VALIDACION DE MODO DE IMPRESION
         MODOIMPRE07.Check7.Caption = "Imprime Ajuste de Stocks"
         MODOIMPRE07.Label1.Caption = "" 'LIMPIA EL LABEL DONDE PONE OK, SI ACEPTA
         MODOIMPRE07.Check8.Value = 0
         Screen.MousePointer = 1
         MODOIMPRE07.Show 1
         CANCELPRINT% = 0
         If MODOIMPRE07.Label1.Caption = "OK" Then
           If MODOIMPRE07.Check8.Value > 0 Then CANCELPRINT% = 1
         End If
         Screen.MousePointer = 11
         '
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
           FEX = FF%
           Call PRINTFORWIN(FECHATEX$(FEX), "FECH-EMI", 0, 0)
           Call PRINTFORWIN(Time$, "HORA-EMI", 0, 0)
           Call PRINTFORWIN(USERNAME$, "USUARIO", 0, 0)
           Call PRINTFORWIN(Str$(NRO&), "NUME-COM", 0, 0)
           Call PRINTFORWIN(TRIM$(CODEXT$(CIII%)), "COD-MAT", 0, 0)
           Call PRINTFORWIN(DESCRIT$(CIII%), "DES-MAT", 0, 0)
           Call PRINTFORWIN(Unimed$(CIII%), "UNIMED", 0, 0)
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
   Call CIERRARCH("*.*")
End Sub

Sub AjusteA0SegunCodigoOfabricacionDeposito(CI1%, NORFA$, DEPO%)
 If DERACCE%("AJUCERO", "Ajuste de Stocks a Cero") < 2 Then
     GoTo 99
   End If
   '
10 Call SELECHO("SELFECHA")
   sFechaCorte$ = VALACT1$("SELFECHA")
   iFechaCorte% = FECHANUM(sFechaCorte$)
   If iFechaCorte% < 0 Then
     Call MENSERR(24, "Fecha de Ajuste no Válida")
     GoTo 10
   End If
   If iFechaCorte% = 0 Then
    GoTo 99
   End If
   '
   
20 If DEPO% < 1 Then 'SI NO VIENE PARAMETRIZADO LO PIDE
        Call SELECHO("TADEPOSI")
        iDeposito% = XVALO(VALACT1$("TADEPOSI"))
        If iDeposito% <= 0 Then
          GoTo 10
        End If
   Else
        iDeposito% = DEPO%
   End If
   '
   sDeposito$ = FORMATNUM$(iDeposito%, "F3.0") + " " + ECOARCH$("TADEPOSI", Chr$(iDeposito%))
   If COMALTER%("Esta Seguro de Ajustar a Cero\Todos los Stocks del Depósito Nro. : " & sDeposito$ & "\ al " & sFechaCorte$ & " \\No, Cancelar\Si, Ajustar a Cero") = 2 Then
        '
        Call ABRESTOCKS
        
        ' *** FILTRO POR DEPOSITO
        Call CARDEPOS
        iDepositoPrincipal% = 0
        If iDeposito% = COSU%(1) Then iDepositoPrincipal% = 1
        If iDepositoPrincipal% = 1 Then
            sFiltroDeposito$ = " AND (MOVISTO.DepoMovi = " & CStr(iDeposito%) & " OR MOVISTO.DepoMovi = 0) "
        Else
            sFiltroDeposito$ = " AND MOVISTO.DepoMovi = " & CStr(iDeposito%) & " "
        End If
        
        ' *** FILTRO POR ARTICULO
        If CI1% < 1 Then
            iPorArticulo% = ALTERNA%("Un Articulo\Sin Filtro Por Articulo")
            sFiltroArticulo$ = ""
            If iPorArticulo% = 1 Then
                Call SELECHO("MASTER")
                iArticulo% = CODINT(VALACT1$("MASTER"))
                If iArticulo% > 0 Then
                    sFiltroArticulo$ = " AND MASTER.CODINT = " & CStr(iArticulo%)
                End If
            End If
        Else
           iPorArticulo% = CI1%
        End If
        
        ' *** FILTRO POR STATUS DE ARTICULO
        sFiltroStatus$ = ""
        If iPorArticulo% > 1 Then
            iItemsBaja% = ALTERNA%("Solo Articulos De Baja\Todos Los Articulos")
            If iItemsBaja% = 1 Then sFiltroStatus$ = " AND MASTER.STATUS < 0 "
        End If
        
        Screen.MousePointer = 11
        
        ' *** PROXIMO NUMERO DE AJUSTE Y FORMATEO DEL NUMERO DE COMPROBANTE
        sCodMovimiento$ = TRIM$(Control$("", "CODAJUST"))
        If sCodMovimiento$ = "" Then sCodMovimiento$ = "AJ"
        NRO& = ProNroComprobante(sCodMovimiento$)
        sDocSecundarioCorto$ = "AJ." + Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Left$(HOS$, 2)
        sDocSecundarioLargo$ = sDocSecundarioCorto$
        sDocPrimarioCorto$ = "AS." + TRIM$(Str$(NRO&))
        sDocPrimarioLargo$ = sDocPrimarioCorto$

        ' ***********************************************
        ' *** AJUSTE DE STOCK A CERO PARA "" NO TRAZABLES ""
        ' ***********************************************
        Dim RstAjusteNoTrazables As ADODB.Recordset
        ConsultaSql$ = "SELECT SUM(CANTINGRE) AS INGRESOS, SUM(CANTSALID) AS SALIDAS, MASTER.CODINT, MOVISTO.COD_INTE, MOVISTO.DEPOMOVI "
        ConsultaSql$ = ConsultaSql$ & " FROM MASTER INNER JOIN MOVISTO WITH(NOLOCK) ON MASTER.CODINT = MOVISTO.COD_INTE WHERE MASTER.TRAZABLE = 0 "
        ConsultaSql$ = ConsultaSql$ & " AND MOVISTO.FECHMOV <= '" & FETEXCOR1$(iFechaCorte%) & "'"      ' FECHA DE CORTE
        ConsultaSql$ = ConsultaSql$ & sFiltroDeposito$                                                  ' DEPOSITO
        If sFiltroStatus$ <> "" Then ConsultaSql$ = ConsultaSql$ & sFiltroStatus$                       ' STATUS ARTICULO
        If sFiltroArticulo$ <> "" Then ConsultaSql$ = ConsultaSql$ & sFiltroArticulo$                   ' UN ARTICULO
        ConsultaSql$ = ConsultaSql$ + " AND MOVISTO.DOSECLAR = 'OF-000140-1'"
        ConsultaSql$ = ConsultaSql$ & " GROUP BY MOVISTO.COD_INTE, MOVISTO.DEPOMOVI, MASTER.CODINT "
        ConsultaSql$ = ConsultaSql$ & " ORDER BY MOVISTO.COD_INTE, MOVISTO.DEPOMOVI"
        
        Set RstAjusteNoTrazables = READSET(ConsultaSql$)
        
        iFecha% = iFechaCorte%
        iNroOrden% = 0
        With RstAjusteNoTrazables
            Do While Not .EOF
                iCodigo% = !CODINT
                iDeposito% = !depomovi
                
                dIngresos# = XVALO(!INGRESOS)
                dSalidas# = XVALO(!SALIDAS)
                dSaldo# = dIngresos# - dSalidas#
                dCantidadAjuste# = (-1) * dSaldo#   ' AJUSTA POR EL VALOR CONTRARIO
                
                If Abs(dCantidadAjuste) >= 0.005 Then
                    iNroOrden% = iNroOrden% + 1
                    dValorUnitario# = 0
                    sReferencia$ = "Inventario de Stocks a Cero"
                    iMoneda% = 0
                    iCliente% = 0
                    SLOTE$ = ""
                    Call MOVISTOK(iFecha%, iCodigo%, SLOTE$, sCodMovimiento$, sDocPrimarioLargo$, sDocPrimarioCorto$, sDocSecundarioLargo$, sDocSecundarioCorto$, sReferencia$, iNroOrden%, dValorUnitario#, iMoneda%, iCliente%, iDeposito%, dCantidadAjuste#)
                End If
                
                .MoveNext
            Loop
        End With
        
        On Error Resume Next
        RstAjusteNoTrazables.Close
        Set RstAjusteNoTrazables = Nothing
        On Error GoTo 0
        
        ' ***********************************************
        ' *** AJUSTE DE STOCK A CERO PARA "" TRAZABLES ""
        ' ***********************************************
        Dim RstAjusteTrazables As ADODB.Recordset
        ConsultaSql$ = "SELECT SUM(CANTINGRE) AS INGRESOS, SUM(CANTSALID) AS SALIDAS, MASTER.CODINT, MOVISTO.COD_INTE, MOVISTO.IDENLOTE, MOVISTO.DEPOMOVI "
        ConsultaSql$ = ConsultaSql$ & " FROM MASTER INNER JOIN MOVISTO WITH(NOLOCK) ON MASTER.CODINT = MOVISTO.COD_INTE WHERE MASTER.TRAZABLE > 0 "
        ConsultaSql$ = ConsultaSql$ & " AND MOVISTO.FECHMOV <= '" & FETEXCOR1$(iFechaCorte%) & "'"                          ' FECHA CORTE
        ConsultaSql$ = ConsultaSql$ & sFiltroDeposito$                                                                      ' DEPOSITO
        If sFiltroStatus$ <> "" Then ConsultaSql$ = ConsultaSql$ & sFiltroStatus$                                           ' STATUS ARTICULO
        If sFiltroArticulo$ <> "" Then ConsultaSql$ = ConsultaSql$ & sFiltroArticulo$ ' UN ARTICULO
        If NORFA$ <> "" Then
           ConsultaSql$ = ConsultaSql$ + " AND MOVISTO.DOSECLAR = " & NORFA$ & "'"
        End If
        ConsultaSql$ = ConsultaSql$ & " GROUP BY MOVISTO.COD_INTE, MOVISTO.IDENLOTE, MOVISTO.DEPOMOVI, MASTER.CODINT "
        ConsultaSql$ = ConsultaSql$ & " ORDER BY MOVISTO.COD_INTE, MOVISTO.IDENLOTE, MOVISTO.DEPOMOVI"
        
        Set RstAjusteTrazables = READSET(ConsultaSql$)
        
        With RstAjusteTrazables
            Do While Not .EOF
                iCodigo% = !CODINT
                SLOTE$ = TRIM$(!idenlote)
                iDeposito% = !depomovi
                
                dIngresos# = XVALO(!INGRESOS)
                dSalidas# = XVALO(!SALIDAS)
                dSaldo# = dIngresos# - dSalidas#
                dCantidadAjuste# = (-1) * dSaldo#   ' AJUSTA POR EL VALOR CONTRARIO
                
                If Abs(dCantidadAjuste) >= 0.005 Then
                    iNroOrden% = iNroOrden% + 1
                    dValorUnitario# = 0
                    sReferencia$ = "Ajuste de Stocks a Cero"
                    iMoneda% = 0
                    iCliente% = 0
                    Call MOVISTOK(iFecha%, iCodigo%, SLOTE$, sCodMovimiento$, sDocPrimarioLargo$, sDocPrimarioCorto$, sDocSecundarioLargo$, sDocSecundarioCorto$, sReferencia$, iNroOrden%, dValorUnitario#, iMoneda%, iCliente%, iDeposito%, dCantidadAjuste#)
                End If
                
                .MoveNext
            Loop
        End With
        
        On Error Resume Next
        RstAjusteTrazables.Close
        Set RstAjusteTrazables = Nothing
        On Error GoTo 0
    
        Screen.MousePointer = 1
 Else
     GoTo 20
   End If
   '
   Call COMUNI("Ajuste a Cero Finalizado Correctamente")
   Screen.MousePointer = 1
99 Command33.Enabled = True
'\\\OT-05-0704-WAR-FIN///
End Sub

Function CMOVDAT&()
   
   NUME_MAX& = 0
   FIN& = ULTREG&("AJUPEND")
   For i& = 1 To FIN&
      TX$ = REGLEIDO("AJUPEND", i&)
      DOPRI$ = Mid$(TX$, 23, 10) 'AJ.000001'
      NUME& = XVALO(Mid$(DOPRI$, 4)) '1
      If NUME_MAX& < NUME& Then
         NUME_MAX& = NUME&
      End If
   Next i&
   '
   CMOVDAT& = NUME_MAX&
   '
End Function

Function VERIFICAAJUSTE%(CI%)
    
    VERIFICAAJUSTE% = 0
    '
    'CONSULTA SOBRE EL MOVISTO AGRUPADA POR CODIGO,LOTE,DEPOSITO
    Call ABRESTOCKS
    SQLX$ = "SELECT COD_INTE,IDENLOTE, DEPOMOVI,"
    SQLX$ = SQLX$ + " SUM(CANTINGRE) AS INGRESO, SUM(CANTSALID)AS SALIDA "
    SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK) WHERE COD_INTE = " & CI%
    SQLX$ = SQLX$ + " GROUP BY COD_INTE, IDENLOTE, DEPOMOVI"
    '
    Dim CONTRLTMP As ADODB.Recordset
    Set CONTRLTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With CONTRLTMP
       Do While Not .EOF
         CANSALIDA = XVALO(!SALIDA)
         CANINGRESO = XVALO(!ingreso)
         If CANSALIDA > CANINGRESO Then
                'MENSAJE DE CONFIRMACION
                OPX% = COMALTER%("ATENCION: Se Han Encontrado Lotes con SobreConsumo\Proceso Necesario Para Continuar\Desea Ajustar a Cero?\\Si, Ajustar\No, Cancelar")
                If OPX% <> 1 Then
                    Text1 = ""
                    Me.Hide
                    Exit Function
                End If
                '
                'VALIDACION DE DERECHOS
                If DERACCE%("AJULOCER", "Ajuste de Lotes A Cero") < 2 Then
                  Text1 = ""
                  Me.Hide
                  Exit Function
                End If
                
                VERIFICAAJUSTE% = 1
                Exit Function
            End If
         .MoveNext
       Loop
    End With
    
End Function

Private Sub Acdir_Click()
   Call Command27_Click
End Sub

Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
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
Private Sub AjusteACero_Click()
    
    Call Command33_Click
    
End Sub

Private Sub AjusteACeroArtDepExcel_Click()

    'VALIDACION DE DERECHOS PARA AUMENTAR O DISMINUIR STOCKS

     If DERACCE%("AJUAUSTO", "Aumento de Stocks") < 2 Then Exit Sub

     If DERACCE%("AJUDISTO", "Disminución de Stocks") < 2 Then Exit Sub

     Dim oLibro As Excel.Workbook
     Dim oHoja As Excel.Worksheet
     
     Set obExcel = New Excel.Application
     CommonDialog1.CancelError = True
     CommonDialog1.DialogTitle = "Archivo Origen de Datos"
     CommonDialog1.FILTER = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
     CommonDialog1.FilterIndex = 1
     On Error GoTo ControlErrores
     CommonDialog1.ShowOpen
     On Error GoTo 0
     '
     DoEvents
     Screen.MousePointer = 11
     '
     Ruta$ = UCase$(CommonDialog1.FileName)
     Set oLibro = obExcel.Workbooks.Open(Ruta$)
     Set oHoja = oLibro.Sheets(1)
     '
     ReferenciaMovimiento$ = InputBox("Referencia", "Motivo del movimiento", "")
     '
     Dim IMPORTACION As Boolean
     UltimaFila& = UltimaFilaHojaExcel(oHoja, 1)
     For FilaLibro& = 2 To UltimaFila&
         '
         Call TRACE(24, FilaLibro&, UltimaFila&)
         On Error Resume Next
         
         CodigoArticulo$ = TRIM$(oHoja.Cells(FilaLibro&, 1))
         CodigoInterno% = CODINT%(CodigoArticulo$)
         DEPOSITO% = XVALO(TRIM$(oHoja.Cells(FilaLibro&, 2)))
         FechaAjuste% = FECHANUM(TRIM$(oHoja.Cells(FilaLibro&, 3)))
         If FechaAjuste% = 0 Then FechaAjuste% = HOY(HO$)
         If Err > 0 Then
            Call MENSERR(24, "Error al Leer Datos Desde la Planilla Excel.\Fila Número: " & CStr(FilaLibro&))
            On Error GoTo 0
            GoTo 25
         End If
         
         On Error GoTo 0
         
         If CodigoInterno% < 1 Then
             Call MENSERR(24, "Código '" & CodigoArticulo$ & "' Inexistente.\Dar de Alta y Repetir el Ajuste Para Este Artículo.")
             GoTo 20
         End If
         
         If DEPOSITO% < 1 And DepositoValido(DEPOSITO%) = True Then
             Call MENSERR(24, "Error en Depósito. Posibles Causas:\ \Falta Depósito Para el Código '" & CodigoArticulo$ & "'\Depósito No Existente\Depósito No Disponible.")
             GoTo 20
         End If
         
         Call AjusteACeroPorArticulo(CodigoInterno%, DEPOSITO%, FechaAjuste%, ReferenciaMovimiento$)
         
20   Next FilaLibro&
     '
25   On Error Resume Next
     oLibro.Close
     obExcel.Quit
     Set oLibro = Nothing
     Set obExcel = Nothing
     On Error GoTo 0
     '
     Screen.MousePointer = 1
     Call MENSERR(24, "Ajuste a Cero Finalizado")

99   Screen.MousePointer = 1
     Exit Sub

ControlErrores:
     Call MENSERR(24, "Error al Abrir Explorador de Archivos")
     GoTo 99
     
End Sub

Function UltimaFilaHojaExcel(oHoja As Excel.Worksheet, ColumnaDeReferencia As Integer) As Long

   Dim valor As Variant
   Dim CantFilasVacias&
   CantFilasVacias& = 0
   UltFila& = 0
   For Fila& = 2 To oHoja.Rows.Count
      If CantFilasVacias >= 100 Then
         GoTo 10
      End If
         valor = TRIM$(oHoja.Cells(Fila&, ColumnaDeReferencia))
         If valor = "" Then 'Si no hay valor
            If UltFila& = 0 Then 'Si no hay registro de ultima fila se registra, si no se mantiene la almacenada
               UltFila& = Fila& - 1
            End If
            CantFilasVacias& = CantFilasVacias& + 1
           Else
            CantFilasVacias& = 0
            UltFila& = 0
          End If
   Next Fila&
   
10 UltimaFilaHojaExcel = UltFila&

End Function

Private Sub AjusteACeroArtDepManual_Click()
    'VALIDACION DE DERECHOS PARA AUMENTAR O DISMINUIR STOCKS

     If DERACCE%("AJUAUSTO", "Aumento de Stocks") < 2 Then Exit Sub

     If DERACCE%("AJUDISTO", "Disminución de Stocks") < 2 Then Exit Sub

10  Call SELECHO("MASTER")
    CodigoArticulo$ = TRIM$(VALACT1$("MASTER"))
    CodigoInterno% = CODINT%(CodigoArticulo$)
    If CodigoInterno% < 1 Then GoTo 99
    
    Call SELECHO("TADEPOSI")
    DEPOSITO% = XVALO(VALACT1$("TADEPOSI"))
    If DEPOSITO% <= 0 Then GoTo 99
         
    Call SELECHO("SELFECHA")
    FechaAjuste% = FECHANUM(VALACT1$("SELFECHA"))
    If FechaAjuste% = 0 Then GoTo 99
    
    Screen.MousePointer = 11
    Call AjusteACeroPorArticulo(CodigoInterno%, DEPOSITO%, FechaAjuste%)
    Screen.MousePointer = 1
    
    Call MENSERR(24, "Ajuste a Cero Finalizado")
    GoTo 10
    
99  Screen.MousePointer = 1

End Sub

Private Sub AjustePorLote_Click()

    Call Command8_Click
    
End Sub

Private Sub Amasto_Click()
    ACONEC$ = "ARCHIG07/AMASTO"
    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
End Sub

Private Sub AnulaAjuste_Click()

    Call CARGALISEL("LISAJUS", "MOVISTO", "FECHMOV/D8;DOPRILAR/A18;REFERCOR/A32", "CODIMOVI='AJ' GROUP BY FECHMOV, DOPRILAR, REFERCOR ORDER BY FECHMOV DESC, DOPRILAR DESC")
    Call SELECHO("LISAJUS/Lista de Ajustes de Stock")
    
    IdAjuste$ = TRIM$(Mid$(VALACT2$("LISAJUS"), 1, 18))
    If IdAjuste$ = "" Then GoTo 99
    
    OpcionAjuste% = COMALTER%("¿Está Seguro Que Desea Anular el Ajuste '" & IdAjuste$ & "'?\\No, Cancelar\Si, Continuar")
    If OpcionAjuste% = 1 Then GoTo 99
    
    Screen.MousePointer = 11
    LOTE$ = TRIM$(SAFETEXT$(VALOBADA("MOVISTO", "IDENLOTE", "DOPRILAR='" & IdAjuste$ & "' AND CODIMOVI='AJ'")))
    CI% = XVALO(VALOBADA("MOVISTO", "COD_INTE", "DOPRILAR='" & IdAjuste$ & "' AND CODIMOVI='AJ'"))
    
    Call BORRAREGISTROS("MOVISTO", "DOPRILAR='" & IdAjuste$ & "' AND CODIMOVI='AJ'", REGAF&)
    
    If LOTE$ <> "" Then Call ACSALOTE(CI%, LOTE$)
    Screen.MousePointer = 1
    
    Call COMUNI("Anulación de Ajuste Realizada")
    
99  Exit Sub
    
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




Private Sub Command10_Click()
    CIXI% = CODINT%(Text1)
    If CIXI% < 1 Then Exit Sub
    If ESTRAZABLE%(CIXI%) < 1 Then Exit Sub
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))

    ILISTA& = List1.ListIndex
    If ILISTA& < 0 Then Exit Sub
    '
    DEPO% = COSU%(ILISTA& + 1)
    
    FAJUI% = HOY(HO$) 'AGREGUE PARA QUE MUESTRE COMO EN PANTALLA DE STOCKS POR QUE ESTABA MUY LENTO
    If InStr(EPAC$, "CAMBACERES") > 0 Or InStr(EPAC$, "QUILMES") > 0 Then
          SALOTES07.Label9 = TRIM$(Str$(DEPO%))
          SALOTES07.Label4 = Text1
          SALOTES07.Height = 5790
          Call SITSTOK07.MuestraLotesExpress(CODINT%(Text1), DEPO%)
          SALOTES07.Show 1
          Exit Sub
     End If
        
    Call ARMALILOT(CIXI%, DEPO%, FAJUI%, 0, 1)
    '
    Call FRESHECHO("!LILOTSE")
    If ULTREG("!LILOTSE") < 1 Then
       Call MENSERR(24, "No Hay Lotes Disponibles\para el Ajuste Propuesto.")
       Exit Sub
    End If
    '
    Call SELECHO("!LILOTSE/Lista de Seleccion de Lotes - Codigo: " + CODEXT$(CIXI%) + " - Deposito: " & DEPO%)
    VDEVU$ = VALACT1$("!LILOTSE")
    VDEV2$ = VALACT2$("!LILOTSE")
    CAMAXI = XVALO(Mid$(VDEV2$, 21, 12))
    'VALIDA SI TIENE CANTIDAD DE SALIDA VERIIFCA QUE NO SEA MAYOR AL LOTE
    If XVALO(Text8) >= FRACMIN(CIXI%) Then
      If XVALO(Text8) > CAMAXI Then
        Call MENSERR(24, "Cantidad del Ajuste Supera\Saldo Total del Lote")
        Exit Sub
      End If
    End If
    '
    Label17 = TRIM$(VDEVU$)
    '

      'End If
      '
    'End If
End Sub

Private Sub Command13_Click()
    Command13.Enabled = False
    QUERSTK07.Show 1
    Command13.Enabled = True
End Sub

Private Sub Command14_Click()
    Command14.Enabled = False
    If App.EXEName = "INVENT07" Then
      Command14.Enabled = True
      Hide
      Exit Sub
    End If
    Call CONECRUN("INVENT07", "Inventarios Físicos y Valorizados")
    Command14.Enabled = True
End Sub

Private Sub Command15_Click()
    Command15.Enabled = False
    FIKARDEX09.Text1 = Text1
    FIKARDEX09.Show 1
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
       IMPSTO07.Show 1
    End If
    Command25.Enabled = True
    '
End Sub

Private Sub Command26_Click()
   Call CONECRUN("FSETUP04/OPTRAJUST", "Configuración de Funcionamiento")
   Call FINAL("")
End Sub

Private Sub Command27_Click()
    Command27.Enabled = False
10  OPX% = ALTERNA%("Aprobación de Ajustes\Pendientes de Aprobación\Anulacion de Ajustes\Otros Accesos Directos")
    If OPX% = 1 Then
        Call AJUSTE_APRO(1)
        GoTo 10
    ElseIf OPX% = 2 Then
        Call AJUSTE_APRO(2)
        GoTo 10
    ElseIf OPX% = 3 Then
        Call ANULAAJUSTE_APRO
    ElseIf OPX% = 4 Then
        Call ACCDIR("AJUSTOCK")
    End If
    Screen.MousePointer = 1
    Command27.Enabled = True
    
End Sub

Sub ANULAAJUSTE_APRO()
   '
   'Call COPYSTRU("LILOTSE", "AJUSTEAP")
   Call TRALOCAL("AJUSTEAP", "")
   FIN& = ULTREG("AJUPEND")
   Call SETULTREG("AJUSTEAP", 0)
   Call TRALOCAL("AJUSTEAP", "")
   
   'CARGO LA LISTA DE SELECCION DESDE EL ARCHIVO AJUPEND
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TXTX$ = REGLEIDO$("AJUPEND", U&)
     
     CIII% = CVI(Mid$(TXTX$, 3, 2))
     If CIII% > 0 Then
        DOPRI$ = TRIM$(Mid$(TXTX$, 23, 10))
        DESCRI$ = AJUSTI$(DESCRIT0$(CIII%), 16)
        CANTINGRE = CVD(Mid$(TXTX$, 88, 8))
        CANTSALI = CVD(Mid$(TXTX$, 96, 8))
        NUMUSU% = Asc(Mid$(TXTX$, 126, 1))
        AJUST = CANTINGRE - CANTSALI
        Call REPLA(TTXX$, TRIM$(DOPRI$), 1, 16)
        Call REPLA(TTXX$, CODEXT$(CIII%), 18, 16)
        Call REPLA(TTXX$, AJUSTI$(DESCRIT0$(CIII%), 24), 36, 24)
        Call REPLA(TTXX$, FORMATNUM$(AJUST, "F10.1"), 62, 10)
        Call REPLA(TTXX$, USERNAM$(NUMUSU%), 74, 16)
        Call REGAPP("!AJUSTEAP", TTXX$)
     End If
   Next U&
   Call TRACENTRAL("AJUSTEAP", "")
   If U& < 2 Then Call COMUNI("No Hay Ajustes Para Autorizar"): Exit Sub
   '
   'PRESENTO LISTA DE SELECCION
50 Call SELECHO("AJUSTEAP/Cod.Movimiento          Código Artículo             Descripción                                                    Cantidad     Usuario")
   '
   'LEO EL NUMERO DE TRANFERENCIA SELECCIONADO
   DOPRI$ = TRIM$(VALACT1$("AJUSTEAP"))
   '
   If DOPRI$ = "" Then Exit Sub
   '
   OPX% = COMALTER%("¿Esta Seguro que Desea Anular el Ajuste de Stock '" & DOPRI$ & "'?\\No, Cancelar\Si, Anular")
   If OPX% = 2 Then
        DOPRI$ = AJUSTI$(DOPRI$, 10)
        Call FILTERDELETE("AJUPEND", 6, DOPRI$)
        Call COMUNI("Ajuste de Stock Anulado")
   End If
   '
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("AJUSTOCK")
End Sub

Private Sub Command3_Click()
    '
    Command3.Enabled = True
    CIXI% = CODINT%(Text1)
    If CIXI% < 1 Then Exit Sub
    On Error GoTo 0
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "CAMBACERES") > 0 Then
       If CIXI% > 0 Then
          If ESTRAZABLE%(CIXI%) = 1 And Unimed$(CIXI%) = "U." Then
             Call COMUNI("Realizar Ajuste a 0, Desde Pantalla de Movimientos por Lote\Ficha Kardex > Seleccionar Lote > Clic Hoja Amarilla > Ajuste a 0")
             GoTo 95
          End If
       End If
    End If
    
    'VALIDACION DE DERECHOS PARA AUMENTAR O DISMINUIR STOCKS
    If XVALO(Text7) > FRACMIN(CIXI%) Then ' ENTRADA
       If DERACCE%("AJUAUSTO", "Aumento de Stocks") < 2 Then Exit Sub
    End If
    If XVALO(Text8) > FRACMIN(CIXI%) Then ' SALIDA
       If DERACCE%("AJUDISTO", "Disminución de Stocks") < 2 Then Exit Sub
    End If
    '
    ILISTA& = List1.ListIndex
    DEPO% = COSU%(ILISTA& + 1)
    CAITMO% = 0
    CAN# = XVALO(Text8)
    If DEPO% > 0 Then
      If CAN# > 0.005 Then
        CAITMO% = 1
        CIMOVIM%(CAITMO%) = CIXI%
        CAMOVIM(CAITMO%) = (-1) * CAN#
        If PUEDEMOVI%(DEPO%, "NODESC") < 0 Then Exit Sub
      End If
    End If
    '
    FAJUI% = FECHANUM(Text10)
    If ESTRAZABLE%(CIXI%) < 1 Then GoTo 80
    '
'ACA SI ES AHP SALTEAR ESTE IF y ademas que sea movimiento de ingreso
    If Label17 = "" Then
       Call MENSERR(24, "Material Trazable - Seleccione Lote / Partida a Ajustar")
       Exit Sub
    End If
    If XVALO(Text8) > 0 Then  ' es un movimiento de salida
         IDLO$ = TRIM$(Label17)
         CANSALI = XVALO(Text8)
         If CANSALI > SaldoLote(CIXI%, IDLO$, DEPO%, FAJUI%) + 0.05 Then
            Call MENSERR(24, "Cantidad del Ajuste Supera\Saldo Total del Lote\en el Depósito y a la Fecha Elegidos")
            Exit Sub
         End If
         GoTo 80
       ElseIf XVALO(Text7) > 0 Then
         IDLOTE$ = TRIM$(Label17)
    End If
    '
80  Call AJUSTO3
    If Control$("STOCKS", "AJUSAPRO") = "SI" Then
       Text6 = ""
       Text9 = ""
       GoTo 95
       
    End If
    
    Call AJUSTO2
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "AHP") > 0 Then
       Call RecuperarLoteNotaDeCreditoParciales(CIXI%)
       Call RecuperarLote
'       If ESTRAZABLE(CIXI%) = 0 Then
'         OPX% = COMALTER%("Generar Lote Automático para el Movimiento de Ajuste ??\\No, Realizar Ajuste sin Lote\Si, Realizar Ajuste con  Lote Automático")
'         If OPX% > 1 Then AgregaLoteaAjusteSinLote_AHP (CIXI%)
'       End If
'Y ACA SI ES AHP SALTEAR REALIZAR LA ASIGNACION DE EL LOTE A LAS SALIDAS SIN RT.
    End If
    Text6.TEXT = Text9.TEXT
95  Text7.TEXT = ""
    Text8.TEXT = ""
    Label17 = ""
    Text1.SetFocus
99  Command3.Enabled = True
End Sub

Private Sub Command33_Click()
'\\\OT-05-0704-WAR-02/11/05///
  Command33.Enabled = False
   '
   'Dim CATEO#(32767)
   '
   If DERACCE%("AJUCERO", "Ajuste de Stocks a Cero") < 2 Then
     GoTo 99
   End If
   '
10 Call SELECHO("SELFECHA")
   sFechaCorte$ = VALACT1$("SELFECHA")
   iFechaCorte% = FECHANUM(sFechaCorte$)
   If iFechaCorte% < 0 Then
     Call MENSERR(24, "Fecha de Ajuste no Válida")
     GoTo 10
   End If
   If iFechaCorte% = 0 Then
    GoTo 99
   End If
   '
20 Call SELECHO("TADEPOSI")
   iDeposito% = XVALO(VALACT1$("TADEPOSI"))
   If iDeposito% <= 0 Then
     GoTo 10
   End If
   '
   sDeposito$ = FORMATNUM$(iDeposito%, "F3.0") + " " + ECOARCH$("TADEPOSI", Chr$(iDeposito%))
   If COMALTER%("Esta Seguro de Ajustar a Cero\Todos los Stocks del Depósito Nro. : " & sDeposito$ & "\ al " & sFechaCorte$ & " \\No, Cancelar\Si, Ajustar a Cero") = 2 Then
        '
        Call ABRESTOCKS
        
        ' *** FILTRO POR DEPOSITO
        Call CARDEPOS
        iDepositoPrincipal% = 0
        If iDeposito% = COSU%(1) Then iDepositoPrincipal% = 1
        If iDepositoPrincipal% = 1 Then
            sFiltroDeposito$ = " AND (MOVISTO.DepoMovi = " & CStr(iDeposito%) & " OR MOVISTO.DepoMovi = 0) "
        Else
            sFiltroDeposito$ = " AND MOVISTO.DepoMovi = " & CStr(iDeposito%) & " "
        End If
        
        ' *** FILTRO POR ARTICULO
        iPorArticulo% = ALTERNA%("Un Articulo\Sin Filtro Por Articulo")
        sFiltroArticulo$ = ""
        If iPorArticulo% = 1 Then
            Call SELECHO("MASTER")
            iArticulo% = CODINT(VALACT1$("MASTER"))
            If iArticulo% > 0 Then
                sFiltroArticulo$ = " AND MASTER.CODINT = " & CStr(iArticulo%)
            End If
        End If
        
        ' *** FILTRO POR STATUS DE ARTICULO
        sFiltroStatus$ = ""
        If iPorArticulo% > 1 Then
            iItemsBaja% = ALTERNA%("Solo Articulos De Baja\Todos Los Articulos")
            If iItemsBaja% = 1 Then sFiltroStatus$ = " AND MASTER.STATUS < 0 "
        End If
        
        Screen.MousePointer = 11
        
        ' *** PROXIMO NUMERO DE AJUSTE Y FORMATEO DEL NUMERO DE COMPROBANTE
        sCodMovimiento$ = TRIM$(Control$("", "CODAJUST"))
        If sCodMovimiento$ = "" Then sCodMovimiento$ = "AJ"
        NRO& = ProNroComprobante(sCodMovimiento$)
        sDocSecundarioCorto$ = "AJ." + Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Left$(HOS$, 2)
        sDocSecundarioLargo$ = sDocSecundarioCorto$
        sDocPrimarioCorto$ = "AS." + TRIM$(Str$(NRO&))
        sDocPrimarioLargo$ = sDocPrimarioCorto$

        ' ***********************************************
        ' *** AJUSTE DE STOCK A CERO PARA "" NO TRAZABLES ""
        ' ***********************************************
        Dim RstAjusteNoTrazables As ADODB.Recordset
        ConsultaSql$ = "SELECT SUM(CANTINGRE) AS INGRESOS, SUM(CANTSALID) AS SALIDAS, MASTER.CODINT, MOVISTO.COD_INTE, MOVISTO.DEPOMOVI "
        ConsultaSql$ = ConsultaSql$ & " FROM MASTER INNER JOIN MOVISTO WITH(NOLOCK) ON MASTER.CODINT = MOVISTO.COD_INTE WHERE MASTER.TRAZABLE = 0 "
        ConsultaSql$ = ConsultaSql$ & " AND MOVISTO.FECHMOV <= '" & FETEXCOR1$(iFechaCorte%) & "'"      ' FECHA DE CORTE
        ConsultaSql$ = ConsultaSql$ & sFiltroDeposito$                                                  ' DEPOSITO
        If sFiltroStatus$ <> "" Then ConsultaSql$ = ConsultaSql$ & sFiltroStatus$                       ' STATUS ARTICULO
        If sFiltroArticulo$ <> "" Then ConsultaSql$ = ConsultaSql$ & sFiltroArticulo$                   ' UN ARTICULO
        ConsultaSql$ = ConsultaSql$ & " GROUP BY MOVISTO.COD_INTE, MOVISTO.DEPOMOVI, MASTER.CODINT "
        ConsultaSql$ = ConsultaSql$ & " ORDER BY MOVISTO.COD_INTE, MOVISTO.DEPOMOVI"
        
        Set RstAjusteNoTrazables = READSET(ConsultaSql$)
        
        iFecha% = iFechaCorte%
        iNroOrden% = 0
        With RstAjusteNoTrazables
            Do While Not .EOF
                iCodigo% = !CODINT
                iDeposito% = !depomovi
                
                dIngresos# = XVALO(!INGRESOS)
                dSalidas# = XVALO(!SALIDAS)
                dSaldo# = dIngresos# - dSalidas#
                dCantidadAjuste# = (-1) * dSaldo#   ' AJUSTA POR EL VALOR CONTRARIO
                
                If Abs(dCantidadAjuste) >= 0.005 Then
                    iNroOrden% = iNroOrden% + 1
                    dValorUnitario# = 0
                    sReferencia$ = "Inventario de Stocks a Cero"
                    iMoneda% = 0
                    iCliente% = 0
                    SLOTE$ = ""
                    Call MOVISTOK(iFecha%, iCodigo%, SLOTE$, sCodMovimiento$, sDocPrimarioLargo$, sDocPrimarioCorto$, sDocSecundarioLargo$, sDocSecundarioCorto$, sReferencia$, iNroOrden%, dValorUnitario#, iMoneda%, iCliente%, iDeposito%, dCantidadAjuste#)
                End If
                
                .MoveNext
            Loop
        End With
        
        On Error Resume Next
        RstAjusteNoTrazables.Close
        Set RstAjusteNoTrazables = Nothing
        On Error GoTo 0
        
        ' ***********************************************
        ' *** AJUSTE DE STOCK A CERO PARA "" TRAZABLES ""
        ' ***********************************************
        Dim RstAjusteTrazables As ADODB.Recordset
        ConsultaSql$ = "SELECT SUM(CANTINGRE) AS INGRESOS, SUM(CANTSALID) AS SALIDAS, MASTER.CODINT, MOVISTO.COD_INTE, MOVISTO.IDENLOTE, MOVISTO.DEPOMOVI "
        ConsultaSql$ = ConsultaSql$ & " FROM MASTER INNER JOIN MOVISTO WITH(NOLOCK) ON MASTER.CODINT = MOVISTO.COD_INTE WHERE MASTER.TRAZABLE > 0 "
        ConsultaSql$ = ConsultaSql$ & " AND MOVISTO.FECHMOV <= '" & FETEXCOR1$(iFechaCorte%) & "'"                          ' FECHA CORTE
        ConsultaSql$ = ConsultaSql$ & sFiltroDeposito$                                                                      ' DEPOSITO
        If sFiltroStatus$ <> "" Then ConsultaSql$ = ConsultaSql$ & sFiltroStatus$                                           ' STATUS ARTICULO
        If sFiltroArticulo$ <> "" Then ConsultaSql$ = ConsultaSql$ & sFiltroArticulo$ ' UN ARTICULO

        ConsultaSql$ = ConsultaSql$ & " GROUP BY MOVISTO.COD_INTE, MOVISTO.IDENLOTE, MOVISTO.DEPOMOVI, MASTER.CODINT "
        ConsultaSql$ = ConsultaSql$ & " ORDER BY MOVISTO.COD_INTE, MOVISTO.IDENLOTE, MOVISTO.DEPOMOVI"
        
        Set RstAjusteTrazables = READSET(ConsultaSql$)
        
        With RstAjusteTrazables
            Do While Not .EOF
                iCodigo% = !CODINT
                SLOTE$ = TRIM$(!idenlote)
                iDeposito% = !depomovi
                
                dIngresos# = XVALO(!INGRESOS)
                dSalidas# = XVALO(!SALIDAS)
                dSaldo# = dIngresos# - dSalidas#
                dCantidadAjuste# = (-1) * dSaldo#   ' AJUSTA POR EL VALOR CONTRARIO
                
                If Abs(dCantidadAjuste) >= 0.005 Then
                    iNroOrden% = iNroOrden% + 1
                    dValorUnitario# = 0
                    sReferencia$ = "Ajuste de Stocks a Cero"
                    iMoneda% = 0
                    iCliente% = 0
                    Call MOVISTOK(iFecha%, iCodigo%, SLOTE$, sCodMovimiento$, sDocPrimarioLargo$, sDocPrimarioCorto$, sDocSecundarioLargo$, sDocSecundarioCorto$, sReferencia$, iNroOrden%, dValorUnitario#, iMoneda%, iCliente%, iDeposito%, dCantidadAjuste#)
                End If
                
                .MoveNext
            Loop
        End With
        
        On Error Resume Next
        RstAjusteTrazables.Close
        Set RstAjusteTrazables = Nothing
        On Error GoTo 0
    
        Screen.MousePointer = 1


'     SQLX$ = " SELECT DoPriCor FROM MOVISTO WITH(NOLOCK) "
'     SQLX$ = SQLX$ + "WHERE CodiMovi = 'AJ' "
'     SQLX$ = SQLX$ + "ORDER BY CAST(SUBSTRING(DoPriCor,4,12) AS INTEGER) DESC"
'     UNUM& = 0
'     Set MoviTemp = READSET(SQLX$)
'     On Error Resume Next
'     MoviTemp.MoveFirst
'     If Err < 1 Then
'       UNUM& = XVALO(Mid$(MoviTemp!DoPriCor, 4))
'     End If
'     On Error GoTo 0
'     NRO& = 1 + UNUM&
'     MoviTemp.Close
'     Set MoviTemp = Nothing
'     '
'     For KI% = 1 To NUMAS%
'       CATEO#(KI%) = 0
'     Next KI%
'     '
'     Call ABRESTOCKS
'     '
'     CMOV$ = TRIM$(CONTROL$("", "CODAJUST"))
'     If CMOV$ = "" Then CMOV$ = "AJ"
'     DOSECO$ = "AJ." + Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Left$(HOS$, 2)
'     DOPRICO$ = "AS." + TRIM$(Str$(NRO&))
'     DOPRILA$ = DOPRICO$
'     DOSELA$ = DOSECO$
'     '
'     Call CARDEPOS
'     DEPOPRIN% = 0
'     If DP1% = COSU%(1) Then DEPOPRIN% = 1
'     '
'     '\\\OT-06-0353-MH-29/08/06///
'     OPXX% = ALTERNA%("Solo Articulos De Baja\Todos Los Articulos")
'     '
'     For CI% = 1 To NUMAS%
'       '
'       If OPXX% = 1 Then
'         If STATITEM%(CI%) <> (-1) Then
'           GoTo 39
'         End If
'       End If
'     '\\\OT-06-0353-MH-29/08/06///
'       '
'       Screen.MousePointer = 11
'       FAXX$ = FETEXCOR1$(FA1%)
'       SQLX$ = " SELECT CANTINGRE , CANTSALID FROM MOVISTO WITH(NOLOCK)"
'       SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CI% & " "
'       'SQLX$ = SQLX$ + "AND INT(CDBL(FechMov))<= " & CDbl(CVDAT(FA1%)) & " "
'       SQLX$ = SQLX$ + " AND FechMov <= '" & FAXX$ & "' "
'       If DEPOPRIN% = 1 Then
'           SQLX$ = SQLX$ + "AND (DepoMovi = " & DP1% & " OR DepoMovi = 0)" & " "
'         Else
'           SQLX$ = SQLX$ + "AND DepoMovi = " & DP1% & " "
'       End If
'       '
'       'Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
''       Set MoviTemp = New ADODB.Recordset
''       MoviTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'       Set MoviTemp = READSET(SQLX$)
'       With MoviTemp
'         On Error Resume Next
'         .MoveFirst
'         If Err < 1 Then
'           Do While Not .EOF
'             CANTI1# = !CANTINGRE - !CANTSALID
'             CATEO#(CI%) = CATEO#(CI%) + CANTI1#
'             .MoveNext
'           Loop
'         End If
'       End With
'       MoviTemp.Close
'       Set MoviTemp = Nothing
'       '
'       If Abs(CATEO#(CI%)) < 0.001 Then GoTo 30
'       'NUORIT% = 0
'       NUORIT% = NUORIT% + 1
'       VUNI# = 0
'       MONEII% = 0
'       NC% = 0
'       DEPO% = DP1%
'       CANTI = CATEO#(CI%) * (-1)
'       REFE$ = "Ajuste de Stocks a Cero"
'       '
'       Call MOVISTOK(FA1%, CI%, LTX$, CMOV$, DOPRILA$, DOPRICO$, DOSELA$, DOSECO$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPO%, CANTI)
'       '
'30     IU& = XVALO(CI%)
'       FIN& = XVALO(NUMAS%)
'       Call TRACE(24, IU&, FIN&)
'       'Screen.MousePointer = 1
'39  Next CI%
'    '
   Else
     GoTo 20
   End If
   '
   Call COMUNI("Ajuste a Cero Finalizado Correctamente")
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
    Text2.TEXT = Unimed$(CI%)
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

Private Sub Command6_Click()
    CIXI% = CODINT%(Text1)
    If CIXI% < 1 Then Exit Sub
    If ESTRAZABLE%(CIXI%) < 1 Then Exit Sub
    '
    ILISTA& = List1.ListIndex
    If ILISTA& < 0 Then Exit Sub
    '
    DEPO% = COSU%(ILISTA& + 1)
    
    FAJUI% = FECHANUM(Text10)
    Call ARMALILOT(CIXI%, DEPO%, FAJUI%)
    '
    Call FRESHECHO("!LILOTSE")
    If ULTREG("!LILOTSE") < 1 Then
       Call MENSERR(24, "No Hay Lotes Disponibles\para el Ajuste Propuesto.")
       Exit Sub
    End If
    '
    Call SELECHO("!LILOTSE/Lista de Seleccion de Lotes - Codigo: " + CODEXT$(CIXI%) + " - Deposito: " & DEPO%)
    VDEVU$ = VALACT1$("!LILOTSE")
    VDEV2$ = VALACT2$("!LILOTSE")
    CAMAXI = XVALO(Mid$(VDEV2$, 21, 12))
    'VALIDA SI TIENE CANTIDAD DE SALIDA VERIIFCA QUE NO SEA MAYOR AL LOTE
    If XVALO(Text8) >= FRACMIN(CIXI%) Then
      If XVALO(Text8) > CAMAXI Then
        Call MENSERR(24, "Cantidad del Ajuste Supera\Saldo Total del Lote")
        Exit Sub
      End If
    End If
    '
    Label17 = TRIM$(VDEVU$)
    '

      'End If
      '
    'End If
 End Sub

Private Sub Command7_Click()
    '
    DESNUM& = 1: If XVALO(Text14.TEXT) > 0 Then DESNUM& = XVALO(Text14.TEXT)
    DESNUME$ = DESNUM&
    While Len(DESNUME$) < 6: DESNUME$ = "0" + DESNUME$: Wend
    'DESNUX$ = "AS." + TRIM$(Str$(DESNUME$))
    DESNUX$ = "AJ-" + DESNUME$               'PARA LA CONSULTA SQL EL DoPriCor DEBE SER AJ.XXXXXX
    HASNUM& = 999999
    HASNUME$ = HASNUM&
    If TRIM$(Text13.TEXT) <> "" Then
        If XVALO(Text13.TEXT) >= DESNUM& Then HASNUM& = XVALO(Text13.TEXT)
        HASNUME$ = HASNUM&
        While Len(HASNUME$) < 6: HASNUME$ = "0" + HASNUME$: Wend
    End If
    'HASNUX$ = "AS." + TRIM$(Str$(HASNUM&))
    HASNUX$ = "AJ-" + HASNUME$              'PARA LA CONSULTA SQL EL DoPriCor DEBE SER AJ.XXXXXX
    '
    DESFE% = 33: If FECHANUM(TRIM$(Text11.TEXT)) > 0 Then DESFE% = FECHANUM(TRIM$(Text11.TEXT))
    HASFE% = FECHANUM("31/12/49") '''cambiar para el sql max fecha admitida 2049
    If TRIM$(Text12.TEXT) <> "" Then
        If FECHANUM(TRIM$(Text12.TEXT)) >= DESFE% Then HASFE% = FECHANUM(TRIM$(Text12.TEXT))
    End If
    '
    Screen.MousePointer = 11
    Call BLANARCH("LISTRAJ")
    '
'    FEC1$ = FECHATEX$(DESFE%)
'    FEC2$ = FECHATEX$(HASFE%)
    
    FEC1$ = FETEXCOR1$(DESFE%)          'EL FORMATO DE LA FECHA DEBE SER  FETEXCOR1$()
    FEC2$ = FETEXCOR1$(HASFE%)
    
    '
    Call ABRESTOCKS
    'SQLX$ = " SELECT * FROM MOVISTO "
    'SQLX$ = SQLX$ + " WHERE CodiMovi =  'AJ' "
    'SQLX$ = SQLX$ + " and DoPriCor >= '" & DESNUX$ & "' "
    'SQLX$ = SQLX$ + " AND DoPriCor <= '" & HASNUX$ & "' "
    'SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) >= " & FEC1# & " "
    'SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & FEC2# & " "
    'SQLX$ = SQLX$ + "ORDER BY VAL(MID$(DoPriCor,4,7)) ASC"
    '
    'Set AJUSMOVTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    
    'cast(substring(DoPriCor,4,12)
    
    'sql
    SQLX$ = " SELECT ReferCor, COD_INTE, DEPOMOVI, CANTINGRE, FECHMOV, DoPriCor, CANTSALID  FROM MOVISTO WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE CodiMovi =  'AJ' "
    SQLX$ = SQLX$ + " AND DoPriCor >= '" & DESNUX$ & "' "
    SQLX$ = SQLX$ + " AND DoPriCor <= '" & HASNUX$ & "' "
    SQLX$ = SQLX$ + " AND FechMov >= '" & FEC1$ & "' "
    SQLX$ = SQLX$ + " AND FechMov <= '" & FEC2$ & "' "
    SQLX$ = SQLX$ + " ORDER BY cast(substring(DoPriCor,4,12) as integer) ASC"
'    Dim AJUSMOVTMP As ADODB.Recordset
'    Set AJUSMOVTMP = New ADODB.Recordset
'    AJUSMOVTMP.CursorType = adOpenDynamic
'    AJUSMOVTMP.LockType = adLockReadOnly
'    AJUSMOVTMP.Open FILTSQL$(SQLX$), FLEXCONN, , , adCmdText

    Set AJUSMOVTMP = READSET(SQLX$)
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
      CI1% = !cod_inte
      DP1% = !depomovi
      CAN# = !CANTINGRE
      FF% = CVINT(!FechMov)
      NRO& = Mid(!dopricor, 4)
      Y$ = REGBLAN$("LISTRAJ")
      Call REPLA(Y$, MKI$(FF%), 1, 2)
      Call REPLA(Y$, REFE$, 3, 48)
      Call REPLA(Y$, MKS$(NRO&), 51, 4)
      Call REPLA(Y$, MKI$(CI1%), 55, 2)
      Call REPLA(Y$, Unimed$(CI1%), 57, 4)
      If CAN# > 0 Then
          Call REPLA(Y$, MKD$(CAN#), 61, 8)
          Call REPLA(Y$, MKD$(!CANTSALID), 69, 8)
        Else
          Call REPLA(Y$, MKD$(CAN#), 61, 8)
          Call REPLA(Y$, MKD$(!CANTSALID), 69, 8)
      End If
      Call REPLA(Y$, Chr$(DP1%), 80, 1)
      Call REGAPP("LISTRAJ", Y$)
      
      .MoveNext
9      Loop
    End With
    '
    AJUSMOVTMP.Close
    Set AJUSMOVTMP = Nothing
    
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

Private Sub Command8_Click()
    Command8.Enabled = False
    '
    'VALIDACION DE DERECHOS PARA AUMENTAR O DISMINUIR STOCKS

    If DERACCE%("AJUAUSTO", "Aumento de Stocks") < 2 Then Exit Sub

    If DERACCE%("AJUDISTO", "Disminución de Stocks") < 2 Then Exit Sub

10  Call SELECHO("MASTER")
    COD$ = VALACT1$("MASTER")
    CI% = CODINT%(COD$)
    If CI% < 1 Then GoTo 99
    '
    If ESTRAZABLE(CI%) <> 1 Then
      Call MENSERR(24, "Nivel de Trazabilidad No Válido Para Ajustes por Lotes")
      GoTo 10
    End If
    '
20  Call SELECHO("TADEPOSI")
    VDEV$ = VALACT1$("TADEPOSI")
    If VDEV$ = "" Then GoTo 10
    DEPOSIT$ = VDEV$
    '
    DEPO% = XVALO(DEPOSIT$)
    
    Call ARMALILOT(CI%, DEPO%)
    '
    Call FRESHECHO("!LILOTSE")
    If ULTREG("!LILOTSE") < 1 Then
       Call MENSERR(24, "No Hay Lotes Disponibles\para el Ajuste Propuesto.")
       Exit Sub
    End If
    '
30  Call SELECHO("!LILOTSE/Lista de Seleccion de Lotes - Codigo: " + CODEXT$(CI%) + " - Deposito: " & DEPO%)
    VDEVU$ = VALACT1$("!LILOTSE")
    VDEV2$ = VALACT2$("!LILOTSE")
    CAMAXI = XVALO(Mid$(VDEV2$, 21, 12))
    REFER$ = "Diferencia de Inventario"
    
    If VDEVU$ = "" Then GoTo 20
    '
    Call REPLA(VDEF$, VDEVU$, 1, 12) 'LOTE
    Call REPLA(VDEF$, MKS$(CAMAXI), 13, 4) 'SALDO
    Call REPLA(VDEF$, MKS$(CAMAXI), 17, 4)
    Call REPLA(VDEF$, REFER$, 21, 35)
    '
40  OBX$ = OBJPLA$("AJUPOLOTE", VDEF$)
    If TRIM$(OBX$) = "" Then
         GoTo 30
    End If
    '
    LOTEX$ = Left$(OBX$, 12)
    SALDO = CVS(Mid$(OBX$, 13, 4))
    SAJUSTADO = CVS(Mid$(OBX$, 17, 4))
    REFE$ = Mid$(OBX$, 21, 35)
    
'    If XVALO(SAJUSTADO) > XVALO(SALDO) Then
'       Call MENSERR(24, "Cantidad del Ajuste Supera\Saldo Total del Lote")
'       Call REPLA(VDEF$, MKS$(SAJUSTADO), 17, 4)
'       GoTo 40
'    End If
    '
    VDEF$ = OBX$
    '
    If XVALO(SAJUSTADO) < 0 Then
       Call COMUNI("El Saldo Ajustado No Puede ser Menor a 0")
      
       GoTo 40
    End If
    '
    CANTI = XVALO(SAJUSTADO) - XVALO(SALDO)
    '
    If Abs(CANTI) < 0.05 Then
       Call COMUNI("Ajuste igual a Saldo")
       GoTo 40
    End If
    '
    NVM$ = TRIM$(ProNroComprobante("AJ"))
    While Len(NVM$) < 6: NVM$ = "0" + NVM$: Wend
    NRO& = XVALO(NVM$)
    FECHA% = HOY(HO$)
    '
    CMOV$ = TRIM$(Control$("", "CODAJUST"))
    If CMOV$ = "" Then CMOV$ = "AJ"
    DOSECO$ = "AJ." + Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Left$(HOS$, 2)
    DOPRICO$ = CMOV$ + "." + NVM$
    NUORIT% = 0
    VUNI# = 0
    MONEII% = 0
    NC% = 0
    '
    If TRIM$(REFE$) = "" Then
       REFE$ = "Diferencia de Inventario"
    End If
    '
    Call MOVISTOK(FECHA%, CI%, LOTEX$, CMOV$, DOPRICO$, DOPRICO$, DOSECO$, DOSECO$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPO%, CANTI)
    '
    If ESTRAZABLE(CI%) > 0 Then
       Call ACSALOTE(CI%, LOTEX$, "NOMESS")
    End If
    '
    Static YAFORI%, FORIPRE$
    '
    FORIPRE$ = TRIM$(Control$("", "FORAJUST"))
    YAFORI% = 1
    If FORIPRE$ <> "" Then
      LU$ = LUDAT$
      If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
        RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
        Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Vale de Ajuste de Stocks:\  \'" + RMCC$ + "'.")
        YAFORI% = (-1)
      End If
    End If

    If TRIM$(FORIPRE$) <> "" Then
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
               Call PRINTFORWIN(TRIM$(CODEXT$(CI%)), "COD-MAT", 0, 0)
               Call PRINTFORWIN(DESCRIT$(CI%), "DES-MAT", 0, 0)
               Call PRINTFORWIN(Unimed$(CI%), "UNIMED", 0, 0)
               Call PRINTFORWIN(ECOARCH$("ECODEP", Chr$(DEPO%)), "DEPOSITO", 0, 0)
               Call PRINTFORWIN(LOTEX$, "NRO-LOTE", 0, 0)
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
    Call CIERRARCH("*.*")
    
99  Screen.MousePointer = 1
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    '
    Command9.Enabled = False
    If DERACCE%("AJUACERO", "Ajustes de Stock / Lotes a Cero") >= 2 Then
       OPX% = COMALTER%("Ajuste de Lote a Cero.\\Ajustar Lotes a Cero\Anular Ajuste a Cero")
       If OPX% = 1 Then            'REALIZA AJUSTE DE LOTE A CERO
           XX$ = OBJPLA$("AJUSCERO", XX$)
           If XX$ = "" Then GoTo 99
           TIPMAT% = Asc(Mid$(XX$, 1, 1))
           SALAJUS = CVS(Mid$(XX$, 2, 4))
           '
           'OBTIENE NUMERO DEL COMPROBANTE
           NVM$ = TRIM$(ProNroComprobante("AJ"))
           While Len(NVM$) < 6: NVM$ = "0" + NVM$: Wend
           CMOV$ = TRIM$(Control$("", "CODAJUST"))
           If CMOV$ = "" Then CMOV$ = "AJ"
           NUCOMP$ = CMOV$ + "." + NVM$
           '
           CADECI% = XVALO(CNTDC$)
           '
           Screen.MousePointer = 11
           ' redondea movimientos de stock
           CONDI$ = "ROUND(cantingre,1) <> cantingre OR ROUND(cantsalid,1) <> cantsalid"
           FIN& = CantRegistros("MOVISTO", CONDI$)
           SQLX$ = "SELECT * FROM MOVISTO WHERE " + CONDI$
           Dim MOVISEL As ADODB.Recordset
           Set MOVISEL = New ADODB.Recordset
           MOVISEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           With MOVISEL
              JJ& = 0
              Do While Not .EOF
                 JJ& = JJ& + 1
                 Call TRACE(20, JJ&, FIN&)
                 CAIN = ROUND(XVALO(!CANTINGRE), CADECI%)
                 CASAL = ROUND(XVALO(!CANTSALID), CADECI%)
                 IDLO$ = SAFETEXT$(!idenlote)
                 CIXI% = XVALO(!cod_inte)
                 !CANTINGRE = CAIN
                 !CANTSALID = CASAL
                 .Update
                 Call ACSALOTE(CIXI%, IDLO$, "NOMESS")
              .MoveNext
              Loop
           End With
           MOVISEL.Close
           Set MOVISEL = Nothing
           '
           ' RECOLECTA LOS ARTICULOS QUE TIENEN EL TIPO DE MATERIAL ELEGIDO
           '
           Dim RSTIMATE As ADODB.Recordset
           Set RSTIMATE = FLEXCONN.Execute("SELECT CODINT FROM MASTER WHERE TIMATE='" & CStr(TIPMAT%) & "'")
           CONDI$ = ""
           With RSTIMATE
            Do While Not .EOF
                CONDI$ = CONDI$ & "cod_inte='" & SAFETEXT$(!CODINT) & "' or "
                .MoveNext
            Loop
           End With
           CONDI$ = AJUSTI$(CONDI$, Len(CONDI$) - 3)
           RSTIMATE.Close
           Set RSTIMATE = Nothing
           
           If TRIM$(CONDI$) = "" Then
                Call COMUNI("No Se Han Encontrado Artículos Que Pertenezcan al Tipo de Material Elegido")
                GoTo 99
           End If
           
           SQLX$ = "SELECT COD_INTE, IDENLOTE, DEPOMOVI, "
           SQLX$ = SQLX$ + " SUM(CANTINGRE) AS CAIN, SUM(CANTSALID) AS CASAL "
           SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK) "
           SQLX$ = SQLX$ + " WHERE " + CONDI$
           SQLX$ = SQLX$ + " GROUP BY COD_INTE, IDENLOTE, DEPOMOVI "
           
           Screen.MousePointer = 11
           
           Dim ACTCEROTMP As ADODB.Recordset
           Set ACTCEROTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
           With ACTCEROTMP
             Do While Not .EOF
                CI% = XVALO(!cod_inte)
                CANSALD = XVALO(!CAIN) - XVALO(!CASAL)
                If Abs(CANSALD) > 0.00000001 Then  ' FRACMIN(CI%) Then
                    If Abs(CANSALD) <= SALAJUS Then
                     LOTT$ = SAFETEXT(!idenlote)
                     DEPP% = XVALO(!depomovi)
                     Call AJUACERO(CI%, LOTT$, DEPP%, CANSALD, NUCOMP$)
                    End If
                End If

19             .MoveNext
             Loop
           End With
           ACTCEROTMP.Close
           Set ACTCEROTMP = Nothing
           '
           Call MENSERR(24, "Proceso de Ajuste de Lotes a Cero Finalizado")
           '
       ElseIf OPX% = 2 Then 'ANULA AJUSTE DE LOTE A CERO
           Call ANUAJUACERO
       End If
99     Screen.MousePointer = 1
    End If
    Command9.Enabled = True
    '
End Sub

Private Sub Exit_Click()
   Call Command1_Click
End Sub

Private Sub Form_Load()
    Call TRANSLABELS(Name)
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)

    '\\\OT-08-0990-OD-30/10/08///
    LOMA% = XVALO(Control$("MASTER", "LOMAXCOD"))
    If LOMA% > 15 Then Text1.MaxLength = LOMA%
    '\\\OT-08-0990-OD-FIN///
    Call AJUSTO1
    
    'Me.BackColor = RGB(218, 218, 205)
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

Private Sub ImportaInventario_Click()

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
    Label17 = "" ' label de lo
    If XVALO(Text9.TEXT) < 0 Then
        Text7.TEXT = FORMATNUM$((-1) * XVALO(Text9.TEXT), "F13." + CNTDC$(CODINT%(Text1)))
        Text9.TEXT = ""
    End If
    Text9.SetFocus
End Sub

Private Sub mnuAjusteStockBobinas_Click()

    ' LEE CODIGO QR
    
    
End Sub

Private Sub mnuAjusteCeroLecturaEtiqueta_Click()
    'VALIDACION DE DERECHOS PARA AUMENTAR O DISMINUIR STOCKS
'    If XVALO(Text7) > FRACMIN(CIXI%) Then ' ENTRADA
       If DERACCE%("AJUAUSTO", "Aumento de Stocks") < 2 Then Exit Sub
'    End If
'    If XVALO(Text8) > FRACMIN(CIXI%) Then ' SALIDA
       If DERACCE%("AJUDISTO", "Disminución de Stocks") < 2 Then Exit Sub
'    End If
    ' LEER ETIQUETA
    LecturaEtiqueta$ = TRIM$(InputBox("Lectura", "Lectura Etiqueta", ""))
    LecturaEtiqueta$ = REPLACAR$(LecturaEtiqueta$, "'", "-")
    LecturaEtiqueta$ = REPLACAR$(LecturaEtiqueta$, "]", "|")
    If LecturaEtiqueta$ = "" Then GoTo 99
    n% = InStr(1, LecturaEtiqueta$, "|")
    If n% < 1 Then GoTo 99
    CodigoProducto$ = TRIM$(Mid$(LecturaEtiqueta$, 1, n% - 1))
    LoteBobina$ = TRIM$(Mid$(LecturaEtiqueta$, n% + 1))
    
    ' PRESENTAR DATOS EN PANTALLA
    Text1 = CodigoProducto$
    CodigoProductoInterno% = CODINT%(CodigoProducto$)
    Label17 = UCase(LoteBobina$)
    
    Call CARDEPOS
    ILISTA& = List1.ListIndex
    DEPOSITO% = COSU%(ILISTA& + 1)
    SaldoBobinaTeorico# = SaldoLote(CodigoProductoInterno%, LoteBobina$, DEPOSITO%)
    MODENTRAJU% = 8
    Text8 = TRIM$(FORMATNUM$(SaldoBobinaTeorico#, "F10.1"))

99  Exit Sub

End Sub

Private Sub mnuAjusteStockLecturaEtiqueta_Click()
    'VALIDACION DE DERECHOS PARA AUMENTAR O DISMINUIR STOCKS
'    If XVALO(Text7) > FRACMIN(CIXI%) Then ' ENTRADA
       If DERACCE%("AJUAUSTO", "Aumento de Stocks") < 2 Then Exit Sub
'    End If
'    If XVALO(Text8) > FRACMIN(CIXI%) Then ' SALIDA
       If DERACCE%("AJUDISTO", "Disminución de Stocks") < 2 Then Exit Sub
'    End If
    '
    ' LEER ETIQUETA
    LecturaEtiqueta$ = TRIM$(InputBox("Lectura", "Lectura Etiqueta", ""))
    LecturaEtiqueta$ = REPLACAR$(LecturaEtiqueta$, "'", "-")
    LecturaEtiqueta$ = REPLACAR$(LecturaEtiqueta$, "]", "|")
    If LecturaEtiqueta$ = "" Then GoTo 99
    n% = InStr(1, LecturaEtiqueta$, "|")
    If n% < 1 Then GoTo 99
    CodigoProducto$ = TRIM$(Mid$(LecturaEtiqueta$, 1, n% - 1))
    LoteBobina$ = TRIM$(Mid$(LecturaEtiqueta$, n% + 1))
    
    If CantRegistros("lotingre", "idenlote='" & LoteBobina$ & "'") < 1 Then
        Call COMUNI("Imposible Continuar. Lote '" & LoteBobina$ & "' Inexistente.")
        GoTo 99
    End If
    
    ' PRESENTAR DATOS EN PANTALLA
    Text1 = CodigoProducto$
    Label17 = UCase(LoteBobina$)
    
    ' LECTURA PESO DESDE BALANZA
    CodigoProductoInterno% = CODINT(CodigoProducto$)
    PesoPorMetro# = CoeficientePesoPorMetroSegunLote#(LoteBobina$)
    BALANZA.txtCodigo = CodigoProducto$
    BALANZA.txtDescripcion = DESCRIT0$(CodigoProductoInterno%)
    BALANZA.Show 1
    If BALANZA.Respuesta = "APROBO" Then
        PESOLEIDO# = XVALO(BALANZA.txtPeso) - 2
        If PESOLEIDO# < 0.005 Then GoTo 80

        Dim ListaTaras As New DatosAuxiliares
        ListaTaras.ListaSelDataSql "TABLATARAS"
        PesoTara# = XVALO(RESP_ZELE_VECT(3))
        If PesoTara# > 0.005 Then
            If PESOLEIDO# > PesoTara# Then
                PESOLEIDO# = PESOLEIDO# - PesoTara#
            Else
                PESOLEIDO# = 0
            End If
        End If

        Set ListaTaras = Nothing

50      On Error Resume Next
        RstTaras.Close
        Set RstTaras = Nothing
        On Error GoTo 0
        
        Call CARDEPOS
        ILISTA& = List1.ListIndex
        DEPOSITO% = COSU%(ILISTA& + 1)
        SaldoBobinaTeorico# = SaldoLote(CodigoProductoInterno%, LoteBobina$, DEPOSITO%)
    
        ' CALCULAR METROS LINEALES
        SaldoBobinaEnMetrosLineales# = PESOLEIDO# / PesoPorMetro#
        SALIDA# = SaldoBobinaTeorico# - SaldoBobinaEnMetrosLineales#
        MODENTRAJU% = 8
        Text8 = TRIM$(FORMATNUM$(SALIDA#, "F10.1"))
    End If
    
80  Unload BALANZA
   
99  Exit Sub

End Sub

Private Sub mnuTablaTaras_Click()

    Dim TABLA As New DatosAuxiliares
    Call TABLA.CrearTablasAuxiliares
    
    IDLIST& = TABLA.ObtenerIdTablaAuxiliar("TABLATARAS")
    Call ABM_TABLA(IDLIST&, "Tabla de Taras")

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
        Text2.TEXT = Unimed$(CI%)
        Text3.TEXT = DESCRIT$(CI%)
        Call AJUSTO2               ' recalcular existencias a fecha
        '
'        If ESTRAZABLE%(CI%) = 1 Then
'            If VERIFICAAJUSTE%(CI%) > 0 Then
'               Call CORRIGESALDOLOTE(CI%)
'            End If
'        End If
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
    Text2.TEXT = Unimed$(CI%)
    '
End Sub
'

Private Sub Text1_LostFocus()
    COD$ = TRIM$(Text1.TEXT)
    CI% = CODINT%(COD$)
    Text3.TEXT = DESCRIT$(CI%)
    Text2.TEXT = Unimed$(CI%)
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

Private Sub Text2_gotfocus()
    Text7.SetFocus

End Sub

Private Sub Text3_Change()
   On Error Resume Next
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "CAMBACERES") > 0 Then
       COD$ = Text1.TEXT
       CI1% = CODINT%(COD$)
       If CI1% > 0 Then
          If ESTRAZABLE%(CI1%) And Unimed$(CI1%) = "U." Then
             Call COMUNI("Realizar Ajuste a 0, Desde Pantalla de Movimientos por Lote\Ficha Kardex > Seleccionar Lote > Clic Hoja Amarilla > Ajuste a 0")
          End If
       End If
    End If
    On Error GoTo 0
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
        Text9.TEXT = FORMATNUM(Str$(CAJUS#), "F13." & CNTDC$(CODINT(Text1)))
    End If
End Sub
'

Private Sub Text7_LostFocus()
   Text7 = FORMATNUM(XVALO(Text7), "F13." & CNTDC$(CODINT(Text1)))
End Sub

Private Sub Text8_Change()
    If MODENTRAJU% = 8 Then
        Text7.TEXT = ""
        CAJUS# = XVALO(Text6.TEXT) + XVALO(Text7.TEXT) - XVALO(Text8.TEXT)
        Text9.TEXT = FORMATNUM(Str$(CAJUS#), "F13." & CNTDC$(CODINT(Text1)))
    End If
End Sub
'

Private Sub Text8_LostFocus()
   Text8 = FORMATNUM(XVALO(Text8), "F13." & CNTDC$(CODINT(Text1)))
End Sub

Private Sub Text9_Change()
    If MODENTRAJU% = 9 Then
       DIFER# = XVALO(Text9.TEXT) - XVALO(Text6.TEXT)
       Text7.TEXT = ""
       Text8.TEXT = ""
       If DIFER# > 0 Then
           'Text7.TEXT = Str$(DIFER#)
           Text7.TEXT = FORMATNUM(Str$(DIFER#), "F13." & CNTDC$(CODINT(Text1)))
         ElseIf DIFER# < 0 Then
           'Text8.TEXT = Str$(Abs(DIFER#))
           Text8.TEXT = FORMATNUM(Abs(DIFER#), "F13." & CNTDC$(CODINT(Text1)))
           
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
      For KI% = 1 To 64
        CATEO#(KI%) = 0
      Next KI%
      '
      Call ABRESTOCKS
      felim$ = FETEXCOR1$(FF%)
      
      SQLX$ = " SELECT CANTINGRE, CANTSALID, DEPOMOVI, IdenLote FROM MOVISTO WITH(NOLOCK) "
      SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
      SQLX$ = SQLX$ + "AND FechMov<= '" + felim$ + "'"
'      Dim MoviTemp As ADODB.Recordset
'      Set MoviTemp = New ADODB.Recordset
'      MoviTemp.CursorType = adOpenStatic
'      MoviTemp.LockType = adLockReadOnly
'      MoviTemp.Open FILTSQL$(SQLX$), FLEXCONN, , , adCmdText
      Set MoviTemp = READSET(SQLX$)
      '
      With MoviTemp
        On Error Resume Next
        .MoveFirst
        If Err < 1 Then
          Do While Not .EOF
            CANTI# = XVALO(!CANTINGRE) - XVALO(!CANTSALID)
            CANTIS = XVALO(!CANTINGRE) - XVALO(!CANTSALID)
            DEP% = XVALO(!depomovi)
            LOT$ = SAFETEXT$(!idenlote)
            '
            COLUX% = 1
            For KI% = 1 To NUSU%
              If DEP% = COSU%(KI%) Then COLUX% = KI%
            Next KI%
            '
            CATEO#(COLUX%) = CATEO#(COLUX%) + CANTIS
          .MoveNext
          Loop
        End If
      End With
    End If
    '
    MoviTemp.Close
    Set MoviTemp = Nothing
    
    List1.Clear
    For U& = 1 To NUSU%
       NUM$ = FORMATNUM(CATEO#(U&), "F13." & CNTDC$(CI%))
       List1.AddItem SUCU$(U&) + NUM$ ' CSTRING$(MKS$(CAN), 4, 10, 1, 2)
    Next U&
    '
    List1.ListIndex = ILISTA&
    List1.Refresh
    Text4.TEXT = Left$(List1.TEXT, 20)
    
    SALDX = XVALO(TRIM$(Mid$(List1.TEXT, 21)))
    Text6.TEXT = FORMATNUM(SALDX, "F13." & CNTDC$(CI%))
    Text7.TEXT = ""
    Text8.TEXT = ""
    AJUSTADOX = XVALO(TRIM$(Mid$(List1.TEXT, 21)))
    Text9.TEXT = FORMATNUM(AJUSTADOX, "F13." & CNTDC$(CI%))
    If XVALO(Text9.TEXT) < 0 Then
        Text7.TEXT = FORMATNUM$((-1) * XVALO(Text9.TEXT), "f13." + CNTDC$(CI%))
        Text9.TEXT = ""
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
    FORIPRE$ = TRIM$(Control$("", "FORAJUST"))
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
    Call ABRESTOCKS
    '
    CMOV$ = TRIM$(Control$("", "CODAJUST"))
    If CMOV$ = "" Then CMOV$ = "AJ"
    NVM$ = TRIM$(ProNroComprobante("AJ"))
    While Len(NVM$) < 6: NVM$ = "0" + NVM$: Wend
    NRO& = XVALO(NVM$)
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
    If ESTRAZABLE(CIII%) < 1 Then LTX$ = ""
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
        CMOV$ = TRIM$(Control$("", "CODAJUST"))
        If CMOV$ = "" Then CMOV$ = "AJ"
        DOSECO$ = "AJ." + Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Left$(HOS$, 2)
        DOPRICO$ = CMOV$ + "." + NVM$
        NUORIT% = 0
        VUNI# = 0
        MONEII% = 0
        NC% = 0
        REFE$ = Text5.TEXT
        If TRIM$(REFE$) = "" Then
            REFE$ = "Diferencia de Inventario"
        End If
        '
        'SI ESTA CONFIGURADA LA OPCION SETUP STOCKS DE AUSTE CON APROBACION
        'GUARDA EL AJSUTE EN EL DAT MOVISTO, PARA QUE LUEGO SE HAGA LA APROBACION
        If Control$("STOCKS", "AJUSAPRO") = "SI" Then
            '
            Call COPYSTRU("MOVISTO", "AJUPEND")
            '
            MONE$ = ""
            '
            NVM$ = TRIM$(Str$(CMOVDAT& + 1)) ' IDENTIFICADOR CMOVDAT& ES UNA FUNCION
            While Len(NVM$) < 6: NVM$ = "0" + NVM$: Wend
            DOPRICO$ = CMOV$ + "." + NVM$
            DOPRICO$ = TRIM$(DOPRICO$)
            Call MOVISTOCK1(FF%, CIII%, LTX$, CMOV$, DOPRICO$, DOSECO$, REFE$, VUNI#, "", 0, DEPO%, CANTI)
            
            Call CIERRARCH("AJUPEND")
            '
            Call COMUNI("Ajuste Registrado Pendiente de Aprobación")
            '
            'SI ENTRO ACA SE VA
            Exit Sub
        End If
        If InStr(1, UCase(EMPREAC$), "PREMOTEC") > 0 Then
           'SI ES PREMOTEC EL AJUSTE TOMA LA FECHA DE ALTA DEL LOTE.
           CONDI$ = "IDENLOTE = '" & LTX$ & "'"
           FEALT$ = VALOBADA("LOTINGRE", "FECHALTA", CONDI$, "NOMESS")
           FF% = CVINT(FEALT$): If FF% < 1 Then FF% = HOY(HO$)
        End If
        If InStr(1, UCase(EMPREAC$), "CONTEPLAST") > 0 Then
           If LTX$ <> "" And TRIM$(UCase$(Unimed(CIII%))) = "M." Then
            METROS1# = Abs(CANTI)
            PesoXMetro# = CoeficientePesoPorMetroSegunLote#(LTX$)
            PESO1# = Abs(CANTI * PesoXMetro#)
           End If
        End If
        
        Call MOVISTOK(FF%, CIII%, LTX$, CMOV$, DOPRICO$, DOPRICO$, DOSECO$, DOSECO$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPO%, CANTI, , , , , , , , , , , METROS1#, PESO1#)
        
        '
        If ESTRAZABLE(CIII%) > 0 Then
           If Abs(CANTI) >= 0.05 Then
              IDLOTXX$ = LTX$
           End If
           Call ACSALOTE(CIII%, LTX$, "NOMESS")
        End If
        
        ' *** ARCON: TRANSFIERE AJUSTE DE STOCK A LA SUCURSAL DE BUENOS AIRES
        If InStr(1, EMPREAC$, "NEUQUEN") > 0 Then
        
            NROO& = XVALO(NVM$)
            IDOCUM& = 10000& * (FF% Mod 1000) + 100& * (NROO& Mod 100)
20:         If CantRegistros("ENTRECOM", "IDOCUM=" & IDOCUM&) > 0 Then
                IDOCUM& = IDOCUM& + 1
                GoTo 20
            End If
            
            'GUARDA ARCHIVO ENTRECOM
            SQLX$ = "SELECT * FROM ENTRECOM WHERE IDOCUM < 1"
            Dim RST1 As ADODB.Recordset
            Set RST1 = New ADODB.Recordset
26:         RST1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            If Err Then
               On Error GoTo 0
               Call CapturaErrorOpen
               GoTo 26
            End If
            On Error GoTo 0
            '
            With RST1
                .AddNew
                !FECHA = CVDAT(FF%)
                !DOCUMENTO = DOPRICO$
                !TIPODOC = "Ajuste de Stock"
                !nume_cli = 0
                !IDOCUM = IDOCUM&
                !cod_inte = CIII%
                !Cod_Exte = CODEXT$(CIII%)
                !Cantidad = CANTI
                !Status = 1
                !VECESENVIADO = 0
                .Update
            End With
            '
            On Error Resume Next
            RST1.Close
            Set RST1 = Nothing
            On Error GoTo 0
        End If
        
        '
237     If TRIM$(FORIPRE$) <> "" Then
          If YAFORI% > 0 Then
             'VALIDACION DE MODO DE IMPRESION
             MODOIMPRE07.Check7.Caption = "Imprime Ajuste de Stocks"
             MODOIMPRE07.Label1.Caption = "" 'LIMPIA EL LABEL DONDE PONE OK, SI ACEPTA
             MODOIMPRE07.Check8.Value = 0
             Screen.MousePointer = 1
             MODOIMPRE07.Show 1
             CANCELPRINT% = 0
             If MODOIMPRE07.Label1.Caption = "OK" Then
               If MODOIMPRE07.Check8.Value > 0 Then CANCELPRINT% = 1
             End If
             Screen.MousePointer = 11
             '
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
               Call PRINTFORWIN(Unimed$(CIII%), "UNIMED", 0, 0)
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
'        aqui reculote para ahp (generar otra rutina por que esta esta anulada con exit sub
'        Call RECULOTE(CIII%)
        Call COMUNI("Ajuste Número" + Str$(NRO&) + "\Completo y Verificado")
        '
299 End If
    '
    Call CIERRARCH("*.*")
    '
End Sub
Sub MOVISTOCK1(FECHA%, CI%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPO%, CANTI)
   '
   Static CTX%, REBLAN$, RELIBR$, UREMOVI&, DECONSUM$
   '
''   If CANTI < 0 Then
''     If DECONSUM$ = "" Then
''       DECONSUM$ = "MANUA"
''       If CONTROL$("", "DECONSUM") = "AUTOM" Then
''         DECONSUM$ = "AUTOM"
''       End If
''     End If
''   End If
   '
   SMP = Screen.MousePointer
   UREMAS& = ULTREG&("MASTER")
   '
'''   If ULTREG&("AJUPEND") < UREMOVI& Then
'''      Call CIERRARCH("*.*")
'''      Call MENSERR(24, "Imposible Registrar este Movimiento.\  \Cierre la Aplicación y Vuelva a Intentarlo.")
'''      Call FINAL("")
'''   End If
   '
   User% = USNBR% Mod 100
   If User% < 1 Then User% = 1
   Call CARDEPOS
   '
'''   If CTX% < 1 Then
'''      REBLAN$ = REGBLAN$("AJUPEND")
'''      RELIBR$ = String$(128, 45)
'''      Call COPYSTRU("AJUPEND", "BKMOVIST")
'''      CTX% = 1
'''   End If
   '
   ' ********** ARMAR REGISTRO DE MOVIMIENTO **********
9  DPRX$ = DOPRI$
10 PPXX% = InStr(DPRX$, "-0")
   If PPXX% > 0 Then
      DPRX$ = Left$(DPRX$, PPXX%) + Mid$(DPRX$, PPXX% + 2)
      GoTo 10
   End If
   DOPRIX$ = AJUSTI$(DPRX$, 10)
   '
   DPRX$ = DOSEC$
20 PPXX% = InStr(DPRX$, "-0")
   If PPXX% > 0 Then
      DPRX$ = Left$(DPRX$, PPXX%) + Mid$(DPRX$, PPXX% + 2)
      GoTo 20
   End If
   DOSECX$ = AJUSTI$(DPRX$, 10)
   '
   NCLIE% = NC%: NPROV% = 0
   If NC% < 0 Then
      NCLIE% = 0: NPROV% = Abs(NC%)
   End If
   '
   Entrada# = CDbl(CANTI): SALIDA# = 0
   If CANTI < 0 Then
      Entrada# = 0: SALIDA# = CDbl(Abs(CANTI))
   End If
   CANACU = SALDETOT + CANTI
   '
   depomovi% = DEPO%
'''   If DECONSUM$ = "AUTOM" Then
'''     If CANTI < 0 Then
'''       DEPII% = DEPO%
'''       DEPOMOVI% = DEPOSIBLE%(CI%, CANTI, DEPII%)
'''     End If
'''   End If
   '
   RECORD$ = REBLAN$
   Call REPLA(RECORD$, MKI$(FECHA%), 1, 2)
   Call REPLA(RECORD$, MKI$(CI%), 3, 2)
   Call REPLA(RECORD$, LOTE$, 5, 16)
   Call REPLA(RECORD$, CMOV$, 21, 2)
   If CMOV$ = "*R" Then
        Call REPLA(RECORD$, "RP", 21, 2)      ' reparación
      ElseIf CMOV$ = "*D" Then
        Call REPLA(RECORD$, "DR", 21, 2)      ' DESPACHO reparación
   End If
   Call REPLA(RECORD$, DOPRIX$, 23, 10)
   Call REPLA(RECORD$, DOSECX$, 33, 10)
   Call REPLA(RECORD$, REFE$, 43, 30)
   Call REPLA(RECORD$, MKD$(VUNI#), 73, 8)
   Call REPLA(RECORD$, MONE$, 81, 2)
   Call REPLA(RECORD$, Chr$(depomovi%), 83, 1)
   Call REPLA(RECORD$, MKI$(NCLIE%), 84, 2)
   Call REPLA(RECORD$, MKI$(NPROV%), 86, 2)
   Call REPLA(RECORD$, MKD$(Entrada#), 88, 8)
   Call REPLA(RECORD$, MKD$(SALIDA#), 96, 8)
   Call REPLA(RECORD$, String$(8, 0), 104, 8)
   Call REPLA(RECORD$, String$(6, 0), 112, 6)
   Call REPLA(RECORD$, MKS$(CANACU), 118, 4)
   '
'''   If UCase$(CMOV$) = "PP" Or CMOV$ = "DM" Then
'''     CAVERI$ = DOSECX$
'''     CAVERJ$ = MKS$(Val(Left$(CAVERI$, 6))) + MKI$(Val(Mid$(CAVERI$, 8, 3)))
'''     Call REPLA(RECORD$, CAVERJ$, 112, 6)
'''   End If
   '
   HO% = HOY(HOS$)
   HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
   '
   Call REPLA(RECORD$, MKI$(HO%), 122, 2)
   Call REPLA(RECORD$, MKI$(Int(HP)), 124, 2)
   Call REPLA(RECORD$, Chr$(User%), 126, 1)
   Call REPLA(RECORD$, String$(2, 0), 127, 2)
   '
   Dim BUGRA As String * 140
'''   If Left$(CMOV$, 1) <> "*" Then
   Call REGAPP("AJUPEND", RECORD$)
   REMOVIST& = ULTREG&("AJUPEND")
   If REMOVIST& > UREMOVI& Then UREMOVI& = REMOVIST&
   '
   NKX% = FreeFile
   Open LUCOM$ + "\AJUPEND.LOC" For Random Access Read Write As #NKX% Len = 140
   RECONO& = 1 + LOF(NKX%) / 140
   BUGRA$ = "AJUPEND " + MKS$(REMOVIST&) + RECORD$
   Put #NKX%, RECONO&, BUGRA$
   Close #NKX%
   '
   If CI% > 0 Then FLAGMOVI%(CI%) = 1
         '
'''     Else               ' reparaciones
'''         Call COPYSTRU("AJUPEND", "MOVIREPA")
'''         Call REGAPP("MOVIREPA", RECORD$)
'''         REMOVIRE& = ULTREG&("MOVIREPA")
'''         Call CIERRARCH("MOVIREPA")
'''         '
'''         NKX% = FreeFile
'''         Open LUCOM$ + "\AJUPEND.LOC" For Random Access Read Write As #NKX% Len = 140
'''         RECONO& = 1 + LOF(NKX%) / 140
'''         BUGRA$ = "MOVIREPA" + MKS$(REMOVIRE&) + RECORD$
'''         Put #NKX%, RECONO&, BUGRA$
'''         Close #NKX%
'''         '
'''   End If
   '
   Screen.MousePointer = SMP
   '
End Sub

'
Sub AJUACERO(CI%, LOTE$, DEP%, CANAJU, Optional NUCOMP$)
    '
    CADECI% = XVALO(CNTDC$)
    Static NUCOMPA$, NUORIT%
    If NUCOMP$ <> NUCOMPA$ Then NUORIT% = 0
    '
    FEAJUS% = HOY(HO$)
    DOSECO$ = "AJ." + Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Left$(HOS$, 2)
    DOSELA$ = DOSECO$
    REFER$ = "AJ.LT.CERO " + LOTE$
    NUORIT% = NUORIT% + 1
    VUNI# = 0
    MONEII% = 0
    NC% = 0
    '
    CANTTII = CANAJU
    If Abs(CANTTII) >= FRACMIN(CI%) Then
       Call MOVISTOK(FEAJUS%, CI%, LOTE$, "AJ", NUCOMP$, NUCOMP$, DOSECO$, DOSELA$, REFER$, NUORIT%, VUNI#, MONEII%, NC%, DEP%, (-1) * CANTTII)
       Call ACSALOTE(CI%, LOTE$, "NOMESS")
    End If
    '
End Sub

Sub ANUAJUACERO()
    '
    Condi1$ = "CODIMOVI = 'AJ' AND (REFERCOR = 'AJUSTE DE LOTE A CERO' OR REFERCOR LIKE 'AJ.LT.CERO%')"
    Call CARGALISEL("ANUAJUCE", "MOVISTO", "DOPRICOR/A10;FECHMOV/D8;REFERCOR/A24", Condi1$, "DISTINCT")
    Call SELECHO("ANUAJUCE/Indice de Ajustes a Cero por Lote")
    If TRIM$(VALACT1$("ANUAJUCE")) = "" Then Exit Sub
    '
    Screen.MousePointer = 11
    AJUSTNUM$ = VALACT1$("ANUAJUCE")
    CONDI2$ = Condi1$ & " AND DOPRICOR = '" & AJUSTNUM$ & "'"
    FIN& = CantRegistros&("MOVISTO", CONDI2$)
    '
    SQLX$ = "SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ & " WHERE " & CONDI2$
    Dim ANUAJCERTMP As ADODB.Recordset
    Set ANUAJCERTMP = New ADODB.Recordset
    '
25  On Error Resume Next
    ANUAJCERTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
    '
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    '
    U& = 0
    With ANUAJCERTMP
      Do While Not .EOF
         U& = U& + 1
         Call TRACE(20, U&, FIN&)
         CI% = XVALO(!cod_inte)
         LOTTEE$ = SAFETEXT$(!idenlote)
         !CANTINGRE = 0
         !CANTSALID = 0
         .Update
         Call ACSALOTE(CI%, LOTTEE$, "NOMESS")
         .MoveNext
      Loop
    End With
    ANUAJCERTMP.Close
    Set ANUAJCERTMP = Nothing
    
    Call BORRAREGISTROS("MOVISTO", CONDI2$, REG&, "NOMESS")
    '
    Screen.MousePointer = 1
    Call MENSERR(24, "Proceso de Anulacion de Ajuste Terminado")
End Sub
    
Private Sub Text9_LostFocus()
   Text9 = FORMATNUM(XVALO(Text9), "F13." & CNTDC$(CODINT(Text1)))
End Sub
Sub CORRIGESALDOLOTE(CI%)
    '
    'CORRECCION DE SALDSO DE LOTES
    'SI HAY MAS CANTIDAD SALIDA QUE INGRESADA GENERA UN MOVIMIEINTO DE ENTRADA
    'PARA EL  LOTE EN DEPOSITO
    '
    FECHA% = HOY(HO$)
    '
    'CONSULTA SOBRE EL MOVISTO AGRUPADA POR CODIGO,LOTE,DEPOSITO
    Call ABRESTOCKS
    SQLX$ = "SELECT COD_INTE,IDENLOTE, DEPOMOVI,"
    SQLX$ = SQLX$ + " SUM(CANTINGRE) AS INGRESO, SUM(CANTSALID)AS SALIDA "
    SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK) WHERE COD_INTE = " & CI%
    SQLX$ = SQLX$ + " GROUP BY COD_INTE, IDENLOTE, DEPOMOVI"
    '
    Dim CONTRLTMP As ADODB.Recordset
    Set CONTRLTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With CONTRLTMP
       Do While Not .EOF
         CANSALIDA = XVALO(!SALIDA)
         CANINGRESO = XVALO(!ingreso)
         If CANSALIDA > CANINGRESO Then
            '
            'PROCESO DE CORRECCION DE SALDOS
            AJUSTA% = 1
            Screen.MousePointer = 11
            DEPO% = XVALO(!depomovi)
            CANTI = CANSALIDA - CANINGRESO
            LTX$ = SAFETEXT$(!idenlote)
            '
            NVM$ = TRIM$(ProNroComprobante("AJ"))
            While Len(NVM$) < 6: NVM$ = "0" + NVM$: Wend
            '
            CMOV$ = TRIM$(Control$("", "CODAJUST"))
            If CMOV$ = "" Then CMOV$ = "AJ"
            DOSECO$ = "AJ." + Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Left$(HOS$, 2)
            DOPRICO$ = CMOV$ + "." + NVM$
            NUORIT% = 0
            VUNI# = 0
            MONEII% = 0
            NC% = 0
            REFE$ = " Aj. Aut. Lote: " & LTX$ & " US: " & USERNAME$
            '
            Call MOVISTOK(FECHA%, CI%, LTX$, CMOV$, DOPRICO$, DOPRICO$, DOSECO$, DOSECO$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPO%, CANTI)
            '
         End If
       .MoveNext
       Loop
    
    End With
    '
    Screen.MousePointer = 1
    Call COMUNI("Ajustado")
    
End Sub

