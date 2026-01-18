VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CONSUMLOT07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Consumos por Lote"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   12750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Ajuste a 0"
      Height          =   255
      Left            =   360
      TabIndex        =   55
      ToolTipText     =   "Ajusta a 0 el Lote Activo (Si Tiene Ordenes de Fab. Abiertas, Alerta y No se Realiza el Ajuste)"
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "ARTÍCULO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2025
      Left            =   120
      TabIndex        =   33
      Top             =   105
      Width           =   5040
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
         Height          =   300
         Left            =   495
         TabIndex        =   51
         Text            =   " "
         Top             =   960
         Width           =   540
      End
      Begin VB.CommandButton BOTSEL 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1020
         Picture         =   "CONSUMLOT07.frx":0000
         TabIndex        =   50
         Top             =   975
         Width           =   175
      End
      Begin VB.Label Label32 
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
         Left            =   180
         TabIndex        =   53
         Top             =   1680
         Width           =   4725
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dep"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   0
         TabIndex        =   52
         Top             =   1050
         Width           =   420
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   360
         TabIndex        =   39
         Top             =   600
         Width           =   1080
      End
      Begin VB.Label Label6 
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
         Height          =   975
         Left            =   1320
         TabIndex        =   38
         Top             =   600
         Width           =   3585
      End
      Begin VB.Label Label5 
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
         Left            =   4200
         TabIndex        =   37
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label4 
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
         Left            =   1320
         TabIndex        =   36
         Top             =   240
         Width           =   2370
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
         Left            =   3720
         TabIndex        =   35
         Top             =   315
         Width           =   375
      End
      Begin VB.Label Label3 
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
         Height          =   330
         Left            =   360
         TabIndex        =   34
         Top             =   315
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1470
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   31
      Top             =   525
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   8
      Left            =   11330
      TabIndex        =   22
      Top             =   2130
      Width           =   15
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   105
      MultiSelect     =   2  'Extended
      TabIndex        =   21
      Top             =   5580
      Visible         =   0   'False
      Width           =   10455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   4
      Left            =   7455
      TabIndex        =   14
      Top             =   2130
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   5
      Left            =   7900
      TabIndex        =   13
      Top             =   2130
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   6
      Left            =   9070
      TabIndex        =   12
      Top             =   2130
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   7
      Left            =   10365
      TabIndex        =   11
      Top             =   2010
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   3
      Left            =   4500
      TabIndex        =   10
      Top             =   2130
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   2
      Left            =   3045
      TabIndex        =   9
      Top             =   2130
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   1
      Left            =   1525
      TabIndex        =   8
      Top             =   2130
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3050
      Index           =   0
      Left            =   1080
      TabIndex        =   7
      Top             =   2130
      Width           =   15
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00CDDADA&
      Caption         =   "IDENTIFICADOR DE PARTIDA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2025
      Left            =   6360
      TabIndex        =   4
      Top             =   105
      Width           =   5475
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   1200
         TabIndex        =   54
         Top             =   1680
         Width           =   1590
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   3750
         TabIndex        =   30
         Top             =   1365
         Width           =   1590
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Destino"
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
         Left            =   2700
         TabIndex        =   29
         Top             =   1455
         Width           =   960
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Left            =   3405
         TabIndex        =   28
         Top             =   735
         Width           =   1035
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   4485
         TabIndex        =   27
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicación"
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
         Left            =   2700
         TabIndex        =   26
         Top             =   1110
         Width           =   960
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   3750
         TabIndex        =   25
         Top             =   1020
         Width           =   1590
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lote Original"
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
         Left            =   -45
         TabIndex        =   24
         Top             =   735
         Width           =   1110
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   1200
         TabIndex        =   23
         Top             =   660
         Width           =   1590
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   1200
         TabIndex        =   20
         ToolTipText     =   "Doble Click para Ajustar Saldos por Lote Según Ingresos y Consumos"
         Top             =   1365
         Width           =   1590
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo"
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
         Left            =   315
         TabIndex        =   19
         Top             =   1440
         Width           =   750
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   1200
         TabIndex        =   18
         Top             =   1020
         Width           =   1590
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   4125
         TabIndex        =   17
         Top             =   315
         Width           =   1200
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Alta"
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
         Left            =   2730
         TabIndex        =   16
         Top             =   405
         Width           =   1350
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   1200
         TabIndex        =   15
         Top             =   315
         Width           =   1590
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Identificador"
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
         Left            =   -90
         TabIndex        =   6
         Top             =   405
         Width           =   1155
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
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
         Left            =   -45
         TabIndex        =   5
         Top             =   1110
         Width           =   1110
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2724
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      ToolTipText     =   "Boton Derecho Sobre Ajustes  Elimina el Mismo)"
      Top             =   2445
      Width           =   12500
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      FillColor       =   &H00400000&
      ForeColor       =   &H80000008&
      Height          =   1870
      Left            =   11940
      ScaleHeight     =   1845
      ScaleWidth      =   2745
      TabIndex        =   0
      Top             =   0
      Width           =   2775
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
         Height          =   735
         Left            =   75
         Picture         =   "CONSUMLOT07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Terminar - Salir de Consulta de Stocks"
         Top             =   210
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   75
         Picture         =   "CONSUMLOT07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Re-Impresión de Etiqueta"
         Top             =   960
         Width           =   650
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "CONSUMLOT07.frx":0ABE
      Top             =   360
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   2280
      OleObjectBlob   =   "CONSUMLOT07.frx":0CF2
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label34 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Kg."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11400
      TabIndex        =   56
      Top             =   2205
      Width           =   975
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10320
      TabIndex        =   49
      Top             =   2205
      Width           =   975
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Salida"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   48
      Top             =   2205
      Width           =   975
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Entrada"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   47
      Top             =   2205
      Width           =   975
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Dep"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7530
      TabIndex        =   46
      Top             =   2205
      Width           =   975
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Referencia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   45
      Top             =   2205
      Width           =   975
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   44
      Top             =   2205
      Width           =   975
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1575
      TabIndex        =   43
      Top             =   2205
      Width           =   975
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "CM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1185
      TabIndex        =   42
      Top             =   2205
      Width           =   975
   End
   Begin VB.Label Label76 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   165
      TabIndex        =   41
      Top             =   2205
      Width           =   975
   End
   Begin VB.Label Label37 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   120
      TabIndex        =   40
      Top             =   2160
      Width           =   12500
   End
End
Attribute VB_Name = "CONSUMLOT07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTSEL_Click()
        Call SELECHO("TADEPOSI")
        If VALACT1$("TADEPOSI") <> "" Then
          Text4 = TRIM$(VALACT1$("TADEPOSI"))
        End If

End Sub

Private Sub Command1_Click()
   Text4 = "" 'limpia casillero de deposito

   Hide
End Sub

Sub Command2_Click()
  If DERACCE%("AJUCERO", "Ajuste de Stocks a Cero") < 2 Then
     GoTo 99
   End If
   Command2.Enabled = False
   CIXI% = CODINT%(Label4)
   LOTEX$ = TRIM$(Label1)
   DEPO% = XVALO(Text4)
   If DepositoValido(DEPO%) Then
    If CIXI% > 0 Then
      If LOTEX$ <> "" Then
         If DEPO% > 0 Then
            Call AjustarLoteA0PorDeposito(CIXI%, LOTEX$, DEPO%, 1)
         End If
      End If
    End If
   Else
    Call COMUNI("Falta Depòsito o No es Vàlido")
   End If
99 Command2.Enabled = True
End Sub
Sub AjustarLoteA0PorDeposito(CIXI%, LOTEX$, DEPO%, Optional NoMostrarMensaje%)

        If sCodMovimiento$ = "" Then sCodMovimiento$ = "AJ"
        NRO& = ProNroComprobante(sCodMovimiento$)
        sDocSecundarioCorto$ = "AJ." + Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Left$(HOS$, 2)
        sDocSecundarioLargo$ = sDocSecundarioCorto$
        sDocPrimarioCorto$ = "AS." + TRIM$(str$(NRO&))
        sDocPrimarioLargo$ = sDocPrimarioCorto$
        If ReferenciaMovimiento = "" Then ReferenciaMovimiento = "Inventario de Stocks a Cero"

        sFiltroDeposito$ = ""
        If DEPO% > 0 Then sFiltroDeposito$ = " AND MOVISTO.DepoMovi = " & CStr(DEPO%) & " "
        sFiltroStatus$ = ""
        OF1_ANT$ = ""
        'SI ESTE LOTE POSEE ORDENES DE FABRICACION ABIERTAS NO SE PERMITE
        For UY& = 0 To List1.ListCount - 1
           Linea$ = List1.List(UY&)
           OF1$ = TRIM$(UCase$(Mid$(Linea$, 29, 12)))
           If OF1$ <> "" And OF1$ <> OF1_ANT$ Then
              OF2$ = "O" & Mid$(OF1$, 2)
              If STATORFA%(OF2$) < 3 Then
                 Call COMUNI("Atenciòn !! La Orden de Fabricaciòn: " & OF2$ & " Se Encuentra Pendiente de Cierre\Realice Cierre por Ajuste Considerando Cantidades Pendientes a Consumir.")
                 Exit Sub
              End If
           End If
           OF1_ANT$ = OF1$
        Next UY&

        Dim RstAjusteNoTrazables As ADODB.Recordset
        ConsultaSql$ = "SELECT SUM(CANTINGRE) AS INGRESOS, SUM(CANTSALID) AS SALIDAS, MASTER.CODINT, MOVISTO.COD_INTE, MOVISTO.DEPOMOVI "
        ConsultaSql$ = ConsultaSql$ & " FROM MASTER INNER JOIN MOVISTO WITH(NOLOCK) ON MASTER.CODINT = MOVISTO.COD_INTE WHERE MASTER.TRAZABLE = 1 "
        'ConsultaSql$ = ConsultaSql$ & " AND MOVISTO.FECHMOV <= '" & FETEXCOR1$(iFechaCorte%) & "'"      ' FECHA DE CORTE
        ConsultaSql$ = ConsultaSql$ & " AND MOVISTO.COD_INTE = " & CIXI%
        ConsultaSql$ = ConsultaSql$ & " AND IDENLOTE = '" & LOTEX$ & "'"
        ConsultaSql$ = ConsultaSql$ & sFiltroDeposito$                                                  ' DEPOSITO
        ConsultaSql$ = ConsultaSql$ & " GROUP BY MOVISTO.COD_INTE, MOVISTO.DEPOMOVI, MASTER.CODINT "
        ConsultaSql$ = ConsultaSql$ & " ORDER BY MOVISTO.COD_INTE, MOVISTO.DEPOMOVI"
        
        Set RstAjusteNoTrazables = READSET(ConsultaSql$)
        
        
        iNroOrden% = 0
        With RstAjusteNoTrazables
            Do While Not .EOF
                iCodigo% = XVALO(!CODINT)
                iDeposito% = XVALO(!depomovi)
                
                dIngresos# = XVALO(!INGRESOS)
                dSalidas# = XVALO(!SALIDAS)
                dSaldo# = dIngresos# - dSalidas#
                dCantidadAjuste# = (-1) * dSaldo#   ' AJUSTA POR EL VALOR CONTRARIO
                
                METROX# = SaldoMetrosEnDepoXLote(CIXI%, DEPO%, LOTEX$)
                PESOKG# = SaldoKilosEnDepoXLote(CIXI%, DEPO%, LOTEX$)
                If Abs(dCantidadAjuste) >= 0.005 Then
                    iNroOrden% = iNroOrden% + 1
                    dValorUnitario# = 0
                    sReferencia$ = "Ajuste Lote a Cero"
                    iMoneda% = 0
                    iCliente% = 0
                    SLOTE$ = LOTEX$
                    Call MOVISTOK(HOY(HO$), iCodigo%, SLOTE$, sCodMovimiento$, sDocPrimarioLargo$, sDocPrimarioCorto$, sDocSecundarioLargo$, sDocSecundarioCorto$, sReferencia$, iNroOrden%, dValorUnitario#, iMoneda%, iCliente%, iDeposito%, dCantidadAjuste#, , , , , , , , , , , METROX#, PESOKG#)
                End If
                
                .MoveNext
            Loop
        End With
        
        On Error Resume Next
        
        RstAjusteNoTrazables.Close
        Set RstAjusteNoTrazables = Nothing
        On Error GoTo 0
        If iNroOrden% > 0 Then
           Screen.MousePointer = 1
           If NoMostrarMensaje% < 1 Then Call COMUNI("Ajuste Realizado")
        End If
End Sub
Private Sub Command8_Click()
   '
   Command8.Enabled = False
   CIXI% = CODINT%(Label4)
   LOTEX$ = TRIM$(Label1)
   
   If CIXI% > 0 Then
      If LOTEX$ <> "" Then
        Call PRINSTIKAPRO1(CIXI%, LOTEX$)
      End If
   End If
99 Command8.Enabled = True
   '
End Sub

Private Sub Form_Activate()
   '
   Call label1_Change

End Sub


Private Sub Form_Load()
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    Label33.Caption = ""
    Label33.Visible = False
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "AHP") > 0 Or InStr(EPAC$, "CAMBACERES") > 0 Then Command2.Visible = True
  

End Sub

Private Sub label1_Change()
   Label8 = ""
   Label9 = ""
   Label19 = ""
   Label13 = ""
   Label14 = ""
   Label17 = ""
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   ESENVAPLAST% = 0
   If InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "CAMBACERES") > 0 Then ESENVAPLAST% = 1

   '
   List1.Clear: List2.Clear
   '
   DEPO% = XVALO(Text4)
   IDLOTAB$ = TRIM$(Label1)
   CIXI% = CODINT(Label4)
   '
   If IDLOTAB$ <> "" And CIXI% > 0 Then
     'Call ACSALOTE(CIXI%, IDLOTAB$)
     '
     List1.Clear: List2.Clear
     '
     'sql
     Call ABRESTOCKS
     SQLX$ = " SELECT FECHALTA, CanAlta, CanSaldo   FROM LOTINGRE WITH(NOLOCK)"
     SQLX$ = SQLX$ + "WHERE IdenLote = '" & IDLOTAB$ & "' "
     SQLX$ = SQLX$ + "AND Cod_Inte = " & CIXI% & " "
     
     Set LOTINGTEMP = READSET(SQLX$)
     'Set LotIngTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
     '
     With LOTINGTEMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         FEXALTA = CVINT%(!FechAlta)
         CANORIG = !CanAlta
         CANSALDOX = !CanSaldo
         '
         Label8 = FECHATEX$(FEXALTA)
'         Label9 = FORMATNUM(CANORIG, "F10,1")
         Label9 = FORMATNUM$(CANORIG, "F14." & CNTDC$(CIXI%) & "")
         'Label13 = FORMATNUM(CANSALDOX, "F10,1")
         Label13 = FORMATNUM$(CANSALDOX, "F14." & CNTDC$(CIXI%) & "")
         
         SARRAS# = CANORIG
         '
         Call ABRESTOCKS
         SQLX$ = " SELECT CANTINGRE, CANTSALID,FECHMOV, CodiMovi, DoPriCor, DoSecCor, DoSecCor, ReferCor, DEPOMOVI  "
         If ESENVAPLAST% > 0 Then SQLX$ = SQLX$ + ",PESOKG ,MOVISTO_ID " 'APARTE POR SI NO ESTA CREADO EN OTRA EMPRESA
         SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK)"
         SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
         SQLX$ = SQLX$ + "AND IdenLote = '" & IDLOTAB$ & "' "
         If DEPO% > 0 Then    'filtro por deposito
            SQLX$ = SQLX$ + "AND DEPOMOVI = " & XVALO(Text4)
         End If
         If TRIM$(Label33) <> "" Then
           If Label33.Visible = True Then
             SQLX$ = SQLX$ + " AND DoSecCor = '" & Label33.Caption & "'"
           End If
         End If
         'SQLX$ = SQLX$ + "ORDER BY FechMov ASC"
         'SQLX$ = SQLX$ + "ORDER BY FechMov, FeReMovi ASC "
         If EXISTE_CAMPO%("MOVISTO", "MOVISTO_ID") > 0 Then 'POR QUE FEREMOVI NO GUARDA LOS SEGUNDOS
            SQLX$ = SQLX$ + " ORDER BY FeReMovi, MOVISTO_ID"
         Else
            SQLX$ = SQLX$ + " ORDER BY FechMov , FeReMovi ASC "
         End If

         '
'         Dim MoviTemp As ADODB.Recordset
'         Set MoviTemp = New ADODB.Recordset
'         MoviTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
         Set MoviTemp = READSET(SQLX$)
         With MoviTemp
           On Error Resume Next
           .MoveFirst
           If Err < 1 Then
              Do While Not .EOF
                   CANIN# = !CANTINGRE
                   CACONS# = !CANTSALID
                   TTXX$ = Space$(256)
                   FEX = CVINT(!FECHMOV)
                   KILOS# = 0
                   If ESENVAPLAST% > 0 Then
                    KILOS# = XVALO(!PESOKG): If CACONS# > CANIN# Then KILOS# = -1 * KILOS#
                    If SAFETEXT$(!CODIMOVI) = "BR" Then KILOS# = PesoOriginalSegúnLote#(IDLOTAB$) 'NO VALIDO + O - POR QUE SIEMPRE SERA +
                   End If
                     Call REPLA(TTXX$, FECHATEX$(FEX), 1, 8)
                     Call REPLA(TTXX$, !CODIMOVI, 11, 2)
                     Call REPLA(TTXX$, !DoPriCor, 15, 12)
                     Call REPLA(TTXX$, !DoSecCor, 29, 12)
                     Call REPLA(TTXX$, !ReferCor, 43, 26)
                   DEPOM$ = FORMATNUM$(!depomovi, "I3")
                     Call REPLA(TTXX$, DEPOM$, 70, 3)
                     'Call REPLA(TTXX$, FORMATNUM$(CANIN#, "F10.1"), 75, 10)
                     Call REPLA(TTXX$, FORMATNUM$(CANIN#, "F10." & CNTDC$(CIXI%) & ""), 75, 10)
                     'Call REPLA(TTXX$, FORMATNUM$(CACONS#, "F10.1"), 86, 10)
                     Call REPLA(TTXX$, FORMATNUM$(CACONS#, "F10." & CNTDC$(CIXI%) & ""), 86, 10)
                     
                     If ESENVAPLAST% > 0 Then
                        Call REPLA(TTXX$, FORMATNUM$(KILOS#, "F10." & CNTDC$(CIXI%) & ""), 107, 10)
                        Call REPLA(TTXX$, XVALO(!MOVISTO_ID), 120, 12)
                     End If
                     
                   For KKU& = List2.ListCount To 1 Step -1
                     If FECHANUM(Left$(List2.List(KKU& - 1), 8)) <= FEX Then
                       List2.AddItem TTXX$, KKU&
                       GoTo 39
                     End If
                   Next KKU&
                   List2.AddItem TTXX$, 0
39            .MoveNext
              Loop
           End If
           On Error GoTo 0
         End With
         MoviTemp.Close
         Set MoviTemp = Nothing
         '
         ' RESERVAS
         Call APERBASDAT("RESERVA")
         SQLZ$ = " SELECT CANTRES, CANTCONS, IDENLOTE, FechRes, IDSUBOR FROM RESERVA WITH(NOLOCK)"
         SQLZ$ = SQLZ$ + "WHERE Cod_Inte = " & CIXI% & " "
         SQLZ$ = SQLZ$ + "AND Status = 0 "     ' para no considerar los marcados para anular
         SQLZ$ = SQLZ$ + "ORDER BY FechRes ASC"
         'Set ReserTemp = Stocks.OpenRecordset(SQLZ$, dbOpenSnapshot)
         '
'         Dim ReserTemp As ADODB.Recordset
         Set ReserTemp = New ADODB.Recordset
'         ReserTemp.Open FILTSQL$(SQLZ$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
         Set ReserTemp = READSET(SQLZ$)
         With ReserTemp
           On Error Resume Next
           .MoveFirst
           If Err < 1 Then
              On Error GoTo 0
              Do While Not .EOF
                CSALI# = !CANTRES - !CANTCONS
                If CSALI# >= 0.05 Then
                  LCONSU$ = ""
                  On Error Resume Next
                  LCONSU$ = !idenlote
                  On Error GoTo 0
                  If LCONSU$ = IDLOTAB$ Then
                    CACONS# = ROUND(CSALI#, 1)
                    TTXX$ = Space$(256)
                    FEX = CVINT(!FechRes)
                      Call REPLA(TTXX$, FECHATEX$(FEX), 1, 8)
                      Call REPLA(TTXX$, "RF", 11, 2)
                      Call REPLA(TTXX$, !idsubor, 15, 12)
                      Call REPLA(TTXX$, "", 29, 12)
                      Call REPLA(TTXX$, "RESERVA FABRICACION", 43, 26)
                    DEPOM$ = FORMATNUM$(0, "I3")
                      Call REPLA(TTXX$, DEPOM$, 70, 3)
                      'Call REPLA(TTXX$, FORMATNUM$(CACONS#, "F10.1"), 86, 10)
                      Call REPLA(TTXX$, FORMATNUM$(CACONS#, "F10." & CNTDC$(CIXI%) & ""), 86, 10)
                    For KKU& = List2.ListCount To 1 Step -1
                      If FECHANUM(Left$(List2.List(KKU& - 1), 8)) <= FEX Then
                        List2.AddItem TTXX$, KKU&
                        GoTo 49
                      End If
                    Next KKU&
                    List2.AddItem TTXX$, 0
                  End If
                End If
49            .MoveNext
              Loop
           End If
         End With
         '
         ' BALAREQMO
         Call APERBASDAT("STOCKS")
         SQLZ$ = " SELECT Kilo_Asig, Kilo_Consu, FECHMOV, DESTINO, IdenReque FROM BALAREQMO WITH(NOLOCK) "
         SQLZ$ = SQLZ$ + "WHERE Cod_Inte = " & CIXI% & " "
         SQLZ$ = SQLZ$ + " AND IdenLote = '" & IDLOTAB$ & "'"
         SQLZ$ = SQLZ$ + " AND Kilo_Asig <> 0"
         SQLZ$ = SQLZ$ + "ORDER BY FechMov ASC"
         '
         Set ReserTemp = New ADODB.Recordset
'         ReserTemp.Open FILTSQL$(SQLZ$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
         Set ReserTemp = READSET(SQLZ$)
         With ReserTemp
           On Error Resume Next
           .MoveFirst
           If Err < 1 Then
              On Error GoTo 0
              Do While Not .EOF
                 CSALI# = XVALO(!Kilo_Asig) - XVALO(!Kilo_Consu)
                 CACONS# = ROUND(CSALI#, XVALO(CNTDC$(CIXI%)))
                 If CACONS# > 0 Then
                    TTXX$ = Space$(256)
                    FEX = CVINT(!FECHMOV)
                         Call REPLA(TTXX$, FECHATEX$(FEX), 1, 8)
                         Call REPLA(TTXX$, "RF", 11, 2)
                         Call REPLA(TTXX$, SAFETEXT$(!DESTINO), 15, 12)
                         Call REPLA(TTXX$, "", 29, 12)
                         Call REPLA(TTXX$, SAFETEXT$(!IdenReque), 43, 26)
                    DEPOM$ = FORMATNUM$(0, "I3")
                         Call REPLA(TTXX$, DEPOM$, 70, 3)
                         'Call REPLA(TTXX$, FORMATNUM$(CACONS#, "F10.1"), 86, 10)
                         Call REPLA(TTXX$, FORMATNUM$(CACONS#, "F10." & CNTDC$(CIXI%) & ""), 86, 10)
                    For KKU& = List2.ListCount To 1 Step -1
                       If FECHANUM(Left$(List2.List(KKU& - 1), 8)) <= FEX Then
                         List2.AddItem TTXX$, KKU&
                         GoTo 491
                       End If
                    Next KKU&
                    List2.AddItem TTXX$, 0
                 End If
491            .MoveNext
              Loop
           End If
         End With
         '
         SARRAS# = 0
         For KKU& = 1 To List2.ListCount
            TTXX$ = List2.List(KKU& - 1)
            SARRAS# = SARRAS# + XVALO(Mid$(TTXX$, 75, 10))
            SARRAS# = SARRAS# - XVALO(Mid$(TTXX$, 86, 10))
            'Call REPLA(TTXX$, FORMATNUM$(SARRAS#, "F11.1"), 96, 11)
            Call REPLA(TTXX$, FORMATNUM$(SARRAS#, "F11." & CNTDC$(CIXI%) & ""), 96, 11)
            List1.AddItem TTXX$
         Next KKU&
         LOTINGTEMP.Close
         Set LOTINGTEMP = Nothing
         '
       End If
     End With
     ReserTemp.Close
     Set ReserTemp = Nothing
     
     '
   End If
   
   CIXXI% = XVALO(CODINT(Label4))
   CONDI$ = "COD_INTE = " & CIXXI% & " AND IDENLOTE ='" & TRIM$(Label1) & "'"
   
   Label14 = VALOBADA("LOTINGRE", "LOTECOL_PROV", CONDI$)
   Label19 = VALOBADA("LOTINGRE", "LIBERA", CONDI$)
   Label17 = VALOBADA("LOTINGRE", "UBICFIS", CONDI$)
   Label22 = VALOBADA("LOTINGRE", "Cod_Desti", CONDI$)
   '
   Screen.MousePointer = 1
   '
End Sub

Private Sub Label1_DblClick()
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   
   If InStr(EPAC$, "EYM") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "CAMBACERES") > 0 Or InStr(EPAC$, "AHP") > 0 Or InStr(EPAC$, "ACHERNAR") > 0 Or InStr(EPAC$, "CAMBACERES") > 0 Or InStr(EPAC$, "PERGAMINO") > 0 Then
        
        CIXI% = CODINT%(Label4)
        MsgBox CIXI%
        If CIXI% > 0 Then
         NORFA$ = TRIM$(Label1)
         FE% = HOY(HO$)
         Call CREALOTIN(FE%, FE%, FECHA%, CIXI%, NORFA$, "", "", CDbl(1))
         Call ACSALOTE(CIXI%, NORFA$, "NOMESS")
         Call label1_Change
        End If
   End If
   
End Sub

Private Sub Label13_DblClick()
   CIXI% = CODINT%(Label4)
   If CIXI% > 0 Then
     IDLOTX$ = TRIM$(Label1)
     If IDLOTX$ <> "" Then
       Call ACSALOTE(CIXI%, IDLOTX$)
       Call label1_Change
     End If
   End If
End Sub

Private Sub Label33_DblClick()
    Label33.Caption = ""
    Label33.Visible = False
    Call label1_Change
End Sub

Private Sub Label4_Change()
'   Label5 = ""
'   Label6 = ""
'   Label1 = ""
'   Text4 = ""
'   '
'   CIXI% = CODINT%(Label4)
'   If CIXI% > 0 Then
'     Label5 = Unimed$(CIXI%)
'     Label6 = DESCRIT0$(CIXI%)
'   End If
   '
End Sub

Private Sub List1_DblClick()
   Linea$ = List1.List(List1.ListIndex)
   DOPRI$ = Mid$(Linea$, 15, 12)
   FE% = FECHANUM(Mid$(Linea$, 1, 8))
   CODIMOVI$ = Mid$(Linea$, 11, 2)
   DOSEC$ = Mid$(Linea$, 29, 12)
   DEPO% = Mid$(Linea$, 70, 3)
   LOTE$ = Label1
   CI1% = CODINT%(Label4)
   ID& = XVALO(Mid$(Linea$, 120, 12))
   EPAC$ = UCase$(TRIM$(EMPREAC$))
   If InStr(EPAC$, "CONTEPLAST") > 0 Then
      Call EditarPesosKg(FE%, CODIMOVI$, DOPRI$, DOSEC$, DEPO%, LOTE$, CI1%, ID&)
      Exit Sub
   End If
   If InStr(EPAC$, "CAMBACERES") > 0 Then
'      Call EditarPesosKg(FE%, CODIMOVI$, DOPRI$, DOSEC$, DEPO%, LOTE$, CI1%, ID&)
     OPX% = COMALTER%("Opciones de Corrección: " & DOPRI$ & "\\Cancelar\Cambia Depósito")
     If OPX% < 2 Or OPX% > 3 Then Exit Sub
     If OPX% = 2 Then 'CAMBIO DE DEPOSITO
       If DERACCE%("CAMBIDEP", "CAMBIO DE DEPOSITOS") > 1 Then
        DEPOSITOX$ = TRIM$(InputBox(ID& & " Numero de Depósito:", "Ingrese el Depósito", str$(DEPO%)))
        If DEPOSITOX$ <> "" Then
          DEPO% = XVALO(DEPOSITOX$)
          If ECOARCH$("TADEPOSI", Chr$(DEPO%)) = "" Then
             Call COMUNI("Depósito ingresado no válido")
             Exit Sub
          End If
          If ID& > 0 Then SQLX$ = "UPDATE MOVISTO SET DEPOMOVI = " & DEPO% & " WHERE MOVISTO_ID = " & ID&: FLEXCONN.Execute (SQLX$): Call COMUNI("Cambio Realizado"): Call label1_Change 'REFERESCA
        Else
          Exit Sub
        End If
       End If
      Exit Sub
    End If
   End If
   
   CONDI$ = "DOPRICOR ='" & DOPRI$ & "' AND CODIMOVI='OF' AND CANTINGRE > 0 "
   CAUSA% = XVALO(VALOBADA("MOVISTO", "CAUSA", CONDI$))
   If CAUSA% > 0 Then
      CAUSATEX$ = TRIM$(ECOARCH$("CAUINGRE", MKI$(CAUSA%)))
      If CAUSATEX$ <> "" Then
         Call COMUNI("CAUSA: " & CAUSATEX$)
      End If
   End If
   '
   'si es BLIPACK MUESTRA LOS CONSUMOS CON GENERADOR DE REPORTES
   If InStr(EMPREAC$, "BLIPACK") > 0 Then
       OPRO$ = TRIM$(Mid$(DOPRI$, 4)) 'TOMO SOLO EL NUMERO SIN EL TEXTO OP-
       CONDI$ = "NUORPROD =  '" & OPRO$ & "'"
       NORFA$ = VALOBADA("ORDEFABR", "IDSUBOR", CONDI$, "NOMESS")
       '
       If TRIM$(NORFA$) = "" Then
          Call MENSERR(24, "Registro No Válido Para Visualizar Consumos")
          Exit Sub
       End If
       '
       If EXISTE%("GENREP07.EXE") > 0 Then
           Call FINAL("? MOVCONSU (" & NORFA$ & ")") ' ABRE LISTADO DE G.R. SEGUN LA O.F.
       End If
       '
   End If
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then
    
       If DERACCE%("AJUAUSTO", "Aumento de Stocks") < 2 Then Exit Sub
       If DERACCE%("AJUDISTO", "Disminución de Stocks") < 2 Then Exit Sub
   
        CIXI% = CODINT%(Label4)
        Linea$ = List1.List(List1.ListIndex)
        DOCU1$ = TRIM$(Mid$(Linea$, 15, 12))
        DOCU2$ = TRIM$(Mid$(Linea$, 29, 12))
        If Left$(DOCU1$, 3) = "AJ." Or Left$(DOCU1$, 3) = "AS." Then
         OPX% = COMALTER%("Desea Borrar el Ajuste: " & DOCU1$ & "\\Cancelar\Borrar")
         If OPX% > 1 Then
             CONDI$ = "DoPriCor ='" & DOCU1$ & "' AND DOSECCOR ='" & DOCU2$ & "' AND COD_INTE = " & CIXI%
             CONDI$ = CONDI$ + " AND IDENLOTE = '" & Label1 & "'"
             Call BORRAREGISTROS("MOVISTO", CONDI$, RegAf&, "NOMESS")
             If RegAf& > 0 Then
                Call COMUNI("Ajuste Eliminado con Exito")
             End If
         End If
        End If
   End If
End Sub


Private Sub Text4_Change()
   If Val(Text4) < 256 Then
     If Val(Text4) < 1 Or ECOARCH$("TADEPOSI", Chr$(Val(Text4))) = "" Then
         Label32 = "Todos los Depósitos"
         LABEL12.Visible = True
         Label13.Visible = True
       Else
         Label32 = ECOARCH$("TADEPOSI", Chr$(Val(Text4)))
         LABEL12.Visible = False
         Label13.Visible = False
     End If
   End If
   
    label1_Change
End Sub
Sub EditarPesosKg(FE%, CODIMOVI$, DOPRI$, DOSEC$, DEPO%, LOTE$, CI1%, IDMOVISTO&)
    'If DERACCE%("CAMPECTP", "MODIFICAR PESO A MOVIMIENTO") < 2 Then Exit Sub
    USER1% = USNBR% Mod 100
    If USER1% <> 1 Then
      Call COMUNI("Solo Puede Ingresar el Supervisor del Sistema")
      Exit Sub
    End If

    TOP1$ = "" 'PARA QUE TRAIGA SOLO UNO SI ES TR
    ES_TR% = 0
    If CODIMOVI$ = "TR" Then
      Call COMUNI("Se ha Seleccionado una Transferencia, el Valor Modificado Afectará a los 2 Movimientos")
      TOP1$ = " TOP 1 "
      ES_TR% = 1
    End If
    If Unimed$(CI1%) <> "M." Then
      Call COMUNI("Esta Opción Solo Admite Modificar el peso a Productos Con Unidad de Medida en Metros")
      Exit Sub
    End If
    
    Codigo$ = CODEXT$(CI1%)
    VENTANA.Inicializar = 0
    VENTANA.CampEditables = "12"
    VENTANA.AgregaRegistro = 0
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 1
    VENTANA.HabilitaInsertar = 1
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "Edición de Peso en Kgs. (Código: " & Codigo$ & " Lote: " & LOTE$ & ")"
    
    Campos$ = "FECHA/D8;" '0
    Campos$ = Campos$ + "T.Mov./A6;"      '1
    Campos$ = Campos$ + "T.Docum.1/A16;"      '2
    Campos$ = Campos$ + "T.Docum.2/A16;"   '3
    Campos$ = Campos$ + "Dep./F4;"       '4
    Campos$ = Campos$ + "Referencia/A32;"        '5
    Campos$ = Campos$ + "Ingreso/F8;"     '6
    Campos$ = Campos$ + "Salida/F8;" '7
    Campos$ = Campos$ + "Kgs./F8.1;" '8
    Campos$ = Campos$ + "K.Totales/F10.1;" '9
    Campos$ = Campos$ + "K.Teórico Fracción/F14.1;" '10
    Campos$ = Campos$ + "Kgs.Editar/F8.1;" '11
    Campos$ = Campos$ + "MovistoID/A0;" '12
    
    Erase VectorToolTipText 'EL INDICE ES LA COLUMNA DONDE SE APLICA EL TOOLTIPTEXT
    VENTANA.VentVectorToolTipText(9) = "Kgs Informados Actualmente"
    VENTANA.VentVectorToolTipText(10) = "Kgs Totales Informados en la Recepción"
    VENTANA.VentVectorToolTipText(11) = "Kgs Teóricos Calculados Segun Metros (de esta Fracción) y Peso Informado en la Recepción"
    VENTANA.VentVectorToolTipText(12) = "Kgs Valor Editable"

    VENTANA.Campox = Campos$
    VENTANA.Inicializar = 1
    Dim obj As New RECORXET  ' los campos identificados con 2 comillas simples traen columnas vacias
    CAMPOSTABLA$ = TOP1$ & "FECHMOV,CODIMOVI,DOPRILAR,DOSECLAR,DEPOMOVI,refercor,CANTINGRE,CANTSALID,PESOKG,'','',PESOKG,MOVISTO_ID"
    CONDI$ = "DOPRILAR = '" & DOPRI$ & "' AND IDENLOTE = '" & LOTE$ & "' AND MOVISTO_ID = " & IDMOVISTO&
    
    Set rs1 = obj.RS_GENERAL("MOVISTO", CAMPOSTABLA$, CONDI$)
    Call Carga_MSF_Recordset(rs1, Campos$, VENTABNEW.MSF)
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    Dim MSF1 As msFlexGrid
    Set MSF1 = VENTABNEW.MSF
    '
    'LE ASIGNO LOS VALORES A LAS COLUMNAS QUE TRAJE VACIAS
    PesoPorMetro# = ROUND(CoeficientePesoPorMetroSegunLote#(TRIM$(LOTE$)), 4)
    CantiEnMetros# = XVALO(MSF1.TextMatrix(1, 7)) 'SEGUN EL REGISTRO QUE MUESTRE PUEDE TENER CANTIDAD EN EL INGRESO O EN LA SALIDA
    If CantiEnMetros# < 1 Then                    'SI UNO NO TIENE BUSCA EN LA SIGUENTE COLUMNA DE SALIDA
      CantiEnMetros# = XVALO(MSF1.TextMatrix(1, 8))
    End If
    '
    PesoTeoricoSegunCantiEnMetros# = CantiEnMetros# * PesoPorMetro#
    MSF1.TextMatrix(1, 10) = ROUND(PesoOriginalSegúnLote#(LOTE$), 2)
    MSF1.TextMatrix(1, 11) = ROUND(PesoTeoricoSegunCantiEnMetros#, 2)
    MSF1.Rows = 2 'para que muestre solo encbezado mas una fila, si no muestra 2 filas.
30  VENTABNEW.Show 1
    If VENTABNEW.APROBADO% < 1 Then Exit Sub
    
    
    ''''AGREGADO PARA QUE SELECCIONE TAMBIEN LA OPERACION PARA QUE SE PUEDA APLICAR EL CAMBIO TAMBIEN EN EL INFOFAB
    Campos$ = "Op/F3;Desc.Oper./a30;Peso Kg./F9.2;ID/A0"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = "Confirme Seleccionando a que Operación  Aplica"

    CAMPOSTABLA$ = " NUMEOPER,DESCOPER,Peso,INFOFAB_ID" '
    TABLA$ = " INFOFAB  "
    CONDI$ = " idsubor = '" & DOPRI$ & "'"
    CONDI$ = CONDI$ + " AND CodiElem = " & CI1%
    CONDI$ = CONDI$ + " AND IDENLOTE = '" & LOTE$ & "'"
    CONDI$ = CONDI$ + " AND CANTSALID <= 0"
    REGISTROS& = CantRegistros(TABLA$, CONDI$, "NOMESS")
    If REGISTROS& < 1 Then GoTo 20
    '
    Set rs1 = obj.RS_GENERAL(TABLA$, CAMPOSTABLA$, CONDI$)
    Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.MSF2, 1)
    On Error Resume Next
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
    FRMZELECHO.Width = 12000
    FRMZELECHO.Show 1
    INFOFAB_ID = XVALO(RESP_ZELE_VECT(4))
    If INFOFAB_ID < 1 Then GoTo 99
    
    ''''''''' FIN AGREGADO PARA QUE SELECCIONE TAMBIEN LA OPERACION PARA QUE SE PUEDA APLICAR EL CAMBIO TAMBIEN EN EL INFOFAB
20  PESOEDIT# = XVALO(MSF1.TextMatrix(1, 12))
    If PESOEDIT# < 1 Then
         Call COMUNI("Atención Peso Ingresado No Válido (Debe Ser Mayor a 0&)")
         GoTo 30
    End If
    PESOEDITADO$ = FORMATNUM$(PESOEDIT#, "F15.4")
    'SI PASO VALIDACION GRABA
    'TRANSACCIÓN GRABACION
    On Error GoTo ERROR_GUARDAR
    FLEXCONN.BeginTrans
    CONDI$ = "DOPRILAR = '" & DOPRI$ & "' AND IDENLOTE = '" & LOTE$ & "'"
    If ES_TR% < 1 Then CONDI$ = CONDI$ + " AND MOVISTO_ID = " & IDMOVISTO&     'SI NO ES TR (QUE GRABA 2 MOVIMIENTOS, AGREGA CONDI DE IDMOVISTO
    SQLX$ = "UPDATE MOVISTO SET PESOKG = " & PESOEDITADO$ & " WHERE " & CONDI$ ' POR QUE P.EJ. UN CONSUMO TIENE TODOS LOS MISMOS DATOS Y GRABARIA MAS REGISTROS QUE EL QUE SE MODIFICARIA
    FLEXCONN.Execute (SQLX$)

    If ES_TR% < 1 And INFOFAB_ID > 0 Then
      CONDI$ = "INFOFAB_ID = " & INFOFAB_ID
      SQLX$ = "UPDATE INFOFAB SET PESO = " & PESOEDITADO$ & " WHERE " & CONDI$
      FLEXCONN.Execute (SQLX$)
    End If
    
    Set VENTANA = Nothing
        
ERROR_GUARDAR:
    If Err <> 0 Then
       FLEXCONN.RollbackTrans
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Grabación Realizado Exitosamente")
    End If
    Unload VENTABNEW

'FALTA COMPLETARA LOS DATOS TEORICOS
'Y VALIDAR QUE SOLO LO HAGA LAUTARO
'___________________________________________
      
99
      
End Sub
