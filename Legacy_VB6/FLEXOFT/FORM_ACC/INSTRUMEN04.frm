VERSION 5.00
Begin VB.Form INSTRUMEN04 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Instrmentos de Medición y Ensayo"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11835
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Height          =   7935
      Left            =   10900
      ScaleHeight     =   7875
      ScaleWidth      =   1155
      TabIndex        =   47
      Top             =   -120
      Width           =   1215
      Begin VB.CommandButton Command7 
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
         Left            =   120
         Picture         =   "INSTRUMEN04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Listado Inventario de Aparatos e Instrumentos de Medidion y Ensayo"
         Top             =   5100
         Width           =   650
      End
      Begin VB.CommandButton Command22 
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
         Height          =   690
         Left            =   120
         Picture         =   "INSTRUMEN04.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Imprime Ficha del Instrumento"
         Top             =   3000
         Width           =   650
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
         Height          =   690
         Left            =   120
         Picture         =   "INSTRUMEN04.frx":0AAC
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Alta de Nuevo Instrumento de Medición"
         Top             =   2310
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         Caption         =   "List"
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
         Picture         =   "INSTRUMEN04.frx":0BF6
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Listado Inventario de Aparatos e Instrumentos de Medidion y Ensayo"
         Top             =   3690
         Width           =   650
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
         Height          =   690
         Left            =   120
         Picture         =   "INSTRUMEN04.frx":0F00
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Cierra la Hoja de Procedimientos"
         Top             =   285
         Width           =   650
      End
      Begin VB.CommandButton Command4 
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
         Left            =   120
         Picture         =   "INSTRUMEN04.frx":104A
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Ayuda en Línea"
         Top             =   975
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "INSTRUMEN04.frx":1354
         Top             =   7320
         Width           =   1515
      End
      Begin VB.Label COMMLAB 
         Height          =   225
         Left            =   0
         TabIndex        =   55
         Top             =   4320
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label DOCUORI 
         Height          =   225
         Left            =   0
         TabIndex        =   54
         Top             =   5640
         Visible         =   0   'False
         Width           =   990
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Instrumentos de Medición - Calibración"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   150
      TabIndex        =   30
      Top             =   3885
      Width           =   10620
      Begin VB.CommandButton Command10 
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
         Left            =   9870
         Picture         =   "INSTRUMEN04.frx":3276
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Ficha de Calibración"
         Top             =   1050
         Width           =   540
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1890
         TabIndex        =   38
         Top             =   1260
         Width           =   7890
      End
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
         Height          =   300
         Index           =   0
         Left            =   3675
         TabIndex        =   33
         Top             =   420
         Width           =   175
      End
      Begin VB.CommandButton Command5 
         Height          =   600
         Index           =   0
         Left            =   9870
         Picture         =   "INSTRUMEN04.frx":3580
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Hoja de Procedimientos de Calibración"
         Top             =   420
         Width           =   540
      End
      Begin VB.CommandButton Command6 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   9555
         TabIndex        =   31
         ToolTipText     =   "Desconecta Documento"
         Top             =   420
         Width           =   230
      End
      Begin VB.Label Label18 
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
         Height          =   330
         Left            =   1890
         TabIndex        =   43
         Top             =   840
         Width           =   7890
      End
      Begin VB.Label Label17 
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
         Height          =   300
         Left            =   8190
         TabIndex        =   42
         Top             =   420
         Width           =   1275
      End
      Begin VB.Label Label16 
         Height          =   225
         Left            =   6615
         TabIndex        =   41
         Top             =   210
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label15 
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
         Height          =   300
         Left            =   5565
         TabIndex        =   40
         Top             =   420
         Width           =   750
      End
      Begin VB.Label Label14 
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
         Height          =   300
         Left            =   1890
         TabIndex        =   39
         Top             =   420
         Width           =   1800
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Intervalo de Calibración:"
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
         TabIndex        =   37
         Top             =   1155
         Width           =   1320
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Procedimiento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   -420
         TabIndex        =   36
         Top             =   480
         Width           =   2220
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rev:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   4935
         TabIndex        =   35
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   7350
         TabIndex        =   34
         Top             =   480
         Width           =   750
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Instrumentos de Medición - Datos Complementarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   150
      TabIndex        =   13
      Top             =   1575
      Width           =   10620
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1890
         TabIndex        =   28
         Top             =   1680
         Width           =   3930
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7245
         TabIndex        =   27
         Top             =   1260
         Width           =   3195
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1890
         TabIndex        =   24
         Top             =   1260
         Width           =   3930
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5460
         TabIndex        =   23
         Top             =   840
         Width           =   4980
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1890
         TabIndex        =   20
         Top             =   840
         Width           =   2250
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5460
         TabIndex        =   19
         Top             =   420
         Width           =   4980
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1890
         TabIndex        =   15
         Top             =   420
         Width           =   2250
      End
      Begin VB.CommandButton Command11 
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
         Height          =   330
         Left            =   6930
         TabIndex        =   14
         Top             =   1680
         Width           =   175
      End
      Begin VB.Label Label2 
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
         Height          =   330
         Left            =   7245
         TabIndex        =   45
         Top             =   1680
         Width           =   3165
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   6510
         TabIndex        =   44
         Top             =   1680
         Width           =   435
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicación:"
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
         Left            =   315
         TabIndex        =   29
         Top             =   1755
         Width           =   1425
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Precisión:"
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
         Left            =   5880
         TabIndex        =   26
         Top             =   1365
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Rango:"
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
         Left            =   735
         TabIndex        =   25
         Top             =   1335
         Width           =   1005
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Modelo:"
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
         Left            =   4305
         TabIndex        =   22
         Top             =   945
         Width           =   1005
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro de Serie:"
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
         TabIndex        =   21
         Top             =   915
         Width           =   1635
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
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
         Left            =   4305
         TabIndex        =   18
         Top             =   525
         Width           =   1005
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
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
         Left            =   735
         TabIndex        =   17
         Top             =   525
         Width           =   1005
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gr:"
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
         Left            =   4620
         TabIndex        =   16
         Top             =   1785
         Width           =   1800
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Instrumentos de Medición - Identificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   150
      TabIndex        =   2
      Top             =   105
      Width           =   8310
      Begin VB.CommandButton Command9 
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
         Height          =   330
         Left            =   3465
         TabIndex        =   12
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1890
         TabIndex        =   11
         Top             =   840
         Width           =   6240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo:"
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
         Left            =   795
         TabIndex        =   10
         Top             =   495
         Width           =   1005
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Denominación:"
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
         Left            =   0
         TabIndex        =   9
         Top             =   900
         Width           =   1800
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   7380
         TabIndex        =   8
         Top             =   420
         Width           =   750
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "C.Interno:"
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
         Left            =   6300
         TabIndex        =   7
         Top             =   525
         Width           =   1005
      End
      Begin VB.Label Label37 
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
         Height          =   330
         Left            =   1890
         TabIndex        =   6
         Top             =   420
         Width           =   1590
      End
   End
   Begin VB.Frame MARCOFICHA 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Observaciones - Componentes y/o Accesorios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Index           =   1
      Left            =   150
      TabIndex        =   3
      Top             =   5775
      Width           =   10620
      Begin VB.TextBox Text28 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   315
         Width           =   10275
      End
      Begin VB.Line Line1 
         X1              =   105
         X2              =   10375
         Y1              =   1695
         Y2              =   1695
      End
      Begin VB.Line Line3 
         X1              =   10375
         X2              =   10375
         Y1              =   315
         Y2              =   1700
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   8520
      TabIndex        =   0
      Top             =   120
      Width           =   2220
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
         Height          =   690
         Left            =   240
         Picture         =   "INSTRUMEN04.frx":388A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   360
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Undo"
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
         Left            =   1320
         Picture         =   "INSTRUMEN04.frx":3B94
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Restablece los Valores Anteriores"
         Top             =   360
         Width           =   650
      End
   End
End
Attribute VB_Name = "INSTRUMEN04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MODIDOC%
Dim GRABATODO%
'
Private Sub Command1_Click()
   '
   GRABATODO% = 1
   Call GRADAINSTRUM
   GRABATODO% = 0
   Exit Sub
   
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   If TRIM$(Label37) = "" Or ECOARCH$("INSTRUME", AJUSTI$(Label37, 12)) = "" Then
      Call MENSERR(24, "Falta Seleccionar Instrumento")
      GoTo 99
   End If
   FICALIB04.Label37 = Label37
   FICALIB04.Label14 = Text3
   FICALIB04.Label28 = Label28
   FICALIB04.Label16 = Label16
   '
   FICALIB04.Label2(0) = Text1
   FICALIB04.Label2(1) = Text2
   FICALIB04.Label2(2) = Text4
   FICALIB04.Label2(3) = Text5
   FICALIB04.Label2(4) = Text6
   FICALIB04.Label2(5) = Text7
   FICALIB04.Label2(6) = Text8
   FICALIB04.Label2(7) = Text9
   FICALIB04.Label2(8) = Label14
   '
   FICALIB04.List2.Clear
   TTYZ$ = Text28.TEXT
   LOCADE& = Len(TTYZ$)
   For U& = LOCADE& To 1 Step -1
     If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 10
   Next U&
   U& = 0
   '
10 TEXTO$ = Left$(TTYZ$, U&) + Chr$(13) + Chr$(10)
   Call FRATEXTO(TEXTO$, 82)
   For LL% = 1 To CARETEX%
      FICALIBRA.List2.AddItem RETEX$(LL%)
   Next LL%
   '
   FICALIB04.Show 1
   '
99 Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   XX$ = REGSEL$("GRUINSTR")
   Label7 = TRIM$(Str$(REGNUM&("GRUINSTR")))
End Sub

Private Sub COMMAND21_CLICK()
   Command21.Enabled = False
   Call GRADAINSTRUM
   OPX% = ALTERNA%("Instrumento\Grupo de Instrumentos\Procedimiento de Calibración")
   Select Case OPX%
     Case 1
       Call ALTAINSTRU
     Case 2
       CONFINSTR04.Show 1
     Case 3
       Call DOCUMI04.NUEINSTRU
   End Select
   Command21.Enabled = True
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
     Call GRADAINSTRUM
     PRINTFORM
   Command22.Enabled = True
End Sub

Private Sub Command3_Click()
   Call GRADAINSTRUM
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub COMMAND31_CLICK(Index As Integer)
   XX$ = REGSEL$("INSTRUCI")
   If Len(XX$) > 4 Then
     CIIINSTRU% = REGNUM("INSTRUCI")
     Label16 = CIIINSTRU%
   End If
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call MENSERR(24, "Ayuda en Línea no Disponible")
   Command4.Enabled = True
End Sub

Private Sub Command5_Click(Index As Integer)
   Call GRADAINSTRUM
   CIIINSTRU% = Val(Label16)
   If CIIINSTRU% < 1 Or CIIINSTRU% > ULTREG&("INSTRUCI") Then
      Call MENSERR(24, "Falta Seleccionar Procedimiento")
      Exit Sub
   End If
   '
   On Error Resume Next
   INSTRUCCI.Label28 = TRIM$(Str$(CIIINSTRU%))
   INSTRUI04.Label28 = TRIM$(Str$(CIIINSTRU%))
   INSTRUCCI.Show 1
   If Err Then
     On Error GoTo 0
     INSTRUI04.Show 1
   End If
   On Error GoTo 0
   '
   Label16 = ""
   Label16 = CIIINSTRU%
   '
   Call CARDAINSTRUM
End Sub

Private Sub Command6_Click(Index As Integer)
   Label16 = ""
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   Call GRADAINSTRUM
   CONFINSTR04.Show 1
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   '
   Command8.Enabled = False
5  XX$ = REGSEL$("GRUINSTR")
   If Len(XX$) <= 4 Then GoTo 99
   VDEF$ = XX$
   '
   INDIN& = REGNUM&("GRUINSTR")
   NGRU% = INDIN&
   '
   FORIPRE$ = TRIM$(CONTROL$("FORINMED", "DOCUMED0"))
   If FORIPRE$ = "" Then
     Call MENSERR(24, "Falta Definir Formulario de Impresión")
     CONFINSTRUM.Show 1
     GoTo 5
   End If
   '
   Screen.MousePointer = 11
   For KKL& = 1 To 4
     OBX$ = TRIM$(Mid$(VDEF$, 65 + 64 * KKL&, 64))
     Call GRAREG("!OBSERVOF", OBX$, KKL&)
   Next KKL&
   Call SETULTREG("!OBSERVOF", 4)
   Call CIERRARCH("!OBSERVOF")
   '
   FEX = HOY(HOS$)
   FECHDOC$ = HOS$
   NUMEDOC$ = ""
   TIPODOC$ = "Inventario de Instrumentos - " + TRIM$(Mid$(VDEF$, 3, 62))
   '
   CODPARAM$(1) = "UNIMED"
   CODPARAM$(2) = "COND-PAG"
   CODPARAM$(3) = "VENDEDOR"
   CODPARAM$(4) = "FECH-VEN"
   CAPARDOC% = 4
   '
   DOCPARAM$(1) = TRIM$(Str$(INDIN&)) + " - " + Mid$(VDEF$, 3, 62)
   DOCPARAM$(2) = TRIM$(Mid$(VDEF$, 65, 30))
   DOCPARAM$(3) = TRIM$(Mid$(VDEF$, 97, 30))
   FEX = CVI(Mid$(VDEF$, 127, 2))
   DOCPARAM$(4) = FECHATEX$(FEX)
   '
   CODTABU1$(1) = "REF-MAT1"
   CODTABU1$(2) = "REF-MAT2"
   CODTABU1$(3) = "REF-MAT3"
   CODTABU1$(4) = "NRO-LOTE"
   CODTABU1$(5) = "TALLE"
   CODTABU1$(6) = "COLOR"
   CACOLTAB1% = 6
   '
   CAITAB1% = 0
   For U& = 1 To ULTREG&("INSTRUME")
     Z$ = REGLEIDO$("INSTRUME", U&)
     If CVI(Mid$(Z$, 67, 2)) = NGRU% Then
       CAITAB1% = CAITAB1% + 1
       TETABU1$(1, CAITAB1%) = Mid$(Z$, 1, 12)
       TETABU1$(2, CAITAB1%) = Mid$(Z$, 13, 52)
       TETABU1$(3, CAITAB1%) = Mid$(Z$, 97, 24)
       TETABU1$(4, CAITAB1%) = Mid$(Z$, 121, 24)
       TETABU1$(5, CAITAB1%) = Mid$(Z$, 225, 32)
       TETABU1$(6, CAITAB1%) = Mid$(Z$, 193, 32)
     End If
   Next U&
   '
   Call PRINTDOC("@" + FORIPRE$)
   GoTo 5
   Screen.MousePointer = 1
   '
99 Command8.Enabled = True
   '
End Sub

Private Sub Command9_Click()
   Call GRADAINSTRUM
   XX$ = REGSEL$("INSTRUME")
   If Len(XX$) > 4 Then
     CIIINSTRU% = REGNUM&("INSTRUME")
     Label28 = CIIINSTRU%
   End If
End Sub

Private Sub COMMLAB_Change()
   If TRIM$(UCase$(COMMLAB.Caption)) = "PRINT" Then
      DOCUOR$ = TRIM$(DOCUORI.Caption)
      Call PRINPROCE(DOCUOR$)
      COMMLAB.Caption = ""
      DOCUORI.Caption = ""
   End If
End Sub

Private Sub Form_Load()
    '
    Screen.MousePointer = 11
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    Screen.MousePointer = 1
    '
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Resize()
   DoEvents
   If WindowState = 1 Then
      Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Label14_DblClick()
   XX$ = REGSEL$("INSTRUCI")
   If Len(XX$) > 4 Then
     CIIINSTRU% = REGNUM("INSTRUCI")
     Label16 = CIIINSTRU%
   End If
End Sub

Private Sub Label16_Change()
  '
  MODIDOC% = 1
  REINSTRUC& = Val(Label16)
  Label14 = ""
  Label15 = ""
  Label17 = ""
  Label18 = ""
  '
  If REINSTRUC& > 0 Then
    If REINSTRUC& <= ULTREG&("INSTRUCI") Then
      XX$ = REGLEIDO$("INSTRUCI", REINSTRUC&)
      CODINSTRU$ = Left$(XX$, 16)
      DESINSTRU$ = Mid$(XX$, 17, 48)
      REVINSTRU$ = Mid$(XX$, 65, 4)
      FEAPRO = CVI(Mid$(XX$, 121, 2))
      '
      Label14 = CODINSTRU$
      Label18 = DESINSTRU$
      Label15 = REVINSTRU$
      Label17 = FECHATEX$(FEAPRO)
    End If
  End If
  '
End Sub

Private Sub Label28_Change()
   '
   Call CARDAINSTRUM
   '
End Sub

Private Sub Label37_Click()
   Call GRADAINSTRUM
End Sub

Private Sub Label37_DblClick()
   XX$ = REGSEL$("INSTRUME")
   If Len(XX$) > 4 Then
     CIIINSTRU% = REGNUM&("INSTRUME")
     Label28 = CIIINSTRU%
   End If
End Sub

Private Sub Label7_Change()
  MODIDOC% = 1
  IGRU& = Val(Label7)
  Label2 = ""
  If IGRU& > 0 Then
    If IGRU& <= ULTREG&("GRUINSTR") Then
      XX$ = REGLEIDO$("GRUINSTR", IGRU&)
      Label2 = TRIM$(Mid$(XX$, 3, 62))
    End If
  End If
End Sub

Private Sub Label7_DblClick()
   XX$ = REGSEL$("GRUINSTR")
   Label7 = TRIM$(Str$(REGNUM&("GRUINSTR")))
End Sub

Private Sub Text1_Change()
  MODIDOC% = 1
End Sub

Private Sub Text1_DblClick()
   '
   'Call GRADATITEM
   '
   'Call SELECHO("MASTER")
   'COD$ = TRIM$(VALACT1$("MASTER"))
   'If COD$ <> "" Then
   '  Text1.TEXT = COD$
   'End If
   '
End Sub

Private Sub Text1_GotFocus()
   '
   'Call GRADATITEM
   '
End Sub

Sub CARDAINSTRUM()
   '
   CIIINSTRU% = Val(Label28)
   If CIIINSTRU% < 1 Then Exit Sub
   '
   REGI& = CIIINSTRU%
   URN& = ULTREG&("INSTRUME")
   URX& = ULTREG&("OBINSTRU")
   If REGI& <= URN& Then
     XX$ = REGLEIDO$("INSTRUME", REGI&)
     Label37 = Left$(XX$, 12)
     Text3 = Mid$(XX$, 13, 52)
     Label7 = CVI(Mid$(XX$, 67, 2))
     Text1 = Mid$(XX$, 73, 24)
     Text2 = Mid$(XX$, 97, 24)
     Text5 = Mid$(XX$, 121, 24)
     Text4 = Mid$(XX$, 145, 16)
     Text6 = Mid$(XX$, 161, 32)
     Text7 = Mid$(XX$, 193, 32)
     Text8 = Mid$(XX$, 225, 32)
     Text9 = Mid$(XX$, 257, 32)
     Label16 = CVI(Mid$(XX$, 69, 2))
     '
     Text28.TEXT = "": TTYZ$ = ""
     If CIIINSTRU% > 0 Then
       If CIIINSTRU% <= URN& Then
         Screen.MousePointer = 11
         RNC$ = RECFILT$("OBINSTRU", 1, MKI$(CIIINSTRU%))
         '
         For U& = 1 To Len(RNC$) Step 4
           RENOTA& = CVS(Mid$(RNC$, U&, 4))
           If RENOTA& > 0 Then
             If RENOTA& <= URX& Then
               TTYY$ = RTrim$(Mid$(REGLEIDO$("OBINSTRU", RENOTA&), 3))
               If TTYZ$ <> "" Then
                 TTYZ$ = TTYZ$ + Chr$(13) + Chr$(10)
               End If
               TTYZ$ = TTYZ$ + RTrim$(TTYY$)
             End If
           End If
         Next U&
         Screen.MousePointer = 1
       End If
     End If
     '
     LOCADE& = Len(TTYZ$)
     For U& = LOCADE& To 1 Step -1
       If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 30
     Next U&
     U& = 0
     '
30   Text28.TEXT = Left$(TTYZ$, U&)
     Text28.Refresh
     '
     OPENFORMS = DoEvents
     MODIDOC% = 0
     '
   End If
   '
End Sub
'
Sub GRADAINSTRUM()
   '
   REGI& = Val(Label28)
   If REGI& < 1 Then Exit Sub
   If REGI& > ULTREG&("INSTRUME") Then Exit Sub
   If MODIDOC% <> 1 Then Exit Sub
   '
   If GRABATODO% <> 1 Then
     If COMALTER%("Se han Modificado los Datos del Instrumento\'" + TRIM$(Label37) + " - " + TRIM$(Text3) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       Call CARDAINSTRUM
       MODIDOC% = 0
       Exit Sub
     End If
   End If
   '
   If Val(Label7) < 1 Or Val(Label7) > ULTREG&("GRUINSTR") Then
     Call MENSERR(24, "Grupo de Instrumento no Válido")
     Exit Sub
   End If
   '
   Call BLOQARCH("INSTRUME")
   Call BLOQARCH("OBINSTRU")
   XX$ = REGLEIDO$("INSTRUME", REGI&)
   Call REPLA(XX$, TRIM$(Label37), 1, 12)
   Call REPLA(XX$, TRIM$(Text3), 13, 52)
   Call REPLA(XX$, MKI$(Val(Label7)), 67, 2)
   Call REPLA(XX$, MKI$(Val(Label16)), 69, 2)
   Call REPLA(XX$, TRIM$(Text1), 73, 24)
   Call REPLA(XX$, TRIM$(Text2), 97, 24)
   Call REPLA(XX$, TRIM$(Text5), 121, 24)
   Call REPLA(XX$, TRIM$(Text4), 145, 16)
   Call REPLA(XX$, TRIM$(Text6), 161, 32)
   Call REPLA(XX$, TRIM$(Text7), 193, 32)
   Call REPLA(XX$, TRIM$(Text8), 225, 32)
   Call REPLA(XX$, TRIM$(Text9), 257, 32)
   Call GRAREG("INSTRUME", XX$, REGI&)
   Call CIERRARCH("INSTRUME")
   '
   CIIINSTRU% = REGI&
   URN& = ULTREG&("OBINSTRU")
   RNC$ = RECFILT$("OBINSTRU", 1, MKI$(CIIINSTRU%))
   KK& = (-3)
   '
   TTYZ$ = Text28.TEXT
   LOCADE& = Len(TTYZ$)
   For U& = LOCADE& To 1 Step -1
     If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 10
   Next U&
   U& = 0
   '
10 TEXTO$ = Left$(TTYZ$, U&) + Chr$(13) + Chr$(10)
   Call FRATEXTO(TEXTO$, 82)
   For LL% = 1 To CARETEX%
      TTYY$ = MKI$(CIIINSTRU%) + RETEX$(LL%)
      KK& = KK& + 4
      If KK& <= Len(RNC$) - 3 Then
           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
         Else
           URN& = URN& + 1
           RENOTA& = URN&
      End If
      Call GRAREG("OBINSTRU", TTYY$, RENOTA&)
   Next LL%
   '
20 While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG("OBINSTRU", MKI$(0) + Space$(82), RENOTA&)
   Wend
   '
   Call CIERRARCH("OBINSTRU")
   MODIDOC% = 0
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
End Sub
'

Private Sub Text2_Change()
  MODIDOC% = 1
End Sub

Private Sub Text28_Change()
   MODIDOC% = 1
End Sub

Private Sub Text3_Change()
  MODIDOC% = 1
End Sub

Private Sub Text4_Change()
  MODIDOC% = 1
End Sub

Sub PRINPROCE(DOCUORIGEN$)
  '
End Sub

Sub ALTAINSTRU()
   '
   VDEF$ = REGBLAN$("INSTRUME")
   INDINS% = 1 + ULTREG&("INSTRUME")
   Call REPLA(VDEF$, MKI$(INDINS%), 65, 2)
   '
10 Screen.MousePointer = 1
   OBX$ = OBJPLA$("NUEINSTRUMEN", VDEF$)
   If OBX$ = "" Then
     Exit Sub
   End If
   '
   VDEF$ = OBX$
   CODINS$ = UCase$(TRIM$(Left$(OBX$, 12)))
   If CODINS$ = "" Then
      Call MENSERR(24, "Falta Código de Instrumento")
      GoTo 10
   End If
   Call REPLA(OBX$, CODINS$, 1, 12)
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("INSTRUME")
   For UI& = 1 To FIN&
      XX$ = REGLEIDO$("INSTRUME", UI&)
      If UCase$(TRIM$(Left$(XX$, 12))) = CODINS$ Then
         Call MENSERR(24, "Código Pre-Existente")
         GoTo 10
      End If
   Next UI&
   '
   DESINS$ = TRIM$(Mid$(OBX$, 13, 52))
   If DESINS$ = "" Then
      Call MENSERR(24, "Falta Denominación del Instrumento")
      GoTo 10
   End If
   '
   GRUIN% = CVI(Mid$(OBX$, 67, 2))
   If GRUIN% < 1 Or GRUIN% > ULTREG&("GRUINSTR") Then
      Call MENSERR(24, "Grupo de Instrumento Inexistente o No Válido")
      GoTo 10
   End If
   '
   Call CIERRARCH("INSTRUME")
   Call BLOQARCH("INSTRUME")
   INDINS% = 1 + ULTREG&("INSTRUME")
   Call REPLA(OBX$, MKI$(INDINS%), 65, 2)
   Call REGAPP("INSTRUME", OBX$)
   Call CIERRARCH("INSTRUME")
   Call BAJALDISCO
   Screen.MousePointer = 1
   '
   INSTRUMEN04.Label28 = TRIM$(Str$(INDINS%))
   On Error Resume Next
   INSTRUMEN04.Show 1
   On Error GoTo 0
   '
End Sub

Private Sub Text5_Change()
  MODIDOC% = 1
End Sub

Private Sub Text6_Change()
  MODIDOC% = 1
End Sub

Private Sub Text7_Change()
  MODIDOC% = 1
End Sub

Private Sub Text8_Change()
  MODIDOC% = 1
End Sub

Private Sub Text9_Change()
  MODIDOC% = 1
End Sub
