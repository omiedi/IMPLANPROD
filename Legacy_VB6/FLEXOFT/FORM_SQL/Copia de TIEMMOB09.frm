VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form TIEMMOB09 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "A.B.M. DE OPERACIONES"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   14250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command8 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   6360
      MaskColor       =   &H0084962E&
      TabIndex        =   47
      ToolTipText     =   "Click Aquì Para Ingresar la Operación a la Lista."
      Top             =   2560
      Width           =   1215
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FBCA9D&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2700
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   8580
      Width           =   1545
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Text            =   "Total Tiempo en Minutos:"
      Top             =   8580
      Width           =   2595
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EFF1F1&
      Height          =   2055
      Left            =   120
      ScaleHeight     =   1995
      ScaleWidth      =   7215
      TabIndex        =   42
      Top             =   2640
      Visible         =   0   'False
      Width           =   7270
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Width           =   7215
         WordWrap        =   -1  'True
      End
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   9060
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Text            =   "Total Tiempo en Segundos:"
      Top             =   8580
      Width           =   2595
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   8400
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   34
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Text            =   "F.Alta:"
      Top             =   600
      Width           =   795
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   12060
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   600
      Width           =   1120
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FBCA9D&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   600
      Width           =   7935
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   70
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "Còdigo:"
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text15 
      Appearance      =   0  'Flat
      BackColor       =   &H00FBCA9D&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FBCA9D&
      Height          =   600
      Index           =   3
      Left            =   70
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "OPERACIONES Y TIEMPOS DE MANOS DE OBRA"
      Top             =   0
      Width           =   13155
   End
   Begin VB.TextBox Text161 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FBCA9D&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   11640
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Total de Tiempos de Operaciones del Código Activo"
      Top             =   8580
      Width           =   1425
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   70
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   8520
      Width           =   13095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Height          =   9135
      Left            =   13200
      TabIndex        =   11
      Top             =   -120
      Width           =   75
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Height          =   9015
      Left            =   0
      TabIndex        =   10
      Top             =   -70
      Width           =   80
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00800000&
      FillColor       =   &H00400000&
      ForeColor       =   &H00400000&
      Height          =   8955
      Left            =   13200
      ScaleHeight     =   8895
      ScaleWidth      =   1755
      TabIndex        =   5
      Top             =   0
      Width           =   1815
      Begin VB.CommandButton Command21 
         Caption         =   "Tbls"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   240
         Picture         =   "TIEMMOB09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Tabla de Sectores"
         Top             =   5400
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
         Height          =   620
         Left            =   240
         Picture         =   "TIEMMOB09.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Imprime Tiempos de Mano de Obra del Item Activo"
         Top             =   4650
         UseMaskColor    =   -1  'True
         Width           =   650
      End
      Begin VB.CommandButton Command5 
         Height          =   620
         Left            =   240
         Picture         =   "TIEMMOB09.frx":07B4
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Búsqueda por Texto "
         Top             =   3750
         Width           =   650
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00E0E0E0&
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
         Left            =   240
         Picture         =   "TIEMMOB09.frx":1B1A
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Graba Lista de Operaciones y Tiempos"
         Top             =   7440
         Width           =   650
      End
      Begin VB.CommandButton Command22 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   580
         Left            =   240
         Picture         =   "TIEMMOB09.frx":1E24
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Acceder a Historial de Operaciones x Artículo"
         Top             =   2850
         Width           =   650
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Copy"
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
         Left            =   240
         Picture         =   "TIEMMOB09.frx":212E
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Copiar Tiempos del Código Activo Hacia Otro."
         Top             =   1950
         Width           =   650
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
         Left            =   240
         Picture         =   "TIEMMOB09.frx":2278
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   1050
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   240
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   7
         Top             =   8250
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   12000
         End
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
         Height          =   650
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir Sin Procesar"
         Top             =   100
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -120
         Picture         =   "TIEMMOB09.frx":2582
         Stretch         =   -1  'True
         Top             =   8520
         Width           =   1410
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1605
      Left            =   70
      TabIndex        =   18
      Top             =   960
      Width           =   13185
      Begin VB.TextBox TXTTIEMPO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   7550
         TabIndex        =   1
         Top             =   540
         Width           =   1110
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Ingresar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   10800
         MaskColor       =   &H0084962E&
         TabIndex        =   4
         ToolTipText     =   "Click Aquì Para Ingresar la Operación a la Lista."
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TXTNOPER 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
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
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Doble Click Para Seleccionar Sección"
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton Command2 
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
         Height          =   360
         Left            =   10600
         TabIndex        =   2
         ToolTipText     =   "Click Para Seleccionar Sección"
         Top             =   550
         Width           =   175
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   12075
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "Cant.Op."
         Top             =   960
         Width           =   1000
      End
      Begin VB.TextBox TXTCANTOP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12075
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1230
         Width           =   1000
      End
      Begin VB.TextBox TXTCANTIDAD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10800
         TabIndex        =   3
         Top             =   550
         Width           =   1180
      End
      Begin VB.TextBox TXTDESCRIPCION 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   10
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   540
         Width           =   7500
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "Sector    /   N°"
         ToolTipText     =   "Doble Click Para Seleccionar Sección"
         Top             =   270
         Width           =   2160
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   7550
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "Tiempo Unit."
         Top             =   270
         Width           =   1150
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   10
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "Descripción"
         Top             =   270
         Width           =   7550
      End
      Begin VB.TextBox TXTSECCION 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   23
         ToolTipText     =   "Doble Click Para Seleccionar Sección"
         Top             =   540
         Width           =   1950
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   11970
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "Modificado"
         Top             =   270
         Width           =   1095
      End
      Begin VB.TextBox TXTMODIFICADO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   11970
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   540
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Check1"
         Height          =   255
         Left            =   12400
         TabIndex        =   19
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   10800
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "Cantidad"
         Top             =   270
         Width           =   1195
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "TIEMMOB09.frx":44A4
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "TIEMMOB09.frx":46D8
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   5685
      Left            =   60
      TabIndex        =   41
      Top             =   2805
      Width           =   13110
      _ExtentX        =   23125
      _ExtentY        =   10028
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      SelectionMode   =   2
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   195
      Left            =   720
      TabIndex        =   9
      Top             =   6120
      Width           =   870
   End
End
Attribute VB_Name = "TIEMMOB09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARBUS$, ABUSQUE$
Sub ACOMODARMSF()

    'ACA ACOMODO LOS TEXTOS NUEVAMENETE
'    For I& = 1 To MSF1.Rows - 1
'     TXTX$ = ""
'     DESCRI$ = TRIM$(MSF1.TextMatrix(I&, 2))
''     DESCRI$ = LTrim$(MSF1.TextMatrix(I&, 2))
'     Call FRATEXTO(DESCRI$, 68)
'     For JJ& = 1 To CARETEX%
'       TXTX$ = TXTX$ + RETEX$(JJ&) + vbCrLf
'     Next JJ&
'     Label7.Height = 250
'     Label7 = DESCRI$
     MSF1.RowHeight(MSF1.Row) = Label7.Height + 100
     MSF1.TextMatrix(MSF1.Row, 2) = Label7
'    Next I&

End Sub

Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Private Sub Command1_Click()
  CI1% = CODINT%(ANACOSI09.Text1)
  ANACOSI09.Text4 = FORMATNUM$(MANO_OBRA_REAL(CI1%), "F6.2") ' MANO DE OBRA REAL
  Unload Me
End Sub

Private Sub Command14_Click()

End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Call SELECHO("SECPROD/SECTORES")
    COD$ = VALACT1$("SECPROD")
    If TRIM$(COD$) <> "" Then
      TXTNOPER.TEXT = COD$
    End If
    On Error Resume Next
     TXTCANTIDAD.SetFocus
    On Error GoTo 0
    Command2.Enabled = True
End Sub



Private Sub Command21_Click()
   Call CARSECPROD
End Sub

Private Sub Command22_Click()
  Command22.Enabled = False
  ANOTIEOBRA.Label8 = Text15
  ANOTIEOBRA.Show 1
  Command22.Enabled = True
End Sub

Private Sub Command3_Click()
   Call TXTCANTIDAD_KeyPress(13)
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   CI1% = CODINT(Text15)
   If CI1% < 1 Then GoTo 99
   '
   VDEF$ = String$(4, 0)
   Call REPLA(VDEF$, MKI$(CI1%), 1, 2)
   '
10 OBX$ = OBJPLA$("COPIACOD", VDEF$)
   If OBX$ = "" Then
       GoTo 99
   End If
   '
   CI2% = CVI(Mid$(OBX$, 3, 2))
   If CI2% < 1 Then
      Call COMUNI("Falta Código Destino o Código Inexistente")
      GoTo 10
   End If
   '
   If CantRegistros("TIEMOBRA", "CODICONJ =" & CI2%, "NOMESS") > 0 Then
      OPX% = COMALTER%("Este Código Posee Tiempos de Mano de Obra Cargados\\Mantener Actual\Sobre-escribir")
      If OPX% < 2 Then GoTo 10
      '
      Call GRABATIEMPMOBRA(CI2%)
      'TENER EN CUENTA QUE TAMBIEN SE VA ACTUALIZAR LA MANO DE OBRA REAL GUARDADA PARA QUE NO HAYA DIFERENCIAS ENTRE TIEMPOS Y CARGADOS Y  ESTA.
      Text15 = CODEXT$(CI2%)
   End If
   '
99 Command4.Enabled = True
End Sub

Private Sub Command5_Click()
 Command5.Enabled = False
 ABUSQUE$ = UCase$(TRIM$(InputBox$("Busqueda por Texto", , ARBUS$)))
   If ABUSQUE$ <> "" Then
     ARBUS$ = ABUSQUE$
     For U& = 1 To MSF1.Rows - 1
      For I& = 0 To MSF1.Cols - 1
       If InStr(UCase$(MSF1.TextMatrix(U&, I&)), ABUSQUE$) > 0 Then
        MSF1.Row = U&: MSF1.Col = I&
        If MSF1.TopRow > MSF1.Row Then GRID.TopRow = MSF1.Row
        On Error Resume Next
        VISIROWS = (MSF1.Height / MSF1.RowHeight(j&)) - 4
        If Err > 0 Then
         VISIROWS = 1
        End If
        On Error GoTo 0
         If VISIROWS > 0 Then
            If MSF1.TopRow < MSF1.Row - VISIROWS Then
              MSF1.TopRow = MSF1.Row - VISIROWS
            End If
         End If
        MSF1.SetFocus
        GoTo 99
       End If
       Next I&
     Next U&
     '
     Call MENSERR(24, "Argumento no Encontrado.")
   End If
99 Command5.Enabled = True
End Sub

Private Sub Command6_Click()

   Command6.Enabled = False
   If DERACCE%("MODIMOBRA", "Modificar Datos de Mano de Obra") < 2 Then Exit Sub
   
   MOBRA_REAL = MANO_OBRA_REAL(CODINT(Text15)) ' MANO DE OBRA REAL
   MOBRA_REAL = TRIM$(FORMATNUM$(MOBRA_REAL, "F12.2"))
   TIEMPX = TRIM$(FORMATNUM$((XVALO(Text161) / 60), "F12.2"))
   OPX% = COMALTER%("Se va a Modificar La Mano de Obra Real\Actual: " & MOBRA_REAL & " \Nueva: " & TIEMPX & " \\Cancelar\Realmente Desea Actualizar Mano de Obra")
   If OPX% < 2 Then GoTo 99
   '
   CI1% = CODINT%(TRIM$(Text15))
   '
   Call GRABATIEMPMOBRA(CI1%)
   Call COMUNI("Mano de Obra Actualizada")
   'REFRESCO LA MANO DE OBRA
99  Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    CI1% = CODINT%(Text15)
    If CI1% < 1 Then
       Call COMUNI("Código Inexistente o No Válido")
       GoTo 99
    End If
    
    Call FINAL("?TIEMMOBRA(" & TRIM$(Text15) & ")")
99  Command7.Enabled = True
    
End Sub

Private Sub Command8_Click()
   BUSCA_DES.Show 1
End Sub

Private Sub Form_Load()
   Call APLICACIONSKINS(Me)
   '
   TEPRUEBA.Font = MSF1.Font
   TEPRUEBA.FontSize = MSF1.FontSize
   MSF1.Cols = 9
   NUCOLU% = 1: TEPRUEBA = "99": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = " "
   NUCOLU% = 2: TEPRUEBA = "999999": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = "N.Ord"
   NUCOLU% = 3: TEPRUEBA = "123456789*123456789*123456789*123456789*123456789*123456789*123456789*12": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = "Descripción"
   NUCOLU% = 4: TEPRUEBA = "1234567": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = "T.Un."
   NUCOLU% = 5: TEPRUEBA = "1234567": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = "Qt."
   NUCOLU% = 6: TEPRUEBA = "": MSF1.ColWidth(NUCOLU% - 1) = 0: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = ""
   NUCOLU% = 7: TEPRUEBA = "SECCIONSECC": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = "Secciòn"
   NUCOLU% = 8: TEPRUEBA = "1234567890": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = "Tiem.Tot."
   NUCOLU% = 9: TEPRUEBA = "01'02¡1234": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = "Modif."
   '
   MSF1.MergeCells = flexMergeRestrictColumns
   MSF1.MergeCol(1) = True

   Call CREACAMPO("", "TIEMOBRA", "NUORD", 3, 0)
   Call CREACAMPO("", "TIEMOBRA", "CODICONJ", 3, 0, 1)
   Call CREACAMPO("", "TIEMOBRA", "DESCRIPCION", 10, 1280)
   Call CREACAMPO("", "TIEMOBRA", "TIEMPOUNITARIO", 3, 0)
   Call CREACAMPO("", "TIEMOBRA", "SecciPro", 6, 4)
   Call CREACAMPO("", "TIEMOBRA", "CantiOps", 3, 0)
   Call CREACAMPO("", "TIEMOBRA", "FECHAMODIF", 8, 0)
   Call CREACAMPO("", "TIEMOBRA", "MARBORRA", 3, 0)
   '
   Call CREACAMPO("ARTICULOS", "MASTER", "FECHALTA_MO", 8, 0)
   '
   Call CREACAMPO("", "HISTOMOBRA", "CODICONJ", 3, 0, 1)
   Call CREACAMPO("", "HISTOMOBRA", "ANOTMOBRA", 12, 0)
   Call CREACAMPO("", "HISTOMOBRA", "MARBORRA", 3, 0)
   '
   TXTMODIFICADO = FECHATEX$(HOY(HO$))
   '
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If Cancel = 0 Then
     CI1% = CODINT%(ANACOSI09.Text1)
     ANACOSI09.Text4 = FORMATNUM$(MANO_OBRA_REAL(CI1%), "F6.2") ' MANO DE OBRA REAL
  End If
End Sub


Private Sub MSF1_DblClick()
   If MSF1.Row >= 1 Then
    With EDITMOBRA
         .TXTDESCRIPCION.TEXT = MSF1.TextMatrix(MSF1.Row, 2)
         .TXTTIEMPO.TEXT = MSF1.TextMatrix(MSF1.Row, 3)
         .TXTCANTIDAD.TEXT = MSF1.TextMatrix(MSF1.Row, 4)
         .TXTNOPER.TEXT = MSF1.TextMatrix(MSF1.Row, 5)
         .TXTSECCION.TEXT = MSF1.TextMatrix(MSF1.Row, 6)
         .TXTMODIFICADO.TEXT = MSF1.TextMatrix(MSF1.Row, 8)
         .Show 1
     End With
   End If
End Sub

Private Sub Text15_Change()
   CODE$ = TRIM$(Text15)
   CIXI% = CODINT%(CODE$)
   If CIXI% < 1 Then Exit Sub
   '
   Text1 = DESCRIT$(CIXI%)
   Text3 = CVDAT(CVINT(VALOBADA("MASTER", "FECHALTA_MO", "Codint = " & CIXI%, "NOMESS")))
   '
   MSF1.Rows = 1
   MSF1.Redraw = False
   SQLX$ = "SELECT * FROM  TIEMOBRA WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE CODICONJ = " & CIXI%
   SQLX$ = SQLX$ + " AND MARBORRA = 0"
   SQLX$ = SQLX$ + " AND NUORD > 0"
   SQLX$ = SQLX$ + " ORDER BY NUORD"
   Set RST = READSET(SQLX$)
   With RST
      Do While Not .EOF
        FILA& = FILA& + 1
        MSF1.Rows = MSF1.Rows + 1
        NUOR% = XVALO(!NUORD)
        DESCRI$ = SAFETEXT$(!Descripcion)
        Label7 = DESCRI$

        MSF1.RowHeight(FILA&) = Label7.Height + 100

        TIEM% = XVALO(!TIEMPOUNITARIO)
        SECCI = XVALO(!SecciPro)
        SECTOR$ = TRIM$(ECOARCH$("SECPROD", MKS$(SECCI)))
        CANTI = XVALO(!CantiOps)
        ACUMUTOT = ACUMUTOT + CANTI
        FECHREV = FECHATEX$(CVINT(!FECHAMODIF))
        '
        MSF1.Rows = FILA& + 1
        MSF1.TextMatrix(FILA&, 1) = SAFETEXT$(NUOR%)
        MSF1.TextMatrix(FILA&, 2) = DESCRI$
        MSF1.TextMatrix(FILA&, 3) = SAFETEXT$(TIEM%)
        MSF1.TextMatrix(FILA&, 4) = SAFETEXT$(CANTI)
        MSF1.TextMatrix(FILA&, 5) = SAFETEXT$(SECCI)
        MSF1.TextMatrix(FILA&, 6) = SECTOR$
        TIME_TOTAL = TIEM% * CANTI
        MSF1.TextMatrix(FILA&, 7) = TIME_TOTAL
        MSF1.TextMatrix(FILA&, 8) = SAFETEXT$(FECHREV)
     .MoveNext
     Loop
     End With
     RST.Close
     Set RST = Nothing
     MSF1.Redraw = True
     Call COLOREAR_GRILLASQL(MSF1, &HEFF1F1)
     Call ACTUALIZAR_DATOS
End Sub

Private Sub Text161_Change()
   Text12 = FORMATNUM$(XVALO(Text161) / 60, "F12.2")
End Sub

Private Sub Text9_DblClick()

   Call Command2_Click

End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     Call Command2_Click
  End If

End Sub

Private Sub TXTCANTIDAD_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 Then
       If TRIM$(TXTDESCRIPCION.TEXT) = "" Then
           Call COMUNI("Falta Ingresar Descripción")
           Exit Sub
       End If
      '
        If XVALO(TXTCANTIDAD.TEXT) < 0.05 Then
           Call COMUNI("Falta Ingresar Cantidad")
           Exit Sub
        End If
        '
        If XVALO(TXTTIEMPO) < 0.005 Then
            Call COMUNI("Falta Ingresar Tiempo")
            Exit Sub
        End If
        '
        If XVALO(TXTNOPER) < 0.005 Then
            Call COMUNI("Falta Ingresar Sector")
            Exit Sub
        End If
        '
        If XVALO(TXTTIEMPO) > 32767 Or XVALO(TXTTIEMPO) < -32767 Then
           Call COMUNI("Valor Ingresado No Válido")
           Exit Sub
        End If
        '
        If XVALO(TXTCANTIDAD) > 32767 Or XVALO(TXTCANTIDAD) < -32767 Then
           Call COMUNI("Valor Ingresado No Válido")
           Exit Sub
        End If

        Call cargar_datos
        Call LIMIA_INFO
        On Error Resume Next
        TXTDESCRIPCION.SetFocus
        On Error GoTo 0
   End If

End Sub
Private Sub cargar_datos()
    '
    Label7 = ""

    MSF1.Cols = 9
    MSF1.Rows = MSF1.Rows + 1
    MSF1.Row = MSF1.Rows - 1
    DESCRI$ = TXTDESCRIPCION.TEXT
    Call FRATEXTO(DESCRI$, 69)
    For JJ& = 1 To CARETEX%
      TXTX$ = TXTX$ + RETEX$(JJ&) + vbCrLf
    Next JJ&
    Label7.Height = 250
    Label7 = TXTX$
    MSF1.RowHeight(MSF1.Rows - 1) = Label7.Height + 200
    MSF1.TextMatrix(MSF1.Rows - 1, 1) = XVALO(TXTCANTOP) + 1
    MSF1.TextMatrix(MSF1.Rows - 1, 2) = TXTX$
    MSF1.TextMatrix(MSF1.Rows - 1, 3) = XVALO(TXTTIEMPO)
    MSF1.TextMatrix(MSF1.Rows - 1, 4) = XVALO(TXTCANTIDAD)
    MSF1.TextMatrix(MSF1.Rows - 1, 5) = TRIM$(TXTNOPER)
    MSF1.TextMatrix(MSF1.Rows - 1, 6) = SAFETEXT$(TXTSECCION)
    TIME_TOTAL = TIME_TOTAL + (XVALO(TXTTIEMPO) * XVALO(TXTCANTIDAD))
    MSF1.TextMatrix(MSF1.Rows - 1, 7) = TIME_TOTAL
    
    MSF1.TextMatrix(MSF1.Rows - 1, 8) = TXTMODIFICADO()
    Call COLOREAR_GRILLASQL(MSF1, &HEFF1F1)
    Call ACTUALIZAR_DATOS
    '
    On Error Resume Next
    VISIROWS = (MSF1.Height / MSF1.RowHeight(j&)) - 4
    If Err > 0 Then
      VISIROWS = 1
    End If
    On Error GoTo 0
    If VISIROWS > 0 Then
       If MSF1.TopRow < MSF1.Row - VISIROWS Then
          MSF1.TopRow = MSF1.Row - VISIROWS
       End If
    End If
'    Label7.Height = 195
End Sub

 Sub ACTUALIZAR_DATOS()
   '
   'ACTUALIZAR EL TOTAL DEL TIEMPO
   TOTALTIEMPO% = 0
   For I& = 1 To MSF1.Rows - 1
     TIEM% = XVALO(MSF1.TextMatrix(I&, 3))
     CANTI = XVALO(MSF1.TextMatrix(I&, 4))
     TOTALTIEMPO% = TOTALTIEMPO% + (TIEM% * CANTI)
     '
     'CONTROLO QUE EL ORDEN SE MANTENGA
     MSF1.TextMatrix(I&, 1) = I&
   Next I&
   Text161 = FORMATNUM$(TOTALTIEMPO%, "F12.0")
  
  'ACTUALIZAR LA CANTIDAD DE OPERACIONES
   TXTCANTOP = I& - 1
End Sub

Private Sub TXTDESCRIPCION_Change()
   Label7 = TXTDESCRIPCION
End Sub

Private Sub TXTNOPER_Change()
      CS1 = XVALO(TXTNOPER.TEXT)
      If CS1 > 0 Then
         TXTSECCION = TRIM$(ECOARCH$("SECPROD", MKS$(CS1)))
      Else
         TXTSECCION = ""
      End If

End Sub

Sub LIMIA_INFO()
  TXTCANTIDAD = ""
  TXTDESCRIPCION = ""
  TXTNOPER = ""
  TXTSECCION = ""
  TXTTIEMPO = ""
  
End Sub

Private Sub TXTNOPER_DblClick()
     Call Command2_Click
End Sub

Private Sub TXTNOPER_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     Call Command2_Click
  End If

End Sub

Private Sub TXTSECCION_DblClick()
   Call Command2_Click
End Sub

Private Sub TXTSECCION_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     Call Command2_Click
  End If
End Sub
Sub GRABATIEMPMOBRA(CI1%)
   CONDI$ = "CODICONJ = " & CI1%
   '
   'SI ES LA PRIMERAVEZ QUE SE GENERA LA LISTA LE AGREGO EN EL MASTER LA FECHA DE CREACION
   FEHO% = HOY(HO$)
   If CantRegistros("TIEMOBRA", CONDI$, "NOMESS") < 1 Then
      Call ACTUREGISTRO("MASTER", "FECHALTA_MO", "Codint = " & CI1%, FETEXCOR1$(FEHO%), REGAF&, "NOMESS")
   End If
   '
   'MARCO LOS REGISTROS CON MARBORRA
   Call ACTUREGISTRO("TIEMOBRA", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
   '
   SQLX$ = "SELECT * FROM TIEMOBRA "
   SQLX$ = SQLX$ + " WHERE CODICONJ < 1"
   '
   Set RST = New ADODB.Recordset
   RST.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   
   With RST
       For I& = 1 To MSF1.Rows - 1
        NUOR% = TRIM$(MSF1.TextMatrix(I&, 1))
        DESCRIX$ = MSF1.TextMatrix(I&, 2)
        TIEM% = XVALO(MSF1.TextMatrix(I&, 3))
        CANTI = XVALO(MSF1.TextMatrix(I&, 4))
        CODSECCION = XVALO(MSF1.TextMatrix(I&, 5))
        FEMODI% = CVINT(MSF1.TextMatrix(I&, 8))
        '
        .AddNew
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
        '
        !CODICONJ = CI1%
        !NUORD = NUOR%
        !Descripcion = DESCRIX$
        !TIEMPOUNITARIO = TIEM%
        !CantiOps = CANTI
        !SecciPro = CODSECCION
        !FECHAMODIF = FETEXCOR1$(FEMODI%)
        !MARBORRA = 0
        .Update
       Next I&
    End With
    RST.Close
    Set RST = Nothing
    TIEMPX = FORMATNUM$(XVALO(Text161) / 60, "F12.2")
    Call ACTUREGISTRO("MASTER", "MANO_OBRA_REAL", "CODINT = " & CI1%, TIEMPX, REGAF&, "NOMESS")
    '
    Call BORRAREGISTROS("TIEMOBRA", CONDI$ & " AND MARBORRA = 1", REGAF&, "NOMESS")
    '
End Sub

Private Sub TXTTIEMPO_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     On Error Resume Next
     Command2.SetFocus
     On Error GoTo 0
   End If
End Sub
Private Sub MSF1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then ' BUSQUEDA  POR TEXTO CON F3
       If ARBUS$ <> "" Then
          VUELTA% = 0
          j& = MSF1.Row

          H& = 0: If MSF1.Col < MSF1.Cols Then H& = MSF1.Col + 1
10        For U& = j& To MSF1.Rows - 1
             HI& = H&: If U& > j& Then HI& = 0
             For I& = HI& To MSF1.Cols - 1
              If InStr(UCase$(MSF1.TextMatrix(U&, I&)), ABUSQUE$) > 0 Then
                MSF1.Row = U&: MSF1.Col = I&
                On Error Resume Next
                 VISIROWS = (MSF1.Height / MSF1.RowHeight(j&)) - 4
                If Err > 0 Then
                 VISIROWS = 1
                End If
                On Error GoTo 0
                 If VISIROWS > 0 Then
                    If MSF1.TopRow < MSF1.Row - VISIROWS Then
                      MSF1.TopRow = MSF1.Row - VISIROWS
                    End If
                 End If
                MSF1.SetFocus
                Exit Sub
              End If
              Next I&
          Next U&
          '
          Call MENSERR(24, "Argumento no Encontrado.")
       End If
   End If

End Sub


Sub CARSECPROD()
    '
    Call BLANARCH("!SECPROD")
    j& = 0
    For U& = 1 To ULTREG&("SECPROD")
        x$ = REGLEIDO$("SECPROD", U&)
        COSECC = CVS(Left$(x$, 4))
        If COSECC >= 1 Then
            Call REGAPP("!SECPROD", x$)
        End If
    Next U&
    Call CIERRARCH("SECPROD")
    Call CIERRARCH("!SECPROD")
    '
    VTB% = VENTABU%("SECPROD")
    If VTB% < 0 Then Exit Sub
    '
    j& = 0
    For U& = 1 To ULTREG&("!SECPROD")
        x$ = REGLEIDO$("!SECPROD", U&)
        COSECC = CVS(Left$(x$, 4))
        If COSECC >= 1 Then
            j& = j& + 1
            Call GRAREG("SECPROD", x$, j&)
        End If
    Next U&
    '
    Call SETULTREG("SECPROD", j&)
    Call CIERRARCH("SECPROD")
    Call CIERRARCH("!SECPROD")
    '
End Sub
