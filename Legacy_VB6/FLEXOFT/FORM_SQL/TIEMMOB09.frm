VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form TIEMMOB09 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "A.B.M. DE OPERACIONES"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   15465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   15465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox TXTFECHAVISTO 
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
      Height          =   330
      Left            =   13140
      Locked          =   -1  'True
      TabIndex        =   52
      TabStop         =   0   'False
      ToolTipText     =   "Fecha Visto - Revisión "
      Top             =   600
      Width           =   1120
   End
   Begin VB.TextBox text17 
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
      Index           =   0
      Left            =   12360
      Locked          =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Text            =   "F.Visto:"
      ToolTipText     =   "Fecha Visto - Revisión "
      Top             =   600
      Width           =   795
   End
   Begin VB.TextBox UsuarioModifico 
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
      Left            =   10740
      Locked          =   -1  'True
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   2550
      Width           =   3435
   End
   Begin VB.TextBox Text13 
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
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   49
      TabStop         =   0   'False
      Text            =   "Responsable Última Modificación:"
      Top             =   2550
      Width           =   3050
   End
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
      TabIndex        =   46
      ToolTipText     =   "Presenta Lista de Selección (F2)"
      Top             =   2560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   5655
      Left            =   120
      TabIndex        =   40
      ToolTipText     =   "Doble Click Editar/Eliminar - Boton Derecho Mover Filas"
      Top             =   2805
      Width           =   14145
      _ExtentX        =   24950
      _ExtentY        =   9975
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EFF1F1&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      ScaleHeight     =   1995
      ScaleWidth      =   7215
      TabIndex        =   41
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
         Height          =   225
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Visible         =   0   'False
         Width           =   7215
         WordWrap        =   -1  'True
      End
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
      TabIndex        =   44
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
      TabIndex        =   43
      TabStop         =   0   'False
      Text            =   "Total Tiempo en Minutos:"
      Top             =   8580
      Width           =   2595
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
      Left            =   10140
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Text            =   "Total Tiempo en Segundos:"
      Top             =   8580
      Width           =   2595
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   8640
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   33
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
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Text            =   "F.Alta:"
      Top             =   600
      Width           =   795
   End
   Begin VB.TextBox Text3 
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
      Height          =   330
      Left            =   11220
      Locked          =   -1  'True
      TabIndex        =   29
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   600
      Width           =   7095
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
      TabIndex        =   15
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox text17 
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
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "OPERACIONES Y TIEMPOS DE MANOS DE OBRA"
      Top             =   0
      Width           =   14220
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
      Left            =   12720
      Locked          =   -1  'True
      TabIndex        =   12
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   8520
      Width           =   14175
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
      Height          =   9075
      Left            =   14280
      ScaleHeight     =   9015
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
         Height          =   650
         Left            =   240
         Picture         =   "TIEMMOB09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Tabla de Sectores"
         Top             =   5500
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
         Left            =   240
         Picture         =   "TIEMMOB09.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Imprime Tiempos de Mano de Obra del Item Activo"
         Top             =   4600
         UseMaskColor    =   -1  'True
         Width           =   650
      End
      Begin VB.CommandButton Command5 
         Height          =   650
         Left            =   240
         Picture         =   "TIEMMOB09.frx":07B4
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Búsqueda por Texto "
         Top             =   3700
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
         TabIndex        =   36
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
         Height          =   650
         Left            =   240
         Picture         =   "TIEMMOB09.frx":1E24
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Acceder a Historial de Operaciones x Artículo"
         Top             =   2800
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
         Height          =   650
         Left            =   240
         Picture         =   "TIEMMOB09.frx":212E
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Copiar Tiempos del Código Activo Hacia Otro."
         Top             =   1900
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
         Height          =   650
         Left            =   240
         Picture         =   "TIEMMOB09.frx":2278
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   1000
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
      TabIndex        =   17
      Top             =   960
      Width           =   14265
      Begin VB.CommandButton Command10 
         Caption         =   "Selección "
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
         Left            =   9480
         MaskColor       =   &H0084962E&
         TabIndex        =   48
         Top             =   960
         Width           =   1800
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Importar desde Excel"
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
         Left            =   7560
         MaskColor       =   &H0084962E&
         TabIndex        =   47
         Top             =   960
         Width           =   1800
      End
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
         Left            =   12300
         MaskColor       =   &H0084962E&
         TabIndex        =   4
         ToolTipText     =   "Click Aquì Para Ingresar la Operación a la Lista."
         Top             =   960
         Width           =   1800
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
         TabIndex        =   28
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
         Left            =   13070
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "Cant.Op."
         Top             =   270
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
         Height          =   405
         Left            =   13070
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   510
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
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "Sector    /   N°"
         ToolTipText     =   "Doble Click Para Seleccionar Sección"
         Top             =   270
         Width           =   2160
      End
      Begin VB.TextBox Text7 
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
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "Tiem. Unit.''"
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "Modificado"
         Top             =   270
         Width           =   1095
      End
      Begin VB.TextBox TXTMODIFICADO 
         Alignment       =   2  'Center
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
         TabIndex        =   19
         Top             =   540
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Check1"
         Height          =   255
         Left            =   12400
         TabIndex        =   18
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
         TabIndex        =   21
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
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   195
      Left            =   960
      TabIndex        =   9
      Top             =   9200
      Width           =   870
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
      Begin VB.Menu mnuCopiarOperaciones 
         Caption         =   "Copiar Operaciones (Selección de Código Origen)"
      End
   End
End
Attribute VB_Name = "TIEMMOB09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARBUS$, ABUSQUE$
Public MODIFICATIEMMOBRA%
Sub ACOMODARMSF()
     MSF1.RowHeight(MSF1.Row) = Label7.Height + 100
     MSF1.TextMatrix(MSF1.Row, 2) = Label7
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

Private Sub Command1_Click()
  CI1% = CODINT%(Text15)
  If CI1% > 0 Then
     If MODIFICATIEMMOBRA% > 0 Then
       If COMALTER%("Se ha Modificado Datos del Artículo\'" + TRIM$(CODEXT$(CI1%)) + " - " + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") = 1 Then
           Call Command6_Click
           CI1% = CODINT%(ANACOSI09.Text1)
           ANACOSI09.Text4 = FORMATNUM$(MANO_OBRA_REAL(CI1%), "F6.2") ' MANO DE OBRA REAL
       End If
       MODIFICATIEMMOBRA% = 0
     End If
   End If

  
  Unload Me
End Sub

Private Sub Command14_Click()

End Sub

Private Sub Command10_Click()
   CONTENED.TXTCONTENEDOR = ""
   Command10.Enabled = False
   For i& = 1 To MSF1.Rows - 1
     For J& = 2 To MSF1.Cols - 1
         If J& <> 3 Then
          textox$ = textox$ + MSF1.TextMatrix(i&, J&) + Chr$(9)
         End If
     Next J&
     Debug.Print textox$
     ACUMULATEXTO$ = ACUMULATEXTO$ & RTrim$(textox$) & Chr$(13)
     textox$ = ""
   Next i&
   CONTENED.TXTCONTENEDOR = ACUMULATEXTO$
   CI1% = CODINT%(Me.Text15)
   CONTENED.TXTRUTA = RUTA_ARCHIVO_PPF$(CI1%)
   'CONTENED.Command1.Caption = "Selección y Abrir Excel"
   CONTENED.Show 1
   
   Command10.Enabled = True
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
   If CI2% = CI1% Then
      Call COMUNI("Código Origien Igual a Código Destino")
      GoTo 10
   End If
   
   If CantRegistros("TIEMOBRA", "CODICONJ = " & CI2%, "NOMESS") > 0 Then
      OPX% = COMALTER%("Este Código Posee Tiempos de Mano de Obra Cargados\\Mantener Actual\Sobre-escribir")
      If OPX% < 2 Then GoTo 10
   End If
   Call GRABATIEMPMOBRA(CI2%)
   Text15 = CODEXT$(CI2%)
   '
99 Command4.Enabled = True
End Sub

Private Sub Command5_Click()
 Command5.Enabled = False
 ABUSQUE$ = UCase$(TRIM$(InputBox$("Busqueda por Texto", , ARBUS$)))
   If ABUSQUE$ <> "" Then
     ARBUS$ = ABUSQUE$
     For U& = 1 To MSF1.Rows - 1
      For i& = 0 To MSF1.Cols - 1
       If InStr(UCase$(MSF1.TextMatrix(U&, i&)), ABUSQUE$) > 0 Then
        MSF1.Row = U&: MSF1.COL = i&
        If MSF1.TopRow > MSF1.Row Then MSF1.TopRow = MSF1.Row
        On Error Resume Next
        VISIROWS = (MSF1.Height / MSF1.RowHeight(J&)) - 4
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
       Next i&
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
   MODIFICATIEMMOBRA% = 0
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
   On Error Resume Next
   BUSCA_DES.TXTBUSCAR.SetFocus
   On Error GoTo 0

   BUSCA_DES.Show 1
   On Error Resume Next
   Command3.SetFocus
   On Error GoTo 0
   
End Sub

Private Sub Command8_KeyDown(KeyCode As Integer, Shift As Integer)
  Call Command8_Click
End Sub

Private Sub Command9_Click()
      CI1% = CODINT%(Me.Text15)
      
      CONTENED.TXTRUTA = RUTA_ARCHIVO_PPF$(CI1%)
      CONTENED.Show 1
      
      If CONTENED.Label1 = "OK" Then
        MSF1.Rows = 1
        With CONTENED
         Fila = Split(.TXTCONTENEDOR, Chr(13))
         For L& = LBound(Fila) To UBound(Fila)
          FILAXX = Fila(L&)
          CANTCAR% = Len(FILAXX)
          strArray1 = Split(FILAXX, Chr(9))
          For i& = LBound(strArray1) To UBound(strArray1)
            a1$ = TRIM(strArray1(i&))
            J& = J& + 1
            Select Case J&
             Case 1
               Me.TXTDESCRIPCION = a1$
             Case 2
               Me.TXTNOPER = FUN_NUMSECTOR(a1$) 'SECTOR
             Case 3
               Me.TXTCANTIDAD = a1$ '  CANTIDAD
             Case 4
               Me.TXTTIEMPO = a1$  ' SEGUNDOS
             Case 5

             Case 6
               Call Command3_Click 'INGRESA LA INFORMACION
               J& = 0
             End Select
            
         Next i&
        Next L&
      
        End With
      End If

End Sub
Function RUTA_ARCHIVO_PPF$(CI1%)
      SQLX$ = "SELECT * FROM DOCUREL WITH(NOLOCK) WHERE COD_INTE = " & CI1%
      Dim DOCTMP As ADODB.Recordset
      Set DOCTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
      With DOCTMP
        Do While Not .EOF
           DESCRI$ = SAFETEXT$(!Descripcion)
           If Left$(UCase$(DESCRI$), 3) = "PPF" Or Left$(UCase$(DESCRI$), 2) = "PF" Then
               RUTA_ARCHIVO_PPF$ = SAFETEXT$(!Ruta$)
               Exit Do
           End If
          .MoveNext
        Loop
      End With
End Function
Function FUN_NUMSECTOR(TXTX$)
    'LEE EL TEXTO ENTRE PARENTESIS QUE TIENE LA DESCRIPCION
    'EJMPLO: (A) AGRAFADO PARA DEJAR COMO ESTA EN EL EXCEL Y PODER RECUPERAR EL NUMERO DEL SECTOR SEGUN EL FORMATO ORIGINAL DE LA TABLA
    
    TEXPARAM$ = TRIM$(UCase$(TXTX$))
    FUN_NUMSECTOR = 0
    For U& = 1 To ULTREG&("!SECPROD")
        X$ = REGLEIDO$("!SECPROD", U&)
        COSECC = CVS(Left$(X$, 4))
        If COSECC >= 1 Then
            J& = J& + 1
            DESC$ = Mid$(X$, 5, 32)
            POS1% = InStr(DESC$, "(")
          
            If POS1% > 0 Then
              DatoLeido$ = TRIM$(UCase$(Mid$(DESC$, POS1% + 1)))
              POS2% = InStr(DatoLeido$, ")")
              If POS2% > 1 Then
                DatoLeido$ = TRIM$(UCase$(Mid$(DatoLeido$, 1, POS2% - 1)))
'                'SI EL DATO ES DE MAS DE UN CARACTER VALIDA SOLO LOS DOS PRIMEROS
'                If Len(DATOLEIDO$) > 1 Then DATOLEIDO$ = Mid$(DATOLEIDO$, 1, 2)
                If TEXPARAM$ = DatoLeido$ Then
                 FUN_NUMSECTOR = COSECC
                 Exit For
                End If
               
              End If
            End If
            
        End If
    Next U&
    '
    Call CIERRARCH("SECPROD")
    Call CIERRARCH("!SECPROD")

End Function
Private Sub Form_Load()
   Call APLICACIONSKINS(Me)
   '
   TEPRUEBA.Font = MSF1.Font
   TEPRUEBA.FontSize = MSF1.FontSize
   MSF1.Cols = 10
   NUCOLU% = 1: TEPRUEBA = "9": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = " "
   NUCOLU% = 2: TEPRUEBA = "9999": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = "Ord"
   NUCOLU% = 3: TEPRUEBA = "123456789*123456789*123456789*123456789*123456789*123456789*123456789*12": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = "Descripción"
   NUCOLU% = 4: TEPRUEBA = "": MSF1.ColWidth(NUCOLU% - 1) = 0: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = ""
   NUCOLU% = 5: TEPRUEBA = "SECCIONSECCAAAAAA": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = "Sector"
   NUCOLU% = 6: TEPRUEBA = "1234567": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = "Qt."
   NUCOLU% = 7: TEPRUEBA = "1234567": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = "T.Un"
   NUCOLU% = 8: TEPRUEBA = "12345678": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = "T.Tot."
   NUCOLU% = 9: TEPRUEBA = "01'02¡1234": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = "X hs."
   NUCOLU% = 10: TEPRUEBA = "01'02¡1234": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 3: MSF1.TextMatrix(0, NUCOLU% - 1) = "Modif."
   '
   MSF1.MergeCells = flexMergeRestrictColumns
   MSF1.MergeCol(1) = True

   Call CREACAMPO("", "TIEMOBRA", "NUORD", 3, 0, 1)
   Call CREACAMPO("", "TIEMOBRA", "CODICONJ", 3, 0, 1)
   Call CREACAMPO("", "TIEMOBRA", "DESCRIPCION", 10, 1280)
   Call CREACAMPO("", "TIEMOBRA", "TIEMPOUNITARIO", 7, 0)
   Call CREACAMPO("", "TIEMOBRA", "SecciPro", 6, 4)
   Call CREACAMPO("", "TIEMOBRA", "CantiOps", 3, 0)
   Call CREACAMPO("", "TIEMOBRA", "FECHAMODIF", 8, 0)
   Call CREACAMPO("", "TIEMOBRA", "MARBORRA", 3, 0)
   Call CREACAMPO("", "TIEMOBRA", "Usua_Mod", 3, 2)
   Call CREACAMPO("", "TIEMOBRA", "FECHA_VISTO", 8, 0)
   Call CREACAMPO("", "TIEMOBRA", "NuordString", 10, 8, 1)
   
   '
   Call CREACAMPO("ARTICULOS", "MASTER", "FECHALTA_MO", 8, 0)
   '
   Call CREACAMPO("", "HISTOMOBRA", "CODICONJ", 3, 0, 1)
   Call CREACAMPO("", "HISTOMOBRA", "ANOTMOBRA", 12, 0)
   Call CREACAMPO("", "HISTOMOBRA", "MARBORRA", 3, 0)
   '
'   TXTMODIFICADO = FECHATEX$(HOY(HO$)) SE CARGA EN EL CHANGE DEL TXT DE CODIGO
   '
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If Cancel = 0 Then
      Command1_Click
  End If
End Sub


Private Sub mnuCopiarOperaciones_Click()
    If DERACCE%("COPOPERA", "Copiar Operaciones de Un Código a Otro") < 2 Then Exit Sub
    If TRIM$(Text15) = "" Then
       Call COMUNI("Debe Tener Previamente Seleccionado el Código de Destino")
       Exit Sub
    End If
    
    CODE$ = ListaSeleccionArticulos(Descripx$)
    CIXI% = CODINT%(CODE$)
    If CIXI% > 0 Then
            GRILLAGRAL.GRID.MergeCells = flexMergeRestrictColumns
            GRILLAGRAL.GRID.MergeCol(1) = True
            GRILLAGRAL.MSF_2.MergeCells = flexMergeRestrictColumns
            GRILLAGRAL.MSF_2.MergeCol(1) = True

20         Campos$ = "Ord/F4;Descripción/A32;Cod/F4;Sector/A16;Qt./F8.1;T.Un/F7;T.Tot./F9;X hs./A10;Modif./f10"
           Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
           Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
           GRILLAGRAL.Caption = "Origen de Operaciones de " & Descripx$
           GRILLAGRAL.mnuMenu1.Caption = "Copiar Operaciones Seleccionadas"
           GRILLAGRAL.mnuMenu1.Visible = True
           GRILLAGRAL.GRID.ZOrder 0
           GRILLAGRAL.TXTBUSCAR = ""
           '
           Text3 = FECHATEX$(CVINT(VALOBADA("MASTER", "FECHALTA_MO", "Codint = " & CIXI%, "NOMESS")))
           '
           ULTIMAFEMOD% = 0
           GRILLAGRAL.GRID.Rows = 1
           GRILLAGRAL.GRID.Redraw = False
           SQLX$ = "SELECT * FROM  TIEMOBRA WITH(NOLOCK) "
           SQLX$ = SQLX$ + " WHERE CODICONJ = " & CIXI%
           SQLX$ = SQLX$ + " AND MARBORRA = 0"
           SQLX$ = SQLX$ + " AND NUORD > 0"
           SQLX$ = SQLX$ + " ORDER BY NUORD"
           Set Rst = READSET(SQLX$)
           With Rst
              Do While Not .EOF
                Fila& = Fila& + 1
                If Fila& = 1 Then
                  UsuarioModifico = USERNAM$(XVALO(!Usua_Mod))
                  TXTFECHAVISTO = FECHATEX$(CVINT(!FECHA_VISTO))
                End If
                GRILLAGRAL.GRID.Rows = GRILLAGRAL.GRID.Rows + 1
                NUOR% = XVALO(!NUORD)
                DESCRI$ = SAFETEXT$(!Descripcion)
                Label7 = DESCRI$
                
                GRILLAGRAL.GRID.RowHeight(Fila&) = Label7.Height + 100
        
                TIEM1 = XVALO(!TIEMPOUNITARIO)
                SECCI = XVALO(!SecciPro)
                SECTOR$ = TRIM$(ECOARCH$("SECPROD", MKS$(SECCI)))
                CANTI = XVALO(!CantiOps)
                ACUMUTOT = ACUMUTOT + CANTI
                FEMOD% = CVINT(!FECHAMODIF)
                If ULTIMAFEMOD% < FEMOD% Then ULTIMAFEMOD% = FEMOD%
                FECHREV = FECHATEX$(FEMOD%)
                '
                GRILLAGRAL.GRID.Rows = Fila& + 1
                GRILLAGRAL.GRID.TextMatrix(Fila&, 1) = SAFETEXT$(NUOR%)
                GRILLAGRAL.GRID.TextMatrix(Fila&, 2) = DESCRI$
                GRILLAGRAL.GRID.TextMatrix(Fila&, 3) = SAFETEXT$(SECCI)
                GRILLAGRAL.GRID.TextMatrix(Fila&, 4) = SECTOR$
                GRILLAGRAL.GRID.TextMatrix(Fila&, 5) = SAFETEXT$(CANTI)
                GRILLAGRAL.GRID.TextMatrix(Fila&, 6) = TRIM$(FORMATNUM$((TIEM1), "F10.2"))
                TIME_TOTAL = TIEM1 * CANTI
                GRILLAGRAL.GRID.TextMatrix(Fila&, 7) = TIME_TOTAL
                GRILLAGRAL.GRID.TextMatrix(Fila&, 8) = 0
'                If Abs(XVALO(TIEM1)) > 0 Then GRILLAGRAL.GRID.TextMatrix(FILA&, 8) = TRIM$(FORMATNUM$(3600 / XVALO(TIEM1), "F10.0"))
                'jorge pidio cambiar  3600 / TIEMPO TOTLA EN VEZ DE T.UNITARIO 29-06-21
                If Abs(XVALO(TIME_TOTAL)) > 0 Then GRILLAGRAL.GRID.TextMatrix(Fila&, 8) = TRIM$(FORMATNUM$(3600 / XVALO(TIME_TOTAL), "F10.0"))
                GRILLAGRAL.GRID.TextMatrix(Fila&, 9) = SAFETEXT$(FECHREV)
                
             .MoveNext
             Loop
        
             End With
             
             Rst.Close
             Set Rst = Nothing
             GRILLAGRAL.GRID.Redraw = True
             TXTMODIFICADO = FECHATEX$(ULTIMAFEMOD%)
60           GRILLAGRAL.Show 1
             
             If GRILLAGRAL.MENU1 > 0 Then
                Dim MSF As msFlexGrid
                Set MSF = GRILLAGRAL.GRID
                If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
                H& = 0
                For i& = 1 To MSF.Rows - 1
                  If MSF.TextMatrix(i&, 0) = "*" Then
                     H& = H& + 1
                     If H& = 1 Then
                       MSF1.Rows = 1 ' SI ES LA PRIMERA VEZ QUE PASA LIMPIA LA GRILLA DESTINO
                     End If
                     MSF1.Rows = H& + 1
                     MSF1.TextMatrix(H&, 1) = MSF.TextMatrix(i&, 1)
                     MSF1.TextMatrix(H&, 2) = MSF.TextMatrix(i&, 2)
                     MSF1.TextMatrix(H&, 3) = MSF.TextMatrix(i&, 3)
                     MSF1.TextMatrix(H&, 4) = MSF.TextMatrix(i&, 4)
                     MSF1.TextMatrix(H&, 5) = MSF.TextMatrix(i&, 5)
                     MSF1.TextMatrix(H&, 6) = MSF.TextMatrix(i&, 6)
                     MSF1.TextMatrix(H&, 7) = MSF.TextMatrix(i&, 7)
                     MSF1.TextMatrix(H&, 8) = MSF.TextMatrix(i&, 8)
                     MSF1.TextMatrix(H&, 9) = MSF.TextMatrix(i&, 9)
                  End If
                Next
                If H& < 1 Then
                   OPX% = COMALTER%("Atención !!! No se seleccionó Ninguna Operación\\Cancelar\Volver y Seleccionar")
                   If OPX% = 2 Then GoTo 60
                End If
                Call numerargrilla(MSF1, 1)
                Call ACTUALIZAR_DATOS
                GRILLAGRAL.MENU1 = 0
             End If
             
     End If

            
End Sub

Private Sub MSF1_DblClick()
   If MSF1.Row >= 1 Then
    With EDITMOBRA
         .TXTDESCRIPCION.TEXT = MSF1.TextMatrix(MSF1.Row, 2)
         .TXTNOPER.TEXT = MSF1.TextMatrix(MSF1.Row, 3)
         .TXTSECCION.TEXT = MSF1.TextMatrix(MSF1.Row, 4)
         .TXTCANTIDAD.TEXT = MSF1.TextMatrix(MSF1.Row, 5)
         .TXTTIEMPO.TEXT = MSF1.TextMatrix(MSF1.Row, 6)
         .TXTMODIFICADO.TEXT = MSF1.TextMatrix(MSF1.Row, 9)
         .TXTCANTOP.TEXT = TXTCANTOP
         .Show 1
     End With
   End If
End Sub

Private Sub MSF1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
      REG& = MSF1.MouseRow
      If REG& >= 1 Then
          MSF1.RowSel = REG&
          Campos$ = "ORD/F4;Descripción/A24;AA/A0;Sector/A16;Qt./F7;T.Un/F7;T.Tot./F8;X hs./A11;Modif/A11"
          Call INSERTARFILA(MSF1, REG&, TIEMMOB09, Campos$)
          Call numerargrilla(TIEMMOB09.MSF1, 1)
          Call ACTUALIZAR_DATOS
      End If
    End If
    
End Sub

Private Sub Text15_Change()
   CODE$ = TRIM$(Text15)
   CIXI% = CODINT%(CODE$)
   If CIXI% < 1 Then Exit Sub
   '
   Text1 = DESCRIT$(CIXI%)
   Text3 = FECHATEX$(CVINT(VALOBADA("MASTER", "FECHALTA_MO", "Codint = " & CIXI%, "NOMESS")))
   '
   ULTIMAFEMOD% = 0
   MSF1.Rows = 1
   MSF1.Redraw = False
   SQLX$ = "SELECT * FROM  TIEMOBRA WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE CODICONJ = " & CIXI%
   SQLX$ = SQLX$ + " AND MARBORRA = 0"
   SQLX$ = SQLX$ + " AND NUORD > 0"
   SQLX$ = SQLX$ + " ORDER BY NUORD"
   Set Rst = READSET(SQLX$)
   With Rst
      Do While Not .EOF
        Fila& = Fila& + 1
        If Fila& = 1 Then
          UsuarioModifico = USERNAM$(XVALO(!Usua_Mod))
          TXTFECHAVISTO = FECHATEX$(CVINT(!FECHA_VISTO))
        End If
        MSF1.Rows = MSF1.Rows + 1
        NUOR% = XVALO(!NUORD)
        DESCRI$ = SAFETEXT$(!Descripcion)
        Label7 = DESCRI$
        
        MSF1.RowHeight(Fila&) = Label7.Height + 100

        TIEM1 = XVALO(!TIEMPOUNITARIO)
        SECCI = XVALO(!SecciPro)
        SECTOR$ = TRIM$(ECOARCH$("SECPROD", MKS$(SECCI)))
        CANTI = XVALO(!CantiOps)
        ACUMUTOT = ACUMUTOT + CANTI
        FEMOD% = CVINT(!FECHAMODIF)
        If ULTIMAFEMOD% < FEMOD% Then ULTIMAFEMOD% = FEMOD%
        FECHREV = FECHATEX$(FEMOD%)
        '
        MSF1.Rows = Fila& + 1
        MSF1.TextMatrix(Fila&, 1) = SAFETEXT$(NUOR%)
        MSF1.TextMatrix(Fila&, 2) = DESCRI$
        MSF1.TextMatrix(Fila&, 3) = SAFETEXT$(SECCI)
        MSF1.TextMatrix(Fila&, 4) = SECTOR$
        MSF1.TextMatrix(Fila&, 5) = SAFETEXT$(CANTI)
        MSF1.TextMatrix(Fila&, 6) = TRIM$(FORMATNUM$((TIEM1), "F10.2"))
        TIME_TOTAL = TIEM1 * CANTI
        MSF1.TextMatrix(Fila&, 7) = TIME_TOTAL
        MSF1.TextMatrix(Fila&, 8) = 0
        'If Abs(XVALO(TIEM1)) > 0 Then MSF1.TextMatrix(FILA&, 8) = TRIM$(FORMATNUM$(3600 / XVALO(TIEM1), "F10.0"))
        'jorge pidio cambiar  3600 / TIEMPO TOTLA EN VEZ DE T.UNITARIO 29-06-21
        If Abs(XVALO(TIME_TOTAL)) > 0 Then MSF1.TextMatrix(Fila&, 8) = TRIM$(FORMATNUM$(3600 / XVALO(TIME_TOTAL), "F10.0"))
       
        MSF1.TextMatrix(Fila&, 9) = SAFETEXT$(FECHREV)
        
     .MoveNext
     Loop

     End With
     
     Rst.Close
     Set Rst = Nothing
     MSF1.Redraw = True
     Call COLOREAR_GRILLASQL(MSF1, &HEFF1F1)
     Call ACTUALIZAR_DATOS
     TXTMODIFICADO = FECHATEX$(ULTIMAFEMOD%)
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
        '
        'VALIDACION QUE NO SE INGRESEN TEXTOS QUE INCLUYAN COMILLAS SIMPLES.
        If InStr(TXTDESCRIPCION.TEXT, "'") > 0 Then
           Call COMUNI("Se Han Detectado Caracteres Reservados (Comilla Simple)\ El Mismo Serà Reemplazado Por Un Espacio Vacío")
           TXTDESCRIPCION.TEXT = REPLACAR$(TXTDESCRIPCION.TEXT, "'", " ")
        End If
        '
        MODIFICATIEMMOBRA% = 1
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

    MSF1.Cols = 10
    MSF1.Rows = MSF1.Rows + 1
    MSF1.Row = MSF1.Rows - 1
    DESCRI$ = TXTDESCRIPCION.TEXT
'    Call FRATEXTO(DESCRI$, 69)
'    For JJ& = 1 To CARETEX%
'      TXTX$ = TXTX$ + RETEX$(JJ&) + vbCrLf
'    Next JJ&
'    Label7.Height = 100
    Label7 = DESCRI$
    MSF1.RowHeight(MSF1.Rows - 1) = Label7.Height + 100
    MSF1.TextMatrix(MSF1.Rows - 1, 1) = XVALO(TXTCANTOP) + 1
    MSF1.TextMatrix(MSF1.Rows - 1, 2) = DESCRI$
    MSF1.TextMatrix(MSF1.Rows - 1, 3) = TRIM$(TXTNOPER)
    MSF1.TextMatrix(MSF1.Rows - 1, 4) = SAFETEXT$(TXTSECCION)
    MSF1.TextMatrix(MSF1.Rows - 1, 5) = XVALO(TXTCANTIDAD)
    MSF1.TextMatrix(MSF1.Rows - 1, 6) = XVALO(TXTTIEMPO)
    TIME_TOTAL = TIME_TOTAL + (XVALO(TXTTIEMPO) * XVALO(TXTCANTIDAD))
    MSF1.TextMatrix(MSF1.Rows - 1, 7) = TIME_TOTAL
    MSF1.TextMatrix(MSF1.Rows - 1, 8) = 0
    'If Abs(XVALO(TXTTIEMPO)) > 0 Then MSF1.TextMatrix(MSF1.Rows - 1, 8) = TRIM$(FORMATNUM$(3600 / XVALO(TXTTIEMPO), "F12.0"))
    'jorge pidio cambiar  3600 / TIEMPO TOTLA EN VEZ DE T.UNITARIO 29-06-21
    If Abs(XVALO(TIME_TOTAL)) > 0 Then MSF1.TextMatrix(MSF1.Rows - 1, 8) = TRIM$(FORMATNUM$(3600 / XVALO(TIME_TOTAL), "F10.0"))
    
    MSF1.TextMatrix(MSF1.Rows - 1, 9) = FECHATEX$(HOY(HO$))
    Call COLOREAR_GRILLASQL(MSF1, &HEFF1F1)
    Call ACTUALIZAR_DATOS
    '
End Sub

 Sub ACTUALIZAR_DATOS()
   '
   'ACTUALIZAR EL TOTAL DEL TIEMPO
   TOTALTIEMPO = 0
   For i& = 1 To MSF1.Rows - 1
     TIEM = ROUND(XVALO(MSF1.TextMatrix(i&, 6)), 1)
     CANTI = XVALO(MSF1.TextMatrix(i&, 5))
     TOTALTIEMPO = TOTALTIEMPO + (TIEM * CANTI)
     '
     'CONTROLO QUE EL ORDEN SE MANTENGA
     MSF1.TextMatrix(i&, 1) = i&
   Next i&
   Text161 = FORMATNUM$(TOTALTIEMPO, "F12.0")
  
  'ACTUALIZAR LA CANTIDAD DE OPERACIONES
   TXTCANTOP = i& - 1
End Sub

Private Sub TXTDESCRIPCION_Change()
   Label7 = TXTDESCRIPCION
End Sub

Private Sub TXTDESCRIPCION_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 113 Then ' CORRESPONDE TECLA <F2>
     Call Command8_Click
   End If
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
   ULTIMAFEMOD% = 0
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
   Set Rst = New ADODB.Recordset
   Rst.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   
   With Rst
       For i& = 1 To MSF1.Rows - 1
        NUOR% = TRIM$(MSF1.TextMatrix(i&, 1))
        DESCRIX$ = MSF1.TextMatrix(i&, 2)
        DESCRIX$ = LTrim(DESCRIX$)
        DESCRIX$ = RTrim(DESCRIX$)
        TIEM1# = XVALO(MSF1.TextMatrix(i&, 6))
        CANTI = XVALO(MSF1.TextMatrix(i&, 5))
        CODSECCION = XVALO(MSF1.TextMatrix(i&, 3))
        FEMODI% = FECHANUM(MSF1.TextMatrix(i&, 9))
        If ULTIMAFEMOD% < FEMODI% Then ULTIMAFEMOD% = FEMODI%

        '
        .AddNew
        On Error Resume Next
        !CODIEMPR = CodiEmp%
        On Error GoTo 0
        '
        !CODICONJ = CI1%
        !NUORD = NUOR%: NUMEORD& = NUOR% 'LO PASO A LONG PARA QUE NO DE ERROR LINEA ABAJO
        !NuordString = FORMATOCON0$(NUMEORD&, 4)
        !Descripcion = DESCRIX$
        
        !TIEMPOUNITARIO = TIEM1#
        
        !CantiOps = CANTI
        !SecciPro = CODSECCION
        '!FECHAMODIF = FETEXCOR1$(FEMODI%) 'daba error en jym
        !FECHAMODIF = CVDAT(HOY(HO$))
           User% = USNBR% Mod 100
           If User% < 1 Then User% = 1
        !Usua_Mod = User%
        !MARBORRA = 0
        !FECHA_VISTO = CVDAT(HOY(HO$))
        .Update
       Next i&
    End With
    Rst.Close
    Set Rst = Nothing
    TIEMPX = TRIM$(FORMATNUM$(XVALO(Text161) / 60, "F12.2"))
    Call ACTUREGISTRO("MASTER", "MANO_OBRA_REAL", "CODINT = " & CI1%, TIEMPX, REGAF&, "NOMESS")
    '
    Call BORRAREGISTROS("TIEMOBRA", CONDI$ & " AND MARBORRA = 1", REGAF&, "NOMESS")
    ANACOSI09.Text4 = TIEMPX
    '
    'ACUTALIZO ESTOS DATOS EN PANTALLA
    Me.UsuarioModifico = USERNAM$(User%)
    Me.TXTMODIFICADO = FECHATEX$(ULTIMAFEMOD%)
    Me.TXTFECHAVISTO = FECHATEX$(HOY(HO$))
    
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
          J& = MSF1.Row

          H& = 0: If MSF1.COL < MSF1.Cols Then H& = MSF1.COL + 1
10        For U& = J& To MSF1.Rows - 1
             HI& = H&: If U& > J& Then HI& = 0
             For i& = HI& To MSF1.Cols - 1
              If InStr(UCase$(MSF1.TextMatrix(U&, i&)), ABUSQUE$) > 0 Then
                MSF1.Row = U&: MSF1.COL = i&
                On Error Resume Next
                 VISIROWS = (MSF1.Height / MSF1.RowHeight(J&)) - 4
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
              Next i&
          Next U&
          '
          Call MENSERR(24, "Argumento no Encontrado.")
       End If
   End If

End Sub


Sub CARSECPROD()
    '
    Call BLANARCH("!SECPROD")
    J& = 0
    For U& = 1 To ULTREG&("SECPROD")
        X$ = REGLEIDO$("SECPROD", U&)
        COSECC = CVS(Left$(X$, 4))
        If COSECC >= 1 Then
            Call REGAPP("!SECPROD", X$)
        End If
    Next U&
    Call CIERRARCH("SECPROD")
    Call CIERRARCH("!SECPROD")
    '
    VTB% = VENTABU%("SECPROD")
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    For U& = 1 To ULTREG&("!SECPROD")
        X$ = REGLEIDO$("!SECPROD", U&)
        COSECC = CVS(Left$(X$, 4))
        If COSECC >= 1 Then
            J& = J& + 1
            Call GRAREG("SECPROD", X$, J&)
        End If
    Next U&
    '
    Call SETULTREG("SECPROD", J&)
    Call CIERRARCH("SECPROD")
    Call CIERRARCH("!SECPROD")
    '
End Sub
