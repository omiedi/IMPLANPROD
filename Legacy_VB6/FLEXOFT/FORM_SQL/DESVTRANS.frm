VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form DESVTRANS 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MODIFICAR DESVIOS TRANSITORIOS"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   2790
   ClientWidth     =   16290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   16290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.TextBox TxtLoteOriginal 
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
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      ToolTipText     =   "Fecha Visto - Revisión "
      Top             =   600
      Width           =   1275
   End
   Begin VB.TextBox TXTID 
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
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      ToolTipText     =   "Fecha Visto - Revisión "
      Top             =   600
      Width           =   1275
   End
   Begin VB.TextBox TxtDesvioEdit 
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
      Height          =   330
      Left            =   13560
      TabIndex        =   33
      Top             =   600
      Width           =   1455
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
      Height          =   360
      Left            =   3360
      TabIndex        =   24
      ToolTipText     =   "Click Para Seleccionar Sección"
      Top             =   600
      Width           =   175
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
      Left            =   12240
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Text            =   "N° de Desvio"
      ToolTipText     =   "Fecha Visto - Revisión "
      Top             =   600
      Width           =   1275
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   5055
      Left            =   120
      TabIndex        =   20
      ToolTipText     =   "Doble Click Editar"
      Top             =   4125
      Width           =   14920
      _ExtentX        =   26326
      _ExtentY        =   8916
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
      Height          =   1575
      Left            =   120
      ScaleHeight     =   1515
      ScaleWidth      =   3255
      TabIndex        =   21
      Top             =   4920
      Visible         =   0   'False
      Width           =   3315
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
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   1000
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   8640
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
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
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "Desvio"
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   600
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Height          =   9015
      Left            =   0
      TabIndex        =   8
      Top             =   -70
      Width           =   80
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00800000&
      FillColor       =   &H00400000&
      ForeColor       =   &H00400000&
      Height          =   9195
      Left            =   15120
      ScaleHeight     =   9135
      ScaleWidth      =   1755
      TabIndex        =   3
      Top             =   0
      Width           =   1815
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
         Picture         =   "DESVTRANS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Graba Lista de Operaciones y Tiempos"
         Top             =   7440
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
         Picture         =   "DESVTRANS.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   17
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
         TabIndex        =   5
         Top             =   8250
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   6
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
         TabIndex        =   4
         ToolTipText     =   "Salir Sin Procesar"
         Top             =   100
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -120
         Picture         =   "DESVTRANS.frx":0614
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
      Height          =   3045
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   15105
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
         Height          =   360
         Left            =   11820
         TabIndex        =   36
         ToolTipText     =   "Click Para Seleccionar Sección"
         Top             =   840
         Width           =   175
      End
      Begin VB.TextBox TXTFECHA 
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
         Height          =   330
         Left            =   10800
         TabIndex        =   35
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox TXTLBLFECHA 
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
         Left            =   10830
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "Fecha"
         Top             =   480
         Width           =   840
      End
      Begin VB.TextBox Text5 
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
         Left            =   9270
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "Serie"
         Top             =   480
         Width           =   840
      End
      Begin VB.CommandButton Command3 
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
         Left            =   9000
         TabIndex        =   31
         ToolTipText     =   "Click Para Seleccionar Sección"
         Top             =   840
         Width           =   175
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
         Height          =   330
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "Lote"
         Top             =   480
         Width           =   840
      End
      Begin VB.TextBox TXTDESCRIPCION 
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
         Height          =   330
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   840
         Width           =   3975
      End
      Begin VB.TextBox TXTCI 
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   840
         Width           =   1120
      End
      Begin VB.TextBox Text4 
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
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "Ci"
         Top             =   480
         Width           =   435
      End
      Begin VB.TextBox Text10 
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
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "Situación"
         Top             =   1320
         Width           =   2000
      End
      Begin VB.TextBox TXTSITUACION 
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
         Height          =   1335
         Left            =   7680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   1620
         Width           =   7500
      End
      Begin VB.TextBox TXTSERIE 
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
         Height          =   330
         Left            =   9240
         TabIndex        =   1
         Top             =   840
         Width           =   1455
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
         Left            =   3285
         TabIndex        =   2
         ToolTipText     =   "Click Para Seleccionar Sección"
         Top             =   840
         Width           =   175
      End
      Begin VB.TextBox TXTREFERENCIA 
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
         Height          =   1335
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   1620
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
         Height          =   330
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "Código"
         Top             =   480
         Width           =   840
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
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Referencia"
         Top             =   1320
         Width           =   2000
      End
      Begin VB.TextBox TXTCODIGO 
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
         Height          =   330
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   840
         Width           =   1950
      End
      Begin VB.TextBox TXTLOTE 
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
         Height          =   330
         Left            =   7530
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   840
         Width           =   1455
      End
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
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "DESVIOS TRANSITORIOS"
      Top             =   0
      Width           =   15180
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   195
      Left            =   4560
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
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
Attribute VB_Name = "DESVTRANS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARBUS$, ABUSQUE$
Public MODIFICATIEMMOBRA%
Sub ACOMODARMSF()
     MSF1.RowHeight(MSF1.row) = Label7.Height + 100
     MSF1.TextMatrix(MSF1.row, 2) = Label7
End Sub

Sub BorraTxtAsociados()

  Me.TXTFECHA = FECHATEX$(HOY(HO$))
  Me.TXTLOTE = ""
  Me.TXTSERIE = ""
  Me.TXTREFERENCIA.TEXT = ""
  Me.TXTSITUACION = ""
  Me.TXTCI = ""
  Me.TXTID = ""
  Me.TxtLoteOriginal = ""
  Me.TxtDesvioEdit = ""
  
  
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


Private Sub Command11_Click()
    Screen.MousePointer = 11
20  Campos$ = "N° Desvío/F8;Referencia/A32;CI/A0;Código/A16;Descripción/A24;Lote/A16;N° Serie/A16"


    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = "LISTA DE DESVIOS TRANISTORIOS"


    J& = 0
       SQLX$ = " select DESV_TRANS, REF_DESVIO,IDENLOTE,NUMEROSERIE,COD_EXTE,COD_INTE  "
       SQLX$ = SQLX$ + " FROM des_trans WITH(NOLOCK) order by cod_inte"
       Set rs = READSET(SQLX$)
       With rs
        Do While Not .EOF
        '
        NRO& = XVALO(!DESV_TRANS)
        REFDES$ = SAFETEXT$(!REF_DESVIO)
        CI1% = XVALO(!COD_INTE)
        CODI$ = SAFETEXT$(!COD_EXTE)
        DESX$ = DESCRIT0$(CI1%)
        LOTE$ = SAFETEXT$(!idenlote)
        NS$ = SAFETEXT$(!NumeroSerie)
        J& = J& + 1
        FRMZELECHO.msf2.Rows = J& + 1
        FRMZELECHO.msf2.TextMatrix(J&, 1) = NRO&
        FRMZELECHO.msf2.TextMatrix(J&, 2) = REFDES$
        FRMZELECHO.msf2.TextMatrix(J&, 3) = CI1%
        FRMZELECHO.msf2.TextMatrix(J&, 4) = CODI$
        FRMZELECHO.msf2.TextMatrix(J&, 5) = DESX$
        FRMZELECHO.msf2.TextMatrix(J&, 6) = LOTE$
        FRMZELECHO.msf2.TextMatrix(J&, 7) = NS$
  
       .MoveNext
     Loop
   End With
   rs.Close
   Set rs = Nothing
   Screen.MousePointer = 1

   FRMZELECHO.Show 1
   
DesvioSeleccionado& = XVALO(RESP_ZELE_VECT(1))
If DesvioSeleccionado& > 0 Then
    Call BorraTxtAsociados

    Text15 = DesvioSeleccionado&
    Me.TXTCODIGO = ""
    Me.TXTDESCRIPCION = ""

Else
GoTo 99
End If

99
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    COD$ = ListaSeleccionArticulos("", True)
    ActualizaListaSeleccion = False
    CODIGOENPANTALLA$ = TXTCODIGO
    If COD$ <> "" Then
     TXTCODIGO.TEXT = COD$
     
     If CODIGOENPANTALLA$ <> TXTCODIGO.TEXT Then
        Me.TXTLOTE = ""
        Me.TXTSERIE = ""
        Me.TXTREFERENCIA.TEXT = ""
        Me.TXTSITUACION = ""

     End If
    End If
    TXTCI = CODINT%(TXTCODIGO)

    Command2.Enabled = True
End Sub






Private Sub Command3_Click()
    CI1% = XVALO(Me.TXTCI)
    '
    SQLX$ = "SELECT IdenLote   FROM MOVISTO WHERE Cod_Inte = CI1%"
    SQLX$ = SQLX$ + " GROUP BY IDENLOTE"
    SQLX$ = SQLX$ + " ORDER BY IDENLOTE DESC"
    
    FRMZELECHO.Caption = "Lotes de Código Activo"
    Campos$ = "LOTE/A20"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)

    Dim obj As New RECORXET
    CONDI$ = "Cod_Inte = " & CI1%
    CONDI$ = CONDI$ + " GROUP BY IDENLOTE ORDER BY IDENLOTE DESC"
    Set rs1 = obj.RS_GENERAL("MOVISTO", " IDENLOTE", CONDI$)
    Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
    '
    '
    FRMZELECHO.Width = 6000
50  FRMZELECHO.Show 1
     If RESP_ZELE_VECT(1) <> "" Then
       TXTLOTE = RESP_ZELE_VECT(1)
     End If
End Sub





Private Sub Command4_Click()
Call CommSel("SELFECHA", Command4, TXTFECHA)
End Sub


Private Sub Command6_Click()

    Command6.Enabled = False
    nDesvioOriginal& = XVALO(Text15)
    nDesvio& = nDesvioOriginal&
    If nDesvioOriginal& <> XVALO(Me.TxtDesvioEdit) And XVALO(Me.TxtDesvioEdit) > 0 Then  'si se le agrego un numero diferente en TxtDesvioEdit PREGUNTA
       OPT1% = COMALTER%("Se ha Seleccionado un N° de Desvio Transitorio\Diferente al Registrado en la B.Datos\Desea Registrar con Este Nuevo Número:\\Cancelar\Guardar con Nuevo Número")
       If OPT1% < 2 Then GoTo 99
       nDesvio& = XVALO(Me.TxtDesvioEdit) 'si paso los controles lo guarda con el nuevo numero, si no entro en el if lo guarda con el numero que le asigno incialmente
    End If
   
    On Error GoTo ERROR_GUARDAR
    FLEXCONN.BeginTrans

    SQLX$ = "UPDATE  DES_TRANS "
    SQLX$ = SQLX$ + " SET CODIEMPR = " & CodiEmpr%
    SQLX$ = SQLX$ + " , DESCRIPCION = '" & Me.TXTDESCRIPCION & "'"
    SQLX$ = SQLX$ + " , SITUACION = '" & Me.TXTSITUACION & "'"
    SQLX$ = SQLX$ + " , REF_DESVIO = '" & TXTREFERENCIA & "'"
    SQLX$ = SQLX$ + " , IDENLOTE = '" & TRIM$(TXTLOTE) & "'"
    SQLX$ = SQLX$ + " , NUMEROSERIE = '" & TRIM$(Me.TXTSERIE) & "'"
    SQLX$ = SQLX$ + " , COD_INTE = '" & XVALO(Me.TXTCI) & "'"
    SQLX$ = SQLX$ + " , COD_EXTE = '" & TRIM$(Me.TXTCODIGO) & "'"
    SQLX$ = SQLX$ + " , IDSUBOR = ''"
    SQLX$ = SQLX$ + " , FecheEmis = '" & FETEXCOR1$(FECHANUM(Me.TXTFECHA)) & "'"
    SQLX$ = SQLX$ + " , DESV_TRANS = " & nDesvio&
    SQLX$ = SQLX$ + " WHERE  DES_TRANS_ID  =  " & XVALO(Me.TXTID)

    FLEXCONN.Execute (SQLX$)
    
    SQLX$ = "UPDATE  TRAZAXSERIE "
    SQLX$ = SQLX$ + " SET DESV_TRANS = " & nDesvio&
    SQLX$ = SQLX$ + " WHERE  DESV_TRANS  =  " & nDesvioOriginal&
    SQLX$ = SQLX$ + " AND  IDENLOTE  =  '" & TRIM$(Me.TxtLoteOriginal) & "'"
    
    FLEXCONN.Execute (SQLX$)
    
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
       GoTo 99
    Else
      FLEXCONN.CommitTrans
      TxtDesvioEdit = 0
      Call BorraTxtAsociados
      Text15 = "" 'PARA REFERESCAR
      Text15 = nDesvio&
   End If

99  Command6.Enabled = True
End Sub







Private Sub Form_Load()
        
        Campos$ = "N° Desvio/A7;Referencia/A24;Lote/A12;Serie/A16;Ci./A8;Código/A16;Descripción/A24;Situación/A24;Fecha/D8;ID/A0"
        Call CARGA_ENCABEZADO(Me.MSF1, Campos$, Me)
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If Cancel = 0 Then
      Command1_Click
  End If
End Sub





Private Sub MSF1_DblClick()
   If MSF1.row >= 1 Then
         Me.TxtDesvioEdit.TEXT = MSF1.TextMatrix(MSF1.row, 1)
         Me.TXTREFERENCIA = MSF1.TextMatrix(MSF1.row, 2)
         Me.TXTLOTE = MSF1.TextMatrix(MSF1.row, 3)
         Me.TXTSERIE = MSF1.TextMatrix(MSF1.row, 4)
         Me.TXTCI = MSF1.TextMatrix(MSF1.row, 5)
         Me.TXTCODIGO = MSF1.TextMatrix(MSF1.row, 6)
         Me.TXTSITUACION = MSF1.TextMatrix(MSF1.row, 8)
         Me.TXTFECHA = MSF1.TextMatrix(MSF1.row, 9)
         Me.TXTID = MSF1.TextMatrix(MSF1.row, 10)
         
         Me.TxtLoteOriginal = MSF1.TextMatrix(MSF1.row, 3)
         
   End If
        

End Sub


Private Sub MSF1_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Verificar si se presionó Ctrl+C
    If (Shift And vbCtrlMask) And (KeyCode = vbKeyC) Then
        ' Obtener la celda seleccionada
        Dim row As Integer, col As Integer
        row = MSF1.row
        col = MSF1.col
        
        ' Copiar el contenido de la celda al portapapeles
        Clipboard.Clear
        Clipboard.SetText MSF1.TextMatrix(row, col)
    End If

End Sub


Private Sub Text15_Change()

   

   ULTIMAFEMOD% = 0
   MSF1.Rows = 1
   MSF1.Redraw = False
   Call CREATABLA_EQUIPOS
   DESVIO_ORIGINAL = XVALO(Text15)
   
   SQLX$ = "SELECT * FROM  DES_TRANS WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE DESV_TRANS = " & DESVIO_ORIGINAL
   Set Rst = READSET(SQLX$)
   With Rst
      Do While Not .EOF
        Fila& = Fila& + 1

        MSF1.Rows = MSF1.Rows + 1
        REF_DESVIO$ = SAFETEXT$(!REF_DESVIO)
        
        Label7 = REF_DESVIO$
        MSF1.RowHeight(Fila&) = Label7.Height + 100
        
        Situacion$ = SAFETEXT$(!Situacion)
        Label7 = Situacion$
        If MSF1.RowHeight(Fila&) < Label7.Height + 100 Then 'si el alto de la situacion es mayor al que quedo de la referencia de desvio
           MSF1.RowHeight(Fila&) = Label7.Height + 100
        End If

        Descripcion$ = DESCRIT0$(XVALO(!COD_INTE))
        Label7 = Descripcion$
        If MSF1.RowHeight(Fila&) < Label7.Height + 100 Then 'si el alto de la situacion es mayor al que quedo de la referencia de desvio
           MSF1.RowHeight(Fila&) = Label7.Height + 100
        End If
        FEMOD% = CVINT(!FecheEmis)
        FECHA1 = FECHATEX$(FEMOD%)
        '
        MSF1.Rows = Fila& + 1
        MSF1.TextMatrix(Fila&, 1) = DESVIO_ORIGINAL
        MSF1.TextMatrix(Fila&, 2) = REF_DESVIO$
        MSF1.TextMatrix(Fila&, 3) = SAFETEXT$(!idenlote)
        MSF1.TextMatrix(Fila&, 4) = SAFETEXT$(!NumeroSerie)
        MSF1.TextMatrix(Fila&, 5) = XVALO(!COD_INTE)
        MSF1.TextMatrix(Fila&, 6) = SAFETEXT$(!COD_EXTE)
        MSF1.TextMatrix(Fila&, 7) = Descripcion$
        MSF1.TextMatrix(Fila&, 8) = SAFETEXT$(!Situacion)
        MSF1.TextMatrix(Fila&, 9) = FECHA1
        MSF1.TextMatrix(Fila&, 10) = XVALO(!DES_TRANS_ID)
        
     .MoveNext
     Loop

     End With
     
     Rst.Close
     Set Rst = Nothing
     MSF1.Redraw = True
End Sub







Private Sub cargar_datos()
    '
    Label7 = ""

    MSF1.Cols = 10
    MSF1.Rows = MSF1.Rows + 1
    MSF1.row = MSF1.Rows - 1
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

Private Sub TXTCODIGO_Change()
   TXTDESCRIPCION.TEXT = DESCRIT0$(CODINT%(TXTCODIGO.TEXT))

End Sub

Private Sub TXTDESCRIPCION_Change()
   Label7 = TXTDESCRIPCION
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

Private Sub UsuarioModifico_Change()

End Sub

