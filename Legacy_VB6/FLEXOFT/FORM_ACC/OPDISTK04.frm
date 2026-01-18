VERSION 5.00
Begin VB.Form OPDISTK04 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Disponibilidad General de Stocks - Configuración"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   6405
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture6 
      Height          =   3900
      Left            =   5460
      ScaleHeight     =   3840
      ScaleWidth      =   1035
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   1100
      Begin VB.Image Image3 
         Height          =   3850
         Left            =   0
         Picture         =   "OPDISTK04.frx":0000
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Autonomia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   3315
      TabIndex        =   39
      Top             =   1680
      Width           =   1140
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   135
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   " "
         Top             =   375
         Width           =   885
      End
   End
   Begin VB.OptionButton Option7 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Por Costo Calculado"
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
      Left            =   2940
      TabIndex        =   38
      Top             =   3100
      Width           =   2220
   End
   Begin VB.OptionButton Option6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Por Costo de Reposición"
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
      Left            =   210
      TabIndex        =   37
      Top             =   3100
      Width           =   2535
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Costeo de Movimiento de Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   105
      TabIndex        =   36
      Top             =   2730
      Width           =   5280
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   4485
      TabIndex        =   34
      Top             =   1680
      Width           =   885
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   60
         Picture         =   "OPDISTK04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Realiza Control de Movimientos de Stock Repetidos o Duplicados"
         Top             =   270
         Width           =   750
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Registro de Trazabilidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   105
      TabIndex        =   29
      Top             =   1680
      Width           =   3180
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Por Declaración de Consumo"
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
         Left            =   1425
         TabIndex        =   31
         Top             =   315
         Width           =   1710
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Automática por F.I.F.O."
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
         Left            =   105
         TabIndex        =   30
         Top             =   315
         Value           =   -1  'True
         Width           =   1350
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   3900
      Left            =   5460
      ScaleHeight     =   3840
      ScaleWidth      =   1110
      TabIndex        =   26
      Top             =   0
      Width           =   1170
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   32
         Top             =   2835
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton command1 
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
         Height          =   610
         Left            =   120
         Picture         =   "OPDISTK04.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   315
         Width           =   645
      End
      Begin VB.CommandButton Command2 
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
         Left            =   120
         Picture         =   "OPDISTK04.frx":27EC
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1000
         Width           =   645
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -290
         Picture         =   "OPDISTK04.frx":2AF6
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   3150
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Presentación de Costo y Precios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   210
      TabIndex        =   22
      Top             =   7770
      Width           =   5160
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cos.Repos. Calculado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1920
         TabIndex        =   25
         Top             =   360
         Width           =   1305
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Precio Base (C.Rep. + 15 %)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3360
         TabIndex        =   24
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cos.Repos. Informado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   360
         TabIndex        =   23
         Top             =   360
         Value           =   -1  'True
         Width           =   1905
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Impresión de Situación de Stocks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   105
      TabIndex        =   16
      Top             =   210
      Width           =   5265
      Begin VB.CommandButton Command5 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   4095
         Picture         =   "OPDISTK04.frx":4A18
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   315
         Width           =   750
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
         Height          =   360
         Left            =   210
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   3750
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
         Left            =   3780
         TabIndex        =   18
         Top             =   315
         Width           =   175
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
         Height          =   360
         Left            =   2415
         TabIndex        =   17
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         TabIndex        =   21
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Origen de Stocks Remotos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   210
      TabIndex        =   7
      Top             =   3990
      Width           =   5160
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
         Height          =   360
         Index           =   3
         Left            =   735
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1995
         Width           =   4230
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
         Height          =   360
         Index           =   2
         Left            =   735
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1470
         Width           =   4230
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
         Height          =   360
         Index           =   1
         Left            =   735
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   945
         Width           =   4230
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
         Height          =   360
         Index           =   0
         Left            =   735
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   420
         Width           =   4230
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R-4:"
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
         Index           =   3
         Left            =   0
         TabIndex        =   15
         Top             =   2100
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R-3:"
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
         Index           =   2
         Left            =   0
         TabIndex        =   13
         Top             =   1575
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R-2:"
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
         Index           =   1
         Left            =   0
         TabIndex        =   11
         Top             =   1050
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R-1:"
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
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   525
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Destino Archivo de Transferencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   210
      TabIndex        =   4
      Top             =   6720
      Width           =   5160
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
         Height          =   360
         Left            =   735
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   420
         Width           =   4230
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta:"
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
         TabIndex        =   6
         Top             =   525
         Width           =   645
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Identificador de Sede"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   210
      TabIndex        =   1
      Top             =   5775
      Width           =   5160
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
         Height          =   360
         Left            =   3780
         MaxLength       =   8
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   420
         Width           =   1185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código Identificador de Sede:"
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
         TabIndex        =   3
         Top             =   525
         Width           =   3480
      End
   End
   Begin VB.CommandButton Command7 
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
      Height          =   855
      Left            =   5600
      Picture         =   "OPDISTK04.frx":4D22
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3990
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   5040
      Picture         =   "OPDISTK04.frx":502C
      Top             =   6950
      Width           =   2010
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu grabar 
         Caption         =   "Grabar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "OPDISTK04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    '
    Call GRACONTROL("", "FORDISTO", Text4.TEXT)
    '
    MODOTRAZA$ = "FIFO"
    If Option5.Value = True Then
       If EXISTE%("TRAZAB04.EXE") < 1 Then
          Call MENSERR(24, "Trazabilidad por Declaración de Consumo\No Disponible en Presente Versión.\   \Falta Aplicativo 'TRAZAB04.EXE'.")
          Option4.Value = True
          Exit Sub
       End If
       MODOTRAZA$ = "DECONSU"
    End If
    Call GRACONTROL("", "RETRAZA", MODOTRAZA$)
    '
    '
    ' los que siguen están deshabilitados - son de versión 'DAT'
    IDENSEDE$ = TRIM$(Text1.TEXT)
    Call GRACONTROL("", "IDENSEDE", Text1.TEXT)
    '
    DESTITRA$ = TRIM$(Text2.TEXT)
    Call GRACONTROL("", "DESTITRA", Text2.TEXT)
    '
    For U& = 0 To 3
       ORITRA$(U&) = TRIM$(Text3(U&).TEXT)
    Next U&
    '
    Call GRACONTROL("", "ORITRA-1", Text3(0).TEXT)
    Call GRACONTROL("", "ORITRA-2", Text3(1).TEXT)
    Call GRACONTROL("", "ORITRA-3", Text3(2).TEXT)
    Call GRACONTROL("", "ORITRA-4", Text3(3).TEXT)
    '
    MODOCORE% = 1: MODOCOREP$ = "COREPOS"
    If Option2.Value = True Then
         MODOCORE% = 2: MODOCOREP$ = "PREBASE"
       ElseIf Option3.Value = True Then
         MODOCORE% = 3: MODOCOREP$ = "COSCALCU"
    End If
    Call GRACONTROL("SITUASTO", "MODOCORE", MODOCOREP$)
    '
    REPOSI$ = "COSTOREPOSICION"
    If Option7.Value = True Then REPOSI$ = "COSTOCALCULADO"
    Call GRACONTROL("CSTOCK", "COSMOVI", REPOSI$)
    '
    Call GRACONTROL("", "AUTODISP", Text6)
    '
    Call CIERRARCH("*.*")
    Call COMUNI("Re-Configuración Completada.\Re-Inicie Ahora la Aplicación '" + UCase$(App.EXEName) + "'.")
    Call FINAL("")
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text4.TEXT = VDEVU$
        Text5.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   '
   Call APERBASDAT("STOCKS")
   SQLX$ = " SELECT * FROM MOVISTO "
   SQLX$ = SQLX$ + "order by FeReMovi ASC"
   '
   Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
   With MoviTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
         On Error GoTo 0
         FIN& = .RecordCount
         .MoveFirst
         KU& = 0: APADUP% = 0
         Do While Not .EOF
           KU& = KU& + 1
           Call TRACE(20, KU&, FIN&)
           FEREMO = !FeReMovi
           FEMO = !FECHMOV
           CIXI% = !COD_INTE
           DOPR$ = !DOPRILAR
           CAIN = !CANTINGRE
           CASA = !CANTSALID
           IDLO$ = !IdenLote
           NORI% = !NUORITEM
           '
           If TRIM$(DOPR$) <> "" And CIXI% > 0 Then
              APOS = .AbsolutePosition
              '
              If APOS < FIN& - 1 Then
                 Do While !FeReMovi = FEREMO
                   .MoveNext
                   If FEMO = !FECHMOV Then
                     If CIXI% = !COD_INTE Then
                       If DOPR$ = !DOPRILAR Then
                         If CAIN = !CANTINGRE Then
                           If CASA = !CANTSALID Then
                             If IDLO$ = !IdenLote Then
                               If NORI% = !NUORITEM Then
                                 APADUP% = 1
                                 MsgBox "Aparente Duplicación en " + DOPR$ + " - " + CODEXT$(CIXI%)
                               End If
                             End If
                           End If
                         End If
                       End If
                     End If
                   End If
                 Loop
                 '
                 .AbsolutePosition = APOS
              End If
           End If
         .MoveNext
         Loop
      End If
   End With
   '
   Screen.MousePointer = 1
   TTXX$ = "Fin del Control"
   If APADUP% < 1 Then TTXX$ = TTXX$ + " sin Movimientos Duplicados."
   MsgBox TTXX$
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Call CONECRUN("FLFORMS/" + TRIM$(Text4.TEXT), "Formularios de Impresion")
End Sub

Private Sub Form_Load()
    '
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
      Me.Frame8.BackColor = &HFCD7B6
      Me.Frame9.BackColor = &HFCD7B6
      
      Me.Option1.BackColor = &HFCD7B6
      Me.Option2.BackColor = &HFCD7B6
      Me.Option3.BackColor = &HFCD7B6
      Me.Option4.BackColor = &HFCD7B6
      Me.Option5.BackColor = &HFCD7B6
      Me.Option6.BackColor = &HFCD7B6
      Me.Option7.BackColor = &HFCD7B6
      MARCOBARRA.Top = 3550
      
      Set MARCOBARRA.Container = Me.Picture6
      
    End If

    Call TRANSLABELS(Name)
    '
    Text4.TEXT = CONTROL$("", "FORDISTO")
    Text5.TEXT = ECOARCH$("INDIFRM", Text4.TEXT)
    '
    Text6.TEXT = CONTROL$("", "AUTODISP")
    '
    MODOTRAZA$ = CONTROL$("", "RETRAZA")
    If MODOTRAZA$ = "DECONSU" Then
       Option5.Value = True
    End If
    '
    '
    ' los que siguen están deshabilitados - son de versión 'DAT'
    Call CARUTRA
    '
    Text1.TEXT = IDENSEDE$
    Text2.TEXT = DESTITRA$
    For U& = 0 To 3
       Text3(U&).TEXT = ORITRA$(U&)
    Next U&
    '
    Option1.Value = True
    MODOCORE% = 1: MODOCOREP$ = "COREPOS"
    If TRIM$(UCase$(CONTROL$("SITUASTO", "MODOCORE"))) = "PREBASE" Then
       Option2.Value = True
       MODOCORE% = 2: MODOCOREP$ = "PREBASE"
     ElseIf TRIM$(UCase$(CONTROL$("SITUASTO", "MODOCORE"))) = "COSCALCU" Then
       Option3.Value = True
       MODOCORE% = 3: MODOCOREP$ = "COSCALCU"
    End If
    '
    '//OT-07-0458//
    Option6.Value = True
    If CONTROL$("CSTOCK", "COSMOVI") = "COSTOCALCULADO" Then Option7.Value = True
    '//OT-07-0458//

End Sub

Private Sub grabar_Click()
  Call Command2_Click
End Sub

Private Sub Option4_Click()
    Call COMUNI("Opción Solo Disponible como\Versión Beta (Experimental).")
End Sub

Private Sub Option5_Click()
    If EXISTE%("TRAZAB04.EXE") < 1 Then
       Call MENSERR(24, "No Disponible en Presente Versión.\   \Falta Aplicativo 'TRAZAB04.EXE'.")
       Option4.Value = True
    End If
End Sub

Private Sub salir_Click()
   Call Command1_Click
End Sub

Private Sub Text4_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text4.TEXT = VDEVU$
        Text5.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text5_Click()
   Text4.SetFocus
End Sub

Private Sub Text6_LostFocus()
    '
    VALOR$ = Text6
    Text6 = FORMATNUM$(VALOR, "F5.2")
    '
End Sub
















