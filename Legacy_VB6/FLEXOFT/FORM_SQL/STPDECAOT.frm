VERSION 5.00
Begin VB.Form STPDECAOT 
   BackColor       =   &H00FFFFD0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Central de Pesadas"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame19 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Cant.Dec."
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
      Left            =   3870
      TabIndex        =   15
      Top             =   1305
      Width           =   1095
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   330
         MaxLength       =   1
         TabIndex        =   16
         Text            =   "1"
         Top             =   285
         Width           =   450
      End
   End
   Begin VB.CommandButton Command6 
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
      Left            =   5265
      Picture         =   "STPDECAOT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Guardar"
      Top             =   750
      Width           =   675
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tablas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   135
      TabIndex        =   11
      Top             =   1290
      Width           =   3660
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Depositos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   465
         Width           =   1560
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Ubicaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2025
         TabIndex        =   13
         Top             =   480
         Width           =   1560
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Central de Pesadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   120
      TabIndex        =   6
      Top             =   90
      Width           =   4875
      Begin VB.CommandButton Command5 
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
         Left            =   4485
         TabIndex        =   8
         Top             =   225
         Width           =   175
      End
      Begin VB.TextBox Text1 
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
         Left            =   3570
         MaxLength       =   3
         TabIndex        =   7
         Top             =   225
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito Asoc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1860
         TabIndex        =   10
         Top             =   300
         Width           =   1680
      End
      Begin VB.Label Label24 
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
         Left            =   165
         TabIndex        =   9
         Top             =   570
         Width           =   4500
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   2475
      Left            =   5100
      ScaleHeight     =   2415
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   0
      Width           =   975
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
         Height          =   555
         Left            =   135
         Picture         =   "STPDECAOT.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir"
         Top             =   75
         Width           =   690
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   150
         ScaleHeight     =   75
         ScaleWidth      =   570
         TabIndex        =   1
         Top             =   1455
         Width           =   630
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   330
            TabIndex        =   2
            Top             =   -105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -285
         Picture         =   "STPDECAOT.frx":0454
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   1665
         Width           =   1515
      End
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
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
      Left            =   1470
      TabIndex        =   4
      Top             =   4830
      Width           =   1335
   End
   Begin VB.Label Label18 
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
      Height          =   300
      Left            =   1470
      TabIndex        =   3
      Top             =   4830
      Width           =   1335
   End
End
Attribute VB_Name = "STPDECAOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Call CIERRARCH("*.*")
  Unload Me
  
End Sub






Private Sub Command5_Click()
  Command5.Enabled = False
  '
  Call SELECHO("TADEPOSI")
  NPROV$ = VALACT1$("TADEPOSI")
    
  If NPROV$ <> "" Then
     Text1.TEXT = NPROV$
    ' Label24 = VALACT2$("TADEPOSI")
  End If
  '
  Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    If TRIM$(Label24) = "" Then
       Call COMUNI("Depósito No Válido")
       GoTo 99
    End If
    
    DEPOTRANS% = XVALO(CONTROL$("BOLRECEP", "DEPTRANS")) ' DEPOSITO MATERIAL EN TRANSITO
    If XVALO(Text1) = DEPOTRANS% Then
        COMUNI ("Depósito Seleccionado No Válido\Configurado como Material en Transito ")
       GoTo 99
    End If

    Call GRACONTROL("STOCKS", "DEPCENPE", Text1.TEXT)
    '
    CANDECI$ = TRIM(Text15)
    If XVALO(CANDECI$) < 1 Then CANDECI$ = 1
    If XVALO(CANDECI$) > 5 Then CANDECI$ = 5
    Call GRACONTROL("STOCKS", "CANDECI", CANDECI$)
    '
    Call CIERRARCH("*.*")
    Call FINAL("")
    Unload Me
99  Command6.Enabled = True
End Sub

Private Sub Form_Load()
    
    Text1.TEXT = CONTROL$("STOCKS", "DEPCENPE")
    Text15.TEXT = CONTROL("STOCKS", "CANDECI")
  
End Sub

Private Sub Option1_DblClick()
   If DERACCE%("TAUBICA", "Tabla de Ubicaciones") > 1 Then Call CARUBIPRE  ' MUESTRA VENTANA DE CARGA DE UBICACIONES PREDEFINIDAS
End Sub

Private Sub Option2_DblClick()
   '
   If DERACCE%("TADEPOSI", "Tabla de Depositos") > 1 Then Call CARTADEPOSI ' MUESTRA VENTANA DE CARGA DE DEPOSITOS
   '
End Sub

Private Sub Text1_Change()
    If XVALO(Text1) < 256 Then
       Label24 = ECOARCH("TADEPOSI", Chr$(XVALO(Text1)))
    End If
    
End Sub

Private Sub Text1_DblClick()
  Call Command5_Click
  
End Sub

