VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FORCANEC07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cálculo Necesidad de Componentes según Programa de Entregas"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5835
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   3120
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   40
      Top             =   2520
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00006060&
      Height          =   6000
      Left            =   4935
      ScaleHeight     =   5940
      ScaleWidth      =   1425
      TabIndex        =   28
      Top             =   -50
      Width           =   1485
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   31
         Top             =   4590
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Start"
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
         Picture         =   "Forcanec07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Inicia Proceso de Cálculo"
         Top             =   1050
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
         Height          =   700
         Left            =   105
         Picture         =   "Forcanec07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Abandona Configuración"
         Top             =   350
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "Forcanec07.frx":0454
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4970
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Valores por Omisión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   315
      TabIndex        =   18
      Top             =   6120
      Width           =   6315
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
         Height          =   285
         Left            =   2940
         TabIndex        =   6
         Text            =   " "
         Top             =   2310
         Width           =   750
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
         Height          =   285
         Left            =   2940
         TabIndex        =   5
         Text            =   " "
         Top             =   1995
         Width           =   750
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
         Height          =   285
         Left            =   2940
         TabIndex        =   4
         Text            =   " "
         Top             =   1680
         Width           =   750
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
         Height          =   285
         Left            =   2940
         TabIndex        =   3
         Text            =   " "
         Top             =   1365
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
         Height          =   285
         Left            =   2940
         TabIndex        =   2
         Text            =   " "
         Top             =   945
         Width           =   540
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Mensual"
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
         Left            =   4725
         TabIndex        =   1
         Top             =   525
         Width           =   1485
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Semanal"
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
         Left            =   2940
         TabIndex        =   0
         Top             =   525
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Elem. Comprados"
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
         Left            =   3885
         TabIndex        =   25
         Top             =   2355
         Width           =   2325
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Mecanizados"
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
         Left            =   3885
         TabIndex        =   24
         Top             =   2040
         Width           =   2325
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Subconjuntos"
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
         Left            =   3885
         TabIndex        =   23
         Top             =   1725
         Width           =   2325
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lotes Standard:"
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
         Left            =   420
         TabIndex        =   22
         Top             =   1410
         Width           =   2325
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Prod. Terminados"
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
         Left            =   3885
         TabIndex        =   21
         Top             =   1410
         Width           =   2325
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Plazo de Disposición:"
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
         Left            =   105
         TabIndex        =   20
         Top             =   945
         Width           =   2640
      End
      Begin VB.Label label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de Programación:"
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
         Left            =   105
         TabIndex        =   19
         Top             =   525
         Width           =   2640
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   210
      TabIndex        =   17
      Top             =   210
      Width           =   4515
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cero Mínimos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2760
         TabIndex        =   42
         Top             =   240
         Width           =   1485
      End
      Begin VB.CheckBox Check11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "No Considera Lote Económico"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   500
         TabIndex        =   39
         Top             =   1800
         Width           =   3900
      End
      Begin VB.CheckBox Check9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Suprime Ordenes de Fabricación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   500
         TabIndex        =   37
         Top             =   1560
         Width           =   3900
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Suprime Compras en Curso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   500
         TabIndex        =   36
         Top             =   1320
         Width           =   3900
      End
      Begin VB.CheckBox Check10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Suprime Pedidos Pendientes de Entrega"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   500
         TabIndex        =   38
         Top             =   1065
         Width           =   3900
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cero Stocks y Mínimos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   2325
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "No Considera Negativos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2760
         TabIndex        =   9
         Top             =   480
         Width           =   1485
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Proceso Normal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1860
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         X1              =   0
         X2              =   4560
         Y1              =   950
         Y2              =   950
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Avance de Proceso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2865
      Left            =   210
      TabIndex        =   16
      Top             =   2520
      Width           =   4515
      Begin VB.CheckBox Check7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ordenes de Compra y Fabricación a Emitir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2415
         Width           =   4250
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Generación de Necesidades Secundarias"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2100
         Width           =   4250
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cálculo de Necesidades Netas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1785
         Width           =   4250
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Apertura de Archivos Maestros"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   525
         Width           =   3400
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Existencia Actual de Componentes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   840
         Width           =   3400
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Inicialización Necesidades Secundarias"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1155
         Width           =   4250
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Necesidades según Entregas Pendientes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1470
         Width           =   4250
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3780
         TabIndex        =   27
         Top             =   525
         Width           =   435
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "Forcanec07.frx":2376
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "Forcanec07.frx":25AA
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Height          =   225
      Left            =   4305
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Height          =   225
      Left            =   3360
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Height          =   225
      Left            =   2730
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "FORCANEC07"
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
    BALAREI% = 0
    'If CantRegistros("BALAREQMO") > 0 Then BALAREI% = 1
    '
    If BALAREI% > 0 Then
       If FORCANEC07.Option1.Value = False Then GoTo 50
       If FORCANEC07.Check8.Value <> 1 Then GoTo 50
       If FORCANEC07.Check9.Value <> 1 Then GoTo 50
       If FORCANEC07.Check10.Value <> 1 Then GoTo 50
    End If
    'BLOQUEA LOS OBJETOS CORRESPONDIENTES A LAS OPCIONES DE CALCULO EN PANTALLA
    Option1.Enabled = False
    Option2.Enabled = False
    Option3.Enabled = False
    Option6.Enabled = False
    Check8.Enabled = False
    Check9.Enabled = False
    Check10.Enabled = False
    Check11.Enabled = False
    '
    TIMEANT% = FLEXCONN.CommandTimeout
    FLEXCONN.CommandTimeout = 999
    Call GRAPARCAL
    Call SETCALCU
    Call CANECOM
    FLEXCONN.CommandTimeout = TIMEANT%
    '
    'DESBLOQUEA LOS OBJETOS CORRESPONDIENTES A LAS OPCIONES DE CALCULO EN PANTALLA
    Option1.Enabled = True
    Option2.Enabled = True
    Option3.Enabled = True
    Option6.Enabled = True
    Check8.Enabled = True
    Check9.Enabled = True
    Check10.Enabled = True
    Check11.Enabled = True
    '
 Exit Sub
    '
50  Call MENSERR(24, "Configuración de Cálculo Incompatible\para Programacion por Balance de Requerimientos.\   \Los valores Requeridos\Serán Restaurados Automáticamente.")
    Call Form_Activate
End Sub

Private Sub Form_Activate()
    Refresh
    Call PARACAL
End Sub

Private Sub Form_Load()
   Call APLICACIONSKINS(Me, Picture1)
End Sub
