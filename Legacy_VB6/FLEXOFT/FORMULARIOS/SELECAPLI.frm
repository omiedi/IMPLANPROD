VERSION 5.00
Begin VB.Form SELECAPLI 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
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
      Height          =   800
      Left            =   7800
      Picture         =   "SELECAPLI.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   2220
      Width           =   855
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   7800
      Picture         =   "SELECAPLI.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   1350
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   7800
      Picture         =   "SELECAPLI.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   480
      Width           =   855
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      IntegralHeight  =   0   'False
      Left            =   5640
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   480
      Width           =   5175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ejecutables:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Aplicaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "SELECAPLI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Hide
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command29_Click()
   Command29.Enabled = False
   Unload Me
   Command29.Enabled = True
End Sub

Private Sub Form_Load()
   '
   Screen.MousePointer = 11
   Call AGRELI("Power-Up Industrial/FLEXO04/ARCHIG04/REMITO04/RETRAS04/BRECEP04/AMASTO04/ORCOMP04/INVENT04/CARMAQ04")
   Call AGRELI("Sistema Power-Up Gestión/FLGEST04/ARCHIG04/ACRGES04/VENGES04/LISPRE04")
   Call AGRELI("Gestión de Stocks e Inventarios/CONSTO04/BRECEP04/INVENT04/REMITO04/RETRAS04")
   Call AGRELI("Capacidad de Fábrica/CALCAP04")
   Call AGRELI("Análisis de Costos Standard/ANACOS04")
   Call AGRELI("Programación de Materiales/PROMRP04/ARCHIG04/AMASTO04")
   Call AGRELI("Secuenciacion de Carga de Maquinas/CARMAQ04/ORFABR04")
   Call AGRELI("Gestión de Compras Industriales/ORCOMP04/ARCHIG04/AMASTO04")
   Call AGRELI("Sub-Sistema ISO 9000/ISO9K04/INOCON04/ARCHIG04/AMASTO04")
   Call AGRELI("Planera Electrónica/PLANEL04/ARCHIG04/AMASTO04")
   Call AGRELI("Ordenes de Fabricación/ORFABR04/ARCHIG04/AMASTO04")
   Call AGRELI("Sistema de Reparaciones/FLEXOFT/REPARA02/ARCHIG04/AMASTO04")
   Call AGRELI("Control de Mano de Obra/IFABRI04/EFIMOB04")
   
   Call AGRELI("Sistema de Importacion Temporaria/SIMPTE03/AMASTO02")
   Call AGRELI("Contabilidad General/FLEXOFT/CONTAGEN/CONTAB04/CPLANCUE/AJUINFLA")
   Call AGRELI("Inventario con Ajuste por Inflacion/CONTAGEN/INVENAJU/AJUINFLA")
   '
   Call AGRELI("Gestión de Ventas y Deudores Clásico/FLEXOFT/FACTUR02/RECIBO02/COGEVE02/AMACLI02/ARCHIMA/AMASTO02/LISPRE02")
   Call AGRELI("Acredores Comerciales Clasico/FLEXOFT/AMAPRO/FACOMP02/ORPAGO02/COGECO02/ARCHIMA")
   Call AGRELI("Pedidos de Clientes/PEDIDO02/AMASTO02/AMACLI02/LISPRE02/ODESPA02")
   Call AGRELI("Gestion de Presupuestos/PRESUP02/AMACLI02/AMASTO02/LISPRE02")
   
   Call AGRELI("Seguimiento Comercial/Seguimiento Comercial/FLXMAIL")
   Screen.MousePointer = 1
   '
End Sub

Sub AGRELI(TTXX$)
   PPYY% = InStr(TTXX$, "/")
   If PPYY% < 1 Then PPYY% = 1 + Len(TTXX$)
   LEYEN$ = AJUSTI$(Left$(TTXX$, PPYY% - 1), 128) + Mid$(TTXX$, PPYY% + 1)
   List1.AddItem LEYEN$
End Sub

Private Sub List1_Click()
   '
   List2.Clear
   For U& = 1 To List1.ListCount
     If List1.Selected(U& - 1) = True Then
       LIEXE$ = TRIM$(Mid$(List1.List(U& - 1), 129)) + "/"
       While LIEXE$ <> ""
         PPXX% = InStr(LIEXE$, "/")
         If PPXX% < 1 Then GoTo 29
         APLIX$ = TRIM$(Left$(LIEXE$, PPXX% - 1))
         LIEXE$ = Mid$(LIEXE$, PPXX% + 1)
         For UK& = 1 To List2.ListCount
           If List2.List(UK& - 1) = APLIX$ Then GoTo 19
         Next UK&
         List2.AddItem APLIX$
19     Wend
     End If
29 Next U&
   '
End Sub
