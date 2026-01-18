VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FORDACOD07 
   BackColor       =   &H00C0EEC0&
   Caption         =   "Muestra Datos de Item de Stock"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   11385
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   1365
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   47
      Top             =   -105
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   8100
      Left            =   10440
      ScaleHeight     =   8040
      ScaleWidth      =   1140
      TabIndex        =   40
      Top             =   0
      Width           =   1200
      Begin VB.CommandButton Command5 
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
         Left            =   120
         Picture         =   "FORDACOD07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   870
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Left            =   120
         Picture         =   "FORDACOD07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Terminar - Salir de Consulta de Stocks"
         Top             =   180
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
         Height          =   690
         Left            =   120
         Picture         =   "FORDACOD07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Imprime Situación de Stocks"
         Top             =   5400
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   41
         Top             =   6360
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   42
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -285
         Picture         =   "FORDACOD07.frx":0ABE
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   6840
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0EEC0&
      Caption         =   "Imagen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7450
      Index           =   0
      Left            =   105
      TabIndex        =   38
      Top             =   120
      Width           =   4875
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   7050
         Index           =   0
         Left            =   150
         ScaleHeight     =   6990
         ScaleWidth      =   4485
         TabIndex        =   39
         Top             =   280
         Width           =   4550
      End
      Begin VB.Image Image1 
         Height          =   10005
         Index           =   0
         Left            =   0
         Top             =   105
         Visible         =   0   'False
         Width           =   10005
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0EEC0&
      Caption         =   "Especificaciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   5040
      TabIndex        =   1
      Top             =   3720
      Width           =   5295
      Begin VB.CommandButton Command1 
         Height          =   260
         Left            =   30
         TabIndex        =   5
         Top             =   3420
         Width           =   3375
      End
      Begin VB.TextBox Text43 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3060
         Left            =   30
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   360
         Width           =   5235
      End
      Begin VB.CommandButton Command49 
         Caption         =   "Copiar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   3360
         TabIndex        =   3
         Top             =   3420
         Width           =   765
      End
      Begin VB.CommandButton Command47 
         Caption         =   "Selecc Todo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   4065
         TabIndex        =   2
         Top             =   3420
         Width           =   1200
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0EEC0&
      Caption         =   "Datos del Articulo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   5040
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.Label Label18 
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
         Height          =   255
         Left            =   1560
         TabIndex        =   37
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label17 
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
         Height          =   255
         Left            =   4680
         TabIndex        =   36
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   3840
         TabIndex        =   35
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   3840
         TabIndex        =   34
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   1080
         TabIndex        =   33
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   1080
         TabIndex        =   32
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label12 
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
         Height          =   255
         Left            =   2040
         TabIndex        =   31
         Top             =   2280
         Width           =   3135
      End
      Begin VB.Label Label11 
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
         Height          =   255
         Left            =   2925
         TabIndex        =   30
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   1080
         TabIndex        =   29
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   1080
         TabIndex        =   28
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label7 
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
         Height          =   255
         Left            =   4320
         TabIndex        =   27
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label6 
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
         Height          =   255
         Left            =   2880
         TabIndex        =   26
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   1080
         TabIndex        =   25
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Left            =   1080
         TabIndex        =   24
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   5055
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   3360
         TabIndex        =   22
         Top             =   260
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "L. Máximo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   13
         Left            =   2805
         TabIndex        =   21
         Top             =   3180
         Width           =   1155
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "St. Actual:"
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
         Index           =   12
         Left            =   120
         TabIndex        =   20
         Top             =   2800
         Width           =   1185
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "L.Econom:"
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
         Index           =   11
         Left            =   2805
         TabIndex        =   19
         Top             =   2820
         Width           =   1275
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "St.Mínimo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   10
         Left            =   120
         TabIndex        =   18
         Top             =   3180
         Width           =   1305
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   8
         Left            =   2280
         TabIndex        =   17
         Top             =   1950
         Width           =   1305
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Pr. Hab:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   7
         Left            =   330
         TabIndex        =   16
         Top             =   2320
         Width           =   945
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "$/U:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   6
         Left            =   640
         TabIndex        =   15
         Top             =   1950
         Width           =   435
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Color:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   5
         Left            =   3800
         TabIndex        =   14
         Top             =   1500
         Width           =   495
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Ta/Tp:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   3
         Left            =   2260
         TabIndex        =   13
         Top             =   1500
         Width           =   585
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Kg/U:"
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
         Index           =   2
         Left            =   520
         TabIndex        =   12
         Top             =   1500
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kilos/ Unid:"
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
         Index           =   4
         Left            =   3345
         TabIndex        =   11
         Top             =   4800
         Width           =   870
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "U-M:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   4200
         TabIndex        =   10
         Top             =   1120
         Width           =   465
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "T-M:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   640
         TabIndex        =   9
         Top             =   1125
         Width           =   1500
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kilos/ Metro:"
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
         Index           =   19
         Left            =   1230
         TabIndex        =   8
         Top             =   4800
         Width           =   870
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   39
         Left            =   2640
         TabIndex        =   7
         Top             =   300
         Width           =   675
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   38
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1140
      End
   End
   Begin VB.PictureBox SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   1530
      TabIndex        =   46
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FORDACOD07.frx":29E0
      Top             =   0
   End
End
Attribute VB_Name = "FORDACOD07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
    '
    '\\\OT-1069-FL-25/11/08///Cierra el formulario
    Unload Me
    '\\\OT-1069-FL-FIN///
    '
End Sub

Private Sub Command47_Click()
   '\\\OT-1069-FL-24/11/08/// Comando Seleccionar todo
   Text43.SelStart = 0
   Text43.SelLength = Len(Text43.TEXT)
   Text43.SetFocus
   '\\\OT-1069-FL-FIN///
End Sub



Private Sub Command49_Click()
'\\\OT-1069-FL-24/11/08/// Comando COPIAR
 Clipboard.SetText Mid$(Text43.TEXT, Text43.SelStart + 1, Text43.SelLength)
 '\\\OT-1069-FL-FIN///
End Sub



Private Sub Command8_Click()
   '\\\OT-1069-FL-25/11/08/// Imprime formulario
   ANCHAC = Width
   Width = 10515
   On Error Resume Next
   PRINTFORM
   If Err Then
      On Error GoTo 0
      Call MENSERR(24, "Error de Impresión de Ficha")
   End If
   On Error GoTo 0
   Width = ANCHAC
   '\\\OT-1069-FL-FIN///
End Sub

Private Sub Form_Load()
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
End Sub

Private Sub Label2_Change()
   '\\\OT-1069-FL-25/11/08/// Imprime formulario
   CODDX$ = TRIM$(Label2)
   CICDX% = CODINT%(CODDX$)
   FORDACOD07.Label2.Caption = CODDX$
   FORDACOD07.Label3.Caption = DESCRIT0$(CICDX%)
   FORDACOD07.Label4.Caption = TIMATE%(CICDX%)
   FORDACOD07.Label18.Caption = ECOARCH("TATIMAT", MKI$(TIMATE%(CICDX%)))
   FORDACOD07.Label5.Caption = FORMATNUM$(PESUNI(CICDX%), "F10.4")
   FORDACOD07.Label6.Caption = TALLE$(CICDX%)
   FORDACOD07.Label7.Caption = COLOR$(CICDX%)
   FORDACOD07.Label8.Caption = FORMATNUM$(COSUNI(CICDX%), "F10.4")
   If DERACCE%("VERCOSTO", "Visualizar Costos y Precios") < 1 Then
      FORDACOD07.Label8.Caption = FORMATNUM$(99999999999#, "F10.4")
   End If
   FORDACOD07.Label9.Caption = PROVHABI%(CICDX%)
   FORDACOD07.Label12.Caption = rasocli(PROVHABI%(CICDX%) * (-1))
   FORDACOD07.Label11.Caption = FECHATEX(FECOSTO%(CICDX%))
   FORDACOD07.Label13.Caption = FORMATNUM$(STOCKACT(CICDX%), "F10.1")
   FORDACOD07.Label14.Caption = FORMATNUM$(STOMIN(CICDX%), "F10.1")
   FORDACOD07.Label15.Caption = FORMATNUM$(LOTESTAN(CICDX%), "F10.1")
   FORDACOD07.Label16.Caption = FORMATNUM$(LOREPOMA(CICDX%), "F10.1")
   FORDACOD07.Label17.Caption = UNIMED$(CICDX%)
   FORDACOD07.Text43.TEXT = ESPECIFIT$(CICDX%)
   '
   Call MUESTRAPLANO(0)
End Sub

Sub MUESTRAPLANO(Index)
   '
   Screen.MousePointer = 11
   '
   Picture1(Index).Visible = True
   Refresh
   '
   Image1(Index).Picture = LoadPicture("")
   Picture1(Index).Cls
   COPLANO$ = LUGRA$ + TRIM$(Label2.Caption) + ".JPG" ' LUGRA$ = la ruta estática de las imagenes (en carpeta Datos o Graficos)
   If COPLANO$ <> "" Then
      If EXISTE%(COPLANO$) > 0 Then
         On Error GoTo 90
         Image1(Index).Picture = LoadPicture(COPLANO$)
         Picture1(Index).Picture = LoadPicture("")
         Picture1(Index).Refresh
         '
         HIMAGEN = Image1(Index).Picture.Width
         VIMAGEN = Image1(Index).Picture.Height
         
         HCUADRO = Picture1(Index).Width
         VCUADRO = Picture1(Index).Height
         '
         REDUH = HIMAGEN / HCUADRO
         REDUV = VIMAGEN / VCUADRO
         REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
         Picture1(Index).Height = VCUADRO * REDUV / REDUC
         Picture1(Index).Width = HCUADRO * REDUH / REDUC
         Picture1(Index).Refresh
         '
         HCUADRO = Picture1(Index).Width
         VCUADRO = Picture1(Index).Height
         Picture1(Index).Left = (Frame1(Index).Width - HCUADRO) / 2
         Picture1(Index).Top = (Frame1(Index).Height - VCUADRO) / 2
         '
         Picture1(Index).PaintPicture Image1(Index).Picture, 0, 0, HCUADRO, VCUADRO
         Picture1(Index).Visible = True
         Picture1(Index).Refresh
      End If
   End If
   GoTo 98
   '
90 On Error GoTo 0
   Call MENSERR(24, "Imagen no Válida")
   '
98 Screen.MousePointer = 1
   '
99 End Sub


Private Sub Text43_Click()
Command49.SetFocus
End Sub

Private Sub Text43_GotFocus()
Command49.SetFocus
End Sub

Private Sub Text43_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command49.SetFocus
End Sub
