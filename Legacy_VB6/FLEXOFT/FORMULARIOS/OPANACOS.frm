VERSION 5.00
Begin VB.Form OPANACOS 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Analisis de Costos - Tablas y Configuracion"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Tablas de Validación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   4635
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Criterios y Tablas Generales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   210
         TabIndex        =   10
         Top             =   1575
         Width           =   4215
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Tabla de Unidades de Medida"
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
            Index           =   9
            Left            =   315
            TabIndex        =   16
            Top             =   735
            Width           =   3375
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Tabla de Tipos de Material"
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
            Index           =   0
            Left            =   315
            TabIndex        =   11
            Top             =   420
            Width           =   3480
         End
      End
      Begin VB.CommandButton Command20 
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
         Left            =   3255
         TabIndex        =   1
         Top             =   525
         Width           =   175
      End
      Begin VB.CommandButton Command16 
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
         Height          =   880
         Left            =   3570
         Picture         =   "OPANACOS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   525
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1050
         Width           =   3225
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1680
         TabIndex        =   3
         Text            =   " "
         Top             =   525
         Width           =   1695
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Configuración Industrial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1800
         Left            =   210
         TabIndex        =   2
         Top             =   2835
         Width           =   4215
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Categorias de Operarios Productivos"
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
            Index           =   8
            Left            =   315
            TabIndex        =   15
            Top             =   1365
            Width           =   3690
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Máquinas y Equipos de Producción"
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
            Index           =   7
            Left            =   315
            TabIndex        =   14
            Top             =   1050
            Width           =   3690
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Secciones y Centros de Costos"
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
            Index           =   6
            Left            =   315
            TabIndex        =   13
            Top             =   735
            Width           =   3480
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Depósitos y Almacenes"
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
            Index           =   5
            Left            =   315
            TabIndex        =   12
            Top             =   420
            Width           =   3795
         End
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tabla:"
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
         Left            =   -735
         TabIndex        =   7
         Top             =   630
         Width           =   2325
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   960
      Left            =   4410
      TabIndex        =   37
      Top             =   420
      Width           =   540
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Actualización Monetaria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3700
      Left            =   4935
      TabIndex        =   27
      Top             =   1320
      Width           =   4635
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ajuste por Inflación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   210
         TabIndex        =   35
         Top             =   420
         Width           =   3270
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Ajuste Automático por Inflación (Costos en Moneda Argentina)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   630
            TabIndex        =   36
            Top             =   315
            Width           =   2430
         End
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   210
         TabIndex        =   33
         Top             =   2100
         Width           =   4215
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Decim."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   3570
         TabIndex        =   28
         Top             =   420
         Width           =   855
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "0"
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
            Left            =   210
            TabIndex        =   32
            Top             =   315
            Width           =   435
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "2"
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
            Left            =   210
            TabIndex        =   31
            Top             =   525
            Width           =   435
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "4"
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
            Index           =   4
            Left            =   210
            TabIndex        =   30
            Top             =   735
            Value           =   -1  'True
            Width           =   435
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "5"
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
            Index           =   5
            Left            =   210
            TabIndex        =   29
            Top             =   945
            Width           =   435
         End
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Moneda / Coeficiente         T-C"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   34
         Top             =   1785
         Width           =   4215
      End
      Begin VB.Line Line4 
         X1              =   210
         X2              =   4410
         Y1              =   3510
         Y2              =   3510
      End
      Begin VB.Line Line3 
         X1              =   4425
         X2              =   4425
         Y1              =   1785
         Y2              =   3510
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Prov.Int"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7900
      Picture         =   "OPANACOS.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Tabla de Proveedores Internos"
      Top             =   315
      Width           =   800
   End
   Begin VB.CommandButton Command3 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6220
      Picture         =   "OPANACOS.frx":0614
      TabIndex        =   25
      ToolTipText     =   "Indices de Ajuste por Inflación"
      Top             =   315
      Width           =   800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cotiz."
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7060
      Picture         =   "OPANACOS.frx":075E
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Parametros de Cálculo - Tipo de Cambio Dolar"
      Top             =   315
      Width           =   800
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
      Height          =   855
      Left            =   8750
      Picture         =   "OPANACOS.frx":08A8
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Guarda - Cierra - Abandona la Aplicación"
      Top             =   315
      Width           =   800
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   5120
      ScaleHeight     =   75
      ScaleWidth      =   900
      TabIndex        =   8
      Top             =   350
      Width           =   960
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   9
         Top             =   105
         Width           =   12000
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fuera de uso"
      Height          =   1800
      Left            =   105
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
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
         Left            =   1260
         TabIndex        =   24
         Top             =   525
         Width           =   435
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
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
         Left            =   1260
         TabIndex        =   23
         Top             =   525
         Width           =   435
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Grupos de Venta según Tasa I.V.A."
         Enabled         =   0   'False
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
         Index           =   4
         Left            =   105
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1260
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Criterios Técnicos de Clasificación"
         Enabled         =   0   'False
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
         Index           =   1
         Left            =   105
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   945
         Width           =   3795
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Grupos Contables de Venta"
         Enabled         =   0   'False
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
         Index           =   3
         Left            =   105
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   630
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Tablas de Clasificación Técnica"
         Enabled         =   0   'False
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
         Index           =   2
         Left            =   105
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   315
         Width           =   3690
      End
   End
   Begin VB.Shape Shape1 
      Height          =   545
      Left            =   4950
      Top             =   620
      Width           =   1380
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4620
      Picture         =   "OPANACOS.frx":09F2
      Top             =   630
      Width           =   2010
   End
End
Attribute VB_Name = "OPANACOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TABLA$(16)
Dim MODUCONE$

Private Sub Check1_Click()
   Command3.Enabled = False
   If Check1.Value = 1 Then
     Command3.Enabled = True
   End If
End Sub

Private Sub Command1_Click()
   '
   Call GRACONTROL("ANACOSTO", "DECICAL", TRIM$(Str$(DECICALCOS%)))
   '
   AJUIN$ = "NO"
   If Check1.Value = 1 Then AJUIN$ = "SI"
   Call GRACONTROL$("ANACOSTO", "AJUINFLA", AJUIN$)
   '
   Call CIERRARCH("*.*")
   Hide
   Call FRESHALL
   '
End Sub

Private Sub Command16_Click()
    '
    Command16.Enabled = False
    For Index = 0 To 9
      If Option1(Index) = True Then
        Call CARTABLA(Index)
        GoTo 99
      End If
   Next Index
99 Command16.Enabled = True
   '
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call CARTAMONED
   Command2.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   '
   ' arma tabla indices de ajuste inflacion
   '
   Screen.MousePointer = 11
   U& = 0
   Dim MES$(12)
   '
   MES$(1) = "ENERO"
   MES$(2) = "FEBRERO"
   MES$(3) = "MARZO"
   MES$(4) = "ABRIL"
   MES$(5) = "MAYO"
   MES$(6) = "JUNIO"
   MES$(7) = "JULIO"
   MES$(8) = "AGOSTO"
   MES$(9) = "SETIEMBRE"
   MES$(10) = "OCTUBRE"
   MES$(11) = "NOVIEMBRE"
   MES$(12) = "DICIEMBRE"
   '
   HOO% = HOY(HOS$)
   MMFF% = Val(Mid$(HOS$, 4, 2)) + 2
   AAFF% = Val(Mid$(HOS$, 7, 2))
   If MMFF% > 12 Then
      MMFF% = MMFF% - 12
      AAFF% = AAFF% + 1
   End If
   '
   MMXX% = 12: AAXX% = 1
   U& = 0: UINAJU# = 1
   REBLA$ = REGBLAN$("INDIAJU")
   '
10 R2$ = TRIM$(Str$(AAXX%)): While Len(R2$) < 2: R2$ = "0" + R2$: Wend
   R1$ = MES$(MMXX%)
   '
   REMESA$ = R1$ + " " + R2$
   REMESB$ = MESANO$(REMESA$, DES%, HAS%)
   U& = U& + 1
   If U& <= 1 Then DES% = (-32767)
   If U& <= ULTREG&("INDIAJU") Then
       TTXX$ = REGLEIDO$("INDIAJU", U&)
     Else
       TTXX$ = REBLA$
       Call REPLA(TTXX$, REMESB$, 1, 16)
       Call REPLA(TTXX$, MKI$(DES%), 17, 2)
       Call REPLA(TTXX$, MKI$(HAS%), 19, 2)
       Call REPLA(TTXX$, MKD$(UINAJU#), 21, 8)
   End If
   Call GRAREG("INDIAJU", TTXX$, U&)
   MMXX% = MMXX% + 1
   If MMXX% > 12 Then
     MMXX% = MMXX% - 12
     AAXX% = AAXX% + 1
   End If
   If 13 * AAXX% + MMXX% <= 13 * AAFF% + MMFF% Then GoTo 10
   Call CIERRARCH("INDIAJU")
   Screen.MousePointer = 1
   '
   Call TRALOCAL("INDIAJU", "")
   VTB% = VENTABU%("INDIAJU")
   If VTB% >= 0 Then
     Call TRACENTRAL("INDIAJU", "")
   End If
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call CARTAPROINTER
   Command4.Enabled = True
End Sub


Private Sub Form_Activate()
   List1.Clear
   For U& = 1 To ULTREG&("TAMONED")
     XX$ = REGLEIDO$("TAMONED", U&)
     TTXX$ = Mid$(XX$, 2, 27) + Space$(21)
     TICAX = CVS(Mid$(XX$, 29, 4))
     Call REPLA(TTXX$, CSTRING$(MKS$(TICAX), 3, 12, 4, 2), 21, 12)
     List1.AddItem TTXX$
   Next U&
   Call CIERRARCH("TAMONED")
End Sub

'
Private Sub Form_Load()
  '
  TABLA$(0) = "TATIMAT"       ' tabla de tipos de material
  TABLA$(1) = "TACRITEC"
  TABLA$(2) = "TACLATEC"
  TABLA$(3) = "GRUCOVEN"       ' grupos contables por articulo
  TABLA$(4) = "TASIVA"
  TABLA$(5) = "TADEPOSI"
  TABLA$(6) = "SECPROD"
  TABLA$(7) = "PADMAQUI"
  TABLA$(8) = "CATOPER"
  TABLA$(9) = "UMEDIDA"
  '
  DCC$ = TRIM$(CONTROL$("ANACOSTO", "DECICAL"))
  If DCC$ = "" Then
      DECICALCOS% = 4
      Call GRACONTROL("ANACOSTO", "DECICAL", TRIM$(Str$(DECICALCOS%)))
    Else
      DECICALCOS% = Val(DCC$)
  End If
  If DECICALCOS% < 0 Or DECICALCOS% > 5 Then DECICALCOS% = 4
  Option2(DECICALCOS%).Value = True
  '
  Command3.Enabled = False
  If CONTROL$("ANACOSTO", "AJUINFLA") = "SI" Then
    Check1.Value = 1
    Command3.Enabled = True
  End If
  '
End Sub

Private Sub List1_DblClick()
   Call CARTAMONED
End Sub

'
Private Sub Option1_Click(Index As Integer)
    '
    For KKI% = 0 To 9
      If KKI% <> Index Then
        Option1(KKI%).Value = False
      End If
    Next KKI%
    '
    Text1.TEXT = TABLA(Index)
    Option1(Index).Value = True
    Text2.TEXT = Option1(Index).Caption
    Command16.Enabled = True
    Command16.SetFocus
    '
End Sub

Sub CARTABLA(Index)
   '
   If Index = 0 Then
       Call CARTATIMA
    ElseIf Index = 1 Then
       Call CARCRITEC
    ElseIf Index = 2 Then
       Call CARTACLAS
    ElseIf Index = 3 Then
       Call CARGRUCOVE
    ElseIf Index = 4 Then
       Call CARTASIVA
    ElseIf Index = 5 Then
       Call CARTADEPOSI
    ElseIf Index = 6 Then
       Call CARSECPROD
    ElseIf Index = 7 Then
       Call CARMAQUI
    ElseIf Index = 8 Then
       Call CARCATOPER
    ElseIf Index = 9 Then
       Call CARUMEDIDA
   End If
   '
End Sub
'
Private Sub Option1_DblClick(Index As Integer)
   Call CARTABLA(Index)
End Sub

Private Sub Option2_Click(Index As Integer)
   DECICALCOS% = Index
End Sub

Private Sub Text1_GotFocus()
   Command16.SetFocus
End Sub

Private Sub Text2_GotFocus()
   Command16.SetFocus
End Sub

Sub CARTATIMA()
    '
    Call BLANARCH("!TABTIMA")
    For I& = 1 To ULTREG&("TABTIMA")
       X$ = REGLEIDO$("TABTIMA", I&)
       If Asc(Left$(X$, 1)) > 0 Then
           Call REGAPP("!TABTIMA", X$)
       End If
    Next I&
    Call CIERRARCH("!TABTIMA")
    '
10  VTB% = VENTABU%("TABTIMA")
    If VTB% < 0 Then Exit Sub
    J& = 0
    For I& = 1 To ULTREG&("!TABTIMA")
       X$ = REGLEIDO$("!TABTIMA", I&)
       If Asc(Left$(X$, 1)) > 0 Then
          DPP$ = Mid$(X$, 17, 1)
          If ECOARCH$("ECODEP", DPP$) = "" Then
             Call MENSERR(24, "Depósito no Definido")
             GoTo 10
          End If
          If Asc(DPP$) = 99 Then
             Call MENSERR(24, "Depósito 99 no Válido")
             GoTo 10
          End If
          J& = J& + 1
          Call GRAREG("TABTIMA", X$, J&)
          Call GRAREG("TATIMAT", X$, J&)
       End If
    Next I&
    Call SETULTREG("TABTIMA", J&)
    Call CIERRARCH("TABTIMA")
    Call SETULTREG("TATIMAT", J&)
    Call CIERRARCH("TATIMAT")
    Call CIERRARCH("!TABTIMA")
    '
End Sub
'

Sub CARGRUCOVE()
    '
    Call BLANARCH("!GRUCOVEN")
    For I& = 1 To ULTREG&("GRUCOVEN")
        X$ = REGLEIDO$("GRUCOVEN", I&)
        If CVI(Left$(X$, 2)) > 0 Then
            Call REGAPP("!GRUCOVEN", X$)
        End If
    Next I&
    Call CIERRARCH("GRUCOVEN")
    '
    VTB% = VENTABU%("GRUCOVEN/NC/TITUPAN=Grupos Contables")
    If VTB% < 0 Then Exit Sub
    J& = 0
    For I& = 1 To ULTREG&("!GRUCOVEN")
        X$ = REGLEIDO$("!GRUCOVEN", I&)
        If CVI(Left$(X$, 2)) > 0 Then
            J& = J& + 1
            Call GRAREG("GRUCOVEN", X$, J&)
        End If
    Next I&
    Call SETULTREG("GRUCOVEN", J&)
    Call CIERRARCH("GRUCOVEN")
    '
End Sub

Sub CARTACLAS()
   '
10 X$ = REGSEL$("TACRICLA")
   If X$ = "" Then Exit Sub
   '
   CLASI% = Asc(Left$(X$, 1))
   CLASX$ = TRIM$(Mid$(X$, 2, 31))
   If CLASI% < 1 Or CLASI% > 8 Or CLASX$ = "" Then Exit Sub
   '
   CLASX$ = REPLACAR$(CLASX$, "/", "-")
   ALISX$ = TRIM$(Str$(CLASI%))
   While Len(ALISX$) < 3: ALISX$ = "0" + ALISX$: Wend
   ACLASI$ = "CLASI" + ALISX$
   Call COPYSTRU("CLASISTO", ACLASI$)
   '
   Call TRALOCAL(ACLASI$, "CLASISTO")
   VTB% = VENTABU%("CLASISTO/NC/TITUPAN=Clasificación Técnica: " + CLASX$)
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!CLASISTO")
         X$ = REGLEIDO$("!CLASISTO", U&)
         Call REPLA(X$, Chr$(U&), 1, 1)
         Call GRAREG(ACLASI$, X$, U&)
      Next U&
      Call SETULTREG(ACLASI$, U&)
      Call CIERRARCH(ACLASI$)
   End If
   Call CIERRARCH("*.*")
   GoTo 10
   '
End Sub

Sub CARCRITEC()
   Call TRALOCAL("TACRICLA", "")
   VTB% = VENTABU%("TACRICLA/NC/TITUPAN=Criterios Clasificación Técnica")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TACRICLA")
         X$ = REGLEIDO$("!TACRICLA", U&)
         Call REPLA(X$, Chr$(U&), 1, 1)
         Call GRAREG("TACRICLA", X$, U&)
      Next U&
   End If
   If ULTREG&("TACRICLA") > 8 Then
      Call SETULTREG("TACRICLA", 8)
   End If
   Call CIERRARCH("TACRICLA")
End Sub

Sub CARTASIVA()
   Call TRALOCAL("GRUPIVA", "")
   VTB% = VENTABU%("GRUPIVA/NC/TITUPAN=Tasa IVA por Categoria")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!GRUPIVA")
         X$ = REGLEIDO$("!GRUPIVA", U&)
         Call REPLA(X$, Chr$(U& - 1), 1, 1)
         Call GRAREG("GRUPIVA", X$, U&)
      Next U&
      FIN& = ULTREG&("!GRUPIVA")
      Call SETULTREG("GRUPIVA", FIN&)
   End If
   If ULTREG&("GRUPIVA") > 3 Then
      Call SETULTREG("GRUPIVA", 3)
   End If
   Call CIERRARCH("GRUPIVA")
End Sub

Sub CARTADEPOSI()
    '
    Call BLANARCH("!TADEPOSI")
    J& = 0
    For U& = 1 To ULTREG&("SUCURSAL")
        X$ = REGLEIDO$("SUCURSAL", U&)
        CODEPO% = Asc(Mid$(X$, 64, 1))
        If CODEPO% > 0 Then
           If CODEPO% <> 99 Then
              If J& < 62 Then
                 J& = J& + 1
                 Y$ = REGBLAN$("TADEPOSI")
                 Call REPLA(Y$, Chr$(CODEPO%), 1, 1)
                 Call REPLA(Y$, Left$(X$, 63), 2, 63)
                 Call REGAPP("!TADEPOSI", Y$)
              End If
           End If
        End If
    Next U&
    Call CIERRARCH("SUCURSAL")
    Call CIERRARCH("!TADEPOSI")
    '
    VTB% = VENTABU%("TADEPOSI")
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    For U& = 1 To ULTREG&("!TADEPOSI")
        X$ = REGLEIDO$("!TADEPOSI", U&)
        CODEPO% = Asc(Left$(X$, 1))
        If CODEPO% > 0 Then
           If CODEPO% < 99 Then
             If J& < 62 Then
               J& = J& + 1
               Y$ = REGBLAN$("SUCURSAL")
               Call REPLA(Y$, Mid$(X$, 2, 63), 1, 63)
               Call REPLA(Y$, Chr$(CODEPO%), 64, 1)
               Call GRAREG("SUCURSAL", Y$, J&)
               Y$ = REGBLAN$("ECODEP")
               Call REPLA(Y$, Left$(X$, 21), 1, 21)
               Call GRAREG("ECODEP", Y$, J&)
               Call GRAREG("TADEPOSI", X$, J&)
             End If
           End If
        End If
    Next U&
    Call SETULTREG("SUCURSAL", J&)
    Call CIERRARCH("SUCURSAL")
    Call SETULTREG("TADEPOSI", J&)
    Call CIERRARCH("TADEPOSI")
    '
    Y$ = REGBLAN$("ECODEP")
    Call REPLA(Y$, Chr$(99) + "Predeterminado      ", 1, 21)
    J& = J& + 1
    Call GRAREG("ECODEP", Y$, J&)
    Call SETULTREG("ECODEP", J&)
    Call CIERRARCH("ECODEP")
    Call CIERRARCH("!TADEPOSI")
    '
End Sub
'
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

Sub CARCATOPER()
   '
   Call TRALOCAL("CATOPER", "CATOPER")
   VTB% = VENTABU%("CATOPER")
   If VTB% >= 0 Then
      PUCONTROL$ = String$(255, 0)
      JJ& = 0
      For U& = 1 To ULTREG&("!CATOPER")
         X$ = REGLEIDO$("!CATOPER", U&)
         COCAT% = Asc(Left$(X$, 1))
         If COCAT% > 0 Then
           If COCAT% < 256 Then
             If Mid$(PUCONTROL$, COCAT%, 1) = Chr$(0) Then
               Mid$(PUCONTROL$, COCAT%, 1) = Chr$(1)
               JJ& = JJ& + 1
               Call GRAREG("CATOPER", X$, JJ&)
             End If
           End If
         End If
      Next U&
      Call SETULTREG("CATOPER", JJ&)
      Call CIERRARCH("CATOPER")
   End If
   '
End Sub

Sub CARUMEDIDA()
   '
   Call TRALOCAL("UNIMED", "")
10 VTB% = VENTABU%("TUMEDIDA/NC/TITUPAN=Unidades de Medida")
   If VTB% >= 0 Then
      JJ& = 0
      For U& = 1 To ULTREG&("!UNIMED")
         X$ = REGLEIDO$("!UNIMED", U&)
         If TRIM$(Left$(X$, 4)) <> "" Then
           If Len(TRIM$(Left$(X$, 4))) > 2 Then
             Call MENSERR(24, "Codigo de Unidad no Válido.\Máximo 2 Posiciones")
             GoTo 10
           End If
           Y$ = Left$(X$, 2) + Mid$(X$, 5, 28)
           JJ& = JJ& + 1
           Call GRAREG("UNIMED", X$, JJ&)
           Call GRAREG("UMEDIDA", Y$, JJ&)
         End If
      Next U&
      Call SETULTREG("UNIMED", JJ&)
      Call SETULTREG("UMEDIDA", JJ&)
   End If
   Call CIERRARCH("UNIMED")
   Call CIERRARCH("UMEDIDA")
   '
End Sub

Sub CARTAMONED()
   Call TRALOCAL("TAMONED", "")
   VTB% = VENTABU%("TAMONED/NC/TITUPAN=Monedas y Tipo de Cambio")
   If VTB% >= 0 Then
      JU& = 0
      For U& = 1 To ULTREG&("!TAMONED")
         X$ = REGLEIDO$("!TAMONED", U&)
         If TRIM$(Mid$(X$, 2, 27)) <> "" Then
           JU& = JU& + 1
           Call REPLA(X$, Chr$(JU&), 1, 1)
           Call GRAREG("TAMONED", X$, JU&)
         End If
      Next U&
      Call SETULTREG("TAMONED", JU&)
   End If
   Call CIERRARCH("TAMONED")
   '
   List1.Clear
   For U& = 1 To ULTREG&("TAMONED")
     XX$ = REGLEIDO$("TAMONED", U&)
     TTXX$ = Mid$(XX$, 2, 27) + Space$(21)
     TICAX = CVS(Mid$(XX$, 29, 4))
     Call REPLA(TTXX$, CSTRING$(MKS$(TICAX), 3, 12, 4, 2), 21, 12)
     List1.AddItem TTXX$
   Next U&
   Call CIERRARCH("TAMONED")
End Sub

Sub CARMAQUI()
    '
    J& = 0
    For U& = 1 To ULTREG&("PADMAQ")
        X$ = REGLEIDO$("PADMAQ", U&)
        COMAQ$ = TRIM$(Left$(X$, 6))
        If COMAQ$ <> "" Then
            J& = J& + 1
            Call GRAREG("!PADMAQ", X$, J&)
        End If
    Next U&
    Call SETULTREG("!PADMAQ", J&)
    Call CIERRARCH("PADMAQ")
    Call CIERRARCH("!PADMAQ")
    '
    VTB% = VENTABU%("CARMAQUI")
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    For U& = 1 To ULTREG&("!PADMAQ")
        X$ = REGLEIDO$("!PADMAQ", U&)
        COMAQ$ = TRIM$(Left$(X$, 6))
        If COMAQ$ <> "" Then
            J& = J& + 1
            Call GRAREG("PADMAQ", X$, J&)
        End If
    Next U&
    '
    Call SETULTREG("PADMAQ", J&)
    Call CIERRARCH("PADMAQ")
    Call CIERRARCH("!PADMAQ")
    '
End Sub

Sub CARTAPROINTER()
   Call TRALOCAL("TAPROINT", "")
   VTB% = VENTABU%("TAPROINT/NC/TITUPAN=Tabla de Proveedores Internos")
   If VTB% >= 0 Then
      J& = 0
      For U& = 1 To ULTREG&("!TAPROINT")
         X$ = REGLEIDO$("!TAPROINT", U&)
         If CVI(Left$(X$, 2)) > 0 Then
           J& = J& + 1
           Call GRAREG("TAPROINT", X$, J&)
         End If
      Next U&
      Call SETULTREG("TAPROINT", J&)
      Call CIERRARCH("TAPROINT")
   End If
End Sub


