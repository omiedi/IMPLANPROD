VERSION 5.00
Begin VB.Form MENUACRE 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFE0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión de Compras - Cuentas Corrientes"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   5865
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4830
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3255
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Gestión de Compras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3870
      Left            =   210
      TabIndex        =   4
      Top             =   210
      Width           =   4425
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFE0&
         Caption         =   "Control de Gestión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   420
         TabIndex        =   12
         Top             =   2415
         Width           =   3585
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2580
            TabIndex        =   13
            Top             =   420
            Width           =   2640
            Begin VB.CommandButton Command3 
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
               Left            =   0
               Picture         =   "MENUACRE.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   14
               ToolTipText     =   "Control General de Gestión de Compras"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Control General de Gestión de Compras"
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
               Left            =   630
               TabIndex        =   15
               Top             =   105
               Width           =   1905
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFE0&
         Caption         =   "Operaciones Comerciales"
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
         Left            =   420
         TabIndex        =   5
         Top             =   420
         Width           =   3585
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   8
            Top             =   1050
            Width           =   2325
            Begin VB.CommandButton Command14 
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
               Left            =   0
               Picture         =   "MENUACRE.frx":030A
               Style           =   1  'Graphical
               TabIndex        =   1
               ToolTipText     =   "Emitir Ordenes de Pago"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Orden de Pago en Cta.Corriente"
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
               Left            =   735
               TabIndex        =   9
               Top             =   105
               Width           =   1485
            End
         End
         Begin VB.PictureBox Picture14 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   10
            Top             =   525
            Width           =   2325
            Begin VB.CommandButton Command22 
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
               Left            =   0
               Picture         =   "MENUACRE.frx":0454
               Style           =   1  'Graphical
               TabIndex        =   0
               ToolTipText     =   "Registrar Facturas de Compras"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Facturación de Proveedores"
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
               Left            =   735
               TabIndex        =   11
               Top             =   60
               Width           =   1590
            End
         End
      End
   End
   Begin VB.CommandButton Command1 
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
      Height          =   800
      Left            =   4830
      Picture         =   "MENUACRE.frx":0896
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton Command29 
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
      Left            =   4830
      Picture         =   "MENUACRE.frx":09E0
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   1155
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4275
      Picture         =   "MENUACRE.frx":0CEA
      Top             =   3560
      Width           =   2010
   End
   Begin VB.Menu op_comer 
      Caption         =   "&Operaciones Comerciales"
      Begin VB.Menu fac_provee 
         Caption         =   "Factura de Proveedores"
      End
      Begin VB.Menu orpago 
         Caption         =   "Ordenes de Pago"
      End
      Begin VB.Menu plh1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu cons_inf 
      Caption         =   "&Consultas e Informes"
      Begin VB.Menu cta_cte 
         Caption         =   "Resumen Cuenta Corriente"
         Begin VB.Menu scr1 
            Caption         =   "Individual por pantalla"
         End
         Begin VB.Menu print1 
            Caption         =   "Imprimir / Exportar"
         End
      End
      Begin VB.Menu cta_porpagar 
         Caption         =   "Cuentas por Pagar"
      End
      Begin VB.Menu plh5 
         Caption         =   "-"
      End
      Begin VB.Menu saldos 
         Caption         =   "Saldos de Proveedores"
         Begin VB.Menu scr2 
            Caption         =   "Por Pantalla"
         End
         Begin VB.Menu print2 
            Caption         =   "Imprimir / Exportar"
         End
      End
      Begin VB.Menu res_men 
         Caption         =   "Resúmenes Mensuales"
         Begin VB.Menu iva 
            Caption         =   "Subdiarios"
            Begin VB.Menu sub_ivacom 
               Caption         =   "Subdiario de IVA Compras"
            End
            Begin VB.Menu sub_compras 
               Caption         =   "Subdiario de Compras"
            End
         End
         Begin VB.Menu cons_tipo 
            Caption         =   "Detalle por Tipo de Comprobante"
            Begin VB.Menu ctipo_fac 
               Caption         =   "Facturación de Compras"
            End
            Begin VB.Menu ctipo_ndeb 
               Caption         =   "Notas de Débito"
            End
            Begin VB.Menu ctipo_ncre 
               Caption         =   "Notas de Crédito"
            End
            Begin VB.Menu plh6 
               Caption         =   "-"
            End
            Begin VB.Menu ctipo_opago 
               Caption         =   "Ordenes de Pago"
            End
         End
         Begin VB.Menu plh7 
            Caption         =   "-"
         End
         Begin VB.Menu res_porcta 
            Caption         =   "Resumen Mensual por Cuenta"
         End
      End
   End
   Begin VB.Menu help 
      Caption         =   "A&yuda"
   End
End
Attribute VB_Name = "MENUACRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call FINAL("")
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   On Error Resume Next
   ORPAGO00.Hide
   On Error GoTo 0
   ORPAGO00.Show 1
   Command14.Enabled = True
End Sub

Private Sub Command2_Click()
   FIN& = ULTREG&("cuecoRR")
   For U& = 1 To FIN&
     X$ = REGLEIDO$("CUECORR", U&)
     FEX = CVI(Left$(X$, 2))
     NC1$ = FORMATNUM$(CVI(Mid$(X$, 3, 2)), "F7.0")
     NC2$ = FORMATNUM$(CVI(Mid$(X$, 5, 2)), "F7.0")
     If Val(NC2$) = 67 Then
       TTXX$ = FORMATNUM$(U&, "F6.0") + "   " + FECHATEX$(FEX) + NC1$ + NC2$ + "   " + Mid$(X$, 7, 18)
       List1.AddItem TTXX$
     End If
   Next U&
   MsgBox "COMPLETO"
End Sub

Private Sub Command22_Click()
    '
    Command22.Enabled = False
    FORFAC00.Show 1
    Command22.Enabled = True
    '
End Sub

Private Sub Command29_Click()
   Call HELPONLINE("ACRE_GES")
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   'Call CONECRUN("COGECO02", "Control de Gestión Comercial")
   COGECOM00.Show 1
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   TTXX$ = String$(512, 0)
   Call GRAREG("CUECORR", TTXX$, 2351)
   Call GRAREG("CUECOSS", TTXX$, 2351)
   Call CIERRARCH("*.*")
   Stop
   
   FIN& = ULTREG&("cuecoSS")
   For U& = 1 To FIN&
     X$ = Left$(REGLEIDO$("CUECOSS", U&), 300)
     'FEX = CVI(Left$(X$, 2))
     'NC1$ = FORMATNUM$(CVI(Mid$(X$, 3, 2)), "F7.0")
     'NC2$ = FORMATNUM$(CVI(Mid$(X$, 5, 2)), "F7.0")
     'TTXX$ = FORMATNUM$(U&, "F6.0") + "   " + FECHATEX$(FEX) + NC1$ + NC2$ + "   " + Mid$(X$, 7, 18)
     'List1.AddItem TTXX$
     Call GRAREG("CUECORR", X$, U&)
   Next U&
   Call CIERRARCH("CUECORR")
   MsgBox "COMPLETO"
End Sub

Private Sub cta_porpagar_Click()
   Call COGECOM00.Command3_Click
End Sub

Private Sub ctipo_fac_Click()
   Call FRESHALL
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
     REMECOMP.MESAN = MESASEL$
     Call CARGARESCOM
     RESUMENCO.MESAN = MESASEL$
     RESUMENCO.Option2.Value = True
     Call RESUMENCO.Command3_Click
   End If
End Sub

Private Sub ctipo_ncre_Click()
   Call FRESHALL
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
     REMECOMP.MESAN = MESASEL$
     Call CARGARESCOM
     RESUMENCO.MESAN = MESASEL$
     RESUMENCO.Option2.Value = True
     Call RESUMENCO.Command5_Click
   End If
End Sub

Private Sub ctipo_ndeb_Click()
   Call FRESHALL
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
     REMECOMP.MESAN = MESASEL$
     Call CARGARESCOM
     RESUMENCO.MESAN = MESASEL$
     RESUMENCO.Option2.Value = True
     Call RESUMENCO.Command4_Click
   End If
End Sub

Private Sub ctipo_opago_Click()
   Call FRESHALL
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
     REMECOMP.MESAN = MESASEL$
     Call CARGARESCOM
     RESUMENCO.MESAN = MESASEL$
     RESUMENCO.Option2.Value = True
     Call RESUMENCO.Command2_Click
   End If
End Sub

Private Sub exit_Click()
   Call Command1_Click
End Sub

Private Sub fac_provee_Click()
   Call Command22_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Load()
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   EJER% = 1
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
   End If
   Call GRATITFOR(Caption)
   '
   If Left$(APLINVO$, 6) = "ORPAGO" Then
      ORPAGO00.Show 1
      Call FINAL("")
      ElseIf Left$(APLINVO$, 6) = "FACTURA" Then
        FORFAC00.Show 1
        Call FINAL("")
        ElseIf Left$(APLINVO$, 6) = "CGCOMPRAS" Then
          COGECOM00.Show 1
          Call FINAL("")
   End If
   '
End Sub


Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")
End Sub

Private Sub help_Click()
   Call Command29_Click
End Sub

Private Sub orpago_Click()
   Call Command14_Click
End Sub

Private Sub print1_Click()
   Call RESCUPRO.Command3_Click
End Sub

Private Sub print2_Click()
   COGECOM00.Option2.Value = True
   Call COGECOM00.Command5_Click
End Sub

Private Sub res_porcta_Click()
   Call FRESHALL
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
     REMECOMP.MESAN = MESASEL$
     Call CARGARESCOM
     RESUMENCO.MESAN = MESASEL$
     RESUMENCO.Option2.Value = True
     Call RESUMENCO.Command6_Click
   End If
End Sub

Private Sub scr1_Click()
   Call COGECOM00.Command7_Click
End Sub

Private Sub scr2_Click()
   COGECOM00.Option1.Value = True
   Call COGECOM00.Command5_Click
End Sub

Private Sub sub_compras_Click()
   Call FRESHALL
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
     REMECOMP.MESAN = MESASEL$
     Call CARGARESCOM
     Call REMECOMP.Command6_Click
   End If
End Sub

Private Sub sub_ivacom_Click()
   Call FRESHALL
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
     REMECOMP.MESAN = MESASEL$
     Call CARGARESCOM
     Call REMECOMP.Command9_Click
   End If
End Sub
