VERSION 5.00
Begin VB.Form MENUVEN 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gestión de Ventas - Cuentas Corrientes"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   5865
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4830
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3150
      Width           =   855
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
   Begin VB.CommandButton Command18 
      Caption         =   "Price"
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
      Picture         =   "MENUVEN.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Acceso a Listas de Precios de Venta"
      Top             =   1995
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0FFC0&
      Caption         =   "Gestión de Ventas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3765
      Left            =   210
      TabIndex        =   5
      Top             =   210
      Width           =   4425
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0FFC0&
         Caption         =   "Control de Gestión"
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
         Left            =   420
         TabIndex        =   13
         Top             =   2310
         Width           =   3585
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   14
            Top             =   420
            Width           =   2325
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
               Picture         =   "MENUVEN.frx":030A
               Style           =   1  'Graphical
               TabIndex        =   15
               ToolTipText     =   "Acceso al Control General de Gestión Comercial"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Control General de Gestión Comercial"
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
               TabIndex        =   16
               Top             =   105
               Width           =   1620
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0FFC0&
         Caption         =   "Operaciones Comerciales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   420
         TabIndex        =   6
         Top             =   525
         Width           =   3585
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   9
            Top             =   840
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
               Picture         =   "MENUVEN.frx":0614
               Style           =   1  'Graphical
               TabIndex        =   1
               ToolTipText     =   "Emitir o Registrar un Recibo de Cobranza"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Recibo Cobranza en Cta. Corriente"
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
               TabIndex        =   10
               Top             =   105
               Width           =   1485
            End
         End
         Begin VB.PictureBox Picture14 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   11
            Top             =   420
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
               Picture         =   "MENUVEN.frx":075E
               Style           =   1  'Graphical
               TabIndex        =   0
               ToolTipText     =   "Generar una Factura en Cuenta Corriente"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Facturación C.Cte"
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
               TabIndex        =   12
               Top             =   105
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
      Picture         =   "MENUVEN.frx":0BA0
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
      Picture         =   "MENUVEN.frx":0CEA
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
      Picture         =   "MENUVEN.frx":0FF4
      Top             =   3450
      Width           =   2010
   End
   Begin VB.Menu op_comerc 
      Caption         =   "&Operaciones Comerciales"
      Begin VB.Menu fac_ven 
         Caption         =   "Facturación de Ventas"
      End
      Begin VB.Menu recibo 
         Caption         =   "Recibo de Cobranza"
      End
      Begin VB.Menu phl1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu cons_inf 
      Caption         =   "&Consultas e Informes"
      Begin VB.Menu cta_cte 
         Caption         =   "Resumen de Cuenta Corriente"
         Begin VB.Menu scr1 
            Caption         =   "Individual por Pantalla"
         End
         Begin VB.Menu print1 
            Caption         =   "Imprimir / Exportar"
         End
      End
      Begin VB.Menu fac_pen 
         Caption         =   "Cuentas por Cobrar"
      End
      Begin VB.Menu plh20 
         Caption         =   "-"
      End
      Begin VB.Menu avan_fac 
         Caption         =   "Avance de Facturación"
         Begin VB.Menu avan_fac1 
            Caption         =   "Facturas, N.Debito y N.Credito"
            Begin VB.Menu scr3 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu print3 
               Caption         =   "Imprimir / Exportar"
            End
         End
         Begin VB.Menu avan_fac2 
            Caption         =   "Cobranza (Cuenta Corriente)"
            Begin VB.Menu scr4 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu print4 
               Caption         =   "Imprimir / Exportar"
            End
         End
      End
      Begin VB.Menu res_mens 
         Caption         =   "Resumenes Mensuales"
         Begin VB.Menu sub_iva 
            Caption         =   "Subdiario de IVA Ventas"
         End
         Begin VB.Menu sub_ven 
            Caption         =   "Subdiario de Ventas"
         End
         Begin VB.Menu iva 
            Caption         =   "-"
         End
         Begin VB.Menu cons_tipocom 
            Caption         =   "Consulta por Tipo de Comprobante"
            Begin VB.Menu ctipo_fac 
               Caption         =   "Facturas"
               Begin VB.Menu scr5 
                  Caption         =   "Por Pantalla"
               End
               Begin VB.Menu print5 
                  Caption         =   "Imprimir / Exportar"
               End
            End
            Begin VB.Menu ctipo_ndeb 
               Caption         =   "Notas de Debito"
               Begin VB.Menu scr6 
                  Caption         =   "Por Pantalla"
               End
               Begin VB.Menu print6 
                  Caption         =   "Imprimir / Exportar"
               End
            End
            Begin VB.Menu ctipo_ncre 
               Caption         =   "Notas de Credito"
               Begin VB.Menu scr7 
                  Caption         =   "Por Pantalla"
               End
               Begin VB.Menu print7 
                  Caption         =   "Imprimir / Exportar"
               End
            End
            Begin VB.Menu ctipo_reci 
               Caption         =   "Recibos de Cobranza"
               Begin VB.Menu scr9 
                  Caption         =   "Por Pantalla"
               End
               Begin VB.Menu print9 
                  Caption         =   "Imprimir / Exportar"
               End
            End
         End
         Begin VB.Menu plh21 
            Caption         =   "-"
         End
         Begin VB.Menu resmen_cta 
            Caption         =   "Resumen Mensual por Cuenta"
            Begin VB.Menu scr8 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu print8 
               Caption         =   "Imprimir / Exportar"
            End
         End
      End
      Begin VB.Menu plh2 
         Caption         =   "-"
      End
      Begin VB.Menu otros 
         Caption         =   "Otros"
         Begin VB.Menu lim_cre 
            Caption         =   "Deudas y Limite de Crédito"
            Begin VB.Menu scr2 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu print2 
               Caption         =   "Listado Impreso"
            End
         End
      End
   End
   Begin VB.Menu help 
      Caption         =   "A&yuda"
   End
End
Attribute VB_Name = "MENUVEN"
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
   RECIBO00.Show 1
   Command14.Enabled = True
End Sub


Private Sub Command18_Click()
   Command18.Enabled = False
   Call CONECRUN("LISPRE02", "Listas de Precios de Venta")
   Command18.Enabled = True
End Sub


Private Sub Command22_Click()
    Command22.Enabled = False
    FACTUR00.Show 1
    Command22.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   'Call CONECRUN("COGEVE02", "Control de Gestión Comercial")
   COGEVE00.Show 1
   Command3.Enabled = True
End Sub


Private Sub exit_Click()
   Call Command1_Click
End Sub

Private Sub fac_pen_Click()
   Call COGEVE00.Command3_Click
End Sub

Private Sub fac_ven_Click()
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
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
   End If
   Call GRATITFOR(Caption)
   '
   If Left$(APLINVO$, 6) = "RECIBO" Then
       RECIBO00.Show 1
       Call FINAL("")
     ElseIf Left$(APLINVO$, 7) = "FACTURA" Then
       FACTUR00.Show 1
       Call FINAL("")
     ElseIf Left$(APLINVO$, 8) = "LPRECIOS" Then
       Call CONECRUN("LISPRE02", "Listas de Precios de Venta")
       Call FINAL("")
     ElseIf Left$(APLINVO$, 8) = "CGVENTAS" Then
       COGEVE00.Show 1
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

Private Sub print1_Click()
   RECUEC00.Text1 = ""
   Call RECUEC00.Command3_Click
End Sub

Private Sub print2_Click()
   COGEVE00.Option2.Value = True
   Call COGEVE00.Command20_Click
End Sub

Private Sub print3_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      FORSELME.MESAN = MESASEL$
      FORSELME.Option3.Value = True
      FORSELME.Option2.Value = True
      Call COGEVE00.Command9_Click
    End If
End Sub

Private Sub print4_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      FORSELME.MESAN = MESASEL$
      FORSELME.Option4.Value = True
      FORSELME.Option2.Value = True
      Call COGEVE00.Command9_Click
    End If
End Sub

Private Sub print5_Click()
   'Falta rutina de seleccion del mes
   RESUVE00.Option2.Value = True
   Call RESUVE00.Command3_Click
End Sub

Private Sub print6_Click()
   'Falta rutina de seleccion del mes
   RESUVE00.Option2.Value = True
   Call RESUVE00.Command4_Click
End Sub

Private Sub print7_Click()
   'Falta rutina de seleccion del mes
   RESUVE00.Option2.Value = True
   Call RESUVE00.Command5_Click
End Sub

Private Sub print8_Click()
   RESUVE00.Option2.Value = True
   Call RESUVE00.Command6_Click
End Sub

Private Sub print9_Click()
   'Falta rutina de seleccion del mes
   RESUVE00.Option2.Value = True
   Call RESUVE00.Command2_Click
End Sub

Private Sub recibo_Click()
   Call Command14_Click
End Sub

Private Sub scr1_Click()
  Call COGEVE00.Command7_Click
End Sub

Private Sub scr2_Click()
  COGEVE00.Option1.Value = True
  Call COGEVE00.Command20_Click
End Sub

Private Sub scr3_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      FORSELME.MESAN = MESASEL$
      FORSELME.Option3.Value = True
      FORSELME.Option1.Value = True
      Call COGEVE00.Command9_Click
    End If
End Sub

Private Sub scr4_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      FORSELME.MESAN = MESASEL$
      FORSELME.Option4.Value = True
      FORSELME.Option1.Value = True
      Call COGEVE00.Command9_Click
    End If
End Sub

Private Sub scr5_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      MESAN = MESASEL$
      RESUVE00.Option1.Value = True
      Call RESUVE00.Command3_Click
    End If
End Sub

Private Sub scr6_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      MESAN = MESASEL$
      RESUVE00.Option1.Value = True
      Call RESUVE00.Command4_Click
    End If
End Sub

Private Sub scr7_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      MESAN = MESASEL$
      RESUVE00.Option1.Value = True
      Call RESUVE00.Command5_Click
    End If
End Sub

Private Sub scr8_Click()
   RESUVE00.Option1.Value = True
   Call RESUVE00.Command6_Click
End Sub

Private Sub scr9_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      REMEVE00.MESAN = MESASEL$
      RESUVE00.Option1.Value = True
      Call RESUVE00.Command2_Click
    End If
End Sub

Private Sub sub_iva_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      REMEVE00.MESAN = MESASEL$
      Call REMEVE00.Command9_Click
    End If
End Sub

Private Sub sub_ven_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      REMEVE00.MESAN = MESASEL$
      Call REMEVE00.Command6_Click
    End If
End Sub
