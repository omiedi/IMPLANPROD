VERSION 5.00
Begin VB.Form FOFACOBR04 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturación de Proveedores por Boleta de Recepción"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8500
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   12000
      Begin VB.CommandButton Command12 
         Caption         =   "O/C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   120
         Picture         =   "FOFACOBR04.frx":0000
         TabIndex        =   31
         ToolTipText     =   "Visualizar Ordenes de Compra"
         Top             =   4320
         Width           =   855
      End
      Begin VB.CommandButton BOTNEXT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   11190
         Picture         =   "FOFACOBR04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Grabar - Registrar Documento en la Base de Datos"
         Top             =   4180
         Width           =   615
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8040
         TabIndex        =   27
         Text            =   " "
         ToolTipText     =   "Ingrese la Bonificación"
         Top             =   4560
         Width           =   1275
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   270
         Left            =   120
         TabIndex        =   21
         Top             =   1740
         Width           =   11685
         Begin VB.Line Line2 
            X1              =   6320
            X2              =   6320
            Y1              =   0
            Y2              =   1155
         End
         Begin VB.Line Line6 
            Index           =   0
            X1              =   1980
            X2              =   1980
            Y1              =   0
            Y2              =   1155
         End
         Begin VB.Line Line20 
            X1              =   4920
            X2              =   4920
            Y1              =   0
            Y2              =   1155
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"FOFACOBR04.frx":0614
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   11850
         End
         Begin VB.Line Line23 
            X1              =   7800
            X2              =   7800
            Y1              =   0
            Y2              =   1050
         End
         Begin VB.Line Line6 
            Index           =   1
            X1              =   9300
            X2              =   9300
            Y1              =   0
            Y2              =   1155
         End
         Begin VB.Line Line6 
            Index           =   2
            X1              =   10640
            X2              =   10640
            Y1              =   0
            Y2              =   1155
         End
         Begin VB.Line Line6 
            Index           =   3
            X1              =   3330
            X2              =   3330
            Y1              =   0
            Y2              =   1155
         End
         Begin VB.Line Line6 
            Index           =   4
            X1              =   4080
            X2              =   4080
            Y1              =   0
            Y2              =   1155
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
         Height          =   2085
         Left            =   105
         TabIndex        =   20
         Top             =   1990
         Width           =   11700
      End
      Begin VB.TextBox Text13 
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
         Height          =   300
         Left            =   1365
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1155
         Width           =   3480
      End
      Begin VB.TextBox Text6 
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
         Height          =   300
         Left            =   9825
         TabIndex        =   1
         Text            =   " "
         ToolTipText     =   "fecha REAL del Comprobante Recibido"
         Top             =   650
         Width           =   960
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1190
         Left            =   11160
         Picture         =   "FOFACOBR04.frx":06A6
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   210
         Width           =   650
      End
      Begin VB.TextBox Text11 
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
         Height          =   360
         Left            =   7995
         TabIndex        =   3
         ToolTipText     =   "Número de Comprobante"
         Top             =   1050
         Width           =   2805
      End
      Begin VB.TextBox Text9 
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
         Height          =   300
         Left            =   9825
         TabIndex        =   2
         Text            =   " "
         ToolTipText     =   "Fecha Correspondiente al Mes de Contabilización del I.V.A."
         Top             =   210
         Width           =   960
      End
      Begin VB.TextBox Text5 
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
         Height          =   300
         Left            =   2205
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   2640
      End
      Begin VB.TextBox Text4 
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
         Height          =   300
         Left            =   1365
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   750
      End
      Begin VB.TextBox Text3 
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
         Height          =   300
         Left            =   1365
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   " "
         Top             =   525
         Width           =   3480
      End
      Begin VB.TextBox Text2 
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
         Height          =   300
         Left            =   1365
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   " "
         Top             =   210
         Width           =   3480
      End
      Begin VB.TextBox Text1 
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
         Height          =   300
         Left            =   210
         MaxLength       =   5
         TabIndex        =   0
         Text            =   " "
         ToolTipText     =   "Número de Proveedor"
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Importe Neto: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7560
         TabIndex        =   29
         Top             =   4995
         Width           =   1800
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9360
         TabIndex        =   28
         Top             =   4935
         Width           =   1485
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bonificación: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6600
         TabIndex        =   26
         Top             =   4620
         Width           =   1440
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9360
         TabIndex        =   25
         Top             =   4560
         Width           =   1485
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Importe: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8400
         TabIndex        =   24
         Top             =   4275
         Width           =   960
      End
      Begin VB.Label LBLIMPO 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9360
         TabIndex        =   23
         Top             =   4200
         Width           =   1485
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C000&
         X1              =   240
         X2              =   10740
         Y1              =   1970
         Y2              =   1970
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00808080&
         X1              =   120
         X2              =   11820
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00808080&
         X1              =   11820
         X2              =   11820
         Y1              =   1740
         Y2              =   4080
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5085
         TabIndex        =   18
         ToolTipText     =   "Categoria IVA"
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Emision:"
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
         Left            =   8730
         TabIndex        =   14
         Top             =   585
         Width           =   960
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Factura Nro:"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   4950
         TabIndex        =   12
         Top             =   1050
         Width           =   2955
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Contable:"
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
         Left            =   8745
         TabIndex        =   11
         Top             =   120
         Width           =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   315
         TabIndex        =   5
         Top             =   525
         Width           =   960
      End
   End
   Begin VB.Label ONEGRA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   840
      TabIndex        =   15
      Top             =   0
      Width           =   1515
   End
   Begin VB.Label IMPIVA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   2415
      TabIndex        =   17
      Top             =   0
      Width           =   1230
   End
   Begin VB.Label IMPIVA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   2415
      TabIndex        =   16
      Top             =   0
      Width           =   1230
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Pos.IVA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   960
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu aprobar 
         Caption         =   "Aprobar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu ver 
      Caption         =   "Ver"
      Begin VB.Menu vercompra 
         Caption         =   "Ver Imagen de Ordenes de Compra"
      End
   End
End
Attribute VB_Name = "FOFACOBR04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TOTACU#(32767)

Private Sub Aprobar_Click()
    Call BOTNEXT_Click
End Sub

Private Sub BOTNEXT_Click()
  '
  For J& = 0 To 32767
    TOTACU#(J&) = 0
  Next J&
  '
  For J& = 1 To ULTREG&("!DETFACOM")
    XY$ = REGLEIDO$("!DETFACOM", J&)
    '
    QFAC1# = CVD(Mid$(XY$, 95, 8)) 'Q. Cantidad p/ Facturar
    If Abs(QFAC1#) >= 0.05 Then
       INPU1% = CVI(Mid$(XY$, 143, 2))
       If INPU1% <= 0 Then
          Call MENSERR(24, "Todos los Items deben tener Imputación Contable.")
          Call Label12_DBLClick
          Exit Sub
       End If
    End If
    '
    QPEFAC1 = CVD(Mid$(XY$, 87, 8)) 'Q. Ya Facturada
    QFAC1# = CVD(Mid$(XY$, 95, 8)) 'Q. Cantidad p/ Facturar
    PUFAC1# = CVD(Mid$(XY$, 111, 8)) 'Pr. Unit p/ Facturar
    CIXI% = CVI(Mid$(XY$, 25, 2))
    If CIXI% < 1 Then
       Call MENSERR(24, "Esta Modalidad no Aplicable\a Materiales sin Código.")
       Exit Sub
    End If
    NUOC& = CVS(Mid$(XY$, 67, 4))
    NPROX% = CVI(Mid$(XY$, 1, 2))
    '
    'NUEVO PARA GRABAR EL PRECIO UNITARIO BONIFICADO
    BN1 = COEFDESCU(TRIM$(Text12)) 'BONIFICACION
    PUNIBON# = PUFAC1# - (PUFAC1# * BN1 / 100)
    Call REPLA(XY$, MKD$(PUNIBON#), 119, 8)
    Call GRAREG("!DETFACOM", XY$, J&)
    'HASTA ACA NUEVO PARA GRABAR PRECIO UNITARIO BONIFICADO
    '
    If NUOC& > 0 Then
       REGOC$ = ULTIPRE$(NPROX%, CIXI%, NUOC&)
       CAPEDOC1# = CVS(Mid$(REGOC$, 55, 4)) 'Q. Pedida O/C
       PREUNOC1# = CVD(Mid$(REGOC$, 67, 8)) 'Pr. Unit. O/C
       COEFU = CVS(Mid$(REGOC$, 63, 4))
       If COEFU > 1 Then PREUNOC1# = PREUNOC1# / COEFU
       If PREUNOC1# > 0.005 Then
         If Abs(PUFAC1# - PREUNOC1#) > 0.00005 Then
            ATX% = COMALTER%(TRIM$(CODEXT$(CIXI%)) & ": Precio Unitario Difiere de la Orden de Compra.\\Continuar\Editar")
            If ATX% = 2 Then
              Call Label12_DBLClick
              Exit Sub
            End If
         End If
       End If
    End If
    '
    If QFAC1 > 0 Then
      If QFAC1 > QPEFAC1 Then
        ATX% = COMALTER%(TRIM$(CODEXT$(CIXI%)) & ": Cantidad Facturada Supera a la Cantidad Pendiente.\\Continuar\Editar")
        If ATX% = 2 Then
          Call Label12_DBLClick
          Exit Sub
        End If
      End If
      '
      PREUNOC1# = CVD(Mid$(REGOC$, 67, 8)) 'Pr. Unit. O/C
      If NUOC& < 1 Or PREUNOC1# < 0.00005 Then
         ' COMPARA PRECIO CON COSUNI
         COUN# = COSPESOS(CIXI%)
         If COUN# >= 0.00005 Then
           If PUNIBON# > COUN# + 0.00005 Then
              Call MENSERR(24, "Codigo '" + TRIM$(CODEXT$(CIXI%)) + "' - Precio Unitario\Superior a Costo Autorizado.\   \Imposible Registrar esta Factura.")
              Exit Sub
           End If
         End If
      End If
    End If
    '
    IMPOFAC1# = (Int(100 * PUFAC1# * QFAC1 + 0.5)) / 100
    TOTACU#(INPU1%) = IMPOFAC1 + TOTACU#(INPU1%)
    '
  Next J&
  '
  RECUEP04.LINDICE.Clear
  FORFAC04.List1.Clear
  Call BLANARCH("!CARICOM")
  ITOTA1# = 0
  For JJ% = 1 To 32767 - 1
    If TOTACU#(JJ%) >= 0.005 Then
      'LO AGREGO AL LIST DEL FORM FACTURACION PPAL.
      '
      BON1 = XVALO(Text12)
      '
      CUEB$ = FORMATNUM(JJ%, "F5")
      Z$ = Space$(128)
      Call REPLA(Z$, CUEB$, 1, 5)
      DECUE$ = CODCUE$(JJ%) + " " + DENOCUE$(JJ%)
      Call REPLA(Z$, DECUE$, 8, 30)
      'NUEVO POR BONIFICACION
      'TOTACU#(JJ%) = TOTACU#(JJ%) - (TOTACU#(JJ%) * BON1 / 100)
      '
      IMPO1$ = FORMATNUM(TOTACU#(JJ%), "F13.4")
      Call REPLA(Z$, IMPO1$, 38, 13)
      RECUEP04.LINDICE.AddItem Z$
      '
      CUE% = JJ%
      IPA# = TOTACU#(JJ%)
      ITOTA1# = ITOTA1# + IPA#
      '
      Y$ = REGBLAN$("!CARICOM")
      Call REPLA(Y$, MKI$(CUE%), 1, 2)
      Call REPLA(Y$, MKD$(IPA#), 3, 8)
      Call REGAPP("!CARICOM", Y$)
      '
    End If
  Next JJ%
  '
  For U& = 1 To RECUEP04.LINDICE.ListCount
    RECUEP04.LINDICE.ListIndex = U& - 1
    Z$ = RECUEP04.LINDICE.TEXT
    FORFAC04.List1.AddItem Z$
  Next U&
  '
  FORFAC04.List1.Refresh
  FORFAC04.ITOGRA = FORMATNUM(ITOTA1#, "F10.2")
  FORFAC04.Label36 = FORMATNUM(ITOTA1#, "F10.2")
  TOTIMPU# = Val(Label9)
  FORFAC04.Text19 = Text12
     Call FORFAC04.Text19_Change        ' para refrescar
  FORFAC04.Text19.BackColor = &HE0E0E0
  FORFAC04.Label35 = Label7
  '
  Call CIERRARCH("!DETFACOM")
  Call CIERRARCH("!CARICOM")
  Hide
  '
End Sub

Private Sub Command12_Click()
    Command12.Enabled = False
    '
    Call APERBASDAT("COMPRAS")
    SQLX$ = " SELECT Nume_Ocom, Refe_Ocom FROM OCOMENCA"
    SQLX$ = SQLX$ + " ORDER BY Nume_Ocom "
    Set VRIMPTEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    With VRIMPTEMP
        On Error Resume Next ' si hay algun error msgbox err
        .MoveFirst
        If Err Then
          On Error GoTo 0
          Call MENSERR(24, "No Aparecen Ordenes de Compra Emitidas")
          GoTo 99
        End If
        On Error GoTo 0
        JJ& = 0
        Do While (.EOF = False)
            XXY$ = REGBLAN$("INDIOCOM")
            Call REPLA(XXY$, MKS(!Nume_Ocom), 1, 4)
            Call REPLA(XXY$, !Refe_OcoM, 5, 44)
            JJ& = JJ& + 1
            Call GRAREG("!INDIOCOM", XXY$, JJ&)
            .MoveNext
        Loop
    End With
    Set VRIMPTEMP = Nothing
    Call SETULTREG("!INDIOCOM", JJ&)
    Call CIERRARCH("!INDIOCOM")
    Call FRESHECHO("!INDIOCOM")
    '
3   Call SELECHO("!INDIOCOM/Indice General de Ordenes de Compra (Pendientes-Cumplidas-Anuladas)")
    NPX& = XVALO(VALACT1$("!INDIOCOM")) ' ACA LEO DEL INDIOCOM
    If NPX& < 1 Then GoTo 99
    '
    TIDOCUM$ = "Nota de Pedido"
    NUDOCU$ = TRIM$(Str$(NPX&))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU1$ = " " + NUDOCU$ + " "
    NUDOCU2$ = " " + NUDOCU$ + "-"
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        PPXX% = InStr(XX$, "Nro.")
        If PPXX% > 0 Then
          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
        End If
        If InStr(XX$, NUDOCU1$) > 4 Or InStr(XX$, NUDOCU2$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    GoTo 99
    '
5  Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
'5   For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    '
'    REDOCU$ = REGACT$("PDOCLST")
'    URE& = ULTREG&("PDOCSPL")
'    LI& = 0: LS& = URE&
'6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'    GoTo 8
'    '
'7   Screen.MousePointer = 1
'    Call MENSERR(24, "Documento no Registrado")
'    GoTo 99
'    '
'8   TPSP$ = ""
'    Screen.MousePointer = 11
'    While L& > 1
'      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'      If DCC& < DOCUNU& Then GoTo 9
'      L& = L& - 1
'    Wend
'    '
'9   CAPAGINAS% = 0
'    For KKL& = L& To URE&
'      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'      End If
'    Next KKL&
'    '
'95  SPOOLSTRING$ = TPSP$
'    FIN& = Len(TPSP$)
'    For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).Top = 0
'      EDITFORM.Picture1(KKI%).Refresh
'      EDITFORM.Picture1(KKI%).Height = 7070
'      EDITFORM.Picture1(KKI%).Width = 11700
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    ALTUPAGINA = 7070
'    TOPEPAGINA = 0
'    PAGINACTIVA% = 0
'    CAPAGINAS% = 0
'    '
'10  PPXX% = InStr(TPSP$, Chr$(255))
'    If PPXX% > 0 Then
'      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
'      TPSP$ = Mid$(TPSP$, PPXX% + 1)
'      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
'      '
'      If Len(COPRIAT$) > 0 Then
'        COPRI% = Asc(Left$(COPRIAT$, 1))
'        ATRIP$ = Mid$(COPRIAT$, 2)
'        Call SHOWCOMMAND(COPRI%, ATRIP$)
'      End If
'      GoTo 10
'      '
'    End If
'    '
'    UI& = UI& + 1
'    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'    If InStr(XX$, TIDOCUM$) > 4 Then
'      ' If InStr(XX$, NUDOCU$) > 4 Then
'      If InStr(XX$, NUDOCU1$) > 4 Or InStr(XX$, NUDOCU2$) > 4 Then
'        DOCUNU& = CVS(Left$(XX$, 4))
'        GoTo 7
'      End If
'    End If
'    '
'    Screen.MousePointer = 1
'    PAGINACTIVA% = 0
'    For KKI% = 0 To 5
'       EDITFORM.Picture1(KKI%).Visible = False
'    Next KKI%
'    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
'    EDITFORM.Picture1(PAGINACTIVA%).Refresh
'    EDITFORM.VScroll1.Min = 0
'    EDITFORM.VScroll1.Value = 0
'    EDITFORM.VScroll1.Max = 1
'    If ALTUPAGINA > EDITFORM.Frame1.Height Then
'       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'    End If
'    ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
'    EDITFORM.HScroll1.Min = 0
'    EDITFORM.HScroll1.Value = 0
'    EDITFORM.HScroll1.Max = 1
'    If ANCHPAGINA > EDITFORM.Frame1.Width Then
'      EDITFORM.HScroll1.Max = (ANCHPAGINA - EDITFORM.Frame1.Width) / 100
'    End If
'    EDITFORM.Command3.Enabled = False
'    EDITFORM.Command5.Enabled = False
'    EDITFORM.Show 1
'    '
    GoTo 3
    '
99  Command12.Enabled = True
End Sub

Private Sub Command2_Click()
    '
    Call FORFAC04.BLAFACOM
    List1.Clear
    Unload Me
    '
End Sub

Private Sub Form_Load()
    '
    If Control("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      
    End If
    '
    Call CENTRAFORM(Me)
    '
End Sub

Private Sub Label12_DBLClick()
  '
  IMPOTOT# = 0
  '
  ATRI$ = "/NC"
  ATRI$ = ATRI$ + "/TITUPAN=Imputacion Contable"
  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(FORFAC04.Top + FORFAC04.Picture1.Top))
  ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(FORFAC04.Left + FORFAC04.Picture1.Left))
  '
  VTB% = VENTABU%("FACBORE" + ATRI$)
  '
  If VTB% >= 0 Then
    '
    List1.Clear
     For J& = 1 To ULTREG&("!DETFACOM")
        XY$ = REGLEIDO$("!DETFACOM", J&)
        NPRO1 = CVI(Mid$(XY$, 1, 2))
        CIXI1% = CVI(Mid$(XY$, 25, 2))
        '
        NBORE1 = CVS(Mid$(XY$, 27, 4))
        NUOC1& = CVS(Mid$(XY$, 67, 4))
        QPENFAC1 = CVD(Mid$(XY$, 87, 8))
        QFAC1 = CVD(Mid$(XY$, 95, 8))
        PUFAC1 = CVD(Mid$(XY$, 111, 8))
        INPU1% = CVI(Mid$(XY$, 143, 2))
        IMPOR1# = (Int(100 * PUFAC1 * QFAC1 + 0.5)) / 100
        IMPOTOT# = IMPOTOT# + IMPOR1#
        '
        NPRO2$ = FORMATNUM(NPRO1, "F5")
        CIXI2$ = CODEXT(CIXI1%)
        DESCRI2$ = Left$(DESCRIT0(CIXI1%), 15)
        NBORE2$ = FORMATNUM(NBORE1, "F5")
        NUOC2$ = FORMATNUM(NUOC1&, "F5")
        QPENFAC2$ = FORMATNUM(QPENFAC1, "F10.1")
        QFAC2$ = FORMATNUM(QFAC1, "F10.1")
        PUFAC2$ = FORMATNUM(PUFAC1, "F11.4")
        INPUCO$ = FORMATNUM(INPU1%, "F5")
        IMPOR2$ = FORMATNUM(IMPOR1#, "F10.2")
        '
        LI1$ = Space(128)
        '
        Call REPLA(LI1$, CIXI2$, 1, 16)
        Call REPLA(LI1$, DESCRI2$, 17, 10)
        Call REPLA(LI1$, NBORE2$, 29, 5)
        Call REPLA(LI1$, NUOC2$, 36, 5)
        Call REPLA(LI1$, QPENFAC2$, 43, 10)
        Call REPLA(LI1$, QFAC2$, 55, 10)
        Call REPLA(LI1$, PUFAC2$, 67, 11)
        '
        Call REPLA(LI1$, IMPOR2$, 79, 11)
        Call REPLA(LI1$, INPUCO$, 89, 5) 'LE SAQUE UNO AL 90. ERA 91
        '
        List1.AddItem LI1$
      Next J&
    '
    LBLIMPO = FORMATNUM(IMPOTOT#, "F11.2")
    '
  End If
End Sub

Private Sub Label7_CHANGE()
  '
  TT1# = Val(LBLIMPO) - Val(Label7)
  Label9 = FORMATNUM(TT1#, "F11.2")
  '
End Sub

Private Sub LBLIMPO_CHANGE()
    '
    Call Text12_Change
    TN1# = Val(LBLIMPO) - Val(Label7)
    Label9 = FORMATNUM(TN1#, "F11.2")
    '
End Sub

Private Sub List1_DBLClick()
  '
  Call Label12_DBLClick
  '
End Sub

Private Sub salir_Click()
   Call Command2_Click
End Sub

Private Sub Text1_GotFocus()
  On Error Resume Next
    Command2.SetFocus
  On Error GoTo 0
End Sub

Private Sub Text11_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text12_Change()
    '
    TN1# = Val(LBLIMPO)
    BN1 = COEFDESCU(TRIM$(Text12))
    IMTO1# = (Int(100 * TN1# * (1 - BN1 / 100) + 0.5)) / 100
    BN2# = TN1# - IMTO1#
    Label7 = FORMATNUM(BN2#, "F11.2")
    Label9 = FORMATNUM(IMTO1#, "F11.2")
    '
End Sub
Private Sub Text13_GOTFOCUS()
   Text1.SetFocus
End Sub

Private Sub Text2_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text3_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text4_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text5_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text6_GOTFOCUS()
  Text1.SetFocus
End Sub

Private Sub Text9_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
      Text1.SetFocus
   End If
End Sub
'
Function ULTIPRE$(NPROX%, CIXI%, NUOC&)
   '
   ' trae el registro de OCOMDETA correspondiente a la
   ' última compra no anulada del artículo y proveedor
   ' con precio registrado mayor que cero
   '
   Call APERBASDAT("COMPRAS")
   VUELTA% = 0
   '
20 UCO$ = REGBLAN$("OCOMDETA")
   Call REPLA(UCO$, MKI$(NPROX%), 7, 2)
   Call REPLA(UCO$, MKI$(CIXI%), 9, 2)
   '
   SQLX$ = "SELECT * FROM OCOMDETA "
   SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
   SQLX$ = SQLX$ + "AND Npro_Ocom = " & NPROX% & " "
   SQLX$ = SQLX$ + "AND Nume_Ocom = " & NUOC& & " "
   If VUELTA% = 0 Then SQLX$ = SQLX$ + "AND Pre_Unit >= 0.00005 "
   SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "
   Set OcomdetaTemp = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
   On Error Resume Next
   With OcomdetaTemp
     .MoveLast
     If Err < 1 Then
         Call REPLA(UCO$, MKS$(!Nume_Ocom), 1, 4)
         Call REPLA(UCO$, MKI$(FECHANUM(Format(!Femi_Ocom, "dd/mm/yy"))), 5, 2)
         Call REPLA(UCO$, !Uni_Mas, 11, 4)
         Call REPLA(UCO$, !Uni_Com, 59, 4)
         Call REPLA(UCO$, MKS$(!Coef_Unids), 63, 4)
         Call REPLA(UCO$, Chr$(!Mone_Emis), 122, 1)
         Call REPLA(UCO$, MKD$(!Pre_Unit), 67, 8)
         Call REPLA(UCO$, MKS$(!Desc_Item), 75, 4)
       Else
         If VUELTA% < 1 Then
           VUELTA% = VUELTA% + 1
           GoTo 20
         End If
     End If
   End With
   '
   ULTIPRE$ = UCO$
   '
End Function



Private Sub vercompra_Click()
    Call Command12_Click
End Sub
