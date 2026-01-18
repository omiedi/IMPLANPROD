VERSION 5.00
Begin VB.Form OPORPA04 
   BackColor       =   &H00C0F0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones Complementarias - Ordenes de Pago"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5865
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
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
      Left            =   735
      MaskColor       =   &H00404000&
      Picture         =   "OPORPA04.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Control B.Datos de Ordenes de Pago"
      Top             =   7770
      Width           =   650
   End
   Begin VB.CommandButton Command9 
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
      Left            =   2310
      Picture         =   "OPORPA04.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Analisis de Pagos por Fecha Acreditación de Cheque"
      Top             =   7875
      Width           =   650
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00404000&
      Height          =   6315
      Left            =   4935
      ScaleHeight     =   6255
      ScaleWidth      =   1230
      TabIndex        =   17
      Top             =   -50
      Width           =   1290
      Begin VB.CommandButton Command1 
         Caption         =   "Tbls"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   140
         Picture         =   "OPORPA04.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Acceso Directo a Tabla de Bancos y Valores de Pago"
         Top             =   1810
         Width           =   650
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   140
         Picture         =   "OPORPA04.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   2500
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   140
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   20
         Top             =   4095
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0C0C0&
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
         Height          =   690
         Left            =   140
         MaskColor       =   &H00404000&
         Picture         =   "OPORPA04.frx":09E0
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   900
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Height          =   690
         Left            =   140
         Picture         =   "OPORPA04.frx":0CEA
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   210
         Width           =   650
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -285
         Picture         =   "OPORPA04.frx":0E34
         Top             =   4430
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0F0C0&
      Caption         =   "Emision de Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Index           =   2
      Left            =   210
      TabIndex        =   13
      Top             =   3570
      Width           =   4530
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1050
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   14
         Top             =   420
         Width           =   2430
         Begin VB.CommandButton Command5 
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
            Picture         =   "OPORPA04.frx":2D56
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Listado de Ordenes de Pago Por Rango de Fechas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado de Or- denes de Pago"
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
            Left            =   840
            TabIndex        =   16
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   105
            Width           =   1425
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0F0C0&
      Caption         =   "Operaciones Especiales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   105
      Width           =   4530
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1605
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   27
         Top             =   2520
         Width           =   2430
         Begin VB.CommandButton Command10 
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
            Picture         =   "OPORPA04.frx":3060
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Compensación Cliente-Proveedor"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Compensación Cliente-Proveedor"
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
            Left            =   840
            TabIndex        =   28
            Top             =   80
            Width           =   1530
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1290
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   10
         Top             =   1995
         Width           =   2430
         Begin VB.CommandButton Command4 
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
            Picture         =   "OPORPA04.frx":336A
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Ajuste Directo de Saldos Acreedores"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Ajuste de Sal- dos Acreedores"
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
            Left            =   840
            TabIndex        =   12
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   80
            Width           =   1530
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   975
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   7
         Top             =   1470
         Width           =   2430
         Begin VB.CommandButton Command11 
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
            Left            =   1840
            Picture         =   "OPORPA04.frx":34B4
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Visualizar Imputación de Créditos y Anticipos"
            Top             =   0
            Width           =   560
         End
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
            Picture         =   "OPORPA04.frx":37BE
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Imputación de Créditos y Anticipos a Proveedores"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Imputación de Creditos"
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
            Left            =   780
            TabIndex        =   9
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   80
            Width           =   1110
         End
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   660
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   1
         Top             =   945
         Width           =   2430
         Begin VB.CommandButton Command16 
            Caption         =   "Gan/ IIBB"
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
            Left            =   1800
            TabIndex        =   30
            ToolTipText     =   "Reimpresion de retenciones de Ganancias e IIBB"
            Top             =   0
            Width           =   570
         End
         Begin VB.CommandButton Command23 
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
            Picture         =   "OPORPA04.frx":3908
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Imprimir Copia de Orden de Pago ya Emitida"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Reimprimir Comprobantes"
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
            Left            =   570
            TabIndex        =   3
            Top             =   75
            Width           =   1200
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   4
         Top             =   420
         Width           =   2430
         Begin VB.CommandButton Command25 
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
            Picture         =   "OPORPA04.frx":3F72
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Anulación de Orden de Pago Emitida"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Anulación de Orden de Pago"
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
            TabIndex        =   6
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   75
            Width           =   1530
         End
      End
   End
End
Attribute VB_Name = "OPORPA04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 Sub Command1_Click()
    Command1.Enabled = False
    If DERACCE%("SETUPACR", "Configuración de Funcionamiento ACREEDORES") = 2 Then
       Call STORPA04.Command13_Click
    End If
    '
' esto fue eliminado el 14/09/06 para ser coherente con la otra invocación desde formulario STOPPA04 - EPB
'    If (USNBR% Mod 100) <> 1 Then
'       Call COMUNI("La Opción Elegida Requiere\Privilegios de Supervisor.")
'       Command1.Enabled = True
'       Exit Sub
'    End If
'    '
'    LU$ = LUDAT$
'    Call CIERRARCH("MEPAGO")
'    Call CIERRARCH("!MEPAGO")
'    If COPYFILE%(LU$ + "MEPAGO.DAT", "C:\FLEXOFT\MEPAGO.WKF") <> 1 Then
'        Command1.Enabled = True
'        Exit Sub
'    End If
'    '
'    Dim REMEPA&(1000)
'    '
'360 ATY% = VENTABU%("COMEPAG/NC")
'    If ATY% < 0 Then GoTo 390
'    '
'361 ATX% = 0: J& = 0
'    Screen.MousePointer = 11
'    For U& = 1 To ULTREG&("!MEPAGO")
'       X$ = REGLEIDO$("!MEPAGO", U&)
'       CDX% = CVI(Left$(X$, 2))
'362    If CDX% > 0 Then
'          If CDX% <= 999 Then
'             '
'             If UCase$(Mid$(X$, 57, 4)) = "CH.3" Then
'                 Call MENSERR(24, "Valor '" + Mid$(X$, 57, 4) + "' no Permitido")
'                 ATX% = 1
'             End If
'             If UCase$(Mid$(X$, 57, 4)) = "PROP" Then
'                 Call MENSERR(24, "Valor '" + Mid$(X$, 57, 4) + "' no Permitido")
'                 ATX% = 1
'             End If
'             '
'             J& = J& + 1
'363          Call GRAREG("!MEPAGO", X$, J&)
'             If TRIM$(Mid$(X$, 3, 12)) = "" Then ATX% = 1
'             If ECOARCH$("ECUECON", TRIM$(Mid$(X$, 45, 12))) = "" Then ATX% = 1
'             CCII% = CUEINT%(TRIM$(Mid$(X$, 45, 12)))
'             Call REPLA(X$, MKI$(CCII%), 61, 2)
'             If CCII% < 1 Or (NUCUEN% > 0 And CCII% > NUCUEN%) Then ATX% = 1
'364          If TRIM$(Mid$(X$, 57, 4)) = "" Then ATX% = 1
'             '
'             If REMEPA&(CDX%) < 1 Then REMEPA&(CDX%) = J&
'             '
'          End If
'       End If
'365 Next U&
'    '
'    Call SETULTREG("!MEPAGO", J&)
'    '
'366 Screen.MousePointer = 1
'    If ATX% > 0 Then
'        For UII% = 0 To 1000: REMEPA&(UII%) = 0: Next UII%
'        Call COMUNI("Datos Incompletos o no Válidos.")
'        GoTo 360
'    End If
'    '
'    JJ& = 0
'    For UII% = 1 To 1000
'      If REMEPA&(UII%) > 0 Then
'        If REMEPA&(UII%) <= J& Then
'          X$ = REGLEIDO$("!MEPAGO", REMEPA&(UII%))
'          JJ& = JJ& + 1
'          Call GRAREG("MEPAGO", X$, JJ&)
'        End If
'      End If
'    Next UII%
'    Call SETULTREG("MEPAGO", JJ&)
'    Call CIERRARCH("MEPAGO")
'    Call CIERRARCH("!MEPAGO")
'    Call FRESHECHO("MEPAGO")
'    '
390 Command1.Enabled = True
    '
End Sub



Sub Command10_Click()
  '\\\OT-06-0112-WAR-30/03/06///
  Command10.Enabled = False
    Call COMPENSACP
  Command10.Enabled = True
  '\\\OT-06-0112-WAR-FIN///
End Sub

Sub Command11_Click()
    Command11.Enabled = False
    'Se agregó listado por fechas(opx=5)
    OPX% = COMALTER%("Puede Visualizar Aplicaciones de Diversas Formas:\\Por Orden de Pago o N/C\Por Factura o N/D\Listado por Rango de Fechas")
    If OPX% = 1 Then
        Call VAPLIOP
      ElseIf OPX% = 2 Then
        Call VAPLIFAC
      ElseIf OPX% = 3 Then
         Call LIAPLIFE
    End If
    Command11.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   If DERACCE%("SETUPCHE", "Configuración Funcionamiento Gestion Valores") = 2 Then
      OPCHEQ04.Show 1
   End If
   Command14.Enabled = True
End Sub

Private Sub Command16_Click()
     '
     Command16.Enabled = False
     '
2    OPX% = COMALTER%("Elija el Tipo de Retención\\Ingresos Brutos\Ganancias")
     If OPX% < 1 Or OPX% > 2 Then GoTo 99
     
     If OPX% = 1 Then
        Screen.MousePointer = 11
        '
        Call COPYSTRU("INDIPRES", "INDIRETE")
        '
        Call APERBASDAT("CABAN")
        SQLX$ = "SELECT * FROM CAJABANC "
        SQLX$ = SQLX$ + "WHERE TipoMovim = 'RTIB' "
        SQLX$ = SQLX$ + "Order by FechEmisi ASC, CODICOM_LAR ASC "
        '
        Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
        '
        JJ& = 0
        With CABATEMP
          On Error Resume Next
          .MoveLast
          If Err < 1 Then
            FIN& = .RecordCount
            On Error GoTo 0
            .MoveFirst
            Do While Not .EOF
              Call TRACE(20, JJ& + 1, FIN&)
              NCLIE% = !Nuprove
              IMPO# = !ValAbsComp
              If IMPO# >= 0.005 Then
                IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
                NPX& = XVALO(Mid$(!Descrimov, 14, 8))
                FEX = CVINT(!FechEmisi)
                YY$ = MKS$(NPX&) + TRIM$(RASOCLI$((-1) * NCLIE%)) + " - " + FECHATEX$(FEX)
                JJ& = JJ& + 1
                Call GRAREG("!INDIRETE", YY$, JJ&)
              End If
            .MoveNext
            Loop
          End If
        End With
        Call SETULTREG("!INDIRETE", JJ&)
        Call FRESHECHO("!INDIRETE")
        Screen.MousePointer = 1
        '
3       Call SELECHO("!INDIRETE/Indice de Retenciones de Ingresos Brutos")
        NPX& = XVALO(VALACT1$("!INDIRETE"))
        If NPX& < 1 Then GoTo 99
        '
        TIPODOC$ = "Retencion I.Brutos"
        NUME$ = TRIM$(Str$(NPX&))
        Do While Len(NUME$) < 8: NUME$ = "0" & NUME$: Loop
        NUMEDOC$ = "0001-" & NUME$
        Screen.MousePointer = 11
        '
        For UI& = ULTREG&("PDOCLST") To 1 Step -1
          XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
          If InStr(XX$, TIPODOC$) > 4 Then
            If InStr(XX$, NUMEDOC$) > 4 Then
              DOCUNU& = CVS(Left$(XX$, 4))
              GoTo 5
            End If
          End If
        Next UI&
        Screen.MousePointer = 1
        Call COMUNI("Documento no Encontrado")
        GoTo 99
        '
5       Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
        GoTo 2
        
    ElseIf OPX% = 2 Then
        Screen.MousePointer = 11
        '
        Call COPYSTRU("INDIPRES", "INDIRETE")
        '
        Call APERBASDAT("CABAN")
        SQLX$ = "SELECT * FROM CAJABANC "
        SQLX$ = SQLX$ + "WHERE TipoMovim = 'RTGAN' "
        SQLX$ = SQLX$ + "Order by FechEmisi ASC, CODICOM_LAR ASC "
        '
        Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
        '
        JJ& = 0
        With CABATEMP
          On Error Resume Next
          .MoveLast
          If Err < 1 Then
            FIN& = .RecordCount
            On Error GoTo 0
            .MoveFirst
            Do While Not .EOF
              Call TRACE(20, JJ& + 1, FIN&)
              NCLIE% = !Nuprove
              IMPO# = !ValAbsComp
              If IMPO# >= 0.005 Then
                IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
                NPX& = XVALO(Mid$(!Descrimov, 14, 8))
                FEX = CVINT(!FechEmisi)
                YY$ = MKS$(NPX&) + TRIM$(RASOCLI$((-1) * NCLIE%)) + " - " + FECHATEX$(FEX)
                JJ& = JJ& + 1
                Call GRAREG("!INDIRETE", YY$, JJ&)
              End If
            .MoveNext
            Loop
          End If
        End With
        Call SETULTREG("!INDIRETE", JJ&)
        Call FRESHECHO("!INDIRETE")
        Screen.MousePointer = 1
        '
6       Call SELECHO("!INDIRETE/Indice de Retenciones de Ganancias")
        NPX& = XVALO(VALACT1$("!INDIRETE"))
        If NPX& < 1 Then GoTo 99
        '
        TIPODOC$ = "Retencion Ganancias"
        NUME$ = TRIM$(Str$(NPX&))
        Do While Len(NUME$) < 8: NUME$ = "0" & NUME$: Loop
        NUMEDOC$ = "0000-20" + Mid$(FECHATEX$(HOY(HO$)), 7, 2) + "-" + NUME$
        Screen.MousePointer = 11
        '
        For UI& = ULTREG&("PDOCLST") To 1 Step -1
          XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
          If InStr(XX$, TIPODOC$) > 4 Then
            If InStr(XX$, NUMEDOC$) > 4 Then
              DOCUNU& = CVS(Left$(XX$, 4))
              GoTo 7
            End If
          End If
        Next UI&
        Screen.MousePointer = 1
        Call COMUNI("Documento no Encontrado")
        GoTo 99
        '
7       Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
        GoTo 2
     End If
     '
99   Command16.Enabled = True
     '
End Sub

Private Sub Command2_Click()
    '
    Hide
    '
End Sub

 Sub Command23_Click()
    '
    Command23.Enabled = False
    Screen.MousePointer = 11
    '
    If EXISTE%(LUDAT$ + "INDIORPA.RFS") < 1 Then
      Call COPYSTRU("INDIPRES", "INDIORPA")
    End If
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + "Order by FechEmisi ASC, CODICOM_LAR ASC "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    JJ& = 0
    With CABATEMP
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        FIN& = .RecordCount
        On Error GoTo 0
        .MoveFirst
        Do While Not .EOF
          Call TRACE(20, JJ& + 1, FIN&)
          NCLIE% = !Nuprove
          IMPO# = !ValAbsComp
          If IMPO# >= 0.005 Then
            IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
            NPX& = XVALO(Mid$(!CodiCom_Lar, 11, 8))
            FEX = CVINT(!FechEmisi)
            YY$ = MKS$(NPX&) + TRIM$(RASOCLI$((-1) * NCLIE%)) + " - " + FECHATEX$(FEX)
            JJ& = JJ& + 1
            Call GRAREG("!INDIORPA", YY$, JJ&)
          End If
        .MoveNext
        Loop
      End If
    End With
    Call SETULTREG("!INDIORPA", JJ&)
    Screen.MousePointer = 1
    '
3   Call SELECHO("!INDIORPA/Indice de Ordenes de Pago Emitidas")
    NPX& = XVALO(VALACT1$("!INDIORPA"))
    If NPX& < 1 Then GoTo 99
    '
    TIDOCUM$ = "Orden de Pago"
    NUDOCU$ = " " + TRIM$(Str$(NPX&)) + " -"
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        PPXX% = InStr(XX$, "Nro.")
        If PPXX% > 0 Then
          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
        End If
        If InStr(XX$, NUDOCU$) > 4 Then
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
'    For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).Top = 0
'      EDITFORM.Picture1(KKI%).Refresh
'      EDITFORM.Picture1(KKI%).Height = 7170
'      EDITFORM.Picture1(KKI%).Width = 11700
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    ALTUPAGINA = 7170
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
'      If InStr(XX$, NUDOCU$) > 4 Then
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
'    EDITFORM.Command3.Enabled = False
'    EDITFORM.Command5.Enabled = False
'    EDITFORM.Show 1
'    '
    GoTo 3
    '
99  Command23.Enabled = True
    '
End Sub

 Sub Command25_Click()
    Command25.Enabled = False
    Call ANUORPA         ' Anulación de Orden de Pago
    Command25.Enabled = True
End Sub

 Sub Command3_Click()
    '
    Command3.Enabled = False
    '//ot-07-0403
    If DERACCE%("IMPUCREP", "Imputacion Créditos/Anticipos Acreedores") >= 2 Then
      OPX% = COMALTER%("Opciones de Trabajo y Procesamiento:\\Imputar Créditos\Des-Aplicar  Pagos\Cancelar")
      If OPX% = 1 Then
          Call FRESHALL
          Call FACOCOMP04.Command3_Click
        ElseIf OPX% = 2 Then
          Call DESAPLIOPAG
      End If
    End If
    '//ot-07-0403
    Command3.Enabled = True
    '
End Sub
Sub Command4_Click()
    Command4.Enabled = False
    If DERACCE%("AJUSALDO", "Ajuste de Saldos Acreedores") >= 2 Then
        Call AJUSALP
    End If
    Command4.Enabled = True
End Sub

 Sub Command5_Click()
    
    OPXX% = ALTERNA%("Listados de Ordenes de Pago Registrados\Atraso de Pagos")
    If OPXX% = 1 Then
       GoTo 3
    ElseIf OPXX% = 2 Then
       Call ATRAPAGOS
       GoTo 99
    Else
       GoTo 99
    End If
   
3   Dim CAMP%, CUEMP%(2048), MEPA%(2048)
    Dim NOPFF(2048)
    '
    Command5.Enabled = False
    '
    'Agregue
    NUEFORMATO% = 0
    If EXISTE%(LUDAT$ + "LISOPAG.RFS") > 0 Then
        NUEFORMATO% = 1
        REBLB$ = REGBLAN$("LISOPAG")
    End If
    '
    OPXX% = ALTERNA%("O/P's por Rango de Fechas\ Anulaciones por Rango De Fechas")
    If OPXX% = 1 Or OPXX% = 2 Then
      GoTo 10
     Else
      GoTo 99
    End If
    '
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    DESDAT = Int(CDbl(CVDAT(DES%)))
    HASDAT = Int(CDbl(CVDAT(HAS%)))
    '
    Screen.MousePointer = 11
    JJ& = 0
    '
    If CAMP% <= 0 Then
      For I& = 1 To ULTREG&("MEPAGO")
         F0$ = REGLEIDO$("MEPAGO", I&)
         CAMP% = I&
         MEPA%(CAMP%) = CVI(Left$(F0$, 2))
         CUE$ = Mid$(F0$, 45, 12)
         CUEMP%(CAMP%) = CUEINT%(CUE$)
         If TRIM$(CODCUE$(CUEMP%(CAMP%))) <> TRIM$(CUE$) Then
            Call MENSERR(24, "Cuenta Contable Incorrecta\para Medio de Pago '" + TRIM$(Mid$(F0$, 3, 12)) + "'\-\Resultados Imprevisibles.")
            Call FINAL("")
         End If
     Next I&
    End If
    '
    Screen.MousePointer = 11
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + "AND OPCIMOVIM = 'OP' "
    If OPXX% = 2 Then 'SOLO ANULADAS
      SQLX$ = SQLX$ + "AND MID$(CodiCom_Lar,6,4) = 'anul'"
    End If
    SQLX$ = SQLX$ + "Order by FechConta ASC, VAL(MID(CodiCom_Lar,11)) ASC "
    '\\\OT-06-0162-WAR-04/05/06///
    SQLX$ = SQLX$ + ", NumeComp ASC "
    '\\\OT-06-0162-WAR-FIN///
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    J1& = 0: J2& = 0: REBLA$ = REGBLAN$("DETPAGC")
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      '
      Do While Not .EOF
        If !CodiCom_Lar = NUMEFA$ Then
          If !Nuprove = NCLIE% Then
            If Int(CDbl(!FechConta)) = FECHAD Then
              INETO# = INETO# + !Idebecomp - !IHABECOMP
              ITOTA# = ITOTA# + !Idebecomp - !IHABECOMP
              Call REPLA(ZZ$, MKD$(INETO#), 25, 8)
              Call REPLA(ZZ$, MKD$(IIVA#), 33, 8)
              Call REPLA(ZZ$, MKD$(IVADU#), 41, 8)
              Call REPLA(ZZ$, MKD$(RETEN#), 49, 8)
              Call REPLA(ZZ$, MKD$(IOTRO#), 57, 8)
              Call REPLA(ZZ$, MKD$(ITOTA#), 65, 8)
              Call GRAREG("DETPAGC", ZZ$, J1&)
              '
              If NUEFORMATO% = 1 Then
                TTA$ = REBLB$
                TT1$ = Left$(ZZ$, 32)
                Call REPLA(TTA$, TT1$, 1, 32)
                Call REPLA(TTA$, MKD$(ITOTA#), 65, 8)
                Call REPLA(TTA$, TRIM(!OBSERGEN), 73, 48)
                Call GRAREG("LISOPAG", TTA$, J1&)
              End If
              '
              GoTo 19
            End If
          End If
        End If
        '
        ZZ$ = REBLA$
        NUMEFA$ = !CodiCom_Lar
        NCLIE% = !Nuprove
        FECHAD = Int(CDbl(!FechConta))
        INETO# = !Idebecomp - !IHABECOMP ' !ValAbsComp
        ITOTA# = !Idebecomp - !IHABECOMP ' !ValAbsComp
        '
        Call REPLA(ZZ$, Left$(TRIM$(!TipoMovim), 2), 1, 2)
        Call REPLA(ZZ$, MKI$(CVINT(!FechConta)), 3, 2)
        Call REPLA(ZZ$, !CodiCom_Lar, 5, 18)
        Call REPLA(ZZ$, MKI$(!Nuprove), 23, 2)
        Call REPLA(ZZ$, MKD$(INETO#), 25, 8)
        Call REPLA(ZZ$, MKD$(IIVA#), 33, 8)
        Call REPLA(ZZ$, MKD$(IVADU#), 41, 8)
        Call REPLA(ZZ$, MKD$(RETEN#), 49, 8)
        Call REPLA(ZZ$, MKD$(IOTRO#), 57, 8)
        Call REPLA(ZZ$, MKD$(ITOTA#), 65, 8)
        '
        J1& = J1& + 1
        Call GRAREG("DETPAGC", ZZ$, J1&)
        '
        'AGREGUE ACA
        If NUEFORMATO% = 1 Then
          TTA$ = REBLB$
          TT1$ = Left$(ZZ$, 32)
          Call REPLA(TTA$, TT1$, 1, 32)
          Call REPLA(TTA$, MKD$(ITOTA#), 65, 8)
          Call REPLA(TTA$, TRIM(!OBSERGEN), 73, 48)
          Call GRAREG("LISOPAG", TTA$, J1&)
        End If
19    .MoveNext
      Loop
    End With
    '
    Call SETULTREG("DETPAGC", J1&)
    Call HEADERS("DETPAGC", "")
    If OPXX% = 2 Then Call HEADERS("DETPAGC", "Solo Anuladas")
    Call HEADERS("DETPAGC", "Corresponde a: " + TRIM$(PERIO$))
    '
    If NUEFORMATO% = 1 Then
       Call SETULTREG("LISOPAG", J1&)
       Call HEADERS("LISOPAG", "")
       Call HEADERS("LISOPAG", "Corresponde a: " + TRIM$(PERIO$))
       RTA1% = COMALTER%("Seleccione la Formato de Listado:\\Tradicional\Con Referencia")
       If RTA1% = 2 Then
         Call FINAL("*LISOPAG")
         GoTo 99
       End If
    End If
    Call FINAL("*LIDPAGC")
    '
99  Screen.MousePointer = 1
    Command5.Enabled = True
    '
End Sub

Sub ATRAPAGOS()
    
   Call APERBASDAT("PROVE")
   PRICLI& = 999999: ULTCLI& = 0
   SQLX$ = "SELECT * FROM Proved12 "
   SQLX$ = SQLX$ + "WHERE Stat_Cli > 0 "
   SQLX$ = SQLX$ + "ORDER BY Nume_Cli ASC"
   Set PROVETEMP = Badaprov.OpenRecordset(SQLX$, dbOpenSnapshot)
   With PROVETEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then PRICLI& = !Nume_Cli
      .MoveLast
      If Err < 1 Then ULTCLI& = !Nume_Cli
      On Error GoTo 0
   End With
   '
   If PRICLI& < 999999 Then Call REPLA(VDEF$, MKI$(PRICLI&), 1, 2)
   If ULTCLI& > 0 Then Call REPLA(VDEF$, MKI$(ULTCLI&), 3, 2)
   
10 OPX$ = OBJPLA$("RANGOPROV", VDEF$)
   NCL1% = CVI(Mid$(OPX$, 1, 2))
   NCL2% = CVI(Mid$(OPX$, 3, 2))
   
   If NCL1% < 1 Or NCL2% < 1 Then GoTo 99
   
   Call DESHASFECHA(DES%, HAS%, PERIO$)
   If DES% < 1 Then GoTo 10
   DES1# = Int(CDbl(CVDAT(DES%)))
   HAS1# = Int(CDbl(CVDAT(HAS%)))
   '
   SQLX$ = "SELECT * FROM CAJABANC"
   SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'OPAG'"
   SQLX$ = SQLX$ + " AND NuProve >= " & NCL1% & ""
   SQLX$ = SQLX$ + " AND NuProve <= " & NCL2% & ""
   SQLX$ = SQLX$ + "AND INT(CDBL(FechConta))>= " & DES1# & " "
   SQLX$ = SQLX$ + "AND INT(CDBL(FechConta))<= " & HAS1# & " "
   Set PROATRATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
      '
      With PROATRATEMP
          On Error Resume Next
          .MoveFirst
          If Err Then
             On Error GoTo 0
             Call MENSERR(24, "Ningun Registro Coincide con el Filtro SQL")
             GoTo 99
          End If
          '
          Y$ = REGBLAN("ATRAPROM")
          '
          U& = 0
          Do While Not .EOF
               PROVE% = !Nuprove
               NOPAGO$ = SAFETEXT$(!CodiCom_Cor)
               ATRACHEQQUE% = XVALO(!AtraCheque)
               ATRACCANCE% = XVALO(!AtraCance)
               ATRASSO% = XVALO(!AtraCheque) + XVALO(!AtraCance)
               Call REPLA(Y$, MKI$(PROVE%), 1, 2)
               Call REPLA(Y$, NOPAGO$, 3, 12)
               Call REPLA(Y$, MKS$(ATRACHEQQUE%), 15, 4)
               Call REPLA(Y$, MKS$(ATRACCANCE%), 19, 4)
               Call REPLA(Y$, MKS$(ATRASSO%), 23, 4)
               U& = U& + 1
               Call GRAREG("ATRAPROM", Y$, U&)
            .MoveNext
          Loop
          Call SETULTREG("ATRAPROM", U&)
          Call CIERRARCH("ATRAPROM")
          Call FILESORT("ATRAPROM", "", 3, 3, "ASC")
          Call FILESORT("ATRAPROM", "", 1, 1, "ASC")
     End With
     Call HEADERS("ATRAPROM", "")
     Call HEADERS("ATRAPROM", "Periodo: " + PERIO$)
     Call HEADERS("ATRAPROM", "Rango de Cuentas: " + TRIM$(Str$(NCL1%)) + " - " + TRIM$(Str$(NCL2%)))
     
     Call FINAL("*ATRAPRPA")
99   Screen.MousePointer = 1
   
End Sub


Private Sub Command7_Click()
    Call HELPONLINE("ORPASTA")
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    '
    Dim CAMP%, CUEMP%(2048), MEPA%(2048)
    Dim NOPFF(2048)
    Dim NORP&(8192), NPRO%(8192), FOPA%(8192), IOPA#(8192)
    '
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    JJ& = 0
    '
    If CAMP% <= 0 Then
      For I& = 1 To ULTREG&("MEPAGO")
         F0$ = REGLEIDO$("MEPAGO", I&)
         CAMP% = I&
         MEPA%(CAMP%) = CVI(Left$(F0$, 2))
         CUE$ = Mid$(F0$, 45, 12)
         CUEMP%(CAMP%) = CUEINT%(CUE$)
         If TRIM$(CODCUE$(CUEMP%(CAMP%))) <> TRIM$(CUE$) Then
            Call MENSERR(24, "Cuenta Contable Incorrecta\para Medio de Pago '" + TRIM$(Mid$(F0$, 3, 12)) + "'\-\Resultados Imprevisibles.")
            Call FINAL("")
         End If
     Next I&
    End If
    '
    CANOPFF% = 0
    CANOPAI% = 0
    FIN& = ULTREG&("CUECORR")
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      CC0$ = REGLEIDO$("CUECORR", I&)
      If CVI(Left$(CC0$, 2)) >= DES% Then
        If CVI(Left$(CC0$, 2)) <= HAS% Then
          If UCase$(Mid$(CC0$, 12, 4)) <> "ANUL" Then
            NOPA& = XVALO(Mid$(CC0$, 18, 8))
            PROV% = CVI(Mid$(CC0$, 5, 2))
            FEMI% = CVI(Mid$(CC0$, 1, 2))
            IMPX# = CVD(Mid$(CC0$, 115, 8))
            If Mid$(CC0$, 7, 2) = "FF" Then
                If NOPA <> NOPFF(CANOPFF%) Then
                  CANOPFF% = CANOPFF% + 1
                  NOPFF(CANOPFF%) = NOPA
                End If
              ElseIf Mid$(CC0$, 7, 2) = "OP" Then
                CANOPAI% = CANOPAI% + 1
                NORP&(CANOPAI%) = NOPA&
                NPRO%(CANOPAI%) = PROV%
                FOPA%(CANOPAI%) = FEMI%
                IOPA#(CANOPAI%) = IMPX#
            End If
          End If
        End If
      End If
   Next I&
   '
   NOPANT = 0: NPANT% = 0: UENCU% = 1
   ATX% = 0
   FIN& = ULTREG&("CHEQUEMI")
   For IAX& = 1 To FIN&
      Call TRACE(20, IAX&, FIN&)
      CH0$ = REGLEIDO$("CHEQUEMI", IAX&)
      CH1$ = Mid$(CH0$, 1, 2)
      CH5$ = Mid$(CH0$, 33, 8)
      CH6$ = Mid$(CH0$, 41, 4)
      CH7$ = Mid$(CH0$, 45, 2)
      CH8$ = Mid$(CH0$, 47, 2)
      '
      FEOPA% = CVI(CH8$)
      If FEOPA% >= DES% Then                ' CORRESPONDE FECHA
        If FEOPA% <= HAS% Then
          IM# = CVD(CH5$)
          If Abs(IM#) >= 0.005 Then
            NOPA& = CVS(CH6$)
            '
            For KKK% = 1 To CANOPFF%
               If NOPFF(KKK%) = NOPA Then GoTo 76 ' DESCARTA FONDO FIJO
            Next KKK%
            '
            NP% = CVI(CH1$)         ' NRO DE PROVEEDOR
            '
            VUELTA% = 0
20          For KKI% = UENCU% To CANOPAI%
              If NOPA& = NORP&(KKI%) Then
                If NP% = NPRO%(KKI%) Then
                  If FEOPA% = FOPA%(KKI%) Then
                    IOPA#(KKI%) = IOPA#(KKI%) - IM#
                    UENCU% = KKI%
                    GoTo 76
                  End If
                End If
              End If
            Next KKI%
            If VUELTA% = 0 Then
              VUELTA% = 1
              UENCU% = 1
              GoTo 20
            End If
            CANOPAI% = CANOPAI% + 1
            NORP&(CANOPAI%) = NOPA&
            NPRO%(CANOPAI%) = NP%
            FOPA%(CANOPAI%) = FEOPA%
            IOPA#(CANOPAI%) = 0 - IM#
            '
          End If
        End If
      End If
76 Next IAX&
   '
80 For KKI% = 1 To CANOPAI%
     If Abs(IOPA#(KKI%)) >= 0.005 Then
       NOPA& = NORP&(KKI%)
       NP% = NPRO%(KKI%)
       FEX = FOPA%(KKI%)
       Call MENSERR(24, "Error en Orden de Pago" + Str$(NOPA&) + "\de " + TRIM$(RASOCLI$((-1) * NP%)) + "\por " + TRIM$(CSTRING$(MKD$(IOPA#(KKI%)), 4, 12, 2, 2)) + " del " + FECHATEX$(FEX))
     End If
   Next KKI%
   '
99 Command8.Enabled = True
   Screen.MousePointer = 1
   '
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   Call CONECRUN("ANAPAGOS", "Analisis de Pagos por Cheque")
   Command9.Enabled = True
End Sub

Sub ANUORPA()       ' anulacion de ordenes de pago
    '
Call ANUORPA2
Exit Sub
    If DERACCE%("ANUORPA", "Anulación de Ordenes de Pago") < 2 Then
      Exit Sub
    End If
    '
    ' 1.- Ingresar Numero de Proveedor
    '     ****************************
    '
    FECHALIM% = XVALO(Control$("", "CIEMECOM"))
    FECHALU = Int(CDbl(CVDAT(FECHALIM%)))
    '
10  Screen.MousePointer = 1
    Call SELECHO("PROVED1")
    NC% = XVALO(VALACT1$("PROVED1"))
    If NC% < 1 Then Exit Sub
    '
    Call APERBASDAT("CABAN")
    Screen.MousePointer = 11
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'OPAG' OR (TipoMovim = 'AJUP' AND OpciMovim = 'CC')) "
    SQLX$ = SQLX$ + "AND NuProve = " & NC% & " "
    SQLX$ = SQLX$ + "AND INT(FechEmisi) > " & FECHALU & " "
    SQLX$ = SQLX$ + "Order by CODICOM_LAR ASC "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    JJ& = 0
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          NCLIE% = !Nuprove
          IMPO# = !ValAbsComp
          'If IMPO# >= 0.005 Then ' PARA QUE TAMBIEN TRAIGA LAS ANULADAS
          If UCase$(Left$(!CodiCom_Lar, 2)) = "OP" Or UCase$(Left$(!CodiCom_Lar, 2)) = "CC" Then
            IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
            NPX& = XVALO(Mid$(!CodiCom_Lar, 11, 8))
            Y$ = MKS$(NPX&) + AJUSTI$(!CodiCom_Lar, 18) + IMPOS$
            JJ& = JJ& + 1
            Call GRAREG("!SELCUECO", Y$, JJ&)
          End If
        .MoveNext
        Loop
      End If
    End With
    Call SETULTREG("!SELCUECO", JJ&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUECO") < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no Registra\Comprobantes de Pago que Puedan Anularse.")
       GoTo 10
    End If
    '
    Call CIERRARCH("!SELCUECO")
    Call FRESHECHO("!SELCUECO")
    '
12  Call SELECHO("!SELCUECO")
    If TRIM$(VALACT2$("!SELCUECO")) = "" Then GoTo 10
    '
    DOCULA$ = TRIM$(Left$(VALACT2$("!SELCUECO"), 18))
    NUOPANU& = XVALO(Mid$(DOCULA$, 13, 6)) ' para anular retenciones i.brutos
    '
    MODOANU% = COMALTER%("Modo de Anulación del Comprobante '" + DOCULA$ + "'\\Marca de Anulación\Re-Utilizar Numero")
    If MODOANU% < 1 Or MODOANU% > 2 Then GoTo 12
    '
    CCC$ = ""
    Call APERBASDAT("CABAN")
    Screen.MousePointer = 11
    With CajaBanc
19    .FindFirst "CodiCom_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & " "
20    If .NoMatch = False Then
        ' DATOS PARA LA RECONSTRUCCION
        CodiEmp% = !CodiEmpr
        CodiSuc% = !CodiSucu
        '
        If MODOANU% = 2 Then
           CCC$ = Left$(TRIM$(!CodiCom_Cor), 9) + "-nl"
          .Delete
          GoTo 19
        End If
        '
        .Edit
           CCL$ = !CodiCom_Lar
           Call REPLA(CCL$, "anul", 6, 4)
        !CodiCom_Lar = CCL$
           CCC$ = Left$(TRIM$(!CodiCom_Cor), 9) + "-nl"
        !CodiCom_Cor = CCC$
        !Status = 9
        !ValAbsComp = 0
        !Idebecomp = 0
        !IHABECOMP = 0
        !BASEIMPO = 0
        .Update
        .FindNext "CodiCom_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & ""
        GoTo 20
      End If
    End With
    '
    If CCC$ <> "" Then
      CCC$ = Left$(CCC$, 10) + "*"
      With CajaBanc
        .FindFirst "TipoMovim = 'CHRECI' AND DocumEnt LIKE '" & CCC$ & "' "
22      If .NoMatch = False Then
          .Edit
          !Status = 1     'LO PONGO DEVUELTA COMO PENDIENTE EN CARTERA DE VALORES
          !Document = ""  'QUITO EL NUMERO DE DEPOSITO ASIGNADO
          !DESTICHEQ = ""
          !REFESALCHE = ""
          !IDENSALCHE = ""
          !FEDEPOSENT = CVDAT(0) 'fecha de deposito
          .Update
          .FindNext "TipoMovim = 'CHRECI' AND DocumEnt LIKE '" & CCC$ & "' "
          GoTo 22
        End If
      End With
    End If
    '
    With SacrePen
      .FindFirst "CodiCom_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & " "
25     If .NoMatch = False Then
        .Delete
        .FindNext "CodiCom_Lar = '" & DOCULA$ & "'AND NuProve = " & NC% & ""
        GoTo 25
      End If
    End With
    '
    For U& = 1 To ULTREG("SURETIB")
      Z1$ = REGLEIDO$("SURETIB", U&)
      If CVS(Mid$(Z1$, 7, 4)) = NUOPANU& Then
        Call REPLA(Z1$, String$(16, 0), 79, 16)
        Call GRAREG("SURETIB", Z1$, U&)
      End If
    Next U&
    '
    ' RECONSTRUIR COMPOSICION DE DEUDA
    NumeClie% = 0
    NumeProv% = NC%
    SQLX$ = "SELECT * FROM ApliPago "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NC% & " "
    Set ApliPag1 = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With ApliPag1
      .FindFirst "CoDocAp_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & " "
35     If .NoMatch = False Then
        CoComCor$ = !CoDocOr_Cor
        CoComLar$ = !CoDocOr_Lar
        FechEmi = Now
        IBruTot# = !ItotOrig
        IIvaTot# = !IIVAORIG
        NuorVen% = 0
        FeVenc = Now
        ImVenc# = !ImpoApli
        '
        CajaBanc.FindFirst "CodiCom_Lar = '" & CoComLar$ & "' AND NuProve = " & NC% & " "
        If CajaBanc.NoMatch = False Then
          FechEmi = CajaBanc!FechEmisi
          FeVenc = CajaBanc!FechEmisi
        End If
        '
        Call GRACREPEN
        .FindNext "CoDocAp_Lar = '" & DOCULA$ & "'AND NuProve = " & NC% & ""
        GoTo 35
      End If
    End With
    '
    SQLX$ = "NuProve = " & NC% & "  AND CoDocAp_Lar = '" & DOCULA$ & "'"
    On Error Resume Next
    CueCorri.Execute "DELETE * FROM ApliPago WHERE " & SQLX$
    On Error GoTo 0
    '
    Screen.MousePointer = 1
    Call COMUNI("ANULACIóN COMPLETA.\  \No Fue Posible Reconstruir\las Fechas de Vencimiento Originales.")
    RECUEP04.Text1 = TRIM$(Str$(NC%))
    RECUEP04.Show 1
    '
    Call CIERRARCH("CUECORR")
    Call CIERRARCH("PROVED2")
    '
    GoTo 10
    '
End Sub
'
Sub ANUORPA2()       ' anulacion de ordenes de pago
    '
    If DERACCE%("ANUORPA", "Anulación de Ordenes de Pago") < 2 Then
      Exit Sub
    End If
    '
    ' 1.- Ingresar Numero de Proveedor
    '     ****************************
    '
    FECHALIM% = XVALO(Control$("", "CIEMECOM"))
    FECHALU = Int(CDbl(CVDAT(FECHALIM%)))
    '
10  Screen.MousePointer = 1
    Call SELECHO("PROVED1")
    NC% = XVALO(VALACT1$("PROVED1"))
    If NC% < 1 Then Exit Sub
    '
    Call APERBASDAT("CABAN")
    Screen.MousePointer = 11
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'OPAG' OR (TipoMovim = 'AJUP' AND OpciMovim = 'CC')) "
    SQLX$ = SQLX$ + "AND NuProve = " & NC% & " "
    SQLX$ = SQLX$ + "AND INT(FechEmisi) > " & FECHALU & " "
    SQLX$ = SQLX$ + "Order by CODICOM_LAR ASC "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    JJ& = 0
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          NCLIE% = !Nuprove
          IMPO# = !ValAbsComp
          'If IMPO# >= 0.005 Then ' PARA QUE TAMBIEN TRAIGA LAS ANULADAS
          If UCase$(Left$(!CodiCom_Lar, 2)) = "OP" Or UCase$(Left$(!CodiCom_Lar, 2)) = "CC" Then
            IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
            NPX& = XVALO(Mid$(!CodiCom_Lar, 11, 8))
            Y$ = MKS$(NPX&) + AJUSTI$(!CodiCom_Lar, 18) + IMPOS$
            JJ& = JJ& + 1
            Call GRAREG("!SELCUECO", Y$, JJ&)
          End If
        .MoveNext
        Loop
      End If
    End With
    Call SETULTREG("!SELCUECO", JJ&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUECO") < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no Registra\Comprobantes de Pago que Puedan Anularse.")
       GoTo 10
    End If
    '
    Call CIERRARCH("!SELCUECO")
    Call FRESHECHO("!SELCUECO")
    '
12  Call SELECHO("!SELCUECO")
    If TRIM$(VALACT2$("!SELCUECO")) = "" Then GoTo 10
    '
    DOCULA$ = TRIM$(Left$(VALACT2$("!SELCUECO"), 18))
    DOCULB$ = DOCULA$
    If Left$(DOCULA$, 2) = "OP" Then
       DOCULB$ = "AS" + Mid$(DOCULA$, 3)
    End If
    NUOPANU& = XVALO(Mid$(DOCULA$, 13, 6)) ' para anular retenciones i.brutos
    '
    MODOANU% = COMALTER%("Modo de Anulación del Comprobante '" + DOCULA$ + "'\\Marca de Anulación\Re-Utilizar Numero")
    If MODOANU% < 1 Or MODOANU% > 2 Then GoTo 12
    '
    CCC$ = ""
    Call APERBASDAT("CABAN")
    Screen.MousePointer = 11
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (CodiCom_Lar = '" & DOCULA$ & "' OR CodiCom_Lar = '" & DOCULB$ & "')AND NuProve = " & NC% & " "
    Set CAJABANC1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
    '
    On Error Resume Next
    CAJABANC1.MoveFirst
    If Err Then
      On Error GoTo 0
      Call MENSERR(24, "Imposible Anular la O/Pago Elegida")
      GoTo 10
    End If
    CCC$ = Left$(TRIM$(CAJABANC1!CodiCom_Cor), 9) + "-nl"
    CCCX$ = Left$(TRIM$(CAJABANC1!CodiCom_Cor), 10) + "*"
    '
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI' AND DocumEnt LIKE '" & CCCX$ & "' "
    Set CAJABANC2 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
    '
    SQLX$ = "SELECT * FROM SACREPEN "
    SQLX$ = SQLX$ + "WHERE (CodiCom_Lar = '" & DOCULA$ & "' OR CodiCom_Lar = '" & DOCULB$ & "') AND NuProve = " & NC% & " "
    Set SACREPEN2 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
    '
    With CAJABANC1
19    .FindFirst "(CodiCom_Lar = '" & DOCULA$ & "' OR CodiCom_Lar = '" & DOCULB$ & "') AND NuProve = " & NC% & " "
20    If .NoMatch = False Then
        ' DATOS PARA LA RECONSTRUCCION
        CodiEmp% = !CodiEmpr
        CodiSuc% = !CodiSucu
        '
        If MODOANU% = 2 Then
           CCC$ = Left$(TRIM$(!CodiCom_Cor), 9) + "-nl"
          .Delete
          GoTo 19
        End If
        '
        .Edit
           CCL$ = !CodiCom_Lar
           Call REPLA(CCL$, "anul", 6, 4)
        !CodiCom_Lar = CCL$
           CCC$ = Left$(TRIM$(!CodiCom_Cor), 9) + "-nl"
        !CodiCom_Cor = CCC$
        !Status = 9
        !ValAbsComp = 0
        !Idebecomp = 0
        !IHABECOMP = 0
        !BASEIMPO = 0
        .Update
        .FindNext "(CodiCom_Lar = '" & DOCULA$ & "' OR CodiCom_Lar = '" & DOCULB$ & "') AND NuProve = " & NC% & ""
        GoTo 20
      End If
    End With
    '
    If CCC$ <> "" Then
      CCC$ = Left$(CCC$, 10) + "*"
      With CAJABANC2
        .FindFirst "TipoMovim = 'CHRECI' AND DocumEnt LIKE '" & CCC$ & "' "
22      If .NoMatch = False Then
          .Edit
          !Status = 1     'LO PONGO DEVUELTA COMO PENDIENTE EN CARTERA DE VALORES
          !Document = ""  'QUITO EL NUMERO DE DEPOSITO ASIGNADO
          !DESTICHEQ = ""
          !REFESALCHE = ""
          !IDENSALCHE = ""
          !FEDEPOSENT = CVDAT(0) 'fecha de deposito
          !FECHENTRE = CVDAT(0)  'fecha de ENTREGA
          .Update
          .FindNext "TipoMovim = 'CHRECI' AND DocumEnt LIKE '" & CCC$ & "' "
          GoTo 22
        End If
      End With
    End If
    '
    With SACREPEN2
      .FindFirst "(CodiCom_Lar = '" & DOCULA$ & "' OR CodiCom_Lar = '" & DOCULB$ & "') AND NuProve = " & NC% & " "
25     If .NoMatch = False Then
        .Delete
        .FindNext "(CodiCom_Lar = '" & DOCULA$ & "' OR CodiCom_Lar = '" & DOCULB$ & "') AND NuProve = " & NC% & ""
        GoTo 25
      End If
    End With
    '
    For U& = 1 To ULTREG("SURETIB")
      Z1$ = REGLEIDO$("SURETIB", U&)
      If CVS(Mid$(Z1$, 7, 4)) = NUOPANU& Then
        Call REPLA(Z1$, String$(16, 0), 79, 16)
        Call GRAREG("SURETIB", Z1$, U&)
      End If
    Next U&
    '
    ' RECONSTRUIR COMPOSICION DE DEUDA
    NumeClie% = 0
    NumeProv% = NC%
    SQLX$ = "SELECT * FROM ApliPago "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NC% & " "
    Set ApliPag1 = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With ApliPag1
      .FindFirst "CoDocAp_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & " "
35     If .NoMatch = False Then
        CoComCor$ = !CoDocOr_Cor
        CoComLar$ = !CoDocOr_Lar
        FechEmi = Now
        IBruTot# = !ItotOrig
        IIvaTot# = !IIVAORIG
        NuorVen% = 0
        FeVenc = Now
        ImVenc# = !ImpoApli
        '
        CajaBanc.FindFirst "CodiCom_Lar = '" & CoComLar$ & "' AND NuProve = " & NC% & " "
        If CajaBanc.NoMatch = False Then
          FechEmi = CajaBanc!FechEmisi
          FeVenc = CajaBanc!FechEmisi
        End If
        '
        Call GRACREPEN
        .FindNext "CoDocAp_Lar = '" & DOCULA$ & "'AND NuProve = " & NC% & ""
        GoTo 35
      End If
    End With
    '
    SQLX$ = "NuProve = " & NC% & "  AND CoDocAp_Lar = '" & DOCULA$ & "'"
    On Error Resume Next
    CueCorri.Execute "DELETE * FROM ApliPago WHERE " & SQLX$
    On Error GoTo 0
    '
    Screen.MousePointer = 1
    Call COMUNI("ANULACIóN COMPLETA.\  \No Fue Posible Reconstruir\las Fechas de Vencimiento Originales.")
    RECUEP04.Text1 = TRIM$(Str$(NC%))
    RECUEP04.Show 1
    '
    Call CIERRARCH("CUECORR")
    Call CIERRARCH("PROVED2")
    '
    GoTo 10
    '
End Sub
'
Private Sub Form_Load()
   If EXISTE%("ANAPAGOS.EXE") < 1 Then
     Command9.Visible = False
   End If
End Sub

Sub AJUSALP()
    '
    MODOMONEI% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(Control$("GESACRED", "MODOMONE"))), 8)
    If MODOMONE$ = "MULTIMON" Then
       MODOMONEI% = 2
    End If
    '
    If MODOMONEI% = 2 Then
      OPXX% = ALTERNA%("Ajuste Cuenta Corriente\Saldo Factura Moneda Extranjera")
      If OPXX% = 1 Then
          GoTo 3
        ElseIf OPXX% = 2 Then
          Call AJUSALDOL
          Exit Sub
        Else
          Exit Sub
      End If
    End If
    '
3   Dim FVIK%(10), SD#(10), IMC#(10)
    HOII% = HOY(HO$)
    FAJUI% = HOII%
    '
5   VDEF$ = MKI$(NC%) + MKI$(FAJUI%)
    Screen.MousePointer = 1
    CLSS$ = OBJPLA$("SELPROAJU", VDEF$)
    If CLSS$ = "" Or CLSS$ = "<ESC>" Then GoTo 99
    '
    NC% = CVI(CLSS$)
    If NC% < 1 Then GoTo 5
    CLI$ = RASOCLI$((-1) * NC%)
    FAJUI% = CVI(Mid$(CLSS$, 3, 2))
    If FAJUI% < 1 Or FAJUI% > HOII% Then GoTo 5
    '
    FECHALIM% = XVALO(Control$("", "CIEMECOM"))
    If FAJUI% <= FECHALIM% Then
        Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
        GoTo 5
    End If
    '
    Screen.MousePointer = 11
    '
    SARRASFE# = 0: SALFINAL# = 0: SALDEUTO# = 0
    '
'2.- Determina Saldo por arrastre de Facturas y Debitos
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where NuProve = " & NC% & " "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "or TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + "or TipoMovim = 'AJUP') "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
        On Error GoTo 0
        GoTo 100
      End If
      On Error GoTo 0
      Do While Not .EOF
        IFAC# = !IHABECOMP - !Idebecomp
        FF% = CVINT(!FechConta)
        SALFINAL# = SALFINAL# + IFAC#
        If FF% <= FAJUI% Then
          SARRASFE# = SARRASFE# + IFAC#
        End If
      .MoveNext
      Loop
    End With
    '
'3.- Determina saldo por composicion de deuda
    SQLX$ = "SELECT * FROM SACREPEN "
    SQLX$ = SQLX$ + "where NuProve = " & NC% & " "
    '
    Set CODEUTEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With CODEUTEMP
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
        On Error GoTo 0
        GoTo 100
      End If
      On Error GoTo 0
      Do While Not .EOF
        SALDEUTO# = SALDEUTO# + !SaldAcre - !SaldDebe
      .MoveNext
      Loop
    End With
    '
100 SALAJU# = SARRASFE#
    VDEF$ = CLSS$ + MKD$(SARRASFE#) + MKD$(SALAJU#) + MKI$(0) + Space$(256) + MKI$(0)
    OBJE$ = "CARAJUPRO"
    If MODOMONEI% = 2 Then
       OBJE$ = "CARAJUPRO-DOL"
       VDEF$ = VDEF$ + MKI$(1) + MKS$(1)
    End If
102 CLSS$ = OBJPLA$(OBJE$, VDEF$)
    If CLSS$ = "" Or CLSS$ = "<ESC>" Then GoTo 5
    SALAJU# = CVD(Mid$(CLSS$, 13, 8))
    ICONI% = CVI(Mid$(CLSS$, 21, 2))
    If ICONI% < 1 Or ECOARCH$("CUECON", MKI$(ICONI%)) = "" Then
      Call MENSERR(24, "Falta Imputación Contable Ajuste")
      VDEF$ = CLSS$
      GoTo 102
    End If
    If ICONI% = CUECOCLI%((-1) * NC%) Then
      Call MENSERR(24, "Cuenta Contable no Válida")
      VDEF$ = CLSS$
      GoTo 102
    End If
    '
    DIFESAL# = (Int(100 * (SALAJU# - SARRASFE#) + 0.5)) / 100
    DIFEDEU# = (Int(100 * (SALFINAL# - SALDEUTO# + DIFESAL#) + 0.5)) / 100
    '
    If Abs(DIFESAL#) < 0.005 Then
      If Abs(DIFEDEU#) < 0.005 Then
        Call MENSERR(24, "No hay Ajuste")
        GoTo 5
      End If
    End If
    '
    MONAJU% = 1 ' PESOS
    COTIMONEM# = 1
    If MODOMONEI% = 2 Then
      MONAJU% = CVI(Mid$(CLSS$, 281, 2))
      If MONAJU% < 1 Or MONAJU% > 6 Or ECOARCH$("TAMONED", MKI$(MONAJU%)) = "" Then
         Call MENSERR(24, "Moneda no Válida")
         VDEF$ = CLSS$
         GoTo 102
      End If
      '
      COTIMONEM# = CVS(Mid$(CLSS$, 283, 4))
      If COTIMONEM# < 0.01 Then
         Call MENSERR(24, "Falta Cotización")
         VDEF$ = CLSS$
         GoTo 102
      End If
      If MONAJU% = 1 Then
         If Abs(COTIMONEM# - 1) > 0.00001 Then
            Call MENSERR(24, "Cotización no Válida")
            VDEF$ = CLSS$
            GoTo 102
         End If
      End If
    End If
    '
    FEX = FAJUI%
    FETX$ = FECHATEX$(FEX)
    NRECIB$ = "AS-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'AJUP' "
    SQLX$ = SQLX$ + "AND CodiCom_Cor LIKE '" & NRECIB$ + "*' "
    SQLX$ = SQLX$ + "ORDER BY CodiCom_Cor DESC"
    '
    ORMA% = (-1)
    Set RECITEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    With RECITEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        ORMA% = XVALO(Mid$(!CodiCom_Cor, 11, 2))
      End If
    End With
    ORMX$ = TRIM$(Str$(ORMA% + 1))
    While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
    NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
    '
    MONEMIA% = MONEMI%
    '
20  NC1% = (-1) * NC%
    '
    MONEMI% = MONAJU%
    CoComCor$ = NRECIB$
    CoComLar$ = NRECIB$
    CodiEmp% = 0
    CodiSuc% = SUC%
    NumeClie% = 0
    NumeProv% = NC%
    FechEmi = CVDAT(FAJUI%)
    '
    Referen$ = RASOCLI$(NC1%)
    FeContab = CVDAT(FAJUI%)
    FeAcredi = CVDAT(FAJUI%)
    NumeTal = 0
    NumeCom = XVALO(Mid$(NRECIB$, 4, 6))
    Observa$ = RASOCLI$(NC1%)
    '
    CuentaCo% = CUECOINT%(CUEMADRE$(NC1%))
    TiMovi = "AJUP"            ' AJUSTE DE SALDO PROVEEDORES
    OpMovi = "AS"              ' AJUSTE DE SALDO
    VaMovi = "CCORR"           ' CUENTA CORRIENTE
    Descrip = "Aj.Saldo - " + RASOCLI$(NC1%)
    ImpoDeb# = (-1) * DIFESAL#
    ValoDola# = ROUND#((-1) * DIFESAL# / COTIMONEM#)
    Call GRACABAN
    '
'4.- Graba la Contra-Partida en ICONI%
    CuentaCo% = ICONI%
    TiMovi = "AJUSALPR"        ' AJUSTE DE SALDO PROVEEDORES
    OpMovi = "AS"              ' AJUSTE DE SALDO
    VaMovi = "NO-IMP"          ' NO IMPUTADO
    Descrip = "Aj.Saldo - " + RASOCLI$(NC1%)
    ImpoDeb# = DIFESAL#
    Call GRACABAN
    '
'5.- Grabar Vencimientos del Ajuste ******************************************
    SIGNO% = 1
    IBruTot# = DIFESAL#
    IIvaTot# = 0
    '
    NuorVen% = 1
    FeVenc = CVDAT(FAJUI%)
    ImVenc# = DIFEDEU#
    Call GRACREPEN
    '
    CajaBanc.Close
    Call COMUNI("Ajuste Completado")
    GoTo 5
    '
99  Call CIERRARCH("*.*")
    '
End Sub
'
Sub AJUSALDOL()
    '
    IAJU% = CUEINT%(Control$("", "CUEIVA-7"))
    If IAJU% < 1 Then
       Call MENSERR(24, "Imposible Registrar Esta Transacción.\   \Falta Definir Cuenta Contable\Correspondiente a Diferencia de Cambio.")
       If DERACCE%("SETUPACR", "Configuración de Funcionamiento ACREEDORES") = 2 Then
          OPFACOM04.Show 1
       End If
       Call COMUNI("Re-Inicie Ahora la Aplicación")
       Call FINAL("")
    End If
    '
10  Call SELECHO("ACREDOR")
    NC% = XVALO(VALACT1$("ACREDOR"))
    If NC% < 1 Then Exit Sub
    '
    SQLX$ = "SELECT * FROM SACREPEN "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NC% & " "
    SQLX$ = SQLX$ + "AND (abs(SaldAcre) + ABS(SaldDebe))>= 0.005 "
    SQLX$ = SQLX$ + "AND (ABS(SaldDebe_Dola)+ ABS(SaldAcre_Dola)) < 0.005 "
    SQLX$ = SQLX$ + "ORDER BY FechEmisi ASC;"
    Set FacPenTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    SUDEUDOL# = 0: JJ& = 0
    With FacPenTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        .MoveFirst
        Do While Not .EOF
          SALDO# = !SaldAcre - !SaldDebe
          IMPOS$ = NUMSTRI$(SALDO#, 14, 2, 1)
          NPX& = XVALO(Mid$(!CodiCom_Lar, 11, 8))
          Y$ = MKS$(NPX&) + AJUSTI$(!CodiCom_Lar, 18) + IMPOS$
          JJ& = JJ& + 1
          Call GRAREG("!SELCUERE", Y$, JJ&)
        .MoveNext
        Loop
      End If
    End With
    Call SETULTREG("!SELCUERE", JJ&)
    Call FRESHECHO("!SELCUERE")
    '
    If JJ& < 1 Then
      Call COMUNI("La Cuenta Elegida no Presenta\Comprobantes en Moneda Extranjera\que Puedan Ajustarse.")
      GoTo 10
    End If
    '
    Call SELECHO("!SELCUERE")
    REGCUEC& = XVALO(VALACT1$("!SELCUERE"))
    If XVALO(REGCUEC&) < 1 Then
        GoTo 10
    End If
    '
    DOCUCABE$ = TRIM$(Left$(VALACT2$("!SELCUERE"), 18))
    DIFEX# = XVALO(Mid$(VALACT2$("!SELCUERE"), 19, 14))
    If COMALTER%("Confirme el Ajuste a CERO del Saldo\en Pesos del Documento '" + DOCUCABE$ + "'.\\Confirmar\Cancelar") = 1 Then
        HOII% = HOY(FETX$)
        NRECIB$ = "AS-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
        '
        SQLX$ = "SELECT * FROM CAJABANC "
        SQLX$ = SQLX$ + "WHERE TipoMovim = 'AJUP' "
        SQLX$ = SQLX$ + "AND CodiCom_Cor LIKE '" & NRECIB$ + "*' "
        SQLX$ = SQLX$ + "ORDER BY CodiCom_Cor DESC"
        '
        ORMA% = (-1)
        Set RECITEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
        With RECITEMP
          On Error Resume Next
          .MoveFirst
          If Err < 1 Then
            ORMA% = XVALO(Mid$(!CodiCom_Cor, 11, 2))
          End If
        End With
        ORMX$ = TRIM$(Str$(ORMA% + 1))
        While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
        NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
        '
20      NC1% = (-1) * NC%
        MONEMIA% = MONEMI%
        MONEMI% = 1
        DIFE1# = DIFEX#
        CUECO% = CUECOCLI%(NC1%)
        HOII% = HOY(HOS$)
        ' graba ajuste diferencia de cambio
        CoComCor$ = NRECIB$
        CoComLar$ = NRECIB$
        Referen$ = RASOCLI$((-1) * NC%)
        FeContab = CVDAT(HOII%)
        FeAcredi = CVDAT(HOII%)
        FechEmi = CVDAT(HOII%)
        NumeCom = XVALO(Mid$(NRECIB$, 4, 6))
        CuentaCo% = CUECO%
        NumeClie% = 0
        NumeProv% = NC%
        TiMovi = "AJUP"            ' AJUSTE DE SALDO DEUDORES
        OpMovi = "AS"              ' AJUSTE DE SALDO
        VaMovi = "CCORR"           ' CUENTA CORRIENTE
        Descrip = "Aj.Saldo - " + RASOCLI$(NC1%)
        Observa = "Ajuste a Cero " + DOCUCABE$
        NumeTal = COBRACT%
        CodiEmp% = 0
        CodiSuc% = SUCUI%
        StatInit = 0
        ImpoDeb# = DIFE1#
        Call GRACABAN
        '
        ' graba contrapartida del ajuste
        CuentaCo% = IAJU%
        TiMovi = "AJUSALP"         ' AJUSTE DE SALDO DEUDORES
        OpMovi = "AS"              ' AJUSTE DE SALDO
        VaMovi = "AJU-$"           ' NO IMPUTADO
        ImpoDeb# = (-1) * DIFE1#
        Call GRACABAN
        '
        DocOrLar$ = DOCUCABE$
        DocOrCor$ = Left$(DOCUCABE$, 5) + Mid$(DOCUCABE$, 9, 1) + "-" + Mid$(DOCUCABE$, 14)
        ImpApli# = DIFE1#
        RefApli$ = "Ajuste a Cero " + DOCUCABE$
        TotOrig# = IMPORORIG(DOCUCABE$, NC%)
        MonOri% = 1
        ImpApliOr# = 0
        Call GRAPLIPAG
        '
        With SacrePen
          .FindFirst "NuProve = " & NC% & " AND CODICOM_LAR = '" & DOCUCABE$ & "' AND ABS(SALDACRE) > 0"
          If .NoMatch = False Then
             .Edit
             !SaldAcre = 0
             .Update
          End If
        End With
        '
        MONEMI% = MONEMIA%
        Call COMUNI("Ajuste Procesado")
        GoTo 10
    End If
    GoTo 10
    '
End Sub
'
Sub VAPLIOP()
  '\\\OT-05-0441-WAR-30/06/05///
10  Screen.MousePointer = 1
    Call SELECHO("PROVED1")
    NC% = XVALO(VALACT1$("PROVED1"))
    If NC% < 1 Then
        GoTo 999
    End If
    '
    Screen.MousePointer = 11
    Call BLOQARCH("!SELCUERE")
    Call BLANARCH("!SELCUERE")
    '
301 With ApliPago
      .FindFirst "NUPROVE = " & NC% & " "
      If .NoMatch = True Then
        Call MENSERR(24, "Sin Movimientos")
        GoTo 10
      End If
    End With
    '
    Call APERBASDAT("CABAN")
    SQLX$ = " SELECT  * FROM APLIPAGO "
    SQLX$ = SQLX$ + " WHERE NuPROVE = " & NC% & ""
    SQLX$ = SQLX$ + " ORDER BY CoDocAp_Cor ASC "
    Set APCOTMP = CueCorri.OpenRecordset(SQLX$, 4)
    With APCOTMP
      '
      On Error Resume Next
      .MoveFirst
      If Err Then
       On Error GoTo 0
       GoTo 301
      End If
      On Error GoTo 0
      Y$ = REGBLAN("!SELCUERE")
      JJ& = 0
      CODVIE$ = ""
      Do While Not .EOF
         If CODVIE$ <> !CoDocAp_Cor Then
            NROX$ = !CoDocAp_Cor
           Call REPLA(Y$, NROX$, 5, 18)
           JJ& = JJ& + 1
           J1% = Int(JJ&)
           Call REPLA(Y$, MKS(J1%), 1, 4)
           Call GRAREG("!SELCUERE", Y$, JJ&)
           CODVIE$ = !CoDocAp_Cor
         End If
        .MoveNext
      Loop
      Call SETULTREG("!SELCUERE", JJ&)
      '
      For I& = 1 To ULTREG("!SELCUERE")
        Y$ = REGLEIDO("!SELCUERE", I&)
        NRODO$ = TRIM(Mid$(Y$, 5, 18))
        IMPO# = 0
        With ApliPago
            .FindFirst " CoDocAp_Cor = '" & NRODO$ & "' "
305          If .NoMatch = False Then
              IMPO# = IMPO# + !ImpoApli
              .FindNext " CoDocAp_Cor = '" & NRODO$ & "' "
              GoTo 305
            End If
        End With
        IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
        Call REPLA(Y$, IMPOS$, 23, 13)
        Call GRAREG("!SELCUERE", Y$, I&)
      Next I
    End With
    '
    Call SETULTREG("!SELCUERE", JJ&)
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    Screen.MousePointer = 1
    '
310 RESEL$ = REGSEL$("!SELCUERE/" + TRIM$(RASOCLI$(NC% * -1)))
    If Len(RESEL$) > 4 Then
      DAPLI$ = Mid$(RESEL$, 5, 18)
      ENCAP$ = TRIM$(DAPLI$) + " - $ " + TRIM$(Mid$(RESEL$, 23, 13)) + " - " + RASOCLI$(NC%)
      '
      JJ& = 0
      With ApliPago
        '
        .FindFirst "CoDocAp_Cor= '" & TRIM(DAPLI$) & "'"
307     If .NoMatch = False Then
          Y$ = REGBLAN$("!SELDOCAP")
            NROY$ = !CoDocOr_Lar
            Refy$ = Left(!RefeApli, 18)
            If UCase(Left(Refy$, 5)) = "ORDEN" Then
                Refy$ = Left(!RefeApli, 14)
            End If
            IMPO# = !ImpoApli
            IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
          Call REPLA(Y$, NROY$, 1, 18)
          Call REPLA(Y$, Refy$, 19, 24)
          Call REPLA(Y$, IMPOS$, 44, 13)
          JJ& = JJ& + 1
          Call GRAREG("!SELDOCAP", Y$, JJ&)
          .FindNext "CoDocAp_Cor= '" & TRIM(DAPLI$) & "'"
          GoTo 307
        End If
        '
      End With
      Call SETULTREG("!SELDOCAP", JJ&)
      Call FRESHECHO("!SELDOCAP")
      '
      RESEL$ = REGSEL$("!SELDOCAP/" + ENCAP$)
      GoTo 310
      '
    End If
    GoTo 10
    '
999 End Sub
Sub VAPLIFAC()
  '
    '
10  Screen.MousePointer = 1
    Call SELECHO("PROVED1")
    NC% = XVALO(VALACT1$("PROVED1"))
    If NC% < 1 Then
        GoTo 999
    End If
    '
    Screen.MousePointer = 11
    Call BLOQARCH("!SELCUERE")
    Call BLANARCH("!SELCUERE")
    '
301 With ApliPago
      .FindFirst "NuPROVE = " & NC% & " "
      If .NoMatch = True Then
        Call MENSERR(24, "Sin Movimientos")
        GoTo 10
      End If
    End With
    '
    Call APERBASDAT("CABAN")
    SQLX$ = " SELECT  * FROM APLIPAGO "
    SQLX$ = SQLX$ + " WHERE NuPROVE = " & NC% & ""
    SQLX$ = SQLX$ + " ORDER BY CoDocOr_Cor ASC "
    Set APCOTMP = CueCorri.OpenRecordset(SQLX$, 4)
    With APCOTMP
      '
      On Error Resume Next
      .MoveFirst '.MoveNext
      If Err Then
       On Error GoTo 0
       GoTo 301
      End If
      On Error GoTo 0
      Y$ = REGBLAN("!SELCUERE")
      JJ& = 0
      CODVIE$ = ""
      Do While Not .EOF
         If CODVIE$ <> !CoDocOr_Cor Then '***
            NROX$ = !CoDocOr_Lar
            Call REPLA(Y$, NROX$, 5, 18)
            IMPO# = IMPORORIG(NROX$, NC%)
            IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
            Call REPLA(Y$, IMPOS$, 23, 13)
            JJ& = JJ& + 1
            J1% = Int(JJ&)
            Call REPLA(Y$, MKS(J1%), 1, 4)
            Call GRAREG("!SELCUERE", Y$, JJ&)
            CODVIE$ = !CoDocOr_Cor
         End If
      .MoveNext
      Loop
      Call SETULTREG("!SELCUERE", JJ&)
    End With
    '
    Call SETULTREG("!SELCUERE", JJ&)
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    Screen.MousePointer = 1
    '
310 RESEL$ = REGSEL$("!SELCUERE/" + TRIM$(RASOCLI$(NC% * -1)))
    If Len(RESEL$) > 4 Then
      DAPLI$ = Mid$(RESEL$, 5, 18)
      ENCAP$ = TRIM$(DAPLI$) + " - $ " + TRIM$(Mid$(RESEL$, 23, 13)) + " - " + RASOCLI$(NC%)
      '
      JJ& = 0
      With ApliPago
        '
        .FindFirst "CoDocOR_lar= '" & TRIM(DAPLI$) & "'"
307     If .NoMatch = False Then
          Y$ = REGBLAN$("!SELDOCAP")
          NROY$ = !CoDocAp_Lar
          Refy$ = Left(!RefeApli, 18)
          IMPO# = !ImpoApli
          'CALCULO EL TOTAL
              IMPO# = 0
              .FindFirst " CoDocAp_LAR = '" & NROY$ & "' "
309            If .NoMatch = False Then
                IMPO# = IMPO# + !ImpoApli
                .FindNext " CoDocAp_LAR = '" & NROY$ & "' "
                GoTo 309
              End If
          '
          IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
          Call REPLA(Y$, NROY$, 1, 18)
          Call REPLA(Y$, Refy$, 19, 24)
          Call REPLA(Y$, IMPOS$, 44, 13)
          JJ& = JJ& + 1
          Call GRAREG("!SELDOCAP", Y$, JJ&)
          .FindNext "CoDocOr_Lar= '" & TRIM(DAPLI$) & "'"
          GoTo 307
        End If
        '
      End With
      Call SETULTREG("!SELDOCAP", JJ&)
      Call FRESHECHO("!SELDOCAP")
      '
      RESEL$ = REGSEL$("!SELDOCAP/" + ENCAP$)
      GoTo 310
      '
    End If
    GoTo 10
    '
999 End Sub
Sub CONAPLI()
  '
  Call APERBASDAT("CABAN")
  SQLX1$ = " SELECT FechEmisi, CodiCom_Lar FROM CAJABANC "
  Set NROTMP = CueCorri.OpenRecordset(SQLX1$, 4)
  '
  SQLX$ = " SELECT * FROM APLIPAGO "
  Set DOCAPTMP = CueCorri.OpenRecordset(SQLX$, 4)
  With DOCAPTMP
    '
    On Error Resume Next
    .MoveLast
    If Err Then
      On Error GoTo 0
      Exit Sub
    End If
    On Error GoTo 0
    Max& = .RecordCount
    JJ& = 0
    YY$ = REGBLAN("APLIPAG")
    FEC1% = 0
    CODVIE$ = ""
    .MoveFirst
    Do While Not .EOF
     'If CODVIE$ <> !CoDocAP_LAR Then '***
      Call REPLA(YY$, MKI(!Nuprove), 3, 2)
      Call REPLA(YY$, !CoDocAp_Lar, 5, 18)
      Call REPLA(YY$, MKD(!ImpoApli), 23, 8)
      Call REPLA(YY$, Left(!RefeApli, 24), 31, 24)
      Call REPLA(YY$, !CoDocOr_Lar, 55, 18)
       DOC1$ = !CoDocAp_Lar
       With NROTMP
         On Error Resume Next
         .FindFirst "CodiCom_Lar = '" & DOC1$ & "' "
         If .NoMatch = True Then
          On Error GoTo 0
          GoTo 25
         End If
         On Error GoTo 0
          FEC1% = CVINT(!FechEmisi)
       End With
25    Call REPLA(YY$, MKI(FEC1%), 1, 2)
      JJ& = JJ& + 1
      Call TRACE(24, JJ&, Max&)
      Call GRAREG("APLIPAG", YY$, JJ&)
      'CODVIE$ = !CoDocAP_LAR
     'End If
      .MoveNext
    Loop
    '
  End With
  Call SETULTREG("APLIPAG", JJ&)
  Call CIERRARCH("*.*")
  '
End Sub
'\\\OT-05-0441-WAR-FIN///
Sub APLIFEC()
   '
   Call DESHASFECHA(DES%, HAS%, PERIO$)
   If PERIO$ = "" Then
      GoTo 99
   End If
   '
   Screen.MousePointer = 11
   '
   DESA = Int(CDbl(CVDAT(DES%)))
   HASA = Int(CDbl(CVDAT(HAS%)))
   '
    Call APERBASDAT("CABAN")
    SQLX$ = " SELECT NUPROVE, FechEmisi, CodiCom_Lar, ValAbsComp FROM CAJABANC "
    SQLX$ = SQLX$ + " WHERE INT(CDBL(FechEmisi)) >= " & DESA & "  "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechEmisi)) <= " & HASA & "  "
    SQLX$ = SQLX$ + " AND ( TipoMovim = 'OPAG' OR (TipoMovim = 'FCOM'AND OpciMovim = 'ND' ) ) "
    'SQLX$ = SQLX$ + " AND TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + " Order by FechEmisi ASC "
    Set NROTMP = CueCorri.OpenRecordset(SQLX$, 4)
    '
    SQLX$ = " SELECT * FROM APLIPAGO "
    Set DOCAPTMP = CueCorri.OpenRecordset(SQLX$, 4)
    '
    With NROTMP
    On Error Resume Next
    .MoveFirst
    If Err Then GoTo 99
    On Error GoTo 0
    Do While .EOF = False
      YY$ = REGBLAN("APLIPAG")
      FEDOC% = CVINT(!FechEmisi)
      NCLIE% = !Nuprove
      DOCAPLI$ = !CodiCom_Lar
      '
      With DOCAPTMP
        SQLX2$ = " CODOCAP_LAR = '" & DOCAPLI$ & "' "
       .FindFirst SQLX2$
36     If DOCAPTMP.NoMatch = False Then
         Call REPLA(YY$, MKI(FEDOC%), 1, 2)
         Call REPLA(YY$, MKI(NCLIE%), 3, 2)
         Call REPLA(YY$, DOCAPLI$, 5, 18)
           IAPLI# = !ImpoApli
         Call REPLA(YY$, MKD(!ImpoApli), 23, 8)
           REFEA$ = Left(!RefeApli, 18)
         Call REPLA(YY$, REFEA$, 31, 24)
         Call REPLA(YY$, !CoDocOr_Lar, 55, 18)
         JJ& = JJ& + 1
         'Call TRACE(24, JJ&, Max&)
         Call GRAREG("APLIPAG", YY$, JJ&)
         .FindNext SQLX2$
         GoTo 36
       End If
      End With
26   .MoveNext
    Loop
    '
  End With
90 Screen.MousePointer = 1
    Call SETULTREG("APLIPAG", JJ&)
    Call HEADERS("APLIPAG", "")
    Call HEADERS("APLIPAG", "Corresponde a: " + TRIM$(PERIO$))
    Call FINAL("*LIAPLIPA")
    '
99  On Error GoTo 0
    Screen.MousePointer = 1
    '
End Sub

'\\\OT-06-0405-MH-22/09/06///
Sub LIAPLIFE()
   '
   Dim FEDOCAP%(), NPROVED%(), NUDOCAP$(), IMPORT#()
   '
   Call DESHASFECHA(DES%, HAS%, PERIO$)
   If PERIO$ = "" Then
      GoTo 99
   End If
   '
   Screen.MousePointer = 11
   '
   DESA = Int(CDbl(CVDAT(DES%)))
   HASA = Int(CDbl(CVDAT(HAS%)))
   '
   Call APERBASDAT("CABAN")
   SQLX$ = " SELECT NUPROVE, IDebeComp,IHabeComp, FechEmisi, CodiCom_Lar FROM CAJABANC "
   SQLX$ = SQLX$ + " WHERE INT(CDBL(FechEmisi)) >= " & DESA & " "
   SQLX$ = SQLX$ + " AND INT(CDBL(FechEmisi)) <= " & HASA & " "
   SQLX$ = SQLX$ + " AND ( TipoMovim = 'OPAG' OR (TipoMovim = 'FCOM'AND OpciMovim = 'NC' ) ) "
   SQLX$ = SQLX$ + " Order by FechEmisi ASC, CodiCom_Lar ASC "
   Set NROTMP = CueCorri.OpenRecordset(SQLX$, 4)
   '
   CADOC% = 0: DOCANT$ = ""
   With NROTMP
     On Error Resume Next
     .MoveFirst
     If Err Then GoTo 90
     On Error GoTo 0
     '
     Do While .EOF = False
       FEDOC% = CVINT(!FechEmisi)
       NPROVE% = !Nuprove
       DOCAPLI$ = !CodiCom_Lar
       IMPOTO# = !Idebecomp - !IHABECOMP
       If Abs(IMPOTO#) < 0.005 Then GoTo 26
       '
       If DOCAPLI$ = DOCANT$ Then
         IMPORT#(CADOC%) = IMPORT#(CADOC%) + IMPOTO#
         GoTo 26
       End If
       '
       CADOC% = CADOC% + 1
       ReDim Preserve FEDOCAP%(CADOC%), NPROVED%(CADOC%), NUDOCAP$(CADOC%), IMPORT#(CADOC%)
       FEDOCAP%(CADOC%) = FEDOC%
       NPROVED%(CADOC%) = NPROVE%
       NUDOCAP$(CADOC%) = DOCAPLI$
       IMPORT#(CADOC%) = IMPOTO#
       DOCANT$ = DOCAPLI$
       '
26   .MoveNext
     Loop
     '
   End With
   '
   SQLX$ = " SELECT * FROM APLIPAGO "
   Set DOCAPTMP = CueCorri.OpenRecordset(SQLX$, 4)
   '
   For IDOC% = 1 To CADOC%
      IMPOTO# = IMPORT#(IDOC%)
      RESTO# = IMPOTO#
      '
      PRIREN% = 1
      DOCAPLI$ = NUDOCAP$(IDOC%)
      NPROVE% = NPROVED%(IDOC%)
      FEDOC% = FEDOCAP%(IDOC%)
      SQLX2$ = " CODOCAP_LAR = '" & DOCAPLI$ & "' "
      DOCAPTMP.FindFirst SQLX2$
      '
36    If DOCAPTMP.NoMatch = False Then
         YY$ = REGBLAN("APLIPA1")
         Call REPLA(YY$, MKI(FEDOC%), 1, 2)
         Call REPLA(YY$, DOCAPLI$, 3, 18)
         'AGREGUE(MH)
         'Reemplazaba por 'NCLIE%
         Call REPLA(YY$, MKI(NPROVE%), 21, 2)
         If PRIREN% = 1 Then
           Call REPLA(YY$, MKD(IMPOTO#), 23, 8)
           PRIREN% = 0
         End If
         Call REPLA(YY$, DOCAPTMP!CoDocOr_Lar, 31, 18)
           REFEA$ = Left(DOCAPTMP!RefeApli, 24)
         Call REPLA(YY$, REFEA$, 49, 24)
           IAPLI# = DOCAPTMP!ImpoApli
           If IAPLI# > RESTO# Then IAPLI# = RESTO#
         Call REPLA(YY$, MKD(IAPLI#), 73, 8)
         RESTO# = RESTO# - IAPLI#
         If IAPLI# >= 0.005 Then
           JJ& = JJ& + 1
           Call GRAREG("APLIPA1", YY$, JJ&)
           CAREGRA% = CAREGRA% + 1
         End If
         DOCAPTMP.FindNext SQLX2$
         GoTo 36
      End If
      '
      ' ESTO ES PARA PREVER QUE ESA o/pAGO AUN NO SE APLICó NADA
      If Abs(RESTO#) >= 0.005 Then
         ' buscar si aun está pendiente de aplicacion
         REFEB$ = "*** Pend.Aplic.***"
         SacrePen.FindFirst "NuProve = " & NPROVE% & " AND CodiCom_Lar = '" & DOCAPLI$ & "' AND (SaldDebe + SaldAcre) > 0.005"
         If SacrePen.NoMatch = True Then
            REFEB$ = "Aplic.Automatica"
            NC1% = (-1) * NPROVE%
            If TICUEPRO%(NC1%) = 1 Then ' CUENTA DE GASTOS
              REFEB$ = "Cuenta de Gastos"
            End If
         End If
         '
         YY$ = REGBLAN("APLIPA1")
         Call REPLA(YY$, MKI(FEDOC%), 1, 2)
         Call REPLA(YY$, DOCAPLI$, 3, 18)
         Call REPLA(YY$, MKI(NPROVE%), 21, 2)
         If PRIREN% = 1 Then
            Call REPLA(YY$, MKD(IMPOTO#), 23, 8)
         End If
         Call REPLA(YY$, REFEB$, 31, 18)
           REFEA$ = "Pendiente Aplic."
         Call REPLA(YY$, REFEA$, 49, 24)
         Call REPLA(YY$, MKD(RESTO#), 73, 8)
         JJ& = JJ& + 1
         Call GRAREG("APLIPA1", YY$, JJ&)
      End If
      '
   Next IDOC%
   '
90 Call SETULTREG("APLIPA1", JJ&)
   Call HEADERS("APLIPA1", "")
   Call HEADERS("APLIPA1", "Corresponde a: " + TRIM$(PERIO$))
   Call FINAL("*LIAPLIP1")
   '
99 Screen.MousePointer = 1
   '
End Sub

Sub LIAPLIFE1()
'   '
   Call DESHASFECHA(DES%, HAS%, PERIO$)
   If PERIO$ = "" Then
      GoTo 99
   End If
   '
   Screen.MousePointer = 11
   '
   DESA = Int(CDbl(CVDAT(DES%)))
   HASA = Int(CDbl(CVDAT(HAS%)))
   '
   Call APERBASDAT("CABAN")
   SQLX$ = " SELECT NUPROVE, IDebeComp,IHabeComp, FechEmisi, CodiCom_Lar FROM CAJABANC "
   SQLX$ = SQLX$ + " WHERE INT(CDBL(FechEmisi)) >= " & DESA & " "
   SQLX$ = SQLX$ + " AND INT(CDBL(FechEmisi)) <= " & HASA & " "
   SQLX$ = SQLX$ + " AND ( TipoMovim = 'OPAG' OR (TipoMovim = 'FCOM'AND OpciMovim = 'NC' ) ) "
   SQLX$ = SQLX$ + " Order by FechEmisi,CodiCom_Lar ASC "
   Set NROTMP = CueCorri.OpenRecordset(SQLX$, 4)
   '
   SQLX$ = " SELECT * FROM APLIPAGO "
   Set DOCAPTMP = CueCorri.OpenRecordset(SQLX$, 4)
   '
   With NROTMP
     On Error Resume Next
     .MoveFirst
     If Err Then GoTo 90
     On Error GoTo 0
     '
     Do While .EOF = False
       FEDOC% = CVINT(!FechEmisi)
       NPROVE% = !Nuprove
       DOCAPLI$ = !CodiCom_Lar
       IMPOTO# = !Idebecomp - !IHABECOMP
       RESTO# = IMPOTO#
       If Abs(IMPOTO#) < 0.005 Then GoTo 26
       '
       PRIREN% = 1
       SQLX2$ = " CODOCAP_LAR = '" & DOCAPLI$ & "' "
       DOCAPTMP.FindFirst SQLX2$
36     If DOCAPTMP.NoMatch = False Then
         YY$ = REGBLAN("APLIPA1")
         Call REPLA(YY$, MKI(FEDOC%), 1, 2)
         Call REPLA(YY$, DOCAPLI$, 3, 18)
         'AGREGUE(MH)
         'Reemplazaba por 'NCLIE%
         Call REPLA(YY$, MKI(NPROVE%), 21, 2)
         If PRIREN% = 1 Then
           Call REPLA(YY$, MKD(IMPOTO#), 23, 8)
           PRIREN% = 0
         End If
         Call REPLA(YY$, DOCAPTMP!CoDocOr_Lar, 31, 18)
           REFEA$ = Left(DOCAPTMP!RefeApli, 24)
         Call REPLA(YY$, REFEA$, 49, 24)
           IAPLI# = DOCAPTMP!ImpoApli
           If IAPLI# > RESTO# Then IAPLI# = RESTO#
         Call REPLA(YY$, MKD(IAPLI#), 73, 8)
         RESTO# = RESTO# - IAPLI#
         If IAPLI# >= 0.005 Then
           JJ& = JJ& + 1
           Call GRAREG("APLIPA1", YY$, JJ&)
           CAREGRA% = CAREGRA% + 1
         End If
         DOCAPTMP.FindNext SQLX2$
         GoTo 36
       End If
       '
       ' ESTO ES PARA PREVER QUE ESE RECIBO AUN NO SE APLICó NADA
       If Abs(RESTO#) >= 0.005 Then
         ' buscar si aun está pendiente de aplicacion
         REFEB$ = "*** Pend.Aplic.***"
         'SaDeuPen.FindFirst "NUPROVE = " & NCLIE% & " AND CodiCom_Lar = '" & DOCAPLI$ & "' AND (SaldDebe + SaldAcre) > 0.005"
         SacrePen.FindFirst "NuProve = " & NPROVE% & " AND CodiCom_Lar = '" & DOCAPLI$ & "' AND (SaldDebe + SaldAcre) > 0.005"
         If SacrePen.NoMatch = True Then
            REFEB$ = "Aplic.Automatica"
            NC1% = (-1) * NPROVE%
            If TICUEPRO%(NC1%) = 1 Then ' CUENTA DE GASTOS
              REFEB$ = "Cuenta de Gastos"
            End If
         End If
         '
         YY$ = REGBLAN("APLIPA1")
         Call REPLA(YY$, MKI(FEDOC%), 1, 2)
         Call REPLA(YY$, DOCAPLI$, 3, 18)
         Call REPLA(YY$, MKI(NPROVE%), 21, 2)
         If PRIREN% = 1 Then
            Call REPLA(YY$, MKD(IMPOTO#), 23, 8)
         End If
         Call REPLA(YY$, REFEB$, 31, 18)
           REFEA$ = "Pendiente Aplic."
         Call REPLA(YY$, REFEA$, 49, 24)
         Call REPLA(YY$, MKD(RESTO#), 73, 8)
         JJ& = JJ& + 1
         Call GRAREG("APLIPA1", YY$, JJ&)
       End If
       '
26   .MoveNext
     Loop
     '
   End With
   '
   'Agregue(MH)
90 Call SETULTREG("APLIPA1", JJ&)
   Call HEADERS("APLIPA1", "")
   Call HEADERS("APLIPA1", "Corresponde a: " + TRIM$(PERIO$))
   Call FINAL("*LIAPLIP1")
   '
99 Screen.MousePointer = 1
   '
End Sub
'
Sub DESAPLIOPAG()
    '//ot-07-0403
10  Call SELECHO("PROVED1")
    NC% = XVALO(VALACT1$("PROVED1"))
    If NC% < 1 Then
        GoTo 999
    End If
    '
    Call BLANARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    '
    ' AQUI ESCRIBIR UNA RUTINA PARA SELECCIONAR DE SALDACRE LOS DOCUMENTOS A IMPUTAR Y CARGARLOS EN ARCHIVO 'SELDOCAP'
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT DISTINCT CoDocAp_Lar, FechEmisi FROM APLIPAGO "
    SQLX$ = SQLX$ + "WHERE NUPROVE = " & NC% & " "
    SQLX$ = SQLX$ + "ORDER BY FechEmisi ASC;"
    Set FacPenTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    JJ& = 0
    On Error Resume Next
    With FacPenTemp
      .MoveLast
      If Err < 1 Then
         .MoveFirst
         Do While Not .EOF
            TIX$ = !CoDocAp_Lar
            SUAPLI# = 0
            SQLX$ = "NUPROVE = " & NC% & " AND CoDocAp_Lar = '" & TIX$ & "'"
            '
            ApliPago.FindFirst SQLX$
20          If ApliPago.NoMatch = False Then
               SUAPLI# = SUAPLI# + ApliPago!ImpoApli
               ApliPago.FindNext SQLX$
               GoTo 20
            End If
            '
            IMPOS$ = NUMSTRI$(SUAPLI#, 14, 2, 1)
            NPX& = XVALO(Mid$(!CoDocAp_Lar, 11, 8))
            Y$ = MKS$(NPX&) + AJUSTI$(!CoDocAp_Lar, 18) + IMPOS$
            JJ& = JJ& + 1
            Call GRAREG("!SELCUERE", Y$, JJ&)
            '
         .MoveNext
         Loop
      End If
    End With
    '
    If ULTREG&("!SELCUERE") < 1 Then
       Call MENSERR(24, "No Hay Pagos Aplicados\en la Cuenta Elegida.")
       GoTo 10
    End If
    '
    SPDX$ = REGSEL$("!SELCUERE")
    If Len(SPDX$) <= 4 Then GoTo 10
    '
    TIX$ = Mid$(SPDX$, 5, 18)
    SQLX$ = "NUPROVE = " & NC% & " AND CoDocAp_Lar = '" & TIX$ & "'"
    '
    With ApliPago
30     .FindFirst SQLX$
       If .NoMatch = False Then
          .Delete
          GoTo 30
       End If
    End With
    '
    '
    Call COMUNI("Des-Aplicación de " + TRIM$(TIX$) + " Completa.\Debe Ahora Reconstruir la Composición de Deuda del Proveedor.")
    '
    RECUEP04.Text1.TEXT = TRIM$(Str$(NC%))
    Call FRESHALL
    PROLLAMA$ = Name
    Call RECUEP04.Command4_Click
    RECUEP04.Show 1
    Call FRESHALL
    '
    GoTo 10
    '
999 End Sub
