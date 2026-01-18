VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form AMACLIS07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CLIENTES - Consultas y Listados"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5085
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   5085
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   1470
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1170
      ScaleHeight     =   585
      ScaleWidth      =   2265
      TabIndex        =   12
      Top             =   5370
      Width           =   2325
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
         Left            =   -20
         Picture         =   "AMACLIS07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Listado de Cuentas Inactivas"
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
         TabIndex        =   14
         Top             =   50
         Width           =   1590
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2D3C0&
      Caption         =   "CONSULTAS - LITTADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   120
      TabIndex        =   0
      Top             =   50
      Width           =   4005
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   1245
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   22
         Top             =   2475
         Width           =   2325
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
            Height          =   630
            Left            =   -20
            Picture         =   "AMACLIS07.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Clientes Inactivos / Activos según fecha de corte"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Listados Cuentas Activas/Inactivas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   700
            TabIndex        =   24
            Top             =   80
            Width           =   1800
         End
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   1005
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   1
         Top             =   1965
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
            Height          =   630
            Left            =   -20
            Picture         =   "AMACLIS07.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Listado de Deudas y Límite de Crédito"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Deudas y Límites de Crédito"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   700
            TabIndex        =   3
            Top             =   30
            Width           =   1800
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   780
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   17
         Top             =   1455
         Width           =   2325
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
            Height          =   630
            Left            =   0
            Picture         =   "AMACLIS07.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Libreta de Direcciones de Mails"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Lista de Direcciones de Mail"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   720
            TabIndex        =   18
            Top             =   30
            Width           =   1485
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   570
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   15
         Top             =   960
         Width           =   2325
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
            Height          =   630
            Left            =   0
            Picture         =   "AMACLIS07.frx":0C28
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Listado de Direcciones Postales"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado de Domicilios Postales"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   700
            TabIndex        =   16
            Top             =   30
            Width           =   1750
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   360
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   7
         Top             =   450
         Width           =   2325
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
            Height          =   630
            Left            =   0
            Picture         =   "AMACLIS07.frx":0F32
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Listados Generales y Clasificados de la Base de Datos de Clientes"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado General y Clasificado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   720
            TabIndex        =   9
            Top             =   40
            Width           =   1800
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2D3C0&
         Caption         =   "SALIDA"
         Height          =   600
         Left            =   420
         TabIndex        =   4
         Top             =   3135
         Width           =   3165
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Impresión"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   1800
            TabIndex        =   6
            Top             =   300
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Pantalla"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   315
            TabIndex        =   5
            Top             =   300
            Width           =   1380
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   4200
      ScaleHeight     =   4665
      ScaleWidth      =   1035
      TabIndex        =   10
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton AGenRep 
         Height          =   600
         Left            =   105
         Picture         =   "AMACLIS07.frx":123C
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   920
         Width           =   650
      End
      Begin VB.CommandButton command1 
         Caption         =   "Exit"
         Height          =   690
         Left            =   105
         Picture         =   "AMACLIS07.frx":1546
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   210
         Width           =   650
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "AMACLIS07.frx":1690
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "AMACLIS07.frx":18C4
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "AMACLIS07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub NuevoListadoCuentasInactivasActivas()

    Dim ACTI%(999999)
    If DERACCE%("DEUINAC", "CONSULTA DE CUENTAS INACTIVAS") < 2 Then GoTo 99

10  OPX1% = COMALTER%("Seleccione el Listado:\\Clientes Inactivos\Clientes Activos")
    If OPX1% = 2 Then
      '
      HO1% = HOY(HO$)
      Call REPLA(VDEF$, MKI$(HO1%), 1, 2)
      OBJ1$ = OBJPLA$("ENTRAFECHA1/MOVIMIENTOS POSTERIORES A", VDEF$)
      If TRIM$(OBJ1$) = "" Then GoTo 10
      '
      FLIM% = CVI(Left$(OBJ1$, 2))
      If FLIM% > 0 Then
         Screen.MousePointer = 11
         '
         Call ABRECONEXION
20       Dim FACT As ADODB.Recordset
         SQLX$ = "SELECT MAX(FECHEMISI) AS FechaEmision, NUCLIEN FROM CAJABANC WITH(NOLOCK) WHERE TIPOMOVIM='FVEN' AND NUCLIEN>0 AND NUCLIEN < 999999 GROUP BY NUCLIEN"
         Set FACT = READSET(SQLX$)
         '
         With FACT
            FIN& = .RecordCount
            .MoveFirst
            Do While Not .EOF
              U& = U& + 1
              Call TRACE(20, U&, FIN&)
              NCLIE = XVALO(!nuclien)
              F1% = CVINT(!FechaEmision)
              If F1% >= FLIM% Then
                    If F1% > ACTI%(NCLIE) Then ACTI%(NCLIE) = F1%
              End If
              .MoveNext
            Loop
            .Close
         End With
         '
         '
30       Dim COBRAN As ADODB.Recordset
         SQLX$ = "SELECT MAX(FECHEMISI) AS FechaEmision, NUCLIEN FROM CAJABANC WITH(NOLOCK) WHERE TIPOMOVIM='RCOB' AND NUCLIEN>0 AND NUCLIEN<999999 GROUP BY NUCLIEN"
         Set COBRAN = READSET(SQLX$)
         '
         With COBRAN
            FIN& = .RecordCount
            U& = 0
            .MoveFirst
            Do While Not .EOF
              U& = U& + 1
              Call TRACE(20, U&, FIN&)
              NCLIE = XVALO(!nuclien)
              F1% = CVINT(!FechaEmision)
              If F1% >= FLIM% Then
                    If F1% > ACTI%(NCLIE) Then ACTI%(NCLIE) = F1%
              End If
              .MoveNext
              
            Loop
            .Close
         End With
         '
         '
40       Dim PRESUPUESTO As ADODB.Recordset
         SQLX$ = "SELECT MAX(FECHEMIS) AS FechaEmision, NROCLIE FROM PRESUVEN WITH(NOLOCK) WHERE NROCLIE>0 AND NROCLIE<999999 GROUP BY NROCLIE"
         Set PRESUPUESTO = READSET(SQLX$)
         '
         With PRESUPUESTO
            FIN& = .RecordCount
            U& = 0
            .MoveFirst
            
            Do While Not .EOF
              U& = U& + 1
              Call TRACE(20, U&, FIN&)
              NCLIE = XVALO(!NroClie)
              F1% = CVINT(!FechaEmision)
              If F1% >= FLIM% Then
                    If F1% > ACTI%(NCLIE) Then ACTI%(NCLIE) = F1%
              End If
              .MoveNext
            Loop
            .Close
         End With
         '
         '
50       Dim PEDIDOS As ADODB.Recordset
         SQLX$ = "SELECT MAX(FECHEMIS) AS FechaEmision, NROCLIE FROM PEDENCA WITH(NOLOCK) WHERE NROCLIE>0 AND NROCLIE<999999 GROUP BY NROCLIE"
         Set PEDIDOS = READSET(SQLX$)
         '
         With PEDIDOS
            FIN& = .RecordCount
            U& = 0
            .MoveFirst
            
            Do While Not .EOF
              U& = U& + 1
              Call TRACE(20, U&, FIN&)
              NCLIE = XVALO(!NroClie)
              F1% = CVINT(!FechaEmision)
              If F1% >= FLIM% Then
                    If F1% > ACTI%(NCLIE) Then ACTI%(NCLIE) = F1%
              End If
              .MoveNext
            Loop
            .Close
         End With
         '
60       Campos$ = "Cuenta/I10;Cliente/A48;Telefono/A24;Mail/A48;Saldo/F14.2;Deuda/F14.2;Observaciones/A48"
         Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
         Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
         GRILLAGRAL.Caption = "LISTADO DE CUENTAS ACTIVAS CON MOVIMIENTOS POSTERIORES AL " & FECHATEX$(FLIM%)
         GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
         GRILLAGRAL.mnuMenu1.Visible = True
         GRILLAGRAL.mnuMenu3.Caption = "Invertir selección"
   
         J& = 0
         For JX& = 1 To 999999
            Call TRACE(20, JX&, 999999)
            If JX& = 999999 Then Exit For
            NCLIE = JX&
            If ACTI%(NCLIE) > 0 Then
              SSCC# = SALDCLI#(NCLIE)
              CONDI$ = "Nume_Cli=" + CStr(NCLIE)
              TELEFONO$ = TRIM$(VALOBADA("DEUDOR12", "Tele_Cli", CONDI$))
              mail$ = TRIM$(VALOBADA("DEUDOR12", "Mail_Cli", CONDI$))
              razonSocial$ = TRIM$(VALOBADA("DEUDOR12", "RaSo_Cli", CONDI$))
              
              J& = J& + 1
              GRILLAGRAL.GRID.Rows = J& + 1
              GRILLAGRAL.GRID.TextMatrix(J&, 1) = NCLIE
              GRILLAGRAL.GRID.TextMatrix(J&, 2) = razonSocial$
              GRILLAGRAL.GRID.TextMatrix(J&, 3) = Mid$(TELEFONO$, 1, 24)
              GRILLAGRAL.GRID.TextMatrix(J&, 4) = Mid$(mail$, 1, 48)
              GRILLAGRAL.GRID.TextMatrix(J&, 5) = TRIM$(FORMATNUM$(SSCC#, "F14.2"))
              GRILLAGRAL.GRID.TextMatrix(J&, 6) = TRIM$(FORMATNUM$(SSCC#, "F14.2"))
              GRILLAGRAL.GRID.TextMatrix(J&, 7) = "Último Movimiento: " + FECHATEX$(ACTI%(NCLIE))
            End If
         Next JX&
         
         GRILLAGRAL.GRID.ZOrder 0
         GRILLAGRAL.TXTBUSCAR = ""
         Call POSICIONAR_ULTIMAFILA(GRILLAGRAL.GRID)
         Screen.MousePointer = 1
         
         GRILLAGRAL.Show 1
         Dim MSF As msFlexGrid
         If GRILLAGRAL.MENU1 > 0 Then
             Set MSF = GRILLAGRAL.GRID
             If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
             Call CARGA_TABLA_IMPRE(MSF, Campos$, "CUEACTI", GRILLAGRAL.Caption, "TABLA_IMPRE", RET_NAMECONS$, 1)
             Call FINAL("?" & RET_NAMECONS$)
             GRILLAGRAL.MENU1 = 0
             Screen.MousePointer = 1
         End If
         
      End If
      
    ElseIf OPX1% = 1 Then
      '
      Call ABRECONEXION
      HO1% = HOY(HO$)
      Call REPLA(VDEF$, MKI$(HO1%), 1, 2)
      OBJ1$ = OBJPLA$("ENTRAFECHA1/MOVIMIENTOS ANTERIORES A", VDEF$)
      If TRIM$(OBJ1$) = "" Then GoTo 10
      '
      FLIM% = CVI(Left$(OBJ1$, 2))
      If FLIM% > 0 Then
         Screen.MousePointer = 11
         Dim FACT0 As ADODB.Recordset
         SQLX$ = "SELECT MAX(FECHEMISI) AS FechaEmision, NUCLIEN FROM CAJABANC WITH(NOLOCK) WHERE TIPOMOVIM='FVEN' AND NUCLIEN>0 AND NUCLIEN<999999 GROUP BY NUCLIEN"
         Set FACT0 = READSET(SQLX$)
         '
         With FACT0
            FIN& = .RecordCount
            U& = 0
            .MoveFirst
            
            Do While Not .EOF
              U& = U& + 1
              Call TRACE(20, U&, FIN&)
              NCLIE = XVALO(!nuclien)
              F1% = CVINT(!FechaEmision)
              If F1% > ACTI%(NCLIE) Then
                ACTI%(NCLIE) = F1%
              End If
              .MoveNext
            Loop
            .Close
         End With
         '
80       Campos$ = "Cuenta/I10;Cliente/A48;Telefono/A24;Mail/A48;Saldo/F14.2;Deuda/F14.2;Observaciones/A48"
         Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
         Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
         GRILLAGRAL.Caption = "LISTADO DE CUENTAS INACTIVAS CON MOVIMIENTOS ANTERIORES AL " & FECHATEX$(FLIM%)
         GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
         GRILLAGRAL.mnuMenu1.Visible = True
         GRILLAGRAL.mnuMenu3.Caption = "Invertir selección"
         
         J& = 0
         For IL& = 1 To 999999
            Call TRACE(20, IL&, 999999)
            NCLIE = XVALO(IL&)
            If ACTI%(NCLIE) > 0 And ACTI%(NCLIE) < FLIM% Then
                 Y$ = REGBLAN$("LISACLI")
                 SSCC# = SALDCLI#(NCLIE)
                 CONDI$ = "Nume_Cli=" + CStr(NCLIE)
                 TELEFONO$ = TRIM$(VALOBADA("DEUDOR12", "Tele_Cli", CONDI$))
                 mail$ = TRIM$(VALOBADA("DEUDOR12", "Mail_Cli", CONDI$))
                 razonSocial$ = TRIM$(VALOBADA("DEUDOR12", "RaSo_Cli", CONDI$))
                 
                 J& = J& + 1
                 GRILLAGRAL.GRID.Rows = J& + 1
                 GRILLAGRAL.GRID.TextMatrix(J&, 1) = NCLIE
                 GRILLAGRAL.GRID.TextMatrix(J&, 2) = razonSocial$
                 GRILLAGRAL.GRID.TextMatrix(J&, 3) = Mid$(TELEFONO$, 1, 24)
                 GRILLAGRAL.GRID.TextMatrix(J&, 4) = Mid$(mail$, 1, 48)
                 GRILLAGRAL.GRID.TextMatrix(J&, 5) = TRIM$(FORMATNUM$(SSCC#, "F14.2"))
                 GRILLAGRAL.GRID.TextMatrix(J&, 6) = TRIM$(FORMATNUM$(SSCC#, "F14.2"))
                 GRILLAGRAL.GRID.TextMatrix(J&, 7) = "Inactiva Desde " + FECHATEX$(ACTI%(NCLIE))
            End If
         Next IL&
         '
         GRILLAGRAL.GRID.ZOrder 0
         GRILLAGRAL.TXTBUSCAR = ""
         Call POSICIONAR_ULTIMAFILA(GRILLAGRAL.GRID)
         Screen.MousePointer = 1
         
         GRILLAGRAL.Show 1
         Dim MSF0 As msFlexGrid
         If GRILLAGRAL.MENU1 > 0 Then
             Set MSF0 = GRILLAGRAL.GRID
             If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
             Call CARGA_TABLA_IMPRE(MSF0, Campos$, "CUEINAC", GRILLAGRAL.Caption, "TABLA_IMPRE", RET_NAMECONS$, 1)
             Call FINAL("?" & RET_NAMECONS$)
             GRILLAGRAL.MENU1 = 0
             Screen.MousePointer = 1
         End If
      
      End If
    End If

99  Screen.MousePointer = 1

End Sub

Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

 Sub Command2_Click()
   '
   Command2.Enabled = False
   CMQ% = COMALTER%("Opciones Disponibles\\Listado Impreso\Exportar")
   If CMQ% < 1 Or CMQ% > 2 Then GoTo 99
   '
   Screen.MousePointer = 11
   'Consulta SQL
   Call APERBASDAT("CLIEN")
   SQLX$ = " Select * from deudor12 "
   SQLX$ = SQLX$ + " ORDER BY Nume_Cli "
   'sql
   
   'Set CLITEMP = Badaclie.OpenRecordset(SQLX$, dbOpenSnapshot)
    Set CLITEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    If CLITEMP.EOF And CLITEMP.BOF Then
        Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
         On Error GoTo 0
         Screen.MousePointer = 1
         GoTo 90
    End If
      
     'With CLITEMP
     '  On Error Resume Next
     '  .MoveFirst
     '  'Valida que la tabla no este vacia
     '  If Err Then
     '    Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
     '    On Error GoTo 0
     '    Screen.MousePointer = 1
     '    GoTo 90
     '  End If
     '  On Error GoTo 0
     'End With
   '
   If CMQ% = 1 Then   'IMPRESION
   If EXISTESQR_ENSQL%("R-LITELMAI") Then
        Call FINAL("?R-LITELMAI")
        GoTo 99
   End If

    With CLITEMP
      .MoveFirst
      Call REGBLAN("LITELMAI")
      jj& = 0
      Do While .EOF = False
        NCLIE = !nume_cli
        EMA$ = EMAILCLI$(NCLIE)
        EMB$ = TEL1CLI$(NCLIE)
        EMC$ = FAXCLI$(NCLIE)
        EMD$ = rasocli$(NCLIE)
        If TRIM$(EMA$ + EMB$ + EMC$) <> "" Then
          Call REPLA(TTXX$, MKS$(NCLIE), 1, 4)
          Call REPLA(TTXX$, EMD$, 5, 28)
          Call REPLA(TTXX$, EMB$, 33, 24)
          Call REPLA(TTXX$, EMC$, 57, 24)
          Call REPLA(TTXX$, EMA$, 81, 47)
          jj& = jj& + 1
          Call GRAREG("LITELMAI", TTXX$, jj&)
        End If
        .MoveNext
        Loop
    End With
      '
      Call SETULTREG("LITELMAI", jj&)
      Call CIERRARCH("LITELMAI")
      Call CIERRARCH("*.*")
      '
      Call FINAL("*LITELMAI")
      GoTo 99
      '
   End If
   '
60 With CLITEMP             'EXPORTACIÓN
     .MoveFirst
     jj& = 0
     Call REGBLAN("BCONTACT")
     Do While (.EOF = False)
        NCLIE = !nume_cli
        EMA$ = EMAILCLI$(NCLIE)
        If TRIM$(EMA$) <> "" Then
          If InStr(EMA$, "@") > 0 Then
             TTXX$ = String$(2, 0) + Space$(1022)
             RACLI$ = rasocli$(NCLIE)
             Call REPLA(TTXX$, RACLI$, 3, 78)
             Call REPLA(TTXX$, "Sres.", 81, 8)
             Call REPLA(TTXX$, RACLI$, 145, 64)
             Call REPLA(TTXX$, EMA$, 217, 64)
             Call REPLA(TTXX$, DOMICLI$(NCLIE), 768 + 1, 48)
             Call REPLA(TTXX$, CPOSCLI$(NCLIE), 768 + 49, 16)  ' cpost empresa
             Call REPLA(TTXX$, LOCACLI$(NCLIE), 768 + 65, 48)  ' localidad empresa
             Call REPLA(TTXX$, DEPROVI$(PROVCLI$(NCLIE)), 768 + 113, 32) ' provincia empresa
             Call REPLA(TTXX$, PAISCLI$(NCLIE), 768 + 145, 32) ' pais empresa
             Call REPLA(TTXX$, TEL1CLI$(NCLIE), 768 + 177, 40) ' telef.empresa
             Call REPLA(TTXX$, FAXCLI$(NCLIE), 768 + 217, 20) ' fax empresa
             jj& = jj& + 1
             Call GRAREG("!BCONTACT", TTXX$, jj&)
          End If
        End If
      .MoveNext
     Loop
   End With
   '
   Call SETULTREG("!BCONTACT", jj&)
   Call CIERRARCH("!BCONTACT")
   MSGXX$ = "Se ha Exportado con Éxito en "
   MSGXX$ = MSGXX$ + LUCOM$ + "BCONTACT.WKF"
   MsgBox (MSGXX$)
   Call CIERRARCH("*.*")
   '
   CLITEMP.Close
   Set CLITEMP = Nothing
   
99 Command2.Enabled = True
   Screen.MousePointer = 1
90  End Sub

Sub Command3_Click()
   Command3.Enabled = False
   Call LILIMICRE
   Command3.Enabled = True
End Sub

 Sub Command4_Click()
    Command4.Enabled = False
       '
    Screen.MousePointer = 11
       '
    If EXISTESQR_ENSQL%("R-LIPOSTAC") Then
        Call FINAL("?R-LIPOSTAC")
        GoTo 99
    End If
      
    Call APERBASDAT("CLIEN")
    '    FIN& = ULTREG&("DEUDOR1")
    '    JJ& = 0
    '    Call COPYSTRU("BADACLIE", "BADACLIX")
    
    'Consulta sql
     SQLX$ = " SELECT * FROM DEUDOR12 "
    'Filtro unicamente las cuentas corrientes
     'SQLX$ = SQLX$ + " WHERE Tcpr_Cli <> 1 "
     SQLX$ = SQLX$ + " WHERE Stat_Cli > 0 " ' Que no esten  de baja
     SQLX$ = SQLX$ + " ORDER BY Raso_Cli ASC "
     
     'Set CLITEMP = Badaclie.OpenRecordset(SQLX$, dbOpenSnapshot)
      Set CLITEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
      If CLITEMP.EOF And CLITEMP.BOF Then
         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
         GoTo 90
      End If
      
         With CLITEMP
           'On Error Resume Next
           '.MoveLast
           'Valida que la tabla no este vacia
           'If Err Then
           '  Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
           '  GoTo 90
           'End If
           
           registros& = .RecordCount
           FIN& = registros&
           .MoveFirst
        End With
    With CLITEMP
        jj& = 0
        Call REGBLAN("BADACLIX")
     Do While (.EOF = False)
    '    For U& = 1 To FIN&
    '      Call TRACE(20, U&, FIN&)
    '      X$ = REGLEIDO$("DEUDOR1", U&)
    '      NC% = CVI(Left$(X$, 2))
           NCLIE = XVALO(!nume_cli)
          ' NC% = (-1) * NP%
         'If NC% < 0 Then
           EMA$ = rasocli$(NCLIE)
           EMB$ = DOMICLI$(NCLIE)
           EMC$ = CPOSCLI$(NCLIE)
           EMD$ = LOCACLI$(NCLIE)
           EME$ = DEPROVI$(PROVCLI$(NCLIE))
           EMF$ = PAISCLI$(NCLIE)
           '
           If TRIM$(EMA$ + EMB$ + EMC$ + EMD$ + EME$ + EMF$) <> "" Then
             TTXX$ = String$(2, 0) + Space$(510)
             Call REPLA(TTXX$, MKS$(NCLIE), 1, 4)
             Call REPLA(TTXX$, EMA$, 6, 60)
             Call REPLA(TTXX$, EMB$, 97, 64)
             Call REPLA(TTXX$, EMC$, 161, 16)
             Call REPLA(TTXX$, EMD$, 177, 48)
             Call REPLA(TTXX$, EME$, 225, 32)
             Call REPLA(TTXX$, EMF$, 257, 24)
             jj& = jj& + 1
             Call GRAREG("BADACLIX", TTXX$, jj&)
           End If
         'End If
         .MoveNext
        Loop
    '   Next U&
    End With
       '
       Call SETULTREG("BADACLIX", jj&)
       Call CIERRARCH("BADACLIX")
       Call CIERRARCH("*.*")
       '
       Call FINAL("*LIPOSTAL")
       Screen.MousePointer = 1
       '
       CLITEMP.Close
       Set CLITEMP = Nothing
99     Command4.Enabled = True
       Screen.MousePointer = 1
90     End Sub
Sub Command5_Click()
'Exit Sub
    Command5.Enabled = False
    Call SELECHO("SELFECHA")
    FECHLI$ = VALACT1$("SELFECHA")
    FLIM% = FECHANUM(FECHLI$)
    If FLIM% > 0 Then
       Screen.MousePointer = 11
       Dim ACTI%(32767)
       FIN& = ULTREG&("FACTUR")
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         X$ = REGLEIDO$("FACTUR", U&)
         NC% = CVI(Left$(X$, 2))
         If NC% > 0 Then
           If NC% <= 32767 Then
             If CVI(Mid$(X$, 7, 2)) > ACTI%(NC%) Then
               ACTI%(NC%) = CVI(Mid$(X$, 7, 2))
             End If
           End If
         End If
       Next U&
       '
       RELISA& = 0
       cx& = ULTREG&("DEUDOR1")
       For IL& = 1 To cx&
         Call TRACE(20, IL&, cx&)
         NC% = CVI(Left$(REGLEIDO$("DEUDOR1", IL&), 2))
         If ACTI%(NC%) < FLIM% Then
            Y$ = REGBLAN$("LISACLI")
            SSCC# = SALDCLI#(NC%)
            Call REPLA(Y$, MKI$(NC%), 1, 2)
            Call REPLA(Y$, MKD$(SSCC#), 3, 8)
            Call REPLA(Y$, MKD$(SSCC#), 11, 8)
            FEX = ACTI%(NC%)
            TTXX$ = "Inactiva Desde " + FECHATEX$(FEX)
            Call REPLA(Y$, TTXX$, 19, 30)
            '
            RELISA& = RELISA& + 1
            Call GRAREG("LISACLI", Y$, RELISA&)
         End If
       Next IL&
       '
       Call SETULTREG("LISACLI", RELISA&)
       Call CIERRARCH("*.*")
       Screen.MousePointer = 1
       Call CONECRUN("*DEUINAC", "Listado de Cuentas Inactivas")
    End If
    Command5.Enabled = True
End Sub

Private Sub Command6_RESPALDO_Click()
    'Exit Sub
Dim ACTI%(999999)
    Command5.Enabled = False
    If DERACCE%("DEUINAC", "CONSULTA DE CUENTAS INACTIVAS") < 2 Then GoTo 99
    If DERACCE%("LDEUINM", "CONSULTA DE CUENTAS INACTIVAS") < 2 Then GoTo 99
    If DERACCE%("LDEUINA", "CONSULTA DE CUENTAS INACTIVAS") < 2 Then GoTo 99
    
    MODOMEDI% = 0
    If InStr(1, EMPREAC$, "MEDITEA") > 0 Then MODOMEDI% = 1
    Call BLANARCH("LISACLI")
    Call HEADERS("LISACLI", "")
    Call HEADERS("LISACLI", "Generado el: " + FECHATEX$(HOY(HO$)) + " - " + Left$(Time$, 5) + " hs.")
    If MODOMEDI% > 0 Then
        Call BLANARCH("LISACLM")
        Call HEADERS("LISACLM", "")
        Call HEADERS("LISACLM", "Generado el: " + FECHATEX$(HOY(HO$)) + " - " + Left$(Time$, 5) + " hs.")
    End If
    '\\\OT-06-0194-WAR-29/07/06///
10  OPX1% = COMALTER%("Seleccione el Listado:\\Clientes Inactivos\Clientes Activos")
    If OPX1% = 2 Then
      '
      HO1% = HOY(HO$)
      Call REPLA(VDEF$, MKI$(HO1%), 1, 2)
      OBJ1$ = OBJPLA$("ENTRAFECHA1", VDEF$)
      If TRIM$(OBJ1$) = "" Then GoTo 10
      '
      FLIM% = CVI(Left$(OBJ1$, 2))
      If FLIM% > 0 Then
         Screen.MousePointer = 11
         'FIN& = ULTREG&("FACTUR")
         '
         Call ABRECONEXION
20       Dim FACT As ADODB.Recordset
         Set FACT = New ADODB.Recordset
         SQLX$ = "SELECT * FROM cajabanc where tipomovim ='FVEN'"
         Set FACT = FLEXCONN.Execute(FILTSQL$(SQLX$))
         '
         With FACT
         FIN& = CantRegistros("CAJABANC", "tipomovim ='FVEN'")
         If .BOF = True And .EOF = True Then GoTo 30
         .MoveFirst
         For U& = 1 To FIN&
           Call TRACE(20, U&, FIN&)
           'X$ = REGLEIDO$("FACTUR", U&)
           NCLIE = XVALO(!nuclien) 'CVI(Left$(X$, 2))
           'F1% = FECHANUM(Format(!FechEmisi, "DD/MM/YY")) 'CVI(Mid$(X$, 7, 2))
           F1% = CVINT(!FECHEMISI)
           If F1% >= FLIM% Then
             If NCLIE <= 999999 Then
               If NCLIE > 0 Then
                 If F1% > ACTI%(NCLIE) Then ACTI%(NCLIE) = F1%
               End If
             End If
           End If
           .MoveNext
         Next U&
         .Close
         End With
         '
         '
30       Dim COBRAN As ADODB.Recordset
         Set COBRAN = New ADODB.Recordset
         SQLX$ = "SELECT * FROM CAJABANC WHERE TIPOMOVIM='RCOB'"
         Set COBRAN = FLEXCONN.Execute(FILTSQL$(SQLX$))
         '
         With COBRAN
         FIN& = CantRegistros("CAJABANC", "TIPOMOVIM='RCOB'")
         If .BOF = True And .EOF = True Then GoTo 40
         .MoveFirst
         For U& = 1 To FIN&
           Call TRACE(20, U&, FIN&)
           
           NCLIE = XVALO(!nuclien)
           'F1% = FECHANUM(Format(!FechEmisi, "DD/MM/YY"))
           F1% = CVINT(!FECHEMISI)
           If F1% >= FLIM% Then
             If NCLIE <= 999999 Then
               If NCLIE > 0 Then
                 If F1% > ACTI%(NCLIE) Then ACTI%(NCLIE) = F1%
               End If
             End If
           End If
           .MoveNext
         Next U&
         .Close
         End With
         '
         '
40       Dim PRESUPUESTO As ADODB.Recordset
         Set PRESUPUESTO = New ADODB.Recordset
         SQLX$ = "SELECT * FROM PRESUVEN"
         Set PRESUPUESTO = FLEXCONN.Execute(FILTSQL$(SQLX$))
         '
         With PRESUPUESTO
         FIN& = CantRegistros("PRESUVEN", "")
         If .BOF = True And .EOF = True Then GoTo 50
         .MoveFirst
         For U& = 1 To FIN&
           Call TRACE(20, U&, FIN&)
           'X$ = REGLEIDO$("FACTUR", U&)
           NCLIE = XVALO(!NroClie) 'CVI(Left$(X$, 2))
           'F1% = FECHANUM(Format(!FECHEMIS, "DD/MM/YY"))
           F1% = CVINT(!FechEmis)
           If F1% >= FLIM% Then
             If NCLIE <= 999999 Then
               If NCLIE > 0 Then
                 If F1% > ACTI%(NCLIE) Then ACTI%(NCLIE) = F1%
               End If
             End If
           End If
           .MoveNext
         Next U&
         .Close
         End With
         '
         '
50       Dim PEDIDOS As ADODB.Recordset
         Set PEDIDOS = New ADODB.Recordset
         SQLX$ = "SELECT * FROM PEDENCA"
         Set PEDIDOS = FLEXCONN.Execute(FILTSQL$(SQLX$))
         '
         With PEDIDOS
         FIN& = CantRegistros("PEDENCA", "")
         If .BOF = True And .EOF = True Then GoTo 60
         .MoveFirst
         For U& = 1 To FIN&
           Call TRACE(20, U&, FIN&)
           NCLIE = XVALO(!NroClie)
           'F1% = FECHANUM(Format(!FECHEMIS, "DD/MM/YY"))
           F1% = CVINT(!FechEmis)
           If F1% >= FLIM% Then
             If NCLIE <= 999999 Then
               If NCLIE > 0 Then
                 If F1% > ACTI%(NCLIE) Then ACTI%(NCLIE) = F1%
               End If
             End If
           End If
           .MoveNext
         Next U&
         .Close
         End With
         '
       
60     RELISA& = 0
       For JX = 1 To 999999
         Call TRACE(20, IL&, 999999)
         If JX = 999999 Then Exit For
         NCLIE = JX
         If ACTI%(NCLIE) > 0 Then
           Y$ = REGBLAN$("LISACLI")
           SSCC# = SALDCLI#(NCLIE)
           Call REPLA(Y$, MKS$(NCLIE), 1, 4)
               RASO$ = rasocli$(NCLIE)
           Call REPLA(Y$, RASO$, 5, 30)
           Call REPLA(Y$, MKD$(SSCC#), 35, 8)
           Call REPLA(Y$, MKD$(SSCC#), 43, 8)
           TTXX$ = "Último Movimiento: " + FECHATEX$(ACTI%(NCLIE))
           Call REPLA(Y$, TTXX$, 51, 30)
           '
           RELISA& = RELISA& + 1
           Call GRAREG("LISACLI", Y$, RELISA&)
         End If
       Next JX
        '
        Call SETULTREG("LISACLI", RELISA&)
        Call FILESORT("LISACLI", "", 1, 1, "ASC")
        Call CIERRARCH("*.*")
        Screen.MousePointer = 1
        If Option1(0).Value = True Then
            Call CONECRUN("*DEUACTI", "Consulta de Cuentas Activas")
          Else
            Call CONECRUN("*LDEUINA", "Listado de Cuentas Inactivas")
        End If
        'Call CONECRUN("*DEUACTI", "Listado de Cuentas Activas")
        '
      End If
     '
    ElseIf OPX1% = 1 Then
    '\\\OT-06-0194-WAR-FIN///
      '
      Call ABRECONEXION
      HO1% = HOY(HO$)
      Call REPLA(VDEF$, MKI$(HO1%), 1, 2)
      OBJ1$ = OBJPLA$("ENTRAFECHA1", VDEF$)
      If TRIM$(OBJ1$) = "" Then GoTo 10
      '
      FLIM% = CVI(Left$(OBJ1$, 2))
      If FLIM% > 0 Then
         Screen.MousePointer = 11
         'Dim FACT As ADODB.Recordset
         Set FACT = New ADODB.Recordset
         SQLX$ = "SELECT * FROM cajabanc where tipomovim ='FVEN'"
         Set FACT = FLEXCONN.Execute(FILTSQL$(SQLX$))
         '
         With FACT
         FIN& = CantRegistros("CAJABANC", "tipomovim ='FVEN'")
         If .BOF = True And .EOF = True Then GoTo 80
         .MoveFirst
         For U& = 1 To FIN&
           Call TRACE(20, U&, FIN&)
           X$ = REGLEIDO$("FACTUR", U&)
           NCLIE = XVALO(!nuclien)
           If NCLIE > 0 Then
             If NCLIE <= 999999 Then
               'F1% = FECHANUM(Format(!FechEmisi, "DD/MM/YY"))
               F1% = CVINT(!FECHEMISI)
               If F1% > ACTI%(NCLIE) Then ' If CVI(Mid$(X$, 7, 2)) > ACTI%(NC%) Then
                 ACTI%(NCLIE) = F1% 'CVI(Mid$(X$, 7, 2))
               End If
             End If
           End If
         .MoveNext
         Next U&
         .Close
         End With
         '
80       RELISA& = 0
         cx& = CantRegistros("Deudor12") ' ULTREG&("DEUDOR1")
         Dim clientes As ADODB.Recordset
         Set clientes = FLEXCONN.Execute("select * from deudor12")
         With clientes
         If .BOF = True And .EOF = True Then GoTo 99
         .MoveFirst
         For IL& = 1 To cx&
           Call TRACE(20, IL&, cx&)
           NCLIE = XVALO(!nume_cli) 'CVI(Left$(REGLEIDO$("DEUDOR1", IL&), 2))
           If NCLIE > 0 Then
             If ACTI%(NCLIE) > 0 And ACTI%(NCLIE) < FLIM% Then
                Y$ = REGBLAN$("LISACLI")
                SSCC# = SALDCLI#(NCLIE)
                Call REPLA(Y$, MKS$(NCLIE), 1, 4)
                    RASO$ = rasocli$(NCLIE)
                Call REPLA(Y$, RASO$, 5, 30)
                Call REPLA(Y$, MKD$(SSCC#), 35, 8)
                Call REPLA(Y$, MKD$(SSCC#), 43, 8)
                FEX = ACTI%(NCLIE)
                TTXX$ = "Inactiva Desde " + FECHATEX$(FEX)
                Call REPLA(Y$, TTXX$, 51, 27)
                '
                RELISA& = RELISA& + 1
                Call GRAREG("LISACLI", Y$, RELISA&)
                If MODOMEDI% > 0 Then
                    Z$ = REGBLAN$("LISACLM")
                    Z$ = Y$
                    CONDI$ = "Nume_Cli=" + Str(NC%)
                    TELEFONO$ = TRIM$(VALOBADA("DEUDOR12", "Tele_Cli", CONDI$))
                    mail$ = TRIM$(VALOBADA("DEUDOR12", "Mail_Cli", CONDI$))
                    Call REPLA(Z$, TELEFONO$, 78, 24)
                    Call REPLA(Z$, mail$, 102, 55)
                    Call GRAREG("LISACLM", Z$, RELISA&)
                End If
             End If
           End If
           .MoveNext
         Next IL&
         .Close
         End With
         '
         Call SETULTREG("LISACLI", RELISA&)
         If MODOMEDI% > 0 Then Call SETULTREG("LISACLM", RELISA&)
         Call CIERRARCH("*.*")
         Call FILESORT("LISACLI", "", 1, 1, "ASC")
         If MODOMEDI% > 0 Then Call FILESORT("LISACLM", "", 1, 1, "ASC")
         
         If Option1(0).Value = True Then
            Call CONECRUN("*DEUINAC", "Consulta de Cuentas Inactivas")
          Else
            If MODOMEDI% > 0 Then
                Call CONECRUN("*LDEUINM", "Listado de Cuentas Inactivas")
                GoTo 99
            End If
            Call CONECRUN("*LDEUINA", "Listado de Cuentas Inactivas")
         End If
         '
      End If
    End If
99  Screen.MousePointer = 1
    Command5.Enabled = True

End Sub
Sub Command6_Click()

    Call NuevoListadoCuentasInactivasActivas
    Exit Sub
    
    'GoTo VoyListadoAnterior
    
    ' CODIGO ANTERIOR NO FUNCIONA. SE QUEDA COLGADO.
    On Error GoTo VoyListadoAnterior
    Call VectCliente.cargarvector

    INICIALIZO% = 0
    Dim ACTI%(999999)
    Command5.Enabled = False
    MODOMEDI% = 0
    If InStr(1, EMPREAC$, "MEDITEA") > 0 Then MODOMEDI% = 1
     TITULOX$ = "Generado el: " + FECHATEX$(HOY(HO$)) + " - " + Left$(Time$, 5) + " hs."

    If MODOMEDI% > 0 Then
        TITULOX$ = "Generado el: " + FECHATEX$(HOY(HO$)) + " - " + Left$(Time$, 5) + " hs."
    End If
10  OPX1% = COMALTER%("Seleccione el Listado:\\Clientes Inactivos\Clientes Activos (posterior a)")
    If OPX1% = 2 Then
      '
      TITULOX$ = "Clientes Activos " & TITULOX$
      HO1% = HOY(HO$)
      Call REPLA(VDEF$, MKI$(HO1%), 1, 2)
      OBJ1$ = OBJPLA$("ENTRAFECHA1", VDEF$)
      If TRIM$(OBJ1$) = "" Then GoTo 10
      '
      FLIM% = CVI(Left$(OBJ1$, 2))
      If FLIM% > 0 Then
         Screen.MousePointer = 11
         '
         Call ABRECONEXION
20       Dim FACT As ADODB.Recordset
         Set FACT = New ADODB.Recordset
         SQLX$ = "SELECT * FROM cajabanc where tipomovim ='FVEN'"
         Set FACT = FLEXCONN.Execute(FILTSQL$(SQLX$))
         '
         With FACT
         FIN& = CantRegistros("CAJABANC", "tipomovim ='FVEN'")
         If .BOF = True And .EOF = True Then GoTo 30
         .MoveFirst
         For U& = 1 To FIN&
           Call TRACE(20, U&, FIN&)
           NCLIE = XVALO(!nuclien) 'CVI(Left$(X$, 2))
           F1% = CVINT(!FECHEMISI)
           If F1% >= FLIM% Then
             If NCLIE <= 999999 Then
               If NCLIE > 0 Then
                 If F1% > ACTI%(NCLIE) Then ACTI%(NCLIE) = F1%
               End If
             End If
           End If
           .MoveNext
         Next U&
         .Close
         End With
         '
         '
30       Dim COBRAN As ADODB.Recordset
         Set COBRAN = New ADODB.Recordset
         SQLX$ = "SELECT * FROM CAJABANC WHERE TIPOMOVIM='RCOB'"
         Set COBRAN = FLEXCONN.Execute(FILTSQL$(SQLX$))
         '
         With COBRAN
         FIN& = CantRegistros("CAJABANC", "TIPOMOVIM='RCOB'")
         If .BOF = True And .EOF = True Then GoTo 40
         .MoveFirst
         For U& = 1 To FIN&
           Call TRACE(20, U&, FIN&)
           
           NCLIE = XVALO(!nuclien)
           F1% = CVINT(!FECHEMISI)
           If F1% >= FLIM% Then
             If NCLIE <= 999999 Then
               If NCLIE > 0 Then
                 If F1% > ACTI%(NCLIE) Then ACTI%(NCLIE) = F1%
               End If
             End If
           End If
           .MoveNext
         Next U&
         .Close
         End With
         '
         '
40       Dim PRESUPUESTO As ADODB.Recordset
         Set PRESUPUESTO = New ADODB.Recordset
         SQLX$ = "SELECT * FROM PRESUVEN"
         Set PRESUPUESTO = FLEXCONN.Execute(FILTSQL$(SQLX$))
         '
         With PRESUPUESTO
         FIN& = CantRegistros("PRESUVEN", "")
         If .BOF = True And .EOF = True Then GoTo 50
         .MoveFirst
         For U& = 1 To FIN&
           Call TRACE(20, U&, FIN&)
           NCLIE = XVALO(!NroClie) 'CVI(Left$(X$, 2))
           F1% = CVINT(!FechEmis)
           If F1% >= FLIM% Then
             If NCLIE <= 999999 Then
               If NCLIE > 0 Then
                 If F1% > ACTI%(NCLIE) Then ACTI%(NCLIE) = F1%
               End If
             End If
           End If
           .MoveNext
         Next U&
         .Close
         End With
         '
         '
50       Dim PEDIDOS As ADODB.Recordset
         Set PEDIDOS = New ADODB.Recordset
         SQLX$ = "SELECT * FROM PEDENCA"
         Set PEDIDOS = FLEXCONN.Execute(FILTSQL$(SQLX$))
         '
         With PEDIDOS
         FIN& = CantRegistros("PEDENCA", "")
         If .BOF = True And .EOF = True Then GoTo 60
         .MoveFirst
         For U& = 1 To FIN&
           Call TRACE(20, U&, FIN&)
           NCLIE = XVALO(!NroClie)
           F1% = CVINT(!FechEmis)
           If F1% >= FLIM% Then
             If NCLIE <= 999999 Then
               If NCLIE > 0 Then
                 If F1% > ACTI%(NCLIE) Then ACTI%(NCLIE) = F1%
               End If
             End If
           End If
           .MoveNext
         Next U&
         .Close
         End With
         '
       
60     RELISA& = 0
       For JX = 1 To 999999
         Call TRACE(20, IL&, 999999)
         If JX = 999999 Then Exit For
         NCLIE = JX
         If ACTI%(NCLIE) > 0 Then
           Y$ = REGBLAN$("LISACLI")
           SSCC# = SALDCLI#(NCLIE)
               'RASO$ = rasocli$(NCLIE)
               RASO$ = VectCliente.nombre(NCLIE)
           TTXX$ = "Último Movimiento: " + FECHATEX$(ACTI%(NCLIE))
           Call REPLA(Y$, TTXX$, 51, 30)
           '
           RELISA& = RELISA& + 1
           TEXTOGRILLA$ = SAFETEXT$(NCLIE) & "|" & RASO$ & "|" & FORMATNUM$(SSCC#, "F12.2") & "|" & FORMATNUM$(SSCC#, "F12.2") & "|" & TTXX$
           Campos$ = "Cta/F7;Razon Social/A32;Campo/F12.2;Campo/F12.2;Movimiento/A30"
           Call CARGA_GRILLA_SHOWREP(SHOWREP07, SHOWREP07.GRID, Campos$, TEXTOGRILLA$, INICIALIZO%)
         End If
       Next JX
        '
        Call CIERRARCH("*.*")
        Screen.MousePointer = 1
        '
      End If
         SHOWREP07.Label7 = SHOWREP07.GRID.Rows - 1
         Screen.MousePointer = 1
         SHOWREP07.Frame1.Visible = False
         SHOWREP07.Label3 = TITULOX$

         SHOWREP07.Show 1

     '
    ElseIf OPX1% = 1 Then
      '
      TITULOX$ = "Clientes Inactivos " & TITULOX$
      If MODOMEDI% > 0 Then
        Campos$ = "Cta/F7;Razon Social/A32;Campo/F12.2;Campo/F12.2;Movimiento/A30;Teléfono/A24;Mail/A55"
      Else
        Campos$ = "Cta/F7;Razon Social/A32;Campo/F12.2;Campo/F12.2;Movimiento/A30"
      End If
      
      Call ABRECONEXION
      HO1% = HOY(HO$)
      Call REPLA(VDEF$, MKI$(HO1%), 1, 2)
      OBJ1$ = OBJPLA$("ENTRAFECHA1", VDEF$)
      If TRIM$(OBJ1$) = "" Then GoTo 10
      '
      FLIM% = CVI(Left$(OBJ1$, 2))
      If FLIM% > 0 Then
         Screen.MousePointer = 11
         Set FACT = New ADODB.Recordset
         SQLX$ = "SELECT * FROM cajabanc where tipomovim ='FVEN'"
         Set FACT = FLEXCONN.Execute(FILTSQL$(SQLX$))
         '
         With FACT
         FIN& = CantRegistros("CAJABANC", "tipomovim ='FVEN'")
         If .BOF = True And .EOF = True Then GoTo 80
         .MoveFirst
         For U& = 1 To FIN&
           Call TRACE(20, U&, FIN&)
           X$ = REGLEIDO$("FACTUR", U&)
           NCLIE = XVALO(!nuclien)
           If NCLIE > 0 Then
             If NCLIE <= 999999 Then
               F1% = CVINT(!FECHEMISI)
               If F1% > ACTI%(NCLIE) Then ' If CVI(Mid$(X$, 7, 2)) > ACTI%(NC%) Then
                 ACTI%(NCLIE) = F1% 'CVI(Mid$(X$, 7, 2))
               End If
             End If
           End If
         .MoveNext
         Next U&
         .Close
         End With
         '
80       RELISA& = 0
         cx& = CantRegistros("Deudor12") ' ULTREG&("DEUDOR1")
         Dim clientes As ADODB.Recordset
         Set clientes = FLEXCONN.Execute("select * from deudor12")
         With clientes
         If .BOF = True And .EOF = True Then GoTo 99
         .MoveFirst
         For IL& = 1 To cx&
           Call TRACE(20, IL&, cx&)
           NCLIE = XVALO(!nume_cli) 'CVI(Left$(REGLEIDO$("DEUDOR1", IL&), 2))
           If NCLIE > 0 Then
             If ACTI%(NCLIE) > 0 And ACTI%(NCLIE) < FLIM% Then
                Y$ = REGBLAN$("LISACLI")
                SSCC# = SALDCLI#(NCLIE)
                    'RASO$ = rasocli$(NCLIE)
                    RASO$ = VectCliente.nombre(NCLIE)
                FEX = ACTI%(NCLIE)
                TTXX$ = "Inactiva Desde " + FECHATEX$(FEX)
                '
                RELISA& = RELISA& + 1
                Call GRAREG("LISACLI", Y$, RELISA&)
                TEXTOGRILLA$ = SAFETEXT$(NCLIE) & "|" & RASO$ & "|" & FORMATNUM$(SSCC#, "F12.2") & "|" & FORMATNUM$(SSCC#, "F12.2") & "|" & TTXX$
                
                If MODOMEDI% > 0 Then
                    Z$ = REGBLAN$("LISACLM")
                    Z$ = Y$
                    'CONDI$ = "Nume_Cli=" + Str(NC%)
                    'TELEFONO$ = TRIM$(VALOBADA("DEUDOR12", "Tele_Cli", CONDI$))
                    TELEFONO$ = VectCliente.TeleCli(NCLIE)
'                    mail$ = TRIM$(VALOBADA("DEUDOR12", "Mail_Cli", CONDI$))
                    mail$ = VectCliente.mailcliente(NCLIE)
                    TEXTOGRILLA$ = TEXTOGRILLA$ & "|" & TELEFONO$ + "|" & mail$
                End If
                Call CARGA_GRILLA_SHOWREP(SHOWREP07, SHOWREP07.GRID, Campos$, TEXTOGRILLA$, INICIALIZO%)
             End If
           End If
           .MoveNext
         Next IL&
         .Close
         End With
         '
         SHOWREP07.Label7 = SHOWREP07.GRID.Rows - 1
         Screen.MousePointer = 1
         SHOWREP07.Frame1.Visible = False
         SHOWREP07.Label3 = TITULOX$
         SHOWREP07.Show 1
         '
      End If
    End If
    GoTo 99
    
VoyListadoAnterior:
    Call Command6_RESPALDO_Click
    
99  Screen.MousePointer = 1
    Command5.Enabled = True
 
End Sub

Sub Command8_Click()

    If Option1(0).Value = True Then
        Option2.Value = True
    End If
    Command8.Enabled = False
    FORLICLI07.Show 1
    Command8.Enabled = True
End Sub

Sub LILIMICRE_RES()
    '
    Dim FacPenTemp As ADODB.Recordset
    Dim CABATEMP As ADODB.Recordset
    '
    ' falta considerar deuda en cheques en cartera
    '
    Screen.MousePointer = 11
    Dim DEUCUECO#(999999), CHEPEAC#(999999)
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM SADEUPEN "
    Set FacPenTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With FacPenTemp
       If Not (.BOF And .EOF) Then
          Do While Not .EOF
            NCLIE = !nuclien
            If NCLIE >= 0 Then
              SALDEU# = !SaldDebe - !SaldAcre
              DEUCUECO#(NCLIE) = DEUCUECO#(NCLIE) + SALDEU#
            End If
          .MoveNext
          Loop
       End If
    End With
    '
    ' INCLUYE CHEQUES PENDIENTES DE ACREDITACION
    HOYDIA = FETEXCOR1$(CVINT(Int(Now)))
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
    SQLX$ = SQLX$ + "AND Status < 9 "
    SQLX$ = SQLX$ + "AND FechAcred >= '" & HOYDIA & "' "
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        '
    With CABATEMP
       If Not (.BOF And .EOF) Then
          Do While Not .EOF
              If TRIM$(!OpciMovim) <> "CH" Then GoTo 19 ' no es un cheque
              '
              ICHEQ# = !ValAbsComp
              NCLIE = !nuclien
              '
              If NCI% > 0 Then
                If NCI% <= 999999 Then
                  CHEPEAC#(NCLIE) = CHEPEAC#(NCLIE) + ICHEQ#
                End If
              End If
              '
19        .MoveNext
          Loop
       End If
    End With
    '
    FIN& = 0
    Dim rstdeudor As ADODB.Recordset
    Set rstdeudor = FLEXCONN.Execute("select max(nume_cli) as max from deudor12")
    If Not (rstdeudor.EOF And rstdeudor.BOF) Then FIN& = rstdeudor!Max
    jj& = 0
    For IL& = 1 To FIN&
       Call TRACE(20, IL&, FIN&)
       NCLIE = IL&    '   CVI(Left$(XX$, 2))
       If Abs(DEUCUECO#(NCLIE)) + Abs(CHEPEAC#(NCLIE)) >= 0.005 Then
          CRELIM# = LIMICRE(NCLIE) 'CVS(Mid$(YY$, 121, 4))
          DISPO# = CRELIM# - DEUCUECO#(NCLIE) - CHEPEAC#(NCLIE)
          EXCES# = 0
          If DISPO# < 0 Then
            EXCES# = Abs(DISPO#)
            DISPO# = 0
          End If
          ZZ$ = REGBLAN$("LILICRE")
          Call REPLA(ZZ$, MKS$(NCLIE), 1, 4)
            RASO$ = rasocli$(NCLIE)
          Call REPLA(ZZ$, RASO$, 5, 30)
          Call REPLA(ZZ$, MKD$(CRELIM#), 35, 8)
          If CRELIM# >= 0 Then
            Call REPLA(ZZ$, MKD$(DEUCUECO#(NCLIE)), 43, 8)
            Call REPLA(ZZ$, MKD$(CHEPEAC#(NCLIE)), 51, 8)
            Call REPLA(ZZ$, MKD$(DISPO#), 59, 8)
            Call REPLA(ZZ$, MKD$(EXCES#), 67, 8)
          End If
          jj& = jj& + 1
          Call GRAREG("LILICRE", ZZ$, jj&)
       End If
    Next IL&
    '
    Call SETULTREG("LILICRE", jj&)
    Call CIERRARCH("LILICRE")
    '
    OPSALI$ = "IMPRE"
    On Error GoTo 40
    If Option1(0).Value = True Then OPSALI$ = "SCREEN"
    GoTo 41
    '
40  Resume 41
41  On Error GoTo 50
    If Option1(0).Value = True Then OPSALI$ = "SCREEN"
    GoTo 51
    '
50  Resume 51
51  On Error GoTo 0
    If OPSALI$ = "SCREEN" Then
        Call CONECRUN("*COLICRE", "Consulta de Límites de Crédito")
      Else
        Call CONECRUN("*LILICRE", "Listado de Límites de Crédito")
    End If
    Screen.MousePointer = 1
   '
End Sub

Sub LILIMICRE()
    '
    If DERACCE%("COLICRE", "Consulta de Límites de Crédito") < 2 Then Exit Sub
    If DERACCE%("LILICRE", "Listado de Límites de Crédito") < 2 Then Exit Sub
    
    Dim FacPenTemp As ADODB.Recordset
    Dim CABATEMP As ADODB.Recordset
    '
    ' falta considerar deuda en cheques en cartera
    '
    Screen.MousePointer = 11
    Dim DEUCUECO#(999999), CHEPEAC#(999999)
    '
    Call VectCliente.cargarvector

    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM SADEUPEN "
    Set FacPenTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With FacPenTemp
       If Not (.BOF And .EOF) Then
          Do While Not .EOF
            NCLIE = XVALO(!nuclien)
            If NCLIE >= 0 Then
              SALDEU# = XVALO(!SaldDebe) - XVALO(!SaldAcre)
              DEUCUECO#(NCLIE) = DEUCUECO#(NCLIE) + SALDEU#
            End If
          .MoveNext
          Loop
       End If
    End With
    '
    ' INCLUYE CHEQUES PENDIENTES DE ACREDITACION
    HOYDIA = FETEXCOR1$(CVINT(Int(Now)))
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
    SQLX$ = SQLX$ + "AND Status < 9 "
    SQLX$ = SQLX$ + "AND FechAcred >= '" & HOYDIA & "' "
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        '
    With CABATEMP
       If Not (.BOF And .EOF) Then
          Do While Not .EOF
              If TRIM$(!OpciMovim) <> "CH" Then GoTo 19 ' no es un cheque
              '
              ICHEQ# = !ValAbsComp
              NCLIE = !nuclien
              '
              If NCI% > 0 Then
                If NCI% <= 999999 Then
                  CHEPEAC#(NCLIE) = CHEPEAC#(NCLIE) + ICHEQ#
                End If
              End If
              '
19        .MoveNext
          Loop
       End If
    End With
    '
    FIN& = 0
    Dim rstdeudor As ADODB.Recordset
    Set rstdeudor = FLEXCONN.Execute("select max(nume_cli) as max from deudor12")
    If Not (rstdeudor.EOF And rstdeudor.BOF) Then FIN& = rstdeudor!Max
    jj& = 0
    For IL& = 1 To FIN&
       Call TRACE(20, IL&, FIN&)
       NCLIE = IL&    '   CVI(Left$(XX$, 2))
       If Abs(DEUCUECO#(NCLIE)) + Abs(CHEPEAC#(NCLIE)) >= 0.005 Then
          CRELIM# = LIMICRE(NCLIE) 'CVS(Mid$(YY$, 121, 4))
          DISPO# = CRELIM# - DEUCUECO#(NCLIE) - CHEPEAC#(NCLIE)
          EXCES# = 0
          If DISPO# < 0 Then
            EXCES# = Abs(DISPO#)
            DISPO# = 0
          End If
            RASO$ = VectCliente.nombre(NCLIE)
          If CRELIM# >= 0 Then
          End If
          jj& = jj& + 1
TEXTOGRILLA$ = SAFETEXT$(NCLIE) & "|" & RASO$ & "|" & FORMATNUM$(CRELIM#, "F12.2") & "|"
TEXTOGRILLA$ = TEXTOGRILLA$ & FORMATNUM$(DEUCUECO#(NCLIE), "F12.2") & "|" & FORMATNUM$(CHEPEAC#(NCLIE), "F12.2") & "|"
TEXTOGRILLA$ = TEXTOGRILLA$ & FORMATNUM$(DISPO#, "F12.2") & "|" & FORMATNUM$(EXCES#, "F12.2")
Campos$ = "Cta/F7;Razon Social/A32;L.Crédito/F12.2;Cta.Corr./F12.2;Cheques/F12.2;Cr.Dispon/F12.2;Excedido/F12.2"
Call CARGA_GRILLA_SHOWREP(SHOWREP07, SHOWREP07.GRID, Campos$, TEXTOGRILLA$, INICIALIZO%)
       End If
    Next IL&
    '
SHOWREP07.Label7 = SHOWREP07.GRID.Rows - 1
Screen.MousePointer = 1
SHOWREP07.Frame1.Visible = False
SHOWREP07.Label3 = "Listado de Límites de Crédito"
Screen.MousePointer = 1
SHOWREP07.Show 1


   '
End Sub


Sub GELIGEPRO()
    '
    Dim TICUE%(32767)
    Dim DETICUE$(2)
    Screen.MousePointer = 11
    '
    For U& = 1 To 32767: TICUE%(U&) = 1: Next U&
    DETICUE$(0) = "Cuenta Corriente"
    DETICUE$(1) = "Gastos"
    FIN& = ULTREG&("ACREDOR")
    '
    For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        NCX% = CVI(Left$(REGLEIDO$("ACREDOR", U&), 2))
        If NCX% > 0 Then
            TICUE%(NCX%) = 0
        End If
    Next U&
    '
    FIN& = ULTREG&("PROVED1")
    ULI& = 0
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("PROVED1", U&)
        NP% = CVI(Left$(X$, 2))
        NC% = (-1) * NP%
        '
        Y$ = REGBLAN$("LIGEPRO")
        Call REPLA(Y$, MKI$(NP%), 1, 2)
        CUECOX% = TICUE%(NP%)
        Call REPLA(Y$, DETICUE$(CUECOX%), 3, 30)
        ULI& = ULI& + 1
        Call GRAREG("LIGEPRO", Y$, ULI&)
        '
    Next U&
    '
    Call SETULTREG("LIGEPRO", ULI&)
    Call CIERRARCH("LIGEPRO")
    Screen.MousePointer = 1
    Call FINAL("*LIGEPRO")
    '
End Sub
'




Private Sub Form_Load()
    
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    
    Dim CLITEMP As ADODB.Recordset
End Sub


