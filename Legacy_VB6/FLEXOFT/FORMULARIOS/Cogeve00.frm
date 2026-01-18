VERSION 5.00
Begin VB.Form COGEVE00 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0F0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Gestión Comercial - Cuentas por Cobrar"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5850
   FillStyle       =   0  'Solid
   Icon            =   "Cogeve00.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0F0FF&
      Caption         =   "Ventas y Deudores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6525
      Left            =   210
      TabIndex        =   3
      Top             =   210
      Width           =   4425
      Begin VB.Frame Frame7 
         BackColor       =   &H00D0F0FF&
         Caption         =   "Resúmenes Mensuales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   420
         TabIndex        =   24
         Top             =   5145
         Width           =   3585
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   25
            Top             =   370
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
               Height          =   600
               Left            =   0
               Picture         =   "Cogeve00.frx":030A
               Style           =   1  'Graphical
               TabIndex        =   26
               ToolTipText     =   "Generación y Emisión de Listados y Estadísticas Mensuales."
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Subdiarios y Estadística"
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
               TabIndex        =   27
               Top             =   105
               Width           =   1485
            End
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00D0F0FF&
         Caption         =   "Control de Estado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3270
         Left            =   420
         TabIndex        =   5
         Top             =   1680
         Width           =   3585
         Begin VB.PictureBox Picture11 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   930
            ScaleHeight     =   585
            ScaleWidth      =   2370
            TabIndex        =   28
            Top             =   1680
            Width           =   2430
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
               Height          =   600
               Left            =   -20
               Picture         =   "Cogeve00.frx":0614
               Style           =   1  'Graphical
               TabIndex        =   29
               ToolTipText     =   "Listados de Avance de Facturación y Cobranza por Día"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Avance de Factura- cion y Cobranza"
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
               TabIndex        =   30
               Top             =   105
               Width           =   1710
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00D0F0FF&
            Caption         =   "Salida"
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
            Left            =   210
            TabIndex        =   18
            Top             =   2460
            Width           =   3165
            Begin VB.OptionButton Option2 
               BackColor       =   &H00D0F0FF&
               Caption         =   "Impresión"
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
               Left            =   1785
               TabIndex        =   20
               Top             =   250
               Width           =   1275
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00D0F0FF&
               Caption         =   "Pantalla"
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
               Left            =   315
               TabIndex        =   19
               Top             =   250
               Value           =   -1  'True
               Width           =   1380
            End
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   690
            ScaleHeight     =   585
            ScaleWidth      =   2370
            TabIndex        =   15
            Top             =   1260
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
               Picture         =   "Cogeve00.frx":091E
               Style           =   1  'Graphical
               TabIndex        =   16
               ToolTipText     =   "Listado de Saldos y Deudas de Clientes"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Saldos y Deudas de Clientes"
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
               TabIndex        =   17
               Top             =   0
               Width           =   1695
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   450
            ScaleHeight     =   585
            ScaleWidth      =   2370
            TabIndex        =   12
            Top             =   840
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
               Picture         =   "Cogeve00.frx":0C28
               Style           =   1  'Graphical
               TabIndex        =   13
               ToolTipText     =   "Listado de Planificación Financiera de Cobranza"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Flujo de Caja Proyectado"
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
               TabIndex        =   14
               Top             =   0
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   210
            ScaleHeight     =   585
            ScaleWidth      =   2370
            TabIndex        =   9
            Top             =   420
            Width           =   2430
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
               Picture         =   "Cogeve00.frx":0F32
               Style           =   1  'Graphical
               TabIndex        =   10
               ToolTipText     =   "Listados de Facturas Pendientes de Cobranza"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Cuentas por Cobrar"
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
               TabIndex        =   11
               Top             =   105
               Width           =   1905
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00D0F0FF&
         Caption         =   "Detalle por Cuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   420
         TabIndex        =   4
         Top             =   420
         Width           =   3585
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2895
            TabIndex        =   6
            Top             =   370
            Width           =   2955
            Begin VB.CommandButton Command20 
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
               Left            =   2340
               Picture         =   "Cogeve00.frx":123C
               Style           =   1  'Graphical
               TabIndex        =   34
               ToolTipText     =   "Listado de Límites de Crédito y Deudas"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command7 
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
               Picture         =   "Cogeve00.frx":1546
               Style           =   1  'Graphical
               TabIndex        =   7
               ToolTipText     =   "Consulta de Resumen de Cuenta Corriente por Cliente"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Resúmenes de Cuenta Corriente"
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
               TabIndex        =   8
               Top             =   105
               Width           =   1485
            End
         End
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4800
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   32
      Top             =   2130
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   3990
      Sorted          =   -1  'True
      TabIndex        =   31
      Top             =   105
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Articulos"
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
      Left            =   4800
      Picture         =   "Cogeve00.frx":1850
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Acceso a Maestro de Items de Stock"
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command18 
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
      Height          =   800
      Left            =   4800
      Picture         =   "Cogeve00.frx":199A
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Configuración de Procesos y Formularios"
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Clientes"
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
      Left            =   4800
      Picture         =   "Cogeve00.frx":1DDC
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Acceso a Maestro de Clientes"
      Top             =   5040
      Width           =   855
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
      Left            =   4800
      Picture         =   "Cogeve00.frx":221E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   340
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
      Left            =   4800
      Picture         =   "Cogeve00.frx":2368
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   1180
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Direct"
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
      Left            =   4800
      Picture         =   "Cogeve00.frx":2672
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Acceso Directo a Otros Procesos"
      Top             =   2520
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4260
      Picture         =   "Cogeve00.frx":297C
      Top             =   6210
      Width           =   2010
   End
End
Attribute VB_Name = "COGEVE00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command16_Click()
   Command16.Enabled = False
   Call CIERRARCH("*.*")
   Call CONECRUN("ARCH_GES/AMACLI", "Archivos Maestros")
   Command16.Enabled = True
End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   Call CONECRUN("FLSETUP/OPCLAVEN", "Configuración de Funcionamiento")
   Command18.Enabled = True
End Sub

Private Sub Command19_Click()
    Command19.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("ARCH_GES/AMASTO", "Archivos Maestros")
    Command19.Enabled = True
End Sub

'
Private Sub Command2_Click()
    CARGAMEN% = 0
    REMEVE00.Show 1
End Sub

Sub Command20_Click()
   Command20.Enabled = False
   Call LILIMCRE
   Command20.Enabled = True
End Sub

Private Sub Command29_Click()
   Command29.Enabled = False
   Call HELPONLINE("COGEVEN")
   Command29.Enabled = True
End Sub

Sub Command3_Click()
    Screen.MousePointer = 11
    Dim SALDEUPEN#(32767), PUNSALDEU$(32767)
    '
    'Call CIERRARCH("!SVD202")
    'Call BLANARCH("!SVD202")
10  RECONMOVI& = 0
    '
    If RECLEN%("FAPECOB") < 64 Then
       Call MENSERR(24, "Longitud de Registro Incorrecta\Archivo 'FAPECOB' - Pasar a 64 by.")
       Screen.MousePointer = 1
       Exit Sub
    End If
    '
    FIN& = ULTREG&("SVD202")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("SVD202", U&)
      TIX% = Asc(Mid$(X$, 3, 1))
      If TIX% = 7 Then TIX% = 0
      IM1# = CVD(Mid$(X$, 11, 8))
      IM2# = CVD(Mid$(X$, 21, 8))
      If TIX% > 0 Then
        If IM1# <> Abs(IM1#) Or IM2# <> Abs(IM2#) Then
          IM1# = Abs(IM1#): IM2# = Abs(IM2#)
          Call REPLA(X$, Chr$(TIX%), 3, 1)
          Call REPLA(X$, MKD$(IM1#), 11, 8)
          Call REPLA(X$, MKD$(IM2#), 21, 8)
          Call GRAREG("SVD202", X$, U&)
        End If
      End If
      '
      MONEMX% = Asc(Mid$(X$, 31, 1))
      If MONEMX% < 1 Or MONEMX% > 2 Then GoSub 200
      '
      NC% = CVI(Left$(X$, 2))
      If NC% > 0 Then
        If NC% < 32768 Then
          If Abs(IM1#) >= 0.005 Then
            SIG% = 1: If TIX% > 3 And TIX% < 7 Then SIG% = (-1)
            SALDEUPEN#(NC%) = SALDEUPEN#(NC%) + SIG% * IM1#
            PUNSALDEU$(NC%) = PUNSALDEU$(NC%) + MKS$(U&)
            RECONMOVI& = RECONMOVI& + 1
              'JJ& = JJ& + 1
              'Call GRAREG("!SVD202", X$, JJ&)
          End If
        End If
      End If
    Next U&
    Call CIERRARCH("SVD202")
    'Call CIERRARCH("!SVD202")
    '
    URENUE& = RECONMOVI&
    UREVIE& = ULTREG&("SVD202")
    '
    If URENUE& < UREVIE& * 0.8 Then
       TTXX$ = "Hay Mucho Espacio Desaprovechado\en la Base de Datos de Facturas Pendientes de Cobranza."
       TTXX$ = TTXX$ + "\  \Compactar esta Base de Datos Puede Acelerar\Significativamente el Funcionamiento del Sistema."
       TTXX$ = TTXX$ + "\  \Para Efectuar la Compactación, TODOS los Restantes Usuarios\Deben Salir de FLEXOFT hasta el Escritorio Principal."
       TTXX$ = TTXX$ + "\  \Desea Ahora Compactar la Base de Datos\de Facturas Pendientes de Cobranza ?"
       If COMALTER%(TTXX$ + "\\Continuar\Compactar Ahora") = 2 Then
          Call COMPACDEU
          For KKIL& = 0 To 32767
            SALDEUPEN#(KKIL&) = 0
            PUNSALDEU$(KKIL&) = ""
          Next KKIL&
          GoTo 10
       End If
    End If
    '
    FIN& = ULTREG&("DEUDOR1")
    Screen.MousePointer = 11
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("DEUDOR1", U&)
      NC% = CVI(Left$(X$, 2))
      If NC% > 0 Then
        If NC% <= 32767 Then
          X2$ = REGLEIDO$("DEUDOR2", U&)
          SCLX# = CVD(Mid$(X2$, 109, 8))
          If Abs(SCLX# - SALDEUPEN#(NC%)) >= 0.005 Then
              'If Abs(SALDCLI#(NC%) - SALDEUPEN#(NC%)) >= 0.005 Then
              'SCLX# = SALDCLI#(NC%)
            GoSub 100
                 'If HABIAREPE% = 0 Then
                 '   RESCUDEU.Text1.TEXT = TRIM$(Str$(NC%))
                 '   Call CARESCUE
                 '   RESCUDEU.Show 1
                 '   Call FRESHALL
                 'End If
                 'Screen.MousePointer = 11
          End If
        End If
      End If
    Next U&
    '
    Screen.MousePointer = 1
    '
    CARFAPEN% = 0
    FAPECOBRA.Show 1
    Exit Sub
    '
100 SALDEUPEN#(NC%) = 0
    HABIAREPE% = 0
    List1.Clear
    XXXX = Len(PUNSALDEU$(NC%))
    For U1& = 1 To Len(PUNSALDEU$(NC%)) Step 4
       REG1& = CVS(Mid$(PUNSALDEU$(NC%), U1&, 4))
       X1$ = REGLEIDO$("SVD202", REG1&)
       SUBIT% = 0
       For KKI% = 1 To Len(X1$)
          SUBIT% = SUBIT% + Asc(Mid$(X1$, KKI%, 1))
       Next KKI%
       TTXX$ = Space$(20)
       Call REPLA(TTXX$, Str$(SUBIT%), 1, 10)
       Call REPLA(TTXX$, Str$(U1&), 11, 10)
       Call REPLA(TTXX$, "A", 1, 1)
       List1.AddItem TTXX$
    Next U1&
    '
    For U1& = 1 To List1.ListCount
       Y1$ = List1.List(U1& - 1)
       For U2& = U1& + 1 To List1.ListCount
          Y2$ = List1.List(U2& - 1)
          If Left$(Y1$, 10) <> Left$(Y2$, 10) Then GoTo 110
          REG1& = Val(Mid$(Y1$, 11))
          REG2& = Val(Mid$(Y2$, 11))
          X1$ = REGLEIDO$("SVD202", REG1&)
          X2$ = REGLEIDO$("SVD202", REG2&)
          If X1$ = X2$ Then
            Call REPLA(X2$, String$(30, 0), 3, 30)
            Call GRAREG("SVD202", X2$, REG2&)
            HABIAREPE% = 1
          End If
       Next U2&
110 Next U1&
Return
    
    
    
    For U1& = 1 To Len(PUNSALDEU$(NC%)) Step 4
      REG1& = CVS(Mid$(PUNSALDEU$(NC%), U1&, 4))
      For U2& = U1& + 4 To Len(PUNSALDEU$(NC%)) Step 4
         REG2& = CVS(Mid$(PUNSALDEU$(NC%), U2&, 4))
         X1$ = REGLEIDO$("SVD202", REG1&)
         X2$ = REGLEIDO$("SVD202", REG2&)
         If X1$ = X2$ Then
            Call REPLA(X2$, String$(30, 0), 3, 30)
            Call GRAREG("SVD202", X2$, REG2&)
            HABIAREPE% = 1
         End If
      Next U2&
      '
      X$ = REGLEIDO$("SVD202", U1&)
      TIX% = Asc(Mid$(X$, 3, 1))
      If TIX% = 7 Then TIX% = 0
      IM1# = CVD(Mid$(X$, 11, 8))
      IM2# = CVD(Mid$(X$, 21, 8))
      SIG% = 1: If TIX% > 3 And TIX% < 7 Then SIG% = (-1)
      SALDEUPEN#(NC%) = SALDEUPEN#(NC%) + SIG% * IM1#
      '
    Next U1&
    Return
    '
200 NC% = CVI(Left$(X$, 2))
    IMPO# = CVD(Mid$(X$, 21, 8))
    SALDO# = CVD(Mid$(X$, 11, 8))
    NROL& = CVS(Mid$(X$, 5, 4))
    MONEMX% = Asc(Mid$(X$, 31, 1))
    '
    If MONEMX% > 2 Then
      MONEMX% = 0
      Call REPLA(X$, Chr$(0), 31, 1)
      Call GRAREG("SVD202", X$, U&)
    End If
    '
    If MONEMX% = 0 Then
      RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
      For KUL& = Len(RFG$) - 3 To 1 Step -4
        RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
        XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
        If CVS(Mid$(XXC$, 39, 4)) = NROL& Then
          If Asc(Mid$(XXC$, 37, 1)) = TIX% Then
            MONEMX% = Asc(Mid$(X$, 31, 1))
            GoTo 215
          End If
        End If
      Next KUL&
    End If
    '
215 If MONEMX% <> 2 Then MONEMX% = 1
    Call REPLA(X$, Chr$(MONEMX%), 31, 1)
    Call GRAREG("SVD202", X$, U&)
    '
Return
'
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    If EXISTE%("INFOPAT.EXE") > 0 Then
          Call CONECRUN("INFOPAT", "Informe Patrimonial Diario")
        Else
          Call MENSERR(24, "Opción no Disponible\en la Presente Versión FLEXOFT.\  \Utilice las Opciones Habilitadas\en el Módulo de Cuentas por Cobrar.")
          'Call CONECRUN("#SVD47", "Plan Financiero de Cobranza")
    End If
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call GELISALCLI
    Command5.Enabled = True
End Sub

Sub Command7_Click()
    'Hide
    RECUEC00.Show 1
End Sub

Private Sub Command8_Click()
    Call ACCDIR("COGEVEN")
End Sub

Sub Command9_Click()
    FORSELME.Show 1
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 0
    Call CIERRARCH("*.*")
    Hide
End Sub

Sub COMPACDEU()
   '
   Call CIERRARCH("*.*")
   Call CIERRARCH("!SVD202")
   Call BLANARCH("!SVD202")
   JJ& = 0
   '
   FIN& = ULTREG&("SVD202")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("SVD202", U&)
     IM1# = CVD(Mid$(X$, 11, 8))
     NC% = CVI(Left$(X$, 2))
     If NC% > 0 Then
       If NC% <= 32767 Then
         If Abs(IM1#) >= 0.005 Then
           JJ& = JJ& + 1
           Call GRAREG("!SVD202", X$, JJ&)
         End If
       End If
     End If
   Next U&
   Call CIERRARCH("SVD202")
   Call CIERRARCH("!SVD202")
   '
   On Error GoTo 10
   Kill LUDAT$ + "SVD202.X*"
   Kill LUDAT$ + "SVD202.DAT"
   On Error GoTo 0
   '
   FIN& = ULTREG&("!SVD202")
   JJ& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("!SVD202", U&)
     JJ& = JJ& + 1
     Call GRAREG("SVD202", X$, JJ&)
   Next U&
   Call CIERRARCH("SVD202")
   Call CIERRARCH("*.*")
   Call COMUNI("Compactación Completa.")
   Exit Sub
   '
10 Resume 20
20 On Error GoTo 0
   Call COMUNI("La Compactación no Pudo Completarse.")
   '
End Sub


Sub GELISALCLI()
    '
    HOXX = HOY(HO$)
    Dim NC%(16384), DEUDA#(16384), SALDO#(16384)
    '
    Screen.MousePointer = 11
    Call HEADERS("LISACLI", "")
    Call HEADERS("LISACLI", "Generado el: " + HO$ + " - " + Left$(Time$, 5) + " hs.")
    '
    CX& = ULTREG&("DEUDOR1")
    '
    For IL& = 1 To CX&
      Call TRACE(20, IL&, CX&)
      X$ = REGLEIDO$("DEUDOR1", IL&)
      NC%(IL&) = CVI(Left$(X$, 2))
      SALDO#(IL&) = SALDCLI#(NC%(IL&))
      DEUDA#(IL&) = 0
    Next IL&
    '
    FIN& = ULTREG&("SVD202")
    For IL& = 1 To FIN&
      SS$ = REGLEIDO$("SVD202", IL&)
      NCI% = CVI(Left$(SS$, 2))
      TI% = Asc(Mid$(SS$, 3, 1))
      If TI% > 7 Then TI% = 2
      SALDEU# = CVD(Mid$(SS$, 11, 8))
      If TI% > 3 Then SALDEU# = (-1) * Abs(SALDEU#)
      For JL& = 1 To CX&
        If NC%(JL&) = NCI% Then
          DEUDA#(JL&) = DEUDA#(JL&) + SALDEU#
          GoTo 809
        End If
      Next JL&
809 Next IL&

900 RELISA& = 0
    '
    For IL& = 1 To CX&
      If NC%(IL&) > 0 Then
        If Abs(SALDO#(IL&)) >= 0.005 Or Abs(DEUDA#(IL&)) >= 0.005 Then
          Y$ = REGBLAN$("LISACLI")
          Call REPLA(Y$, MKI$(NC%(IL&)), 1, 2)
          Call REPLA(Y$, MKD$(SALDO#(IL&)), 3, 8)
          Call REPLA(Y$, MKD$(DEUDA#(IL&)), 11, 8)
          If Abs(SALDO#(IL&) - DEUDA#(IL&)) > 0.005 Then
            Call REPLA(Y$, "ATENCION: Saldo <> Deuda", 19, 30)
          End If
          '
          RELISA& = RELISA& + 1
          Call GRAREG("LISACLI", Y$, RELISA&)
        End If
      End If
    Next IL&
    '
    Call SETULTREG("LISACLI", RELISA&)
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    If Option1.Value = True Then
         Call FINAL("*LISACLI")
       Else
         Call FINAL("*LIMPCLI")
    End If
    '
End Sub
'
Sub LILIMCRE()
   '
   Screen.MousePointer = 11
   '
   Dim DEUCUECO#(16384), CHEPEAC#(16384)
   '
   FIN& = ULTREG&("SVD202")
   For IL& = 1 To FIN&
      Call TRACE(20, IL&, FIN&)
      SS$ = REGLEIDO$("SVD202", IL&)
      NCI% = CVI(Left$(SS$, 2))
      If NCI% > 0 Then
         If NCI% <= 16384 Then
            TI% = Asc(Mid$(SS$, 3, 1))
            If TI% > 7 Then TI% = 2
            SALDEU# = CVD(Mid$(SS$, 11, 8))
            If TI% > 3 Then SALDEU# = (-1) * Abs(SALDEU#)
            DEUCUECO#(NCI%) = DEUCUECO#(NCI%) + SALDEU#
            GoTo 19
         End If
      End If
19 Next IL&
   '
   HOII% = HOY(HOS$)
   FIN& = ULTREG&("CHECAR")
   For IL& = 1 To FIN&
      Call TRACE(20, IL&, FIN&)
      XX$ = REGLEIDO$("CHECAR", IL&)
      NCI% = CVI(Mid$(XX$, 3, 2))
      If NCI% > 0 Then
         If NCI% <= 16384 Then
            If CVI(Mid$(XX$, 41, 2)) >= HOII% Then
               CHEPEAC#(NCI%) = CHEPEAC#(NCI%) + CVD(Mid$(XX$, 33, 8))
            End If
         End If
      End If
   Next IL&
   '
   FIN& = ULTREG&("DEUDOR1")
   JJ& = 0
   For IL& = 1 To FIN&
      Call TRACE(20, IL&, FIN&)
      XX$ = REGLEIDO$("DEUDOR1", IL&)
      NCI% = CVI(Left$(XX$, 2))
      If NCI% > 0 Then
         If NCI% <= 16384 Then
            If Abs(DEUCUECO#(NCI%)) + Abs(CHEPEAC#(NCI%)) >= 0.005 Then
               YY$ = REGLEIDO$("DEUDOR2", IL&)
               CRELIM# = CVS(Mid$(YY$, 121, 4))
               DISPO# = CRELIM# - DEUCUECO#(NCI%) - CHEPEAC#(NCI%)
               EXCES# = 0
               If DISPO# < 0 Then
                  EXCES# = Abs(DISPO#)
                  DISPO# = 0
               End If
               ZZ$ = REGBLAN$("LILICRE")
               Call REPLA(ZZ$, MKI$(NCI%), 1, 2)
               Call REPLA(ZZ$, MKD$(CRELIM#), 3, 8)
               Call REPLA(ZZ$, MKD$(DEUCUECO#(NCI%)), 11, 8)
               Call REPLA(ZZ$, MKD$(CHEPEAC#(NCI%)), 19, 8)
               Call REPLA(ZZ$, MKD$(DISPO#), 27, 8)
               Call REPLA(ZZ$, MKD$(EXCES#), 35, 8)
               JJ& = JJ& + 1
               Call GRAREG("LILICRE", ZZ$, JJ&)
            End If
         End If
      End If
   Next IL&
   '
   Call SETULTREG("LILICRE", JJ&)
   Call CIERRARCH("LILICRE")
   '
   If Option1.Value = True Then
        Call CONECRUN("*COLICRE", "Consulta de Límites de Crédito")
      Else
        Call CONECRUN("*LILICRE", "Listado de Límites de Crédito")
   End If
   Screen.MousePointer = 1
   '
End Sub

