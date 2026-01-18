VERSION 5.00
Begin VB.Form COGEVE04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0F0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Gestión Comercial - Cuentas por Cobrar"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5640
   FillStyle       =   0  'Solid
   Icon            =   "Cogeve04.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00404080&
      Height          =   6945
      Left            =   4725
      ScaleHeight     =   6885
      ScaleWidth      =   1635
      TabIndex        =   27
      Top             =   0
      Width           =   1695
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
         Height          =   690
         Left            =   120
         Picture         =   "Cogeve04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Importar, Descargar y Acceso Directo a Otros Procesos"
         Top             =   3390
         Width           =   650
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
         Height          =   690
         Left            =   120
         Picture         =   "Cogeve04.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   1000
         Width           =   650
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
         Height          =   690
         Left            =   120
         Picture         =   "Cogeve04.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   310
         Width           =   650
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Cust."
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
         Left            =   120
         Picture         =   "Cogeve04.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Acceso a Maestro de Clientes"
         Top             =   2010
         Width           =   650
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
         Height          =   690
         Left            =   120
         Picture         =   "Cogeve04.frx":0EAA
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Configuración de Procesos y Formularios"
         Top             =   4410
         Width           =   650
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Mast."
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
         Left            =   120
         Picture         =   "Cogeve04.frx":12EC
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Acceso a Maestro de Items de Stock"
         Top             =   2700
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   28
         Top             =   5880
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "Cogeve04.frx":1436
         Top             =   6300
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0F0FF&
      Caption         =   "Ventas y Deudores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6525
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   4215
      Begin VB.Frame Frame7 
         BackColor       =   &H00D0F0FF&
         Caption         =   "Resúmenes Mensuales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   315
         TabIndex        =   18
         Top             =   5145
         Width           =   3585
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   19
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
               Picture         =   "Cogeve04.frx":3358
               Style           =   1  'Graphical
               TabIndex        =   20
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
               TabIndex        =   21
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3270
         Left            =   315
         TabIndex        =   2
         Top             =   1680
         Width           =   3585
         Begin VB.PictureBox Picture11 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   930
            ScaleHeight     =   585
            ScaleWidth      =   2370
            TabIndex        =   22
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
               Picture         =   "Cogeve04.frx":3662
               Style           =   1  'Graphical
               TabIndex        =   23
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
               TabIndex        =   24
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
            TabIndex        =   15
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
               TabIndex        =   17
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
               TabIndex        =   16
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
            TabIndex        =   12
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
               Picture         =   "Cogeve04.frx":396C
               Style           =   1  'Graphical
               TabIndex        =   13
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
               TabIndex        =   14
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
            TabIndex        =   9
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
               Picture         =   "Cogeve04.frx":3C76
               Style           =   1  'Graphical
               TabIndex        =   10
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
               TabIndex        =   11
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
            TabIndex        =   6
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
               Picture         =   "Cogeve04.frx":3F80
               Style           =   1  'Graphical
               TabIndex        =   7
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
               TabIndex        =   8
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   315
         TabIndex        =   1
         Top             =   420
         Width           =   3585
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2895
            TabIndex        =   3
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
               Picture         =   "Cogeve04.frx":428A
               Style           =   1  'Graphical
               TabIndex        =   26
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
               Picture         =   "Cogeve04.frx":4594
               Style           =   1  'Graphical
               TabIndex        =   4
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
               TabIndex        =   5
               Top             =   105
               Width           =   1485
            End
         End
      End
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   3255
      Sorted          =   -1  'True
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "COGEVE04"
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
   Call CONECRUN("ARCHIG04/AMACLI", "Archivos Maestros")
   Command16.Enabled = True
End Sub

Sub Command18_Click()
   Command18.Enabled = False
   If DERACCE%("STPCOGEV", "Configuración de Resumenes Mensuales de Venta") >= 2 Then
      'Call CONECRUN("FSETUP04/OPCLAVEN", "Configuración de Funcionamiento")
      OPCLAV04.Show 1
      Call FINAL("")
   End If
   Command18.Enabled = True
End Sub

Private Sub Command19_Click()
    Command19.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("ARCHIG04/AMASTO", "Archivos Maestros")
    Command19.Enabled = True
End Sub

'
Sub Command2_Click()
    CARGAMEN% = 0
    REMEVE04.Show 1
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
    '
    Screen.MousePointer = 11
    Dim SALDEUPEN#(32768)
    '
    If ULTREG&("DEUDOR1") > 0 Then
        X$ = REGLEIDO$("DEUDOR1", 1)      ' fuerza la
        Call GRAREG("DEUDOR1", X$, 1)     ' reindexacion
        Call CIERRARCH("DEUDOR1")         ' del archivo
    End If
    '
    CARFAPEN% = 0
    Call FAPECOB04.FAPECOB1
    '
    FAPECOB04.Show 1
    '
End Sub

 Sub Command4_Click()
    Command4.Enabled = False
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM SADEUPEN "
    '18/04/07 (OT-07-0158)
    SQLX$ = SQLX$ + "WHERE (SaldDebe)>= 0.005 "
    SQLX$ = SQLX$ + "OR (SaldAcre)>= 0.005 "
    '18/04/07 (FIN)
    SQLX$ = SQLX$ + "ORDER BY FeVencim ASC "
    Set SADEUTMP = CueCorri.OpenRecordset(SQLX$, 4)
    With SADEUTMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        JJ& = 0
        YYY$ = REGBLAN("FLUCAPRO")
        Do While .EOF = False
          '
          Call REPLA(YYY$, MKI(CVINT(!FEVENCIM)), 1, 2)
          Call REPLA(YYY$, !CodiCom_Cor, 3, 12)
            NC1% = !NuClien
          Call REPLA(YYY$, CODICLI$(NC1%), 15, 12)
          Call REPLA(YYY$, RASOCLI$(NC1%), 27, 30)
          Call REPLA(YYY$, MKD(!SaldDebe), 57, 8)
          Call REPLA(YYY$, MKD(!SaldAcre), 65, 8)
          Call REPLA(YYY$, MKD(0), 53, 8)
          '
          JJ& = JJ& + 1
          Call GRAREG("FLUCAPRO", YYY$, JJ&)
        .MoveNext
        Loop
      End If
    End With
    '
20  SQLX$ = "SELECT * FROM SACREPEN "
    '18/04/07 (OT-07-0158)
    SQLX$ = SQLX$ + "WHERE (SaldDebe)>= 0.005 "
    SQLX$ = SQLX$ + "OR (SaldAcre)>= 0.005 "
    '18/04/07 (FIN)
    SQLX$ = SQLX$ + "ORDER BY FeVencim ASC "
    Set SADEUTMP = CueCorri.OpenRecordset(SQLX$, 4)
    With SADEUTMP
        On Error Resume Next
        .MoveFirst
        If Err Then
            On Error GoTo 0
            Call COMUNI("No se ha Podido Calcular el Flujo de Caja")
            Exit Sub
        End If
        On Error GoTo 0
        Do While .EOF = False
          Call REPLA(YYY$, MKI(CVINT(!FEVENCIM)), 1, 2)
          Call REPLA(YYY$, !CodiCom_Cor, 3, 12)
            NC1% = (-1) * !NuProve
          Call REPLA(YYY$, CODICLI$(NC1%), 15, 12)
          Call REPLA(YYY$, RASOCLI$(NC1%), 27, 30)
          Call REPLA(YYY$, MKD(!SaldDebe), 57, 8)
          Call REPLA(YYY$, MKD(!SaldAcre), 65, 8)
          Call REPLA(YYY$, MKD(0), 73, 8)
          '
          JJ& = JJ& + 1
          Call GRAREG("FLUCAPRO", YYY$, JJ&)
          .MoveNext
        Loop
    End With
    '
    Call SETULTREG("FLUCAPRO", JJ&)
    Call CIERRARCH("FLUCAPRO")
    Call FILESORT("FLUCAPRO", "", 1, 2, "ASC")
    SARRA# = 0: FANTE% = 0
    FIN& = JJ&
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("FLUCAPRO", U&)
      If CVI(Left$(XX$, 2)) = FANTE% Then
        If U& > 1 Then
          YY$ = REGLEIDO$("FLUCAPRO", U& - 1)
          Call REPLA(YY$, MKD$(0), 73, 8)
          Call GRAREG("FLUCAPRO", YY$, U& - 1)
          XX$ = REGLEIDO$("FLUCAPRO", U&)
        End If
      End If
      FANTE% = CVI(Left$(XX$, 2))
      SARRA# = SARRA# + CVD(Mid$(XX$, 57, 8)) - CVD(Mid$(XX$, 65, 8))
      Call REPLA(XX$, MKD$(SARRA#), 73, 8)
      Call GRAREG("FLUCAPRO", XX$, U&)
    Next U&
    Call CIERRARCH("FLUCAPRO")
    Call FINAL("*LIFLUCA")
    Command4.Enabled = True
End Sub

Sub Command5_Click()
    Command5.Enabled = False
    Call GELISALCLI
    Command5.Enabled = True
End Sub

Sub Command7_Click()
    Command7.Enabled = False
    RECUEC04.Show 1
    Command7.Enabled = True
End Sub

Sub Command8_Click()
    Command8.Enabled = False
    '
    If EXISTE%("LCOMIS04.EXE") > 0 Then
      OPXX% = ALTERNA%("Liquidación de Comisiones\Otros Accesos Directos")
      If OPXX% = 1 Then
          Call CONECRUN("LCOMIS04", "Liquidación de Comisiones")
          GoTo 99
        ElseIf OPXX% = 2 Then
          GoTo 20
        Else
          GoTo 99
      End If
    End If
    '
20  Call ACCDIR("COGEVEN")
    '
99  Command8.Enabled = True
End Sub

Sub Command9_Click()
    AVAFAC04.Show 1
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
    On Error Resume Next
    Hide
    On Error GoTo 0
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
    Dim DEUDA#(32767), SALDO#(32767)
    '
    Screen.MousePointer = 11
    Call HEADERS("LISACLI", "")
    Call HEADERS("LISACLI", "Generado el: " + HO$ + " - " + Left$(Time$, 5) + " hs.")
    '
    Screen.MousePointer = 11
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + "or TipoMovim = 'RCOB' "
    SQLX$ = SQLX$ + "or TipoMovim = 'AJUD' "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    CLIMAX% = 0
    '
    With CABATEMP
      On Error Resume Next
      .MoveLast
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos")
        GoTo 99
      End If
      '
      On Error GoTo 0
      FIN& = .RecordCount
      .MoveFirst
      UKL& = 0
      Do While Not .EOF
        UKL& = UKL& + 1
        Call TRACE(20, UKL&, FIN&)
        NC% = !NuClien
        If NC% > 0 Then
            IFAC# = !IDEBECOMP - !IHABECOMP
            SALDO#(NC%) = SALDO#(NC%) + IFAC#
        End If
      .MoveNext
      Loop
    End With
    '
    SQLX$ = "SELECT * FROM SADEUPEN "
    SQLX$ = SQLX$ + "WHERE Status < 3 "
    '
    Set SACTEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    On Error Resume Next
    With SACTEMP
      .MoveLast
      If Err < 1 Then
        FIN& = .RecordCount
        .MoveFirst
        UKL& = 0
        Do While Not .EOF
          UKL& = UKL& + 1
          Call TRACE(20, UKL&, FIN&)
          NC% = !NuClien
          SALPAR# = !SaldDebe - !SaldAcre
          If Abs(SALPAR#) >= 0.005 Then
            DEUDA#(NC%) = DEUDA#(NC%) + SALPAR#
          End If
        .MoveNext
        Loop
      End If
    End With
    '
    cx& = ULTREG&("DEUDOR1")
    RELISA& = 0
    '
    For IL& = 1 To cx&
      Call TRACE(20, IL&, cx&)
      X$ = REGLEIDO$("DEUDOR1", IL&)
      NC% = CVI(Left$(X$, 2))
      If NC% > 0 Then
        ' ajustar saldo de la cuenta
        If Abs(SALDO#(NC%)) >= 0.005 Or Abs(DEUDA#(NC%)) >= 0.005 Then
          Y$ = REGBLAN$("LISACLI")
          Call REPLA(Y$, MKI$(NC%), 1, 2)
          Call REPLA(Y$, MKD$(SALDO#(NC%)), 3, 8)
          Call REPLA(Y$, MKD$(DEUDA#(NC%)), 11, 8)
          If Abs(SALDO#(NC%) - DEUDA#(NC%)) > 0.005 Then
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
99 End Sub
'
   '///////*****RUTINA VIEJA******/////
'    '
'    HOXX = HOY(HO$)
'    Dim NC%(16384), DEUDA#(16384), SALDO#(16384)
'    '
'    Screen.MousePointer = 11
'    Call HEADERS("LISACLI", "")
'    Call HEADERS("LISACLI", "Generado el: " + HO$ + " - " + Left$(Time$, 5) + " hs.")
'    '
'    CX& = ULTREG&("DEUDOR1")
'    '
'    For IL& = 1 To CX&
'      Call TRACE(20, IL&, CX&)
'      X$ = REGLEIDO$("DEUDOR1", IL&)
'      NC%(IL&) = CVI(Left$(X$, 2))
'      SALDO#(IL&) = SALDCLI#(NC%(IL&))
'      DEUDA#(IL&) = 0
'    Next IL&
'    '
'    FIN& = ULTREG&("SVD202")
'    For IL& = 1 To FIN&
'      SS$ = REGLEIDO$("SVD202", IL&)
'      NCI% = CVI(Left$(SS$, 2))
'      TI% = Asc(Mid$(SS$, 3, 1))
'      If TI% > 7 Then TI% = 2
'      SALDEU# = CVD(Mid$(SS$, 11, 8))
'      If TI% > 3 Then SALDEU# = (-1) * Abs(SALDEU#)
'      For JL& = 1 To CX&
'        If NC%(JL&) = NCI% Then
'          DEUDA#(JL&) = DEUDA#(JL&) + SALDEU#
'          GoTo 809
'        End If
'      Next JL&
'809 Next IL&
'
'900 RELISA& = 0
'    '
'    For IL& = 1 To CX&
'      If NC%(IL&) > 0 Then
'        If Abs(SALDO#(IL&)) >= 0.005 Or Abs(DEUDA#(IL&)) >= 0.005 Then
'          Y$ = REGBLAN$("LISACLI")
'          Call REPLA(Y$, MKI$(NC%(IL&)), 1, 2)
'          Call REPLA(Y$, MKD$(SALDO#(IL&)), 3, 8)
'          Call REPLA(Y$, MKD$(DEUDA#(IL&)), 11, 8)
'          If Abs(SALDO#(IL&) - DEUDA#(IL&)) > 0.005 Then
'            Call REPLA(Y$, "ATENCION: Saldo <> Deuda", 19, 30)
'          End If
'          '
'          RELISA& = RELISA& + 1
'          Call GRAREG("LISACLI", Y$, RELISA&)
'        End If
'      End If
'    Next IL&
'    '
'    Call SETULTREG("LISACLI", RELISA&)
'    Call CIERRARCH("*.*")
'    Screen.MousePointer = 1
'    '
'    If Option1.Value = True Then
'         Call FINAL("*LISACLI")
'       Else
'         Call FINAL("*LIMPCLI")
'    End If
'    '

Sub LILIMCRE()
   '
   Screen.MousePointer = 11
   '
   NCMIN% = 32767: NCMAX% = 0
   IXXX$ = LOADFILE$(LUDAT$ + "DEUDOR1.X01")
   '
   For KKL& = 1 To Len(IXXX$) Step 2
     NCMX% = CVI(Mid$(IXXX$, KKL&, 2))
     If NCMX% > 0 Then
       If NCMX% < NCMIN% Then NCMIN% = NCMX%
       If NCMX% > NCMAX% Then NCMAX% = NCMX%
     End If
   Next KKL&
   '
   VDEF$ = MKI$(NCMIN%) + MKI$(NCMAX%) + Chr$(0)

   OBX$ = OBJPLA$("DESHASCLIEVEN", VDEF$)
   If OBX$ = "" Then GoTo 99
   NCL1% = CVI(Left$(OBX$, 2))
   NCL2% = CVI(Mid$(OBX$, 3, 2))
   VENDSEL% = Asc(Mid$(OBX$, 5, 1))
   '
   Screen.MousePointer = 1
   Call SELECHO("GRUCOCLI")
   GRUPCLIEN% = XVALO(TRIM$(VALACT1$("GRUCOCLI")))
   '
   Dim DEUCUECO#(16384), CHEPEAC#(16384)
   '
   Screen.MousePointer = 11
   Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM SADEUPEN "
   SQLX$ = SQLX$ + " WHERE Status < 3 "
   If NCL1% > 0 And NCL2% > 0 Then
     SQLX$ = SQLX$ + " AND NuClien >= " & NCL1%
     SQLX$ = SQLX$ + " AND NuClien <= " & NCL2%
   End If
   '
   Set SACTEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   On Error Resume Next
   With SACTEMP
     .MoveLast
     If Err < 1 Then
       FIN& = .RecordCount
       On Error GoTo 0
       .MoveFirst
       J1& = 0
       Do While Not .EOF
         J1& = J1& + 1
         Call TRACE(20, J1&, FIN&)
         NC% = !NuClien
         SALPAR# = !SaldDebe - !SaldAcre
         If Abs(SALPAR#) >= 0.005 Then
           DEUCUECO#(NC%) = DEUCUECO#(NC%) + SALPAR#
         End If
       .MoveNext
       Loop
     End If
   End With
   '
   HOII% = HOY(HOS$)
   FLACRE = CDbl(CVDAT(HOII%))
   '
   Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
   SQLX$ = SQLX$ + "AND FechAcred >= " & FLACRE & " "
   '
   Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   JJ& = 0: TRANSPOR# = 0
   Z0$ = REGBLAN$("LILICRE")
   With CABATEMP
     On Error Resume Next
     .MoveLast
     If Err < 1 Then
       FIN& = .RecordCount
       On Error GoTo 0
       .MoveFirst
       J1& = 0
       Do While Not .EOF
            J1& = J1& + 1
            Call TRACE(20, J1&, FIN&)
            NCI% = !NuClien
            If NCI% > 0 Then
              If NCI% <= 16384 Then
                CHEPEAC#(NCI%) = CHEPEAC#(NCI%) + !IDEBECOMP
              End If
            End If
       .MoveNext
       Loop
     End If
   End With
   '
   FIN& = ULTREG&("DEUDOR1")
   JJ& = 0
   For IL& = 1 To FIN&
     Call TRACE(20, IL&, FIN&)
     XX$ = REGLEIDO$("DEUDOR1", IL&)
     NCI% = CVI(Left$(XX$, 2))
     If NCI% > 0 Then
       If NCI% <= 16384 Then
         VVEND% = VENDCLI%(NCI%) 'devuelve el vendedor asignado si lo tiene
         GRCCL% = GRUCOCLI%(NCI%) 'devuelve el grupo cliente asignado si lo tiene

         If VENDSEL% <> 0 And VENDSEL% <> VVEND% Then GoTo 50
         If GRUPCLIEN% <> 0 And GRUPCLIEN% <> GRCCL% Then GoTo 50

         If Abs(DEUCUECO#(NCI%)) + Abs(CHEPEAC#(NCI%)) >= 0.005 Then
           YY$ = REGLEIDO$("DEUDOR2", IL&)
           CRELIM# = LIMICRE(NCI%)
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
50 Next IL&
   '
   Call SETULTREG("LILICRE", JJ&)
   Call CIERRARCH("LILICRE")
   '
   If Option1.Value = True Then
        Call CONECRUN("*COLICRE", "Consulta de Límites de Crédito")
      Else
        Call CONECRUN("*LILICRE", "Listado de Límites de Crédito")
   End If
99 Screen.MousePointer = 1
   '
End Sub

