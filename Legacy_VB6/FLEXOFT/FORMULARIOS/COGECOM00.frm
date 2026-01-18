VERSION 5.00
Begin VB.Form COGECOM00 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Gestión de Compras - Cuentas por Pagar"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
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
      Height          =   1275
      Left            =   210
      TabIndex        =   20
      Top             =   315
      Width           =   3585
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   21
         Top             =   400
         Width           =   2325
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
            Picture         =   "COGECOM00.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Consulta de Resumen de Cuenta Corriente por Proveedor"
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
            TabIndex        =   23
            Top             =   105
            Width           =   1485
         End
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0FFFF&
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
      Height          =   2430
      Left            =   210
      TabIndex        =   10
      Top             =   1740
      Width           =   3585
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   2160
         TabIndex        =   14
         Top             =   840
         Width           =   2220
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
            Picture         =   "COGECOM00.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Listado de Saldos y Deudas de Acreedores"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Saldos y Deudas de Acreedores"
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
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   17
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
            Picture         =   "COGECOM00.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Listados de Obligaciones Pendientes de Pago"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Cuentas por Pagar"
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
            TabIndex        =   19
            Top             =   105
            Width           =   1905
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   11
         Top             =   1630
         Width           =   3165
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
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
            TabIndex        =   13
            Top             =   250
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFFF&
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
            TabIndex        =   12
            Top             =   250
            Width           =   1275
         End
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFFF&
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
      Height          =   1380
      Left            =   210
      TabIndex        =   6
      Top             =   4305
      Width           =   3585
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   7
         Top             =   500
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
            Picture         =   "COGECOM00.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   8
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
            TabIndex        =   9
            Top             =   105
            Width           =   1485
         End
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   3990
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   4
      Top             =   4830
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   105
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   3045
      Visible         =   0   'False
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
      Height          =   855
      Left            =   3975
      Picture         =   "COGECOM00.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   420
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
      Height          =   855
      Left            =   3990
      Picture         =   "COGECOM00.frx":0D72
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   1365
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
      Height          =   855
      Left            =   3990
      Picture         =   "COGECOM00.frx":107C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Acceso Directo a Otros Procesos"
      Top             =   2310
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3465
      Picture         =   "COGECOM00.frx":1386
      Top             =   5160
      Width           =   2010
   End
End
Attribute VB_Name = "COGECOM00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

'
Private Sub Command2_Click()
    CARGAMEN% = 0
    REMECOMP.Show 1
End Sub

Private Sub Command29_Click()
   Call HELPONLINE("COGECOM")
End Sub

Sub Command3_Click()
    '
    Screen.MousePointer = 11
    Dim SALDEUPEN#(32768)
    '
    If ULTREG&("ACREDOR") > 0 Then
        X$ = REGLEIDO$("ACREDOR", 1)      ' fuerza la
        Call GRAREG("ACREDOR", X$, 1)     ' reindexacion
        Call CIERRARCH("ACREDOR")         ' del archivo
    End If
    '
    NCMIN% = 32767: NCMAX% = 0
    FIN& = ULTREG&("ACREDOR")
    For U& = 1 To FIN&
        NCX% = CVI(Left$(REGLEIDO$("ACREDOR", U&), 2))
        If NCX% > 0 Then
            If NCX% > NCMAX% Then NCMAX% = NCX%
            If NCX% < NCMIN% Then NCMIN% = NCX%
        End If
    Next U&
    '
    FAPEPAG00.Text10.TEXT = Str$(NCMIN%)
    FAPEPAG00.Text8.TEXT = Str$(NCMAX%)
    '
    Screen.MousePointer = 1
    '
    CARFAPEN% = 0
    FAPEPAG00.Show 1
    '
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    'Call OPNOIN("Planificacion Finaciera de Pagos")
    '
    Screen.MousePointer = 11
    Dim SALDEUPEN#(32768)
    '
    If ULTREG&("ACREDOR") > 0 Then
        X$ = REGLEIDO$("ACREDOR", 1)      ' fuerza la
        Call GRAREG("ACREDOR", X$, 1)     ' reindexacion
        Call CIERRARCH("ACREDOR")         ' del archivo
    End If
    '
    NCMIN% = 32767: NCMAX% = 0
    FIN& = ULTREG&("ACREDOR")
    For U& = 1 To FIN&
        NCX% = CVI(Left$(REGLEIDO$("ACREDOR", U&), 2))
        If NCX% > 0 Then
            If NCX% > NCMAX% Then NCMAX% = NCX%
            If NCX% < NCMIN% Then NCMIN% = NCX%
        End If
    Next U&
    '
    FAPEPAGO.Text10.TEXT = Str$(NCMIN%)
    FAPEPAGO.Text8.TEXT = Str$(NCMAX%)
    '
    Screen.MousePointer = 1
    '
    CARFAPEN% = 0
    FAPEPAGO.Show 1
    '
    Command4.Enabled = True
End Sub

Sub Command5_Click()
    Command5.Enabled = False
    Call GELISALPRO
    Command5.Enabled = True
End Sub

Sub Command7_Click()
    Command7.Enabled = False
    RESCUPRO.Show 1
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Call ACCDIR("COGECOM")
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call BAJALDISCO
   End If
End Sub

Sub GELISALPRO()
    '
    HOXX = HOY(HO$)
    Dim DEUDA#(32767), SALDO#(32767), TICUE%(32767)
    '
    Screen.MousePointer = 11
    Call HEADERS("LISAPRO", "")
    Call HEADERS("LISAPRO", "Generado el: " + HO$ + " - " + Left$(Time$, 5) + " hs.")
    '
    FIN& = ULTREG&("PROVED1")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       PP1$ = REGLEIDO$("PROVED1", U&)
       NC% = CVI(Left$(PP1$, 2))
       If NC% > 0 Then
          PP2$ = REGLEIDO$("PROVED2", U&)
          TICUE%(NC%) = Asc(Mid$(PP2$, 119, 1))
          If TICUE%(NC%) < 0 Or TICUE%(NC%) > 1 Then
             Call MENSERR(24, "Tipo de Cuenta no Válido\para Cuenta " + TRIM$(Str$(NC%)) + " " + TRIM$(Mid$(PP2$, 3, 30)) + ".")
             TICUE%(NC%) = 0
          End If
       End If
    Next U&
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("CUECORR")
    '
    For IX& = 1 To FIN&
       Call TRACE(20, IX&, FIN&)
       SS$ = REGLEIDO$("CUECORR", IX&)
       SS4$ = Mid$(SS$, 5, 2)
       SD1$ = Mid$(SS$, 115, 8)
       SH1$ = Mid$(SS$, 123, 8)
       SALDEU$ = Mid$(SS$, 139, 60)
       '
       NC% = CVI(Mid$(SS$, 5, 2))
       If NC% > 0 Then
          If TICUE%(NC%) = 0 Then   ' saltea cuentas de gastos
             TOT# = CVD(SH1$) - CVD(SD1$)
             SALDO#(NC%) = SALDO#(NC%) + TOT#
             '
             For KKI% = 1 To 6
                FVI% = CVI(Mid$(SALDEU$, 10 * KKI% - 9, 2))
                SALPAR# = CVD(Mid$(SALDEU$, 10 * KKI% - 7, 8))
                If FVI% > 0 Then
                  If Abs(SALPAR#) >= 0.005 Then
                    DEUDA#(NC%) = DEUDA#(NC%) + SALPAR#
                  End If
                End If
             Next KKI%
          End If
       End If
    Next IX&
    '
    CX& = ULTREG&("PROVED1")
    RELISA& = 0
    '
    For IL& = 1 To CX&
       Call TRACE(20, IL&, CX&)
       X$ = REGLEIDO$("PROVED1", IL&)
       NC% = CVI(Left$(X$, 2))
       If NC% > 0 Then
          ' ajustar saldo de la cuenta
          If Abs(SALDO#(NC%)) >= 0.005 Or Abs(DEUDA#(NC%)) >= 0.005 Then
             Y$ = REGBLAN$("LISAPRO")
             Call REPLA(Y$, MKI$(NC%), 1, 2)
             Call REPLA(Y$, MKD$(SALDO#(NC%)), 3, 8)
             Call REPLA(Y$, MKD$(DEUDA#(NC%)), 11, 8)
             If Abs(SALDO#(NC%) - DEUDA#(NC%)) > 0.005 Then
                Call REPLA(Y$, "ATENCION: Saldo <> Deuda", 19, 30)
             End If
             '
             RELISA& = RELISA& + 1
             Call GRAREG("LISAPRO", Y$, RELISA&)
          End If
       End If
    Next IL&
    '
    Call SETULTREG("LISAPRO", RELISA&)
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    If Option1.Value = True Then
          Call FINAL("*LISAPRO")
        Else
          Call FINAL("*LIMPPRO")
    End If
    '
End Sub
'

