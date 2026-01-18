VERSION 5.00
Begin VB.Form COGECOM04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Gestión de Compras - Cuentas por Pagar"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00004040&
      Height          =   5475
      Left            =   3780
      ScaleHeight     =   5415
      ScaleWidth      =   915
      TabIndex        =   19
      Top             =   0
      Width           =   975
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   140
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   4515
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -420
            TabIndex        =   24
            Top             =   0
            Width           =   12000
         End
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
         Height          =   690
         Left            =   140
         Picture         =   "COGECOM04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Acceso Directo a Otros Procesos"
         Top             =   2310
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
         Left            =   140
         Picture         =   "COGECOM04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   890
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
         Left            =   140
         Picture         =   "COGECOM04.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   200
         Width           =   650
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -290
         Picture         =   "COGECOM04.frx":075E
         Top             =   4920
         Width           =   1515
      End
   End
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
      Left            =   105
      TabIndex        =   15
      Top             =   105
      Width           =   3585
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   16
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
            Picture         =   "COGECOM04.frx":2680
            Style           =   1  'Graphical
            TabIndex        =   17
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
            TabIndex        =   18
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
      Left            =   105
      TabIndex        =   5
      Top             =   1470
      Width           =   3585
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   2160
         TabIndex        =   9
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
            Picture         =   "COGECOM04.frx":298A
            Style           =   1  'Graphical
            TabIndex        =   10
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
            TabIndex        =   11
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
         TabIndex        =   12
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
            Picture         =   "COGECOM04.frx":2C94
            Style           =   1  'Graphical
            TabIndex        =   13
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
            TabIndex        =   14
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
         TabIndex        =   6
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
            TabIndex        =   8
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
            TabIndex        =   7
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
      Left            =   105
      TabIndex        =   1
      Top             =   3990
      Width           =   3585
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   2
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
            Picture         =   "COGECOM04.frx":2F9E
            Style           =   1  'Graphical
            TabIndex        =   3
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
            TabIndex        =   4
            Top             =   105
            Width           =   1485
         End
      End
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   105
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   2730
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "COGECOM04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub



'
 Sub Command2_Click()
    CARGAMEN% = 0
    REMECO04.Show 1
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
    FAPEPAG04.Text10.TEXT = Str$(NCMIN%)
    FAPEPAG04.Text8.TEXT = Str$(NCMAX%)
    '
    Screen.MousePointer = 1
    '
    CARFAPEN% = 0
    FAPEPAG04.Show 1
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
    RECUEP04.Show 1
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
        TICUE%(NC%) = TICUEPRO%((-1) * NC%)     ' Asc(Mid$(PP2$, 119, 1))
        If TICUE%(NC%) < 0 Or TICUE%(NC%) > 1 Then
          Call MENSERR(24, "Tipo de Cuenta no Válido\para Cuenta " + TRIM$(Str$(NC%)) + " " + TRIM$(Mid$(PP2$, 3, 30)) + ".")
          TICUE%(NC%) = 0
        End If
      End If
    Next U&
    '
    Screen.MousePointer = 11
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "or TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + "or TipoMovim = 'AJUP' "
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
      FIN& = .RecordCount
      JJ& = 0
      '
      On Error GoTo 0
      .MoveFirst
      Do While Not .EOF
        JJ& = JJ& + 1
        Call TRACE(20, JJ&, FIN&)
        NC% = !Nuprove
        If NC% > 0 Then
          If TICUE%(NC%) = 0 Then ' solo cuentas corrientes
            IFAC# = !IHABECOMP - !Idebecomp
            SALDO#(NC%) = SALDO#(NC%) + IFAC#
          End If
        End If
      .MoveNext
      Loop
    End With
    '
    SQLX$ = "SELECT * FROM SACREPEN "
    SQLX$ = SQLX$ + "WHERE Status < 3 "
    '
    Set SACTEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    On Error Resume Next
    With SACTEMP
      .MoveLast
      If Err < 1 Then
        FIN& = .RecordCount
        JJ& = 0
        On Error GoTo 0
        .MoveFirst
        Do While Not .EOF
          JJ& = JJ& + 1
          Call TRACE(20, JJ&, FIN&)
          NC% = !Nuprove
          SALPAR# = !SaldAcre - !SaldDebe
          If Abs(SALPAR#) >= 0.005 Then
            DEUDA#(NC%) = DEUDA#(NC%) + SALPAR#
          End If
        .MoveNext
        Loop
      End If
    End With
    '
    cx& = ULTREG&("PROVED1")
    RELISA& = 0
    '
    For IL& = 1 To cx&
      Call TRACE(20, IL&, cx&)
      X$ = REGLEIDO$("PROVED1", IL&)
      NC% = CVI(Left$(X$, 2))
      If NC% > 0 Then
        ' ajustar saldo de la cuenta
        If TICUE%(NC%) = 0 Then
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
99 End Sub
'

