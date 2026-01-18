VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form COGECOM07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C7D9CC&
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
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   24
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Control Arrastre Saldos"
      Height          =   225
      Left            =   1470
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H0050765A&
      ForeColor       =   &H80000008&
      Height          =   5475
      Left            =   3840
      ScaleHeight     =   5445
      ScaleWidth      =   945
      TabIndex        =   19
      Top             =   0
      Width           =   975
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H000C93DC&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   26
         Top             =   5040
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0076C9F5&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   27
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Direct"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "COGECOM07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Acceso Directo a Otros Procesos"
         Top             =   2310
         Width           =   650
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "COGECOM07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   890
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "COGECOM07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   200
         Width           =   650
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C7D9CC&
      Caption         =   "DETALLE POR CUENTA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
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
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   630
         ScaleHeight     =   615
         ScaleWidth      =   2295
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
            Height          =   630
            Left            =   0
            Picture         =   "COGECOM07.frx":075E
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
            TabIndex        =   18
            Top             =   105
            Width           =   1485
         End
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C7D9CC&
      Caption         =   "CONTROL DE ESTADO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
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
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   945
         ScaleHeight     =   615
         ScaleWidth      =   2190
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
            Height          =   630
            Left            =   0
            Picture         =   "COGECOM07.frx":0A68
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Listado de Saldos y Deudas de Acreedores"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Saldos y Deudas    de Acreedores"
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
            TabIndex        =   11
            Top             =   105
            Width           =   1500
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   525
         ScaleHeight     =   615
         ScaleWidth      =   2400
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
            Height          =   630
            Left            =   0
            Picture         =   "COGECOM07.frx":0D72
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
            TabIndex        =   14
            Top             =   105
            Width           =   1905
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C7D9CC&
         Caption         =   "Salida"
         BeginProperty Font 
            Name            =   "Arial"
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
            BackColor       =   &H00C7D9CC&
            Caption         =   "Pantalla"
            BeginProperty Font 
               Name            =   "Arial"
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
            BackColor       =   &H00C7D9CC&
            Caption         =   "Impresión"
            BeginProperty Font 
               Name            =   "Arial"
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
      BackColor       =   &H00C7D9CC&
      Caption         =   "RESUMENES MENSUALES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
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
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   630
         ScaleHeight     =   615
         ScaleWidth      =   2295
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
            Height          =   630
            Left            =   0
            Picture         =   "COGECOM07.frx":107C
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
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "COGECOM07.frx":1386
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "COGECOM07.frx":1400
      Top             =   1200
   End
End
Attribute VB_Name = "COGECOM07"
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
    REMECO07.Show 1
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
    FAPEPAG07.Text10.TEXT = Str$(NCMIN%)
    FAPEPAG07.Text8.TEXT = Str$(NCMAX%)
    '
    Screen.MousePointer = 1
    '
    CARFAPEN% = 0
    FAPEPAG07.Show 1
    '
End Sub

Private Sub SaldoYDeudaDeProveedores()

    HOXX = HOY(HO$)
    Dim DEUDA#(32767), SALDO1#(32767)
    
    Dim oReporte As cGenrep
    Dim Rst As ADODB.Recordset
    
    Set oReporte = New cGenrep
    Set Rst = oReporte.GeneraReporteConTablaTemporal("SALPRO", "SALDO DE PROVEEDORES", "Nro(T=Cuenta)/I10;Proveedor(T=Razon Social)/A48;Saldo(TT=SI)/F14.2;Deuda(TT=SI)/F14.2;Observacion/A24")

    Call APERBASDAT("CABAN")
    Dim CABATEMP As ADODB.Recordset
    SQLX$ = "SELECT NuProve, sum(IHabeComp) as habe, sum(IDebeComp) as debe, STAT_CLI FROM CAJABANC INNER JOIN PROVED12 WITH(NOLOCK) ON NUPROVE = NUME_CLI "
    SQLX$ = SQLX$ + "WHERE nuprove > 0 and (TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "or TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + "or TipoMovim = 'AJUP')"
    SQLX$ = SQLX$ + "GROUP BY NUPROVE, STAT_CLI"
    '
    Set CABATEMP = READSET(SQLX$)
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
      NIVCONXXX% = NIVCONFI%
      '
      On Error GoTo 0
      .MoveFirst
      Do While Not .EOF
        JJ& = JJ& + 1
        Call TRACE(20, JJ&, FIN&)
        NC% = !NUPROVE
        If NC% > 0 Then
          STAT% = !Stat_Cli
          If STAT% > 0 Then
            IFAC# = !HABE - !DEBE
            SALDO1#(NC%) = SALDO1#(NC%) + IFAC#
          End If
        End If
70    .MoveNext
      Loop
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    '
    Dim SACTEMP As ADODB.Recordset
    SQLX$ = "SELECT NuProve, SaldAcre, SaldDebe FROM SACREPEN WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE Status < 3 "
    Set SACTEMP = READSET(SQLX$)
    '
    With SACTEMP
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        FIN& = .RecordCount
        JJ& = 0
        On Error GoTo 0
        .MoveFirst
        Do While Not .EOF
          JJ& = JJ& + 1
          Call TRACE(20, JJ&, FIN&)

          NC% = !NUPROVE
          SALPAR# = !SaldAcre - !SaldDebe
          If Abs(SALPAR#) >= 0.005 Then
            DEUDA#(NC%) = DEUDA#(NC%) + SALPAR#
          End If
80      .MoveNext
        Loop
      End If
    End With
    SACTEMP.Close
    Set SACTEMP = Nothing
    '
    cx& = ULTREG&("PROVED1")
    RELISA& = 0
    '
    With Rst
        For IL& = 1 To cx&
          Call TRACE(20, IL&, cx&)
          X$ = REGLEIDO$("PROVED1", IL&)
          NC% = CVI(Left$(X$, 2))
          If NC% > 0 Then
            If Abs(SALDO1#(NC%)) >= 0.005 Or Abs(DEUDA#(NC%)) >= 0.005 Then
               If XVALO(VALOBADA("proved12", "Tcpr_Cli", "nume_cli=" & NC%)) = 0 Then
               
                    .AddNew
                    .Fields(0) = NC%
                    .Fields(1) = Left$(rasocli$(NC% * -1), 48)
                    .Fields(2) = DEUDA#(NC%) ' SALDO1#(NC%)
                    .Fields(3) = DEUDA#(NC%)
                    If Abs(SALDO1#(NC%) - DEUDA#(NC%)) > 0.005 Then
                       '.Fields(4) = "ATENCION: Saldo <> Deuda"
                    End If
                    .Update
                    
               End If
            End If
          End If
        Next IL&
    End With
    
    oReporte.MostrarReporte
    
99  Exit Sub

End Sub




Sub Command5_Click()
    Command5.Enabled = False
    Call GELISALPRO
    Command5.Enabled = True
End Sub


Sub Command7_Click()
    Command7.Enabled = False
    RECUEP07.Show 1
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
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture4)
    '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call BAJALDISCO
   End If
End Sub

Sub GELISALPRO()
    '
    Call SaldoYDeudaDeProveedores
    Exit Sub
    '
    '--------------------------------
'    HOXX = HOY(HO$)
'    Dim DEUDA#(32767), SALDO#(32767), TICUE%(32767)
'    '
'    Screen.MousePointer = 11
'    Call HEADERS("LISAPRO", "")
'    Call HEADERS("LISAPRO", "Generado el: " + HO$ + " - " + Left$(Time$, 5) + " hs.")
'    '
'''    FIN& = ULTREG&("PROVED1")
'''    For U& = 1 To FIN&
'''      Call TRACE(20, U&, FIN&)
'''      PP1$ = REGLEIDO$("PROVED1", U&)
'''      NC% = CVI(Left$(PP1$, 2))
'''      If NC% > 0 Then
'''        PP2$ = REGLEIDO$("PROVED2", U&)
'''        TICUE%(NC%) = TICUEPRO%((-1) * NC%)     ' Asc(Mid$(PP2$, 119, 1))
'''        If TICUE%(NC%) < 0 Or TICUE%(NC%) > 1 Then
'''          Call MENSERR(24, "Tipo de Cuenta no Válido\para Cuenta " + TRIM$(Str$(NC%)) + " " + TRIM$(Mid$(PP2$, 3, 30)) + ".")
'''          TICUE%(NC%) = 0
'''        End If
'''      End If
'''    Next U&
'    '
'    Screen.MousePointer = 11
'    'JANDY SQL VER DE HACER JOIN Y SELECT ...SUM()
'    Call APERBASDAT("CABAN")
'    Dim CABATEMP As ADODB.Recordset
'    SQLX$ = "SELECT NuProve, sum(IHabeComp) as habe, sum(IDebeComp) as debe FROM CAJABANC WITH(NOLOCK) "
'    SQLX$ = SQLX$ + "WHERE nuprove >0 and (TipoMovim = 'FCOM' "
'    SQLX$ = SQLX$ + "or TipoMovim = 'OPAG' "
'    SQLX$ = SQLX$ + "or TipoMovim = 'AJUP')"
'    SQLX$ = SQLX$ + "GROUP BY NUPROVE"
'    '
'    Set CABATEMP = READSET(SQLX$)
'    CLIMAX% = 0
'    '
'    With CABATEMP
'      On Error Resume Next
'      .MoveLast
'      If Err > 0 Then
'        On Error GoTo 0
'        Call MENSERR(24, "Sin Movimientos")
'        GoTo 99
'      End If
'      FIN& = .RecordCount
'      JJ& = 0
'      NIVCONXXX% = NIVCONFI%
'      '
'      On Error GoTo 0
'      .MoveFirst
'      Do While Not .EOF
'        JJ& = JJ& + 1
'        Call TRACE(20, JJ&, FIN&)
'        NC% = !NUPROVE
'        If NC% > 0 Then
'          STAT% = XVALO(VALOBADA("PROVED12", "STAT_CLI", "NUME_CLI=" & NC%))
'          If STAT% > 0 Then
'            IFAC# = !HABE - !DEBE
'            SALDO#(NC%) = SALDO#(NC%) + IFAC#
'          End If
'        End If
'70    .MoveNext
'      Loop
'    End With
'    CABATEMP.Close
'    Set CABATEMP = Nothing
'    '
'    Dim SACTEMP As ADODB.Recordset
'    SQLX$ = "SELECT NuProve, SaldAcre, SaldDebe FROM SACREPEN WITH(NOLOCK) "
'    SQLX$ = SQLX$ + "WHERE Status < 3 "
'    Set SACTEMP = READSET(SQLX$)
'    '
'    With SACTEMP
'      On Error Resume Next
'      .MoveLast
'      If Err < 1 Then
'        FIN& = .RecordCount
'        JJ& = 0
'        On Error GoTo 0
'        .MoveFirst
'        Do While Not .EOF
'          JJ& = JJ& + 1
'          Call TRACE(20, JJ&, FIN&)
'
'          NC% = !NUPROVE
'          SALPAR# = !SaldAcre - !SaldDebe
'          If Abs(SALPAR#) >= 0.005 Then
'            DEUDA#(NC%) = DEUDA#(NC%) + SALPAR#
'          End If
'80      .MoveNext
'        Loop
'      End If
'    End With
'    SACTEMP.Close
'    Set SACTEMP = Nothing
'    '
'    cx& = ULTREG&("PROVED1")
'    RELISA& = 0
'    '
'    For IL& = 1 To cx&
'      Call TRACE(20, IL&, cx&)
'      X$ = REGLEIDO$("PROVED1", IL&)
'      NC% = CVI(Left$(X$, 2))
'      If NC% > 0 Then
'        If Abs(SALDO#(NC%)) >= 0.005 Or Abs(DEUDA#(NC%)) >= 0.005 Then
'           'If TICUEPRO%((-1) * NC%) = 0 Then
'           If XVALO(VALOBADA("proved12", "Tcpr_Cli", "nume_cli=" & NC%)) = 0 Then
'             Y$ = REGBLAN$("LISAPRO")
'             Call REPLA(Y$, MKI$(NC%), 1, 2)
'             Call REPLA(Y$, MKD$(SALDO#(NC%)), 3, 8)
'             Call REPLA(Y$, MKD$(DEUDA#(NC%)), 11, 8)
'             If Abs(SALDO#(NC%) - DEUDA#(NC%)) > 0.005 Then
'                Call REPLA(Y$, "ATENCION: Saldo <> Deuda", 19, 30)
'             End If
'             '
'             RELISA& = RELISA& + 1
'             Call GRAREG("LISAPRO", Y$, RELISA&)
'           End If
'        End If
'      End If
'    Next IL&
'    '
'    Call SETULTREG("LISAPRO", RELISA&)
'    Call CIERRARCH("*.*")
'    Screen.MousePointer = 1
'    '
'    If Option1.Value = True Then
'         Call FINAL("*LISAPRO")
'       Else
'         Call FINAL("*LIMPPRO")
'    End If
'    '
99 End Sub


'
Private Sub Command6_Click()
    'RUTINA PARA CONTROL DE ARRASTRE DE SALDOS REALIZADO PARA DOSIVAC
    Exit Sub
    
    CLAX$ = "": CLACO% = 0
30  PPX% = Int(Rnd * 63)
    If PPX% < 48 Or PPX% > 57 Then GoTo 30
    CLAX$ = CLAX$ + Chr$(PPX%)
    If Len(CLAX$) < 6 Then GoTo 30
    '
    RESPUES$ = CONTRACLA$(CLAX$)
    CONTRASE$ = InputBox$("Ingrese Contraseña", "Santo y Seña (" + CLAX$ + ")")
    If CONTRASE$ <> RESPUES$ Then Exit Sub
    '
    limit$ = InputBox$("Fecha Limite")
    If limit$ = "" Then Exit Sub
    FLIMI% = FECHANUM(limit$)
    If FLIMI% < 1 Then Exit Sub
    '
    FLIM$ = FETEXCOR1$(FLIMI%)
    Dim DEUDA#(32767), SALDO#(32767)
    '
    Screen.MousePointer = 11
    Call HEADERS("LISACLI", "")
    Call HEADERS("LISACLI", "Generado el: " + HO$ + " - " + Left$(Time$, 5) + " hs.")
    '
    Screen.MousePointer = 11
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "or TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + "or TipoMovim = 'AJUP') "
    SQLX$ = SQLX$ + " and FechEmisi <= '" & FLIM$ & "' "
    '
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = New ADODB.Recordset
    CABATEMP.CursorLocation = adUseClient

27  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 27
    End If
    On Error GoTo 0

     
    CLIMAX% = 0
    '
    With CABATEMP
      On Error Resume Next
      .MoveLast
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos")
        Exit Sub
      End If
      '
      On Error GoTo 0
      FIN& = .RecordCount
      .MoveFirst
      UKL& = 0
      Do While Not .EOF
        UKL& = UKL& + 1
        Call TRACE(20, UKL&, FIN&)
        NC% = !NUPROVE
        If NC% > 0 Then
            IFAC# = !IHABECOMP - !Idebecomp
            SALDO#(NC%) = SALDO#(NC%) + IFAC#
        End If
      .MoveNext
      Loop
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    '
    FIN& = ULTREG&("CUECORR")
    For IX& = 1 To FIN&
       Call TRACE(20, IX&, FIN&)
       SS$ = REGLEIDO$("CUECORR", IX&)
       FF% = CVI(Left$(SS$, 2))
       If FF% <= FLIMI% Then
          SS4$ = Mid$(SS$, 5, 2)
          SD1$ = Mid$(SS$, 115, 8)
          SH1$ = Mid$(SS$, 123, 8)
          SALDEU$ = Mid$(SS$, 139, 60)
          '
          NC% = CVI(Mid$(SS$, 5, 2))
          If NC% > 0 Then
             If TICUEPRO%((-1) * NC%) = 0 Then ' saltea cuentas de gastos
                TOT# = CVD(SH1$) - CVD(SD1$)
                DEUDA#(NC%) = DEUDA#(NC%) + TOT#
             End If
          End If
       End If
    Next IX&
    '
    cx& = 32766
    For IL& = 1 To cx&
      Call TRACE(20, IL&, cx&)
      'X$ = REGLEIDO$("DEUDOR1", IL&)
      'NC% = CVI(Left$(X$, 2))
      'If NC% > 0 Then
        ' ajustar saldo de la cuenta
      If Abs(SALDO#(IL&) - DEUDA#(IL&)) >= 0.005 Then
         NC% = IL&
         If TICUEPRO%((-1) * NC%) = 0 Then ' saltea cuentas de gastos
            y$ = REGBLAN$("LISAPRO")
             Call REPLA(y$, MKI$(NC%), 1, 2)
             Call REPLA(y$, MKD$(SALDO#(NC%)), 3, 8)
             Call REPLA(y$, MKD$(DEUDA#(NC%)), 11, 8)
             If Abs(SALDO#(NC%) - DEUDA#(NC%)) > 0.005 Then
                Call REPLA(y$, "ATENCION: Saldo <> Deuda", 19, 30)
             End If
             '
             RELISA& = RELISA& + 1
             Call GRAREG("LISAPRO", y$, RELISA&)
           End If
        End If
    Next IL&
    '
    Call SETULTREG("LISAPRO", RELISA&)
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    Call FINAL("*LISAPRO")
    '
    If COMALTER%("Desea Ajustar los Saldos al " + limit$ + " ? \\Si, Ajustarlos\No, Continuar") <> 1 Then Exit Sub
    '
100 For IL& = 1 To ULTREG&("LISApro")
        y$ = REGLEIDO$("LISAPRO", IL&)
        NC% = CVI(Left$(y$, 2))
        '
        If COMALTER%("Recupera Cuenta " + Str$(NC%) + " - " + rasocli$((-1) * NC%) + " ?\\Si\No") <> 1 Then GoTo 98
        '
        SQLX$ = "SELECT * FROM CAJABANC"
        SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'FCOM'"
        SQLX$ = SQLX$ + " AND NuProve = " & NC% & " "
        SQLX$ = SQLX$ + " AND FECHEMISI <= '" & FLIM$ & " '"
        '
        Dim CONTROLTMP As ADODB.Recordset
        Set CONTROLTMP = New ADODB.Recordset
        
26      On Error Resume Next
        CONTROLTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
        If Err Then
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 26
        End If
        On Error GoTo 0
        '
        With CONTROLTMP
           Do While Not .EOF
             JJ& = JJ& + 1
             Call TRACE(20, JJ&, CantiReg&)
             NC% = !NUPROVE
             FecEmision% = CVINT%(!FECHEMISI)
             ImporteTotal# = XVALO(!IHABECOMP) - XVALO(!Idebecomp)
             NUFACX$ = TRIM$(!CODICOM_LAR)
             '
             RFF$ = RECFILT$("CUECORR", 4, MKI(NC%))
             '
             For U& = 1 To Len(RFF$) Step 4
                REG& = CVS(Mid$(RFF$, U&, 4))
                X$ = REGLEIDO$("CUECORR", REG&)
                If TRIM$(Mid$(X$, 7, 18)) = NUFACX$ And CVI(Mid$(X$, 1, 2)) = CVINT(!FECHEMISI) Then
                  IMPO# = CVD(Mid$(X$, 123, 8)) - CVD(Mid$(X$, 115, 8))
                  Diferencia# = IMPO# - ImporteTotal#
                  If Abs(Diferencia#) >= 0.005 Then
                     '
                     MsgBox "PROVEEDOR: " & NC% & " - COMPROBANTE: " & !CODICOM_LAR
                     '
                     CODILAR$ = !CODICOM_LAR
                     NCLI% = !NUPROVE
                     FEMIX$ = FETEXCOR1$(CVINT(!FECHEMISI))
                     If IMPO# >= 0.005 Then
                          !IHABECOMP = IMPO#
                          !Idebecomp = 0
                        ElseIf IMPO# < (-0.005) Then
                          !Idebecomp = Abs(IMPO#)
                          !IHABECOMP = 0
                        Else   ' IMPO#=0
                          !IHABECOMP = 0
                          !Idebecomp = 0
                     End If
                     .Update
                     CONTROLTMP.Close
                     Set CONTROLTMP = Nothing
                     '
                     SQLX$ = "SELECT * FROM CAJABANC"
                     SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'COMPRA'"
                     SQLX$ = SQLX$ + " AND CodiCom_Lar = '" & CODILAR$ & "'"
                     SQLX$ = SQLX$ + " AND FECHEMISI = '" & FEMIX$ & "'"
                     SQLX$ = SQLX$ + " AND NuProve = " & NCLI% & " "
                     '
                     Dim CONTROL1 As ADODB.Recordset
                     Set CONTROL1 = New ADODB.Recordset
                     '
25                   On Error Resume Next
                     CONTROL1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
                     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                        On Error GoTo 0
                        Call CapturaErrorOpen
                        GoTo 25
                     End If
                     On Error GoTo 0
                     '
                     If IMPO# >= 0.005 Then
                         CONTROL1!Idebecomp = CONTROL1!Idebecomp + Diferencia#
                       Else
                         CONTROL1!IHABECOMP = CONTROL1!IHABECOMP - Diferencia#
                     End If
                     CONTROL1!Idebecomp = CONTROL1!Idebecomp - CONTROL1!IHABECOMP
                     CONTROL1!IHABECOMP = 0
                     If CONTROL1!Idebecomp < 0 Then
                        CONTROL1!IHABECOMP = Abs(CONTROL1!Idebecomp)
                        CONTROL1!Idebecomp = 0
                     End If
                     If CONTROL1!IHABECOMP < 0 Then
                        CONTROL1!Idebecomp = Abs(CONTROL1!IHABECOMP)
                        CONTROL1!IHABECOMP = 0
                     End If
                     '
                     CONTROL1.Update
                     CONTROL1.Close
                     Set CONTROL1 = Nothing
                     GoTo 98
                     '
                  End If
                  GoTo 19
                End If
             Next U&
19         .MoveNext
           Loop

       End With
       CONTROLTMP.Close
       Set CONTROLTMP = Nothing

98  Next IL&
    '
99  Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    MsgBox "FIN DEL CONTROL"
    '
End Sub
    

Sub GELISALPRO2()
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
    NIVCONXXX% = NIVCONFI%
    
    Screen.MousePointer = 11
    'JANDY SQL VER DE HACER JOIN Y SELECT ...SUM()
    CONDI$ = " TipoMovim = 'FCOM' "
    CONDI$ = CONDI$ + " or TipoMovim = 'OPAG' "
    CONDI$ = CONDI$ + " or TipoMovim = 'AJUP' "
    '
    SQLX$ = "SELECT SUM(IHABECOMP) AS IHABE, SUM(IDEBECOMP)AS IDEBE ,NUPROVE, Niv_Clasi FROM CAJABANC WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    SQLX$ = SQLX$ + " GROUP BY NUPROVE,NIV_CLASI"
    '
    FIN& = CantRegistros&("CAJABANC", CONDI$, "NOMESS")
    If FIN& < 1 Then
        Call MENSERR(24, "Sin Movimientos")
        GoTo 99
    End If

    Set CABATEMP = READSET(SQLX$)
    With CABATEMP
      Do While Not .EOF
        JJ& = JJ& + 1
        Call TRACE(20, JJ&, FIN&)
        NC% = XVALO(!NUPROVE)
        If NC% > 0 Then
          If XVALO(!Niv_Clasi) > NIVCONXXX% Then GoTo 70 'VER QUE FACTURAS MOSTRAR EN LISTADO
'          If TICUEPRO%((-1) * NC%) = 0 Then
           If TICUE%(NC%) = 0 Then ' solo cuentas corrientes
            IFAC# = XVALO(!IHABE) - XVALO(!IDEBE)
            SALDO#(NC%) = SALDO#(NC%) + IFAC#
          End If
        End If
70    .MoveNext
      Loop
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    '
    CONDI$ = " Status < 3 "
    '
    FIN& = CantRegistros&("CAJABANC", CONDI$, "NOMESS")
    If FIN& < 1 Then
        Call MENSERR(24, "Sin Movimientos")
        GoTo 99
    End If

    SQLX$ = "SELECT SUM(SALDAcre) AS SACRE ,SUM(SaldDebe) AS SDEBE, NUPROVE, Niv_Clasi FROM SACREPEN WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE  " & CONDI$
    SQLX$ = SQLX$ + " GROUP BY NuProve, Niv_Clasi "
    '
    Set SACTEMP = READSET(SQLX$)
    With SACTEMP

        Do While Not .EOF
          JJ& = JJ& + 1
          Call TRACE(20, JJ&, FIN&)
          '\\\OT-08-1032-OD-11/11/08 VERIFICA EL NIVEL DEL USUARIO PARA
          If XVALO(!Niv_Clasi) > NIVCONXXX% Then GoTo 80 'VER QUE FACTURAS MOSTRAR EN LISTADO
          '\\\OT-08-1032-OD-FIN

          NC% = XVALO(!NUPROVE)
          SALPAR# = XVALO(!SACRE) - XVALO(!SDebe)
          If Abs(SALPAR#) >= 0.005 Then
            DEUDA#(NC%) = DEUDA#(NC%) + SALPAR#
          End If
80      .MoveNext
        Loop
    End With
    SACTEMP.Close
    Set SACTEMP = Nothing
    '
    cx& = ULTREG&("PROVED1")
    RELISA& = 0
    '
    For IL& = 1 To cx&
      Call TRACE(20, IL&, cx&)
      X$ = REGLEIDO$("PROVED1", IL&)
      NC% = CVI(Left$(X$, 2))
      If NC% > 0 Then
        If Abs(SALDO#(NC%)) >= 0.005 Or Abs(DEUDA#(NC%)) >= 0.005 Then
           If TICUEPRO%((-1) * NC%) Then
             y$ = REGBLAN$("LISAPRO")
             Call REPLA(y$, MKI$(NC%), 1, 2)
             Call REPLA(y$, MKD$(SALDO#(NC%)), 3, 8)
             Call REPLA(y$, MKD$(DEUDA#(NC%)), 11, 8)
             If Abs(SALDO#(NC%) - DEUDA#(NC%)) > 0.005 Then
                Call REPLA(y$, "ATENCION: Saldo <> Deuda", 19, 30)
             End If
             '
             RELISA& = RELISA& + 1
             Call GRAREG("LISAPRO", y$, RELISA&)
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




