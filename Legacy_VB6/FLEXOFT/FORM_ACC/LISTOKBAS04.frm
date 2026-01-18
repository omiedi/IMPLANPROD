VERSION 5.00
Begin VB.Form LISTOKBAS04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stocks - Consultas y Listados"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5415
   Icon            =   "LISTOKBAS04.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000080&
      Height          =   4740
      Left            =   4515
      ScaleHeight     =   4680
      ScaleWidth      =   1425
      TabIndex        =   15
      Top             =   0
      Width           =   1485
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
         Left            =   105
         Picture         =   "LISTOKBAS04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   315
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
         Left            =   105
         Picture         =   "LISTOKBAS04.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1000
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   16
         Top             =   3600
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "LISTOKBAS04.frx":075E
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   3980
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0FF&
      Caption         =   "Control de Existencias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   210
      TabIndex        =   7
      Top             =   210
      Width           =   4110
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0FF&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   210
         TabIndex        =   9
         Top             =   1890
         Width           =   3690
         Begin VB.Frame Frame10 
            BackColor       =   &H00E0E0FF&
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
            Left            =   420
            TabIndex        =   12
            Top             =   1320
            Width           =   2850
            Begin VB.OptionButton Option4 
               BackColor       =   &H00E0E0FF&
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
               Left            =   1575
               TabIndex        =   3
               Top             =   250
               Width           =   1170
            End
            Begin VB.OptionButton Option5 
               BackColor       =   &H00E0E0FF&
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
               TabIndex        =   2
               Top             =   250
               Value           =   -1  'True
               Width           =   1380
            End
         End
         Begin VB.PictureBox Picture11 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   2790
            TabIndex        =   10
            Top             =   420
            Width           =   2850
            Begin VB.CommandButton Command30 
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
               Left            =   1690
               Picture         =   "LISTOKBAS04.frx":2680
               Style           =   1  'Graphical
               TabIndex        =   5
               ToolTipText     =   "Listado de Existencias Negativas"
               Top             =   0
               Width           =   560
            End
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
               Left            =   2230
               Picture         =   "LISTOKBAS04.frx":298A
               Style           =   1  'Graphical
               TabIndex        =   6
               ToolTipText     =   "Listado de Items con Stock Sub-Mínimo"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command19 
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
               Picture         =   "LISTOKBAS04.frx":2C94
               Style           =   1  'Graphical
               TabIndex        =   4
               ToolTipText     =   "Existencias Físicas Actuales"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Existencias Negativos Críticos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   630
               TabIndex        =   11
               Top             =   0
               Width           =   1215
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0FF&
         Caption         =   "Consultas / Listados Generales"
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
         TabIndex        =   8
         Top             =   420
         Width           =   3690
         Begin VB.PictureBox Picture12 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   2790
            TabIndex        =   13
            Top             =   420
            Width           =   2850
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
               Picture         =   "LISTOKBAS04.frx":2F9E
               Style           =   1  'Graphical
               TabIndex        =   0
               ToolTipText     =   "Listado de Existencias y Ubicaciones"
               Top             =   0
               Width           =   560
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
               Height          =   600
               Left            =   2230
               Picture         =   "LISTOKBAS04.frx":32A8
               Style           =   1  'Graphical
               TabIndex        =   1
               ToolTipText     =   "Ficha Kardex - Histórico de Movimientos"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Existencias - His- torico - Reservas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   630
               TabIndex        =   14
               Top             =   105
               Width           =   1740
            End
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "LISTOKBAS04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub



Private Sub Command19_Click()
    Command19.Enabled = False
    XX1% = COMALTER%("Opciones de Listado: \\Solo Cantidades\Valorizado a última Compra")
    Select Case XX1%
      Case 1
        Call EXISTO2
      Case 2
        Call EXISTO4
    End Select
    Command19.Enabled = True
End Sub

'
Sub EXISTO2()
    '
    TMTX% = 0: DPSX% = 0
    VDEF$ = MKI$(PRICOD%) + MKI$(ULTCOD%) + Chr$(TMTX%) + Chr$(DPSX%)
    '
10  RACOD$ = OBJPLA$("DESHASCOD4", VDEF$)
    If RACOD$ = "" Then
       Exit Sub
    End If
    CI1% = CVI(Left$(RACOD$, 2)): cod1$ = CODEXT$(CI1%)
    CI2% = CVI(Mid$(RACOD$, 3, 2)): cod2$ = CODEXT$(CI2%)
    TMT% = Asc(Mid$(RACOD$, 5, 1))
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TIMAT$ = "" Then
       TMT% = 0
       Call REPLA(RACOD$, Chr$(0), 5, 1)
       TIMAT$ = "Material No Clasificado"
    End If
    DPS% = Asc(Mid$(RACOD$, 6, 1))
    '
    If cod1$ = "" Or cod2$ = "" Or cod2$ < cod1$ Then
       VDEF$ = RACOD$
       GoTo 10
    End If
    '
    Call HEADERS("NEGASTO", "")
    Call HEADERS("LIMASEL", "Tipo de Material: " + TIMAT$)
    '
    J& = 0
    Screen.MousePointer = 11
    NVII& = NUINV%
    '
    For I& = 1 To NUMAS%
      Call TRACE(20, I&, URE&)
      CIXI% = I&
      CDDX$ = CODEXT$(CIXI%)
      If CDDX$ >= cod1$ Then
        If CDDX$ <= cod2$ Then
          TMT1% = TIMATE%(CIXI%)
          If TMT% < 1 Or TMT1% = TMT% Then
            '
            CII& = CIXI%
            Call VESARRAS(CIXI%)
            'SSTT = STOCKACT(CIXI%)
            SSTT = STODEPOS(CIXI%, DPS%)
            If Abs(SSTT) > 0.05 Then
              TTN$ = MKI$(CIXI%) + MKS$(SSTT) + String$(4, 0)
              J& = J& + 1
              Call GRAREG("NEGASTO", TTN$, J&)
            End If
            '
          End If
        End If
      End If
    Next I&
    '
    Call SETULTREG("NEGASTO", J&)
    Call FILESORT("NEGASTO", "", 1, 1, "ASC")
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    If J& < 1 Then
        Call COMUNI("No se Detectaron Existencias Físicas.\Pulse cualquier tecla para volver al menu.")
      Else
        If Option5.Value = True Then
            Call CONECRUN("*EXISTS", "")
          ElseIf Option4.Value = True Then
            Call CONECRUN("*EXISTL", "")
        End If
    End If
    '
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Call LIUBIEXI
    Command2.Enabled = True
End Sub

Private Sub Command20_Click()
    Command20.Enabled = False
    Call CRITSTO2
    If ULTREG&("CRITSTO") > 0 Then
        Call CIERRARCH("*.*")
        If Option5.Value = True Then
            Call CONECRUN("*CRITSTS", "")
          ElseIf Option4.Value = True Then
            Call CONECRUN("*CRITSTL", "")
        End If
    End If
    Command20.Enabled = True
End Sub


Private Sub Command29_Click()
    Call HELPONLINE("CONSTO00")
End Sub



Private Sub Command30_Click()
    Command30.Enabled = False
    Call NEGASTO2
    If ULTREG&("NEGASTO") > 0 Then
        Call CIERRARCH("*.*")
        If Option5.Value = True Then
            Call CONECRUN("*NEGASTS", "")
          ElseIf Option4.Value = True Then
            Call CONECRUN("*NEGASTL", "")
        End If
    End If
    Command30.Enabled = True
End Sub




Private Sub Command9_Click()
'    If NUEWINTA% = 1 Then
'        Call CONECRUN("*KARDEX", "Ficha Kardex de Movimientos")
'      Else
'        Call GENKARDEX("")
'    End If
    Command9.Enabled = False
    FIKARDEX.Show 1
    Command9.Enabled = True
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    '
    Call TRANSLABELS(Name)
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
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
   Hide
End Sub


'
Sub NEGASTO2()    'DETERMINACION ITEMS CON STOCK NEGATIVO
    '
    J& = 0             'Call BLANARCH("NEGASTO")
    Screen.MousePointer = 11
    '
    FIN& = NUMAS%
    For I& = 1 To NUMAS%
      Call TRACE(20, I&, FIN&)
      CI% = I&
      Call VESARRAS(CI%)
      SSTT = STOCKACT(CI%)
      If SSTT < (-0.05) Then
        TTN$ = MKI$(CI%) + MKS$(SSTT) + String$(4, 0)
        J& = J& + 1
        Call GRAREG("NEGASTO", TTN$, J&)
      End If
    Next I&
    '
    Call SETULTREG("NEGASTO", J&)
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    If J& < 1 Then
        Call COMUNI("No se detectaron items con stock negativo.\Pulse cualquier tecla para volver al menu.")
      Else
        Call COMUNI("Se Detectaron " + TRIM$(Str$(J&)) + " Casos\de Stocks Negativos.")
    End If
    '
End Sub
'

'
Sub EXISTO4()
Dim PUN1#(32767), MOPUN1%(32767), FEPUN1#(32767)
    '
    TMTX% = 0
    VDEF$ = MKI$(PRICOD%) + MKI$(ULTCOD%) + Chr$(TMTX%) + Chr$(DPSX%)
    '
10  RACOD$ = OBJPLA$("DESHASCOD4", VDEF$)
    If RACOD$ = "" Then
       Exit Sub
    End If
    CI1% = CVI(Left$(RACOD$, 2)): cod1$ = CODEXT$(CI1%)
    CI2% = CVI(Mid$(RACOD$, 3, 2)): cod2$ = CODEXT$(CI2%)
    TMT% = Asc(Mid$(RACOD$, 5, 1))
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TIMAT$ = "" Then
       TMT% = 0
       Call REPLA(RACOD$, Chr$(0), 5, 1)
       TIMAT$ = "Material No Clasificado"
    End If
    DPS% = Asc(Mid$(RACOD$, 6, 1))
    '
    If cod1$ = "" Or cod2$ = "" Or cod2$ < cod1$ Then
       VDEF$ = RACOD$
       GoTo 10
    End If
    '
    Call HEADERS("NEGASVAL", "")
    Call HEADERS("LIMASEL", "Tipo de Material: " + TIMAT$)
    '
    J& = 0
    Screen.MousePointer = 11
    NVII& = NUINV%
    '
    Call ABRECOMPRAS
    SQLX$ = "SELECT * FROM OCOMDETA "
    SQLX$ = SQLX$ + " WHERE Stat_Ocom < 9 " 'NO ANULADA
    Set STKVALTMP = Compras.OpenRecordset(SQLX$, 4)
    With STKVALTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        GoTo 11
      End If
      On Error GoTo 0
      .MoveLast
      TOT& = .RecordCount
      .MoveFirst
      RA& = 0
      Do While .EOF = False
        FEPUN1#(!Cod_Inte) = CDbl(!Femi_Ocom)
        RA& = RA& + 1
        Call TRACE(20, RA&, TOT&)
        .MoveNext
      Loop
      .MoveFirst
      RA& = 0
      Do While .EOF = False
        If FEPUN1#(!Cod_Inte) <= CDbl(!Femi_Ocom) Then
          FEPUN1#(!Cod_Inte) = CDbl(!Femi_Ocom)
          MOPUN1%(!Cod_Inte) = !Mone_Emis
          PUN1#(!Cod_Inte) = !PrUn_Neto '!Pre_Unit
        End If
        RA& = RA& + 1
        Call TRACE(20, RA&, TOT&)
        .MoveNext
      Loop
    End With
    '
    Y$ = REGBLAN("NEGASVAL")
    '
11  For I& = 1 To NUMAS%
      URE& = NUMAS%
      Call TRACE(20, I&, URE&)
      CIXI% = I&
      CDDX$ = CODEXT$(CIXI%)
      If CDDX$ >= cod1$ Then
        If CDDX$ <= cod2$ Then
          TMT1% = TIMATE%(CIXI%)
          If TMT% < 1 Or TMT1% = TMT% Then
            '
            CII& = CIXI%
            Call VESARRAS(CIXI%)
            'SSTT = STOCKACT(CIXI%)
            SSTT = STODEPOS(CIXI%, DPS%)
            '
            If Abs(SSTT) > 0.05 Then
              Call REPLA(Y$, MKI$(CIXI%), 1, 2)
              Call REPLA(Y$, MKS$(SSTT), 3, 4)
              Call REPLA(Y$, MKD$(PUN1#(CII&)), 9, 8)
              Call REPLA(Y$, MKI$(MOPUN1%(CII&)), 17, 2)
              Call REPLA(Y$, MKI$(CVINT(FEPUN1#(CII&))), 19, 2)
              IMPORTE1# = SSTT * PUN1#(CII&)
              Call REPLA(Y$, MKD$(IMPORTE1#), 21, 8)
              IMPORTE2# = IMPORTE1# * TICAMONE(MOPUN1%(CII&)) / 1
              Call REPLA(Y$, MKD$(IMPORTE2#), 29, 8)
              J& = J& + 1
              Call GRAREG("NEGASVAL", Y$, J&)
            End If
            '
          End If
        End If
      End If
    Next I&
    '
    Call SETULTREG("NEGASVAL", J&)
    Call FILESORT("NEGASVAL", "", 1, 1, "ASC")
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    If J& < 1 Then
        Call COMUNI("No se Detectaron Existencias Físicas.\Pulse cualquier tecla para volver al menu.")
      Else
        Call FINAL("*LISTVAL")
    End If
    '
End Sub
'
Sub CRITSTO2()    'DETERMINACION ITEMS CON STOCK SUB MINIMO
    '
    J& = 0: JC& = 0          'Call BLANARCH("NEGASTO")
    Screen.MousePointer = 11
    FIN& = NUMAS%
    FF% = HOY(HOS$)
    PERIO$ = "St.Min." + HOS$
    '
    For I& = 1 To NUMAS%
      Call TRACE(20, I&, FIN&)
      CI% = I&
      Call VESARRAS(CI%)
      SSTT = STOCKACT(CI%)
      '
      STMIN = STOMIN(CI%)
      If SSTT < STMIN Then
        '
        COU = COSPESOS(CI%)
        REPOS = STMIN - SSTT
        If REPOS < LOTESTAN(CI%) Then REPOS = LOTESTAN(CI%)
        '
        TTN$ = REGBLAN$("CRITSTO")
        Call REPLA(TTN$, MKI$(CI%), 1, 2)
        Call REPLA(TTN$, UMED$, 3, 2)
        Call REPLA(TTN$, MKS$(SSTT), 5, 4)
        Call REPLA(TTN$, MKS$(STMIN), 9, 4)
        Call REPLA(TTN$, MKS$(STMIN - SSTT), 13, 4)
        Call REPLA(TTN$, MKS$(REPOS), 17, 4)
        Call REPLA(TTN$, MKS$(COU), 21, 4)
        Call REPLA(TTN$, MKD$(COU * REPOS), 25, 8)
        J& = J& + 1
        Call GRAREG("CRITSTO", TTN$, J&)
        '
        Z$ = REGBLAN$("PLACOMP")
        Call REPLA(Z$, MKI$(CI%), 1, 2)
        Call REPLA(Z$, MKI$(FF%), 3, 2)
        Call REPLA(Z$, PERIO$, 5, 16)
           CANTX = LOTESTAN(CI%)
           If CANTX < STMIN - SSTT Then CANTX = STMIN - SSTT
        Call REPLA(Z$, MKS$(CANTX), 21, 4)
        Call REPLA(Z$, MKS$(0), 25, 4) ' tot.progra.por item
        Call REPLA(Z$, MKS$(0), 29, 4) ' filler
           PHB% = PROVHABI%(CI%)
        Call REPLA(Z$, MKI$(PHB%), 33, 2)
        Call REPLA(Z$, MKS$(COU), 35, 4)
        Call REPLA(Z$, MKD$(COU * CANTX), 39, 8)
        '
        JC& = JC& + 1
        Call GRAREG("PLACOMP", Z$, JC&)
        '
      End If
    Next I&
    '
    Call SETULTREG("CRITSTO", J&)
    Call SETULTREG("PLACOMP", JC&)
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    If J& < 1 Then
        Call COMUNI("No se detectaron items con stock sub-mínimo.\Pulse cualquier tecla para volver al menu.")
      Else
        Call COPYSTRU("PROVED1", "PROVEPLA")
        ASELPRO$ = "PROVEPLA"
        Screen.MousePointer = 11
        PPXXZ$ = String$(32767, 0)
        For U& = 1 To ULTREG("PLACOMP")
           XX$ = REGLEIDO$("PLACOMP", U&)
           NPRX% = CVI(Mid$(XX$, 33, 2))
           If NPRX% > 0 Then
              Mid$(PPXXZ$, NPRX%, 1) = Chr$(1)
           End If
        Next U&
        '
        JJ& = 0
        For U& = 1 To ULTREG&("ACREDOR")
          XX$ = REGLEIDO$("ACREDOR", U&)
          NPRX% = CVI(Left$(XX$, 2))
          If NPRX% > 0 Then
            If Asc(Mid$(PPXXZ$, NPRX%, 1)) > 0 Then
              JJ& = JJ& + 1
              Call GRAREG("PROVEPLA", XX$, JJ&)
            End If
          End If
        Next U&
        Call SETULTREG("PROVEPLA", JJ&)
        Screen.MousePointer = 1
        '
        Call COMUNI("Se Detectaron " + TRIM$(Str$(J&)) + " Casos\de Stocks Sub-Mínimo.")
    End If
    '
End Sub

Sub LIUBIEXI()
    '
    Dim PRIRE&(32767)
    Static NUSU%
    '
    TXT$ = "El Listado de Artículos y Ubicaciones Existente\"
    FEX = FEGEAR%("UBIEXI")
    If FEX < 1 Then GoTo 310
    TXT$ = TXT$ + "Corresponde al " + FECHATEX$(FEX) + " a las " + HORATEX$(HOGEAR%("UBIEXI")) + " hs.\"
    TXT$ = TXT$ + "\Consultar\Re-Generar\Abandonar"
    X% = COMALTER%(TXT$)
    On X% GoTo 320, 310
    Exit Sub
    '
310 Screen.MousePointer = 11
    NMS% = NUMAS%
    '
'    If NUSU% = 0 Then
'        NUSU% = 1
'        For I& = 1 To ULTREG&("SUCURSAL")
'            XXX$ = REGLEIDO$("SUCURSAL", I&)
'            If I& <= 62 Then
'                If Asc(Left$(XXX$, 1)) > 32 Then
'                    NUSU% = I&
'                End If
'            End If
'        Next I&
'        Call CIERRARCH("SUCURSAL")
'    End If
'    '
    J& = 0
    FIN& = ULTREG&("UBISTOCK")
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      XX$ = REGLEIDO$("UBISTOCK", I&)
      CI% = CVI(Left$(XX$, 2))
      UBI$ = AJUSTI$(Mid$(XX$, 3, 24), 24) + AJUSTI$(Mid$(XX$, 31, 30), 30)
      If TRIM$(UBI$) <> "" Then
        If CI% > 0 Then
          If CI% <= NMS% Then
            YY$ = REGBLAN$("UBIEXI")
            Call REPLA(YY$, MKI$(CI%), 1, 2)
            Call REPLA(YY$, UBI$, 7, 54)
            J& = J& + 1
            Call GRAREG("UBIEXI", YY$, J&)
            If PRIRE&(CI%) < 1 Then
              PRIRE&(CI%) = J&
            End If
          End If
        End If
      End If
    Next I&
    '
311 FIN& = NUMAS%
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      CI% = I&
      Call VESARRAS(CI%)
      CAN = STOCKACT(CI%)
      '
      If Abs(CAN) >= 0.05 Then
        If PRIRE&(CI%) > 0 Then
            YY$ = REGLEIDO$("UBIEXI", PRIRE&(CI%))
            Call REPLA(YY$, MKS$(CAN), 3, 4)
            Call GRAREG("UBIEXI", YY$, PRIRE&(CI%))
          Else
            YY$ = REGBLAN$("UBIEXI")
            Call REPLA(YY$, MKI$(CI%), 1, 2)
            Call REPLA(YY$, MKS$(CAN), 3, 4)
            J& = J& + 1
            Call GRAREG("UBIEXI", YY$, J&)
            PRIRE&(CI%) = J&
        End If
      End If
    Next I&
    '
    Call SETULTREG("UBIEXI", J&)
    Call CIERRARCH("UBIEXI")
    Call GRARGEN("UBIEXI", "")
    '
320 Screen.MousePointer = 1
    FIN& = ULTREG&("UBIEXI")
    If FIN& < 1 Then
      Call CIERRARCH("*.*")
      Exit Sub
    End If
    '
330 VDEF$ = MKI$(PRICOD%) + MKI$(ULTCOD%) + Chr$(0)
    '
    OBX$ = OBJPLA$("DESHASCOD", VDEF$)
    If OBX$ <> "" Then
      '
      Screen.MousePointer = 11
      CI1% = CVI(Left$(OBX$, 2))
      cod1$ = CODEXT$(CI1%)
      CI2% = CVI(Mid$(OBX$, 3, 2))
      cod2$ = CODEXT$(CI2%)
      '
      J& = 0
      For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        YY$ = REGLEIDO$("UBIEXI", U&)
        CIII% = CVI(Left$(YY$, 2))
        CODX$ = CODEXT$(CIII%)
        If CODX$ >= cod1$ Then
          If CODX$ <= cod2$ Then
            J& = J& + 1
            Call GRAREG("LUBIEXI", YY$, J&)
          End If
        End If
      Next U&
      Call SETULTREG("LUBIEXI", J&)
      Call CIERRARCH("LUBIEXI")
      Call CIERRARCH("*.*")
      Screen.MousePointer = 1
      Call FINAL("*SUBIEXI")
    End If
    '
End Sub
'


