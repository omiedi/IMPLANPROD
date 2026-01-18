VERSION 5.00
Begin VB.Form OPRESERVA 
   BackColor       =   &H00EEEEC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reservas de Materiales"
   ClientHeight    =   6180
   ClientLeft      =   2040
   ClientTop       =   1305
   ClientWidth     =   5610
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4620
      ScaleHeight     =   75
      ScaleWidth      =   705
      TabIndex        =   24
      Top             =   2100
      Width           =   765
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command6 
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
      Left            =   4580
      Picture         =   "OPRESERVA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1155
      Width           =   800
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EEEEC0&
      Caption         =   "Actualización"
      Height          =   2700
      Left            =   210
      TabIndex        =   10
      Top             =   210
      Width           =   4110
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1365
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   11
         Top             =   1845
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
            Left            =   0
            Picture         =   "OPRESERVA.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Depuración de    la Base de Datos"
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
            TabIndex        =   13
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1050
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   21
         Top             =   1370
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
            Picture         =   "OPRESERVA.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Control Según O/Fs en Curso"
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
            Top             =   0
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   14
         Top             =   895
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
            Height          =   600
            Left            =   0
            Picture         =   "OPRESERVA.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Cancelación      de Reservas"
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
            TabIndex        =   16
            Top             =   0
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   17
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
            Picture         =   "OPRESERVA.frx":0C28
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Generar  Re- servas por O/F"
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
            TabIndex        =   19
            Top             =   0
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEC0&
      Caption         =   "Consultas - Listados"
      Height          =   3000
      Left            =   210
      TabIndex        =   1
      Top             =   3045
      Width           =   5160
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
         Height          =   525
         Left            =   4095
         Picture         =   "OPRESERVA.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2280
         Width           =   855
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Totales por Material Reservado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   525
         TabIndex        =   8
         Top             =   525
         Width           =   4110
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00EEEEC0&
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
         TabIndex        =   5
         Top             =   2205
         Width           =   3690
         Begin VB.OptionButton Option5 
            BackColor       =   &H00EEEEC0&
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
            TabIndex        =   7
            Top             =   250
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00EEEEC0&
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
            Left            =   2100
            TabIndex        =   6
            Top             =   250
            Width           =   1275
         End
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Por Material"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   525
         TabIndex        =   4
         Top             =   1680
         Width           =   2955
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Por Número de O/Fabricación "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   525
         TabIndex        =   3
         Top             =   1365
         Width           =   3060
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Consulta General de Reservas por O/F"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   525
         TabIndex        =   2
         Top             =   1050
         Value           =   -1  'True
         Width           =   3900
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
      Left            =   4580
      Picture         =   "OPRESERVA.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   315
      Width           =   800
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3990
      Picture         =   "OPRESERVA.frx":1386
      Top             =   2390
      Width           =   2010
   End
End
Attribute VB_Name = "OPRESERVA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    OPRESERVA.Hide
End Sub

Private Sub Command2_Click()
    OPCI% = 1
    If OPRESERVA.Option2.Value = True Then OPCI% = 2
    If OPRESERVA.Option3.Value = True Then OPCI% = 3
    If OPRESERVA.Option6.Value = True Then OPCI% = 4
    If OPCI% = 1 Then
        If Option4.Value = True Then
             Call CONECRUN("*LIRESE1", "")
           Else
             Call CONECRUN("*CORESE1", "")
        End If
      ElseIf OPCI% = 2 Then
        If EXISTE%(LUDAT$ + "CARFORLI.DAT") > 0 Then
          If Option4.Value = True Then
            Call LISTAMAT
            Exit Sub
          End If
        End If
        Call CONECRUN("*CORESE2", "")
      ElseIf OPCI% = 3 Then
        Call CONECRUN("*CORESE3", "")
      ElseIf OPCI% = 4 Then
        Call TOMATRES
      Else
        OPRESERVA.Hide
    End If
    '
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call GENRESER("")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call ANURESER("")
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    '
    If DERACCE%("DEPURESE", "Depuración de Reservas de Materiales") < 2 Then GoTo 99
    Call BLOQARCH("RESERVA")
    Screen.MousePointer = 11
    Call COPYSTRU("RESERVA", "RESERVB")
    Call BLANARCH("!RESERVA")
    JJ& = 0: KK& = ULTREG&("RESERVB")
    FIN& = ULTREG&("RESERVA")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("RESERVA", U&)
      CAPEN = CVD(Mid$(XX$, 19, 8)) - CVD(Mid$(XX$, 27, 8))
      If CAPEN >= 0.05 Then
          JJ& = JJ& + 1
          Call GRAREG("!RESERVA", XX$, JJ&)
        Else
          Call REGAPP("RESERVB", XX$)
      End If
    Next U&
    Call SETULTREG("!RESERVA", JJ&)
    Call CIERRARCH("!RESERVA")
    Call TRACENTRAL("RESERVA", "RESERVA")
    Call CIERRARCH("RESERVA")
    Call CIERRARCH("RESERVB")
    Screen.MousePointer = 1
99  Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call HELPONLINE("RESERVA")
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   Call CHECKRES
   Command7.Enabled = True
End Sub

Private Sub Form_Load()
   '
   Call TRANSLABELS(Name)
   '
   OPRESERVA.Option1.Value = True
End Sub
'
Sub TOMATRES()
    '
    Dim CARESE#(16384)
    FIN& = ULTREG&("RESERVA")
    If FIN& < 1 Then
        Call MENSERR(24, "No hay Reservas Registradas")
        Exit Sub
    End If
    '
    Screen.MousePointer = 11
    '
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        XX$ = REGLEIDO$("RESERVA", I&)
        CI% = CVI(Left$(XX$, 2))
        If CI% > 0 Then
            If CI% <= NUMAS% Then
                TORE# = CVD(Mid$(XX$, 19, 8))
                CACO# = CVD(Mid$(XX$, 27, 8))
                SIGNO% = Sgn(CVS(Mid$(XX$, 75, 4)))
                SARE# = TORE# - CACO#
                If SARE# * SIGNO% < 0 Then
                    SARE# = 0
                End If
                Call REPLA(XX$, MKD$(SARE#), 35, 8)
                Call GRAREG("RESERVA", XX$, I&)
                CARESE#(CI%) = CARESE#(CI%) + SARE#
            End If
        End If
    Next I&
    '
    FIN& = ULTREG&("REPUREPA")
    For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       XX$ = REGLEIDO$("REPUREPA", I&)
       CI% = CVI(Mid$(XX$, 7, 2))
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           TORE# = CVS(Mid$(XX$, 21, 4))
           CACO# = CVS(Mid$(XX$, 25, 4))
           '
           If UCase$(UNIMED$(CIIX%)) = "GK" Then
             TORE# = TORE# / 1000
             CACO# = CACO# / 1000
           End If
           '
           SARE# = TORE# - CACO#
           If SARE# < 0.05 Then GoTo 29
           '
           NOREPARA& = CVS(Left$(XX$, 4))
           REREP& = REGBUS&("PRESREPA", 1, MKS$(NOREPARA&))
           If REREP& > 0 Then
             YY$ = REGLEIDO$("PRESREPA", REREP&)
             If CVI(Mid$(YY$, 123, 2)) > 0 Then GoTo 29 ' O/Rep Cerrada
             If CVS(Mid$(YY$, 125, 4)) < 1 Then GoTo 29 ' falta O/Venta
           End If
           '
           CARESE#(CI%) = CARESE#(CI%) + SARE#
         End If
       End If
29  Next I&
    '
    JJ& = 0
    FIN& = NUINV%: If FIN& < 1 Then FIN& = 1
    '
    For I& = 1 To NUINV%
        Call TRACE(20, I&, FIN&)
        II$ = REGLEIDO$("INVERT", I& + 1)
        CI% = CVI(Mid$(II$, 17, 2))
        If CI% > 0 Then
            If CI% <= NUMAS% Then
                If Abs(CARESE#(CI%)) > 0.05 Then
                    'STOCK# = STODEPOS(CI%, DEPOEXPE%)
                    STOCK# = STOCKACT(CI%)
                    XX$ = REGBLAN$("TOMATRES")
                    Call REPLA(XX$, MKI$(CI%), 1, 2)
                    Call REPLA(XX$, MKD$(CARESE#(CI%)), 3, 8)
                    Call REPLA(XX$, MKD$(STOCK#), 11, 8)
                    FALTA# = CARESE#(CI%) - STOCK#
                    If FALTA# < 0 Then FALTA# = 0
                    Call REPLA(XX$, MKD$(FALTA#), 35, 8)
                    OBSE$ = "      "
                    If FALTA# > 0 Then OBSE$ = "CRIT. "
                    If STOCK# - STOMIN(CI%) < 0 Then OBSE$ = "CRIT. "
                    If STOCK# < 0 Then OBSE$ = "NEG.  "
                    Call REPLA(XX$, OBSE$, 43, 6)
                    JJ& = JJ& + 1
                    Call GRAREG("TOMATRES", XX$, JJ&)
                End If
            End If
        End If
    Next I&
    '
    Call SETULTREG("TOMATRES", JJ&)
    Call CIERRARCH("TOMATRES")
    Screen.MousePointer = 1
    '
    If OPRESERVA.Option4.Value = True Then
           Call FINAL("*LITOTRE")       ' impresion
         Else
           Call FINAL("*COTOTRE")       ' pantalla
    End If
    '
End Sub
'
Sub LISTAMAT()
    '
    REBLA$ = REGBLAN$("LISTAMAT")
    '
10  XX$ = REGSEL$("ORFANUE")
    If Len(XX$) > 4 Then
      NUOFAC& = CVS(Mid$(XX$, 1, 4))
      CIACT% = CVI(Mid$(XX$, 71, 2))
      DCLIEN$ = TRIM$(Mid$(XX$, 101, 8))
      '
      TITUX$ = "O.F.: " + TRIM$(Str$(NUOFAC&)) + " - " + DESCRIT0$(CIACT%) + "        Cliente / Destino: " + DCLIEN$
      '
      RFF$ = RECFILT$("CARFORLI", 1, MKI$(CIACT%))
      If Len(RFF$) < 4 Then
         Call COMUNI("Código sin Fórmula por Lista")
         Option5.Value = True
         Command2.Enabled = True
         Call Command2_Click
         Exit Sub
       Else
        JJ& = 0: DPRPA% = 0
        For KU& = 1 To Len(RFF$) Step 4
          RECA& = CVS(Mid$(RFF$, KU&, 4))
          ZZ$ = REGLEIDO$("CARFORLI", RECA&)
          '
          CODIEL% = CVI(Mid$(ZZ$, 19, 2))
          CODIMP% = CVI(Mid$(ZZ$, 97, 2))
          If CODIMP% < 1 Then
            If CODIEL% > 0 Then
              If COSUNI(CODIEL%) > 0 Then
                CODIMP% = CODIEL%
              End If
            End If
          End If
          DPRP% = DEPOPRE%(CODIMP%)
          If DPRP% < 1 Then DPRP% = 1
          '
          YY$ = REBLA$
            JI% = JJ&
          Call REPLA(YY$, MKI$(JI%), 1, 2)
            CANTI = CVS(Mid$(ZZ$, 93, 4))
          Call REPLA(YY$, MKS$(CANTI), 3, 4)
            NIVIEL% = Asc(Mid$(ZZ$, 18, 1))
            If NIVIEL% = 1 Then GRUPO% = GRUPO% + 1
            NPLA$ = TRIM$(Mid$(ZZ$, 21, 16))
            NPOS$ = TRIM$(Mid$(ZZ$, 37, 8))
            NDES$ = TRIM$(Mid$(ZZ$, 45, 48))
          Call REPLA(YY$, MKI$(CODIEL%), 7, 2)
          Call REPLA(YY$, NPLA$, 9, 16)
          Call REPLA(YY$, NPOS$, 25, 8)
          Call REPLA(YY$, NDES$, 33, 40)
          Call REPLA(YY$, MKI$(GRUPO%), 73, 2)
          If CODIMP% < 1 Then GoTo 29
          '
          Call REPLA(YY$, DESCRIT0$(CODIMP%), 81, 46)
          Call REPLA(YY$, DEPOPRE%(CODIMP%), 127, 2)
          Call REPLA(YY$, MKI$(CODIMP%), 129, 2)
            CANTID = CANTI
            If CODIMP% < 1 Then CANTID = 0
            SONKILOS% = 0
            MEDI2 = CVS(Mid$(ZZ$, 103, 4))
              If MEDI2 > 0.05 Then CANTID = CANTID * (MEDI2 + 5) / 1000: SONKILOS% = 1
            MEDI3 = CVS(Mid$(ZZ$, 107, 4))
              If MEDI3 > 0.05 Then CANTID = CANTID * (MEDI3 + 5) / 1000: SONKILOS% = 1
            MEDI1 = CVS(Mid$(ZZ$, 99, 4))
              If MEDI1 > 0.05 Then CANTID = CANTID * (MEDI1 + 3) / 1000: SONKILOS% = 1
            KILOUN = PESUNI(CODIMP%)
              If SONKILOS% > 0 Then
                If KILOUN > 0.0005 Then
                  CANTID = CANTID * KILOUN
                End If
              End If
            PORSCRA = CVS(Mid$(ZZ$, 115, 4))
              CANTID = CANTID * (1 + PORSCRA / 100)
          Call REPLA(YY$, MKS$(MEDI1), 131, 4)
          Call REPLA(YY$, MKS$(MEDI2), 135, 4)
          Call REPLA(YY$, MKS$(MEDI3), 139, 4)
          If SONKILOS% > 0 Then
            Call REPLA(YY$, MKS$(KILOUN), 143, 4)
            Call REPLA(YY$, MKS$(PORSCRA), 147, 4)
            Call REPLA(YY$, MKS$(CANTID), 151, 4)
          End If
          Call REPLA(YY$, MKI$(DPRP%), 155, 2)
            OBSER$ = TRIM$(Mid$(ZZ$, 125, 36))
          Call REPLA(YY$, OBSER$, 157, 36)
          '
          JJ& = JJ& + 1
          Call GRAREG("LISTAMAT", YY$, JJ&)
          '
29      Next KU&
        '
        Call SETULTREG("LISTAMAT", JJ&)
        Call CIERRARCH("LISTAMAT")
        '
        Call FILESORT("LISTAMAT", "LISTBMAT", 10, 11, "ASC")
        '
        LL& = 0
        DPRPA% = 0
        For U& = 1 To JJ&
          YY$ = REGLEIDO$("LISTBMAT", U&)
          DPRP% = CVI(Mid$(YY$, 155, 2))
          If DPRP% <> DPRPA% Then
            If DPRPA% > 0 Then
              LL& = LL& + 1
              Call GRAREG("LISTAMAT", String$(512, "="), LL&)
            End If
            DPRPA% = DPRP%
          End If
          LL& = LL& + 1
          Call GRAREG("LISTAMAT", YY$, LL&)
        Next U&
        Call SETULTREG("LISTAMAT", LL&)
        '
        Call HEADERS("LISTAMAT", "")
        Call HEADERS("LISTAMAT", TITUX$)
        Call FINAL("*LISTAMAT")
      End If
      GoTo 10
    End If
    '
End Sub


