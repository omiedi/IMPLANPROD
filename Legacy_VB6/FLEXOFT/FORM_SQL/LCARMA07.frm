VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form LCARMA07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0D0D0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "C.R.P. - Programación de Carga de Máquinas"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4470
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   1
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   1215
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   2640
      Left            =   3600
      ScaleHeight     =   2610
      ScaleWidth      =   1245
      TabIndex        =   14
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton AGenRep 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   635
         Left            =   105
         Picture         =   "LCARMA07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   960
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   635
         Left            =   105
         Picture         =   "LCARMA07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   250
         Width           =   650
      End
   End
   Begin VB.PictureBox Picture11 
      BackColor       =   &H80000016&
      Height          =   645
      Left            =   945
      ScaleHeight     =   585
      ScaleWidth      =   2265
      TabIndex        =   8
      Top             =   5145
      Width           =   2325
      Begin VB.CommandButton Command33 
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
         Picture         =   "LCARMA07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Planilla de Horas por Sección"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Horas - Sección"
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
         TabIndex        =   10
         Top             =   100
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture12 
      BackColor       =   &H80000016&
      Height          =   645
      Left            =   945
      ScaleHeight     =   585
      ScaleWidth      =   2265
      TabIndex        =   5
      Top             =   4620
      Width           =   2325
      Begin VB.CommandButton Command27 
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
         Picture         =   "LCARMA07.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Planilla General de Producción"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Pl. Producción"
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
         TabIndex        =   7
         Top             =   100
         Width           =   1800
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0D0D0&
      Caption         =   "LISTADOS VARIOS DE CONTROL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton Command7 
         Height          =   630
         Left            =   870
         Picture         =   "LCARMA07.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Carga de Trabajo por Maquina"
         Top             =   1315
         Width           =   550
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   840
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   16
         Top             =   1295
         Width           =   2325
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Programa de Carga    por Máquina"
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
            Left            =   735
            TabIndex        =   17
            Top             =   75
            Width           =   1800
         End
      End
      Begin VB.PictureBox Picture10 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   630
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   11
         Top             =   840
         Width           =   2325
         Begin VB.CommandButton Command32 
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
            Picture         =   "LCARMA07.frx":0D72
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Plan de Fechas Proyectadas de Entrega"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Fechas de Entrega Proyectadas"
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
            Left            =   735
            TabIndex        =   13
            Top             =   0
            Width           =   1485
         End
      End
      Begin VB.PictureBox Picture13 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   420
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   2
         Top             =   420
         Width           =   2325
         Begin VB.CommandButton Command28 
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
            Picture         =   "LCARMA07.frx":107C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Planilla de Carga de Máquinas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Planilla de Carga"
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
            Left            =   630
            TabIndex        =   4
            Top             =   100
            Width           =   1695
         End
      End
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
      Height          =   750
      Left            =   840
      Picture         =   "LCARMA07.frx":1386
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   3255
      Width           =   750
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4020
      OleObjectBlob   =   "LCARMA07.frx":1690
      Top             =   2100
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   4080
      OleObjectBlob   =   "LCARMA07.frx":18C4
      TabIndex        =   21
      Top             =   420
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "LCARMA07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AGenRep_Click()
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

Private Sub Command27_Click()
    Call OPNOIN("")
End Sub

Sub Command28_Click()
    Command28.Enabled = False
10  Call PLACARMA
    Command28.Enabled = True
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("CARGMAQ")
End Sub

Sub Command32_Click()
    FPROEN07.Show 1
End Sub

Private Sub Command33_Click()
    Call OPNOIN("")
End Sub

Sub Command7_Click()
   '
   Command7.Enabled = False
   REBLA$ = REGBLAN$("LIPROCAR")
   CONSIMULACION% = 0
     If EXISTE_TABLA("SIMU_ORDEFABR") > 0 Then
        If CantRegistros("SIMU_ORDEFABR", , "NOMESS") > 0 Then
          OPX% = COMALTER%("Opciones de Trabajo\\Listado Cálculo Normal\Listado Considerando Simulación")
          If OPX% > 1 Then
            CONSIMULACION% = XVALO(CONTROL("CARMAQUI", "ULTISIMU"))
            If CONSIMULACION% < 1 Then
               Call COMUNI("Atención!! No Hay Cálculos con Simulación de Fabricación\Se Visualizará Cálculo Normal")
            End If
          End If
        End If
     End If
     ARCHOPERFAP$ = "OPERFAP"
     If CONSIMULACION% < 1 Then
       Call GENEOPERFAP
     Else
       ARCHOPERFAP$ = "SOPERFAP"
     End If
   
   
   '
   JJ& = 0
   For IK& = 1 To ULTREG&(ARCHOPERFAP$)
      ZZ$ = REGLEIDO$(ARCHOPERFAP$, IK&)
      YY$ = REBLA$
      '
        Maquina$ = Mid$(ZZ$, 73, 6)
      Call REPLA(YY$, Maquina$, 1, 6)
        NumeOrFa$ = Mid$(ZZ$, 1, 12)
      Call REPLA(YY$, NumeOrFa$, 7, 12)
         CodPie$ = Mid$(ZZ$, 13, 2)
       Call REPLA(YY$, CodPie$, 19, 2)
         CANTIPIE = CVD(Mid$(ZZ$, 79, 8))
       Call REPLA(YY$, MKS$(CANTIPIE), 21, 4)
         SALDOPIE = CVD(Mid$(ZZ$, 95, 8))
       Call REPLA(YY$, MKS$(SALDOPIE), 25, 4)
         NumeOpe$ = Mid$(ZZ$, 15, 2)
       Call REPLA(YY$, NumeOpe$, 29, 2)
         DesOpe$ = Mid$(ZZ$, 17, 48)
       Call REPLA(YY$, DesOpe$, 31, 34)
         TIEPRO = 60 * CVS(Mid$(ZZ$, 131, 4))
       Call REPLA(YY$, MKS$(TIEPRO), 65, 4)
         DesFec$ = Mid$(ZZ$, 135, 2)
       Call REPLA(YY$, DesFec$, 73, 2)
         DesHor$ = Mid$(ZZ$, 137, 2)
       Call REPLA(YY$, DesHor$, 75, 2)
         HasFec$ = Mid$(ZZ$, 139, 2)
       Call REPLA(YY$, HasFec$, 77, 2)
         HasHor$ = Mid$(ZZ$, 141, 2)
       Call REPLA(YY$, HasHor$, 79, 2)
         Obser$ = Mid$(ZZ$, 151, 48)
       Call REPLA(YY$, Obser$, 81, 16)
       '
       JJ& = JJ& + 1
       Call GRAREG("LIPROCAR", YY$, JJ&)
   Next IK&
   '
   Call SETULTREG("LIPROCAR", JJ&)
   Call FILESORT("LIPROCAR", "", 12, 15, "ASC")
   Call FILESORT("LIPROCAR", "", 1, 1, "ASC")
   Call CIERRARCH("LIPROCAR")
   Call FINAL("*LIPROCA")
   '
   Command7.Enabled = True
End Sub

Private Sub Form_Load()
    Call APLICACIONSKINS(Me, Picture1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
End Sub

Sub PLACARMA()        ' GENERA PLANILLA DE CARGA DE MAQUINA
     '
     Dim MAQUISEL$(512)
     Dim EQSEL$(512), PUNOP$(512)
     '
100  FORSELEQU.Show 1
     Call FRESHALL
     If FORSELEQU.Label2 <> "OK" Then Exit Sub
     '
     CAMAQUI% = 0
     HO% = HOY(HOS$)
     DIA% = FECHANUM(FORSELEQU.Text1)
     DIADOS% = FECHANUM(FORSELEQU.Text2)
     MAQSEL$ = ""
     For IIMAQ% = 0 To FORSELEQU.LIMATSEL.ListCount - 1
        If FORSELEQU.LIMATSEL.Selected(IIMAQ%) Then MAQSEL$ = MAQSEL$ & Mid$(FORSELEQU.LIMATSEL.List(IIMAQ%), 1, 6)
     Next IIMAQ%
     LU$ = LUDAT$
     Modo% = 2
     '
     CONSIMULACION% = 0
     If EXISTE_TABLA("SIMU_ORDEFABR") > 0 Then
        If CantRegistros("SIMU_ORDEFABR", , "NOMESS") > 0 Then
          OPX% = COMALTER%("Opciones de Trabajo\\Listado Cálculo Normal\Listado Considerando Simulación")
          If OPX% > 1 Then
            CONSIMULACION% = XVALO(CONTROL("CARMAQUI", "ULTISIMU"))
            If CONSIMULACION% < 1 Then
               Call COMUNI("Atención!! No Hay Cálculos con Simulación de Fabricación\Se Visualizará Cálculo Normal")
            End If
          End If
        End If
     End If
     ARCHOPERFAP$ = "OPERFAP"
     If CONSIMULACION% < 1 Then
       Call GENEOPERFAP
     Else
       ARCHOPERFAP$ = "SOPERFAP"
     End If
     '
     TITU$ = "INFORMES DE PRODUCCION - GENERALES"
     '
     Screen.MousePointer = 11
     FEX = DIA%
     CORRESP$ = "Corresponde a: " + FECHATEX$(FEX)
     FEX = DIADOS%
     CORRESP$ = CORRESP$ + " - " + FECHATEX$(FEX)
     '
     CAEQSEL% = 0
     For REGI& = 1 To ULTREG&(ARCHOPERFAP$)
         XX$ = REGLEIDO$(ARCHOPERFAP$, REGI&)
         DES% = CVI(Mid$(XX$, 135, 2))
         HAS% = CVI(Mid$(XX$, 139, 2))
         MAQUIY$ = Mid$(XX$, 73, 6)
         If ((DES% >= DIA% And DES% <= DIADOS%) Or (HAS% >= DIA% And HAS% <= DIADOS%)) And InStr(1, MAQSEL$, TRIM$(MAQUIY$)) > 0 Then
            CAPROD# = CVD(Mid$(XX$, 79, 8))
            CAREA# = CVD(Mid$(XX$, 87, 8))
            SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
            Call REPLA(XX$, MKD$(SALDO#), 95, 8)
            Call GRAREG(ARCHOPERFAP$, XX$, REGI&)
            '
            For j% = 1 To CAEQSEL%
               If EQSEL$(j%) = MAQUIY$ Then
                  PUNOP$(j%) = PUNOP$(j%) + MKS$(REGI&)
                  GoTo 120
               End If
            Next j%
            '
            If Modo% = 2 Or MAQUIY$ = MAQUIX$ Then
               CAEQSEL% = CAEQSEL% + 1
               EQSEL$(CAEQSEL%) = MAQUIY$
               PUNOP$(CAEQSEL%) = MKS$(REGI&)
            End If
            '
         End If
120   Next REGI&
      '
      For i% = 1 To CAEQSEL%
         For j% = i% + 1 To CAEQSEL%
            If EQSEL$(i%) > EQSEL$(j%) Then
               EQRES$ = EQSEL$(j%)
               PURES$ = PUNOP$(j%)
               EQSEL$(j%) = EQSEL$(i%)
               PUNOP$(j%) = PUNOP$(i%)
               EQSEL$(i%) = EQRES$
               PUNOP$(i%) = PURES$
            End If
         Next j%
      Next i%
      '
130   FORIPRE$ = TRIM$(CONTROL$("", "PLACARMA"))
      If FORIPRE$ <> "" Then
         GoSub 500
         GoTo 199
      End If
      '
      REGQ& = 0
      Call HEADERS("OPERFAQ", "")
      Screen.MousePointer = 11
      For i% = 1 To CAEQSEL%
        If Len(PUNOP$(i%)) >= 4 Then
          MAQUIX$ = EQSEL$(i%)
          DEQUIX$ = ECOARCH$("PADMAQ", MAQUIX$)
          If REGQ& > 0 Then
            REGQ& = REGQ& + 1
            XX$ = REGBLAN$("OPERFAQ")
            Call GRAREG("OPERFAQ", XX$, REGQ&)
            REGQ& = REGQ& + 1
            Call GRAREG("OPERFAQ", String$(256, "="), REGQ&)
          End If
          REGQ& = REGQ& + 1
          XX$ = REGBLAN$("OPERFAQ")
          Call GRAREG("OPERFAQ", XX$, REGQ&)
          REGQ& = REGQ& + 1
          XX$ = REGBLAN$("OPERFAQ")
          Call REPLA(XX$, "@Equipo: " + TRIM$(MAQUIX$) + " - " + DEQUIX$, 17, 48)
          Call GRAREG("OPERFAQ", XX$, REGQ&)
          REGQ& = REGQ& + 1
          XX$ = REGBLAN$("OPERFAQ")
          Call REPLA(XX$, "@******* ", 17, 48)
          Call GRAREG("OPERFAQ", XX$, REGQ&)
          REGQ& = REGQ& + 1
          XX$ = REGBLAN$("OPERFAQ")
          Call GRAREG("OPERFAQ", XX$, REGQ&)
          '
          For JJ& = 1 To Len(PUNOP$(i%)) Step 4
            REGI& = CVS(Mid$(PUNOP$(i%), JJ&, 4))
            If REGI& > 0 Then
              XX$ = REGLEIDO$(ARCHOPERFAP$, REGI&)
              '
              NORFA$ = TRIM$(Mid$(XX$, 1, 12))
              If Mid$(NORFA$, 1, 2) = "OF" Then
                Referencia$ = SAFETEXT$(VALOBADA("ORDEFABR", "ID_ORFA", "IDSUBOR='" & NORFA$ & "'"))
              Else
                Referencia$ = SAFETEXT$(VALOBADA("SIMU_ORDEFABR", "ID_ORFA", "IDSUBOR='" & NORFA$ & "'"))
              End If
              '
              DES% = CVI(Mid$(XX$, 135, 2))
              HAS% = CVI(Mid$(XX$, 139, 2))
              CAPROD# = CVD(Mid$(XX$, 79, 8))
              CAREA# = CVD(Mid$(XX$, 87, 8))
              SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
              Call REPLA(XX$, MKD$(SALDO#), 95, 8)
              Call REPLA(XX$, Referencia$, 151, 48)
              Call GRAREG(ARCHOPERFAP$, XX$, REGI&)
              '
              'pasar tiempo de minutos a segundos
              For KK1% = 103 To 131 Step 4
                 TIESEGUN = 60 * CVS(Mid$(XX$, KK1%, 4))
                 Call REPLA(XX$, MKS$(TIESEGUN), KK1%, 4)
              Next KK1%
              REGQ& = REGQ& + 1
              Call GRAREG("OPERFAQ", XX$, REGQ&)
            End If
          Next JJ&
        End If
     Next i%
     '
     Screen.MousePointer = 1
     If REGQ& < 1 Then
        Call MENSERR(24, "No hay Operaciones Programadas\para esa Fecha y ese Equipo.")
        GoTo 100
     End If
     '
190  Call SETULTREG("OPERFAQ", REGQ&)
     '
     Call CIERRARCH("OPERFAQ")
     If ULTREG&("OPERFAQ") > 0 Then
        Call HEADERS("OPERFAQ", "")
        Call HEADERS("OPERFAQ", CORRESP$)
        Call CIERRARCH("*.*")
        Call CONECRUN("*OPERFAQ", "Planilla de Carga de Maquina")
     End If
     '
199  If Modo% = 1 Then GoTo 100
     Exit Sub
     '
500  Screen.MousePointer = 11
     HOII% = HOY(HOS$)
     FECHDOC$ = HOS$
     NUMEDOC$ = ""
     TIPODOC$ = "Planilla de Carga de Máquina"
     '
     CODPARAM$(1) = "EQUIPO1"
     CAPARDOC% = 1
     '
     CODTABU1$(1) = "NOP-OPER"
     CODTABU1$(2) = "DES-OPER"
     CODTABU1$(3) = "CANTIDAD"
     CODTABU1$(4) = "TIE-TOTA"
     CODTABU1$(5) = "FECH-VEN"
     CODTABU1$(6) = "TIE-PREP"
     CODTABU1$(9) = "FECH-REC"
     CODTABU1$(7) = "TIE-PROD"
     CODTABU1$(8) = "OBS-OPER"
     CODTABU1$(10) = "TOT-PREP"
     CACOLTAB1% = 10
     '
     For i% = 1 To CAEQSEL%
        If Len(PUNOP$(i%)) >= 4 Then
           MAQUIX$ = EQSEL$(i%)
           '
           For KKU& = 1 To FORSELEQU.LIMATSEL.ListCount
             If TRIM$(Left$(FORSELEQU.LIMATSEL.List(KKU& - 1), 6)) = TRIM$(MAQUIX$) Then
               If FORSELEQU.LIMATSEL.Selected(KKU& - 1) = True Then
                 GoTo 510
               End If
             End If
           Next KKU&
           GoTo 590
           '
510        DEQUIX$ = ECOARCH$("PADMAQ", MAQUIX$)
           DOCPARAM$(1) = DEQUIX$
           '
           CAITAB1% = 0
           For JJ& = 1 To Len(PUNOP$(i%)) Step 4
             REGI& = CVS(Mid$(PUNOP$(i%), JJ&, 4))
             If REGI& > 0 Then
               XX$ = REGLEIDO$(ARCHOPERFAP$, REGI&)
               NOFAX$ = Left$(XX$, 12)
               CIXI% = CVI(Mid$(XX$, 13, 2))
               CAPROD# = CVD(Mid$(XX$, 79, 8))
               CAREA# = CVD(Mid$(XX$, 87, 8))
               SALDOI = CAPROD# - CAREA#: If SALDOI < 0 Then SALDOI = 0
               TPREP = CVS(Mid$(XX$, 103, 4))
               TTOT = CVS(Mid$(XX$, 131, 4))
               FEXDES = CVI(Mid$(XX$, 135, 2))
               HDES% = CVI(Mid$(XX$, 137, 2))
               FEXHAS = CVI(Mid$(XX$, 139, 2))
               HHAS% = CVI(Mid$(XX$, 141, 2))
               '
               CAITAB1% = CAITAB1% + 1
               TETABU1$(1, CAITAB1%) = NOFAX$
               TETABU1$(2, CAITAB1%) = TRIM$(CODEXT$(CIXI%)) + " - " + DESCRIT0$(CIXI%)
               TETABU1$(3, CAITAB1%) = Str$(SALDOI)
                 TT1% = 32767: If TTOT < 32767 Then TT1% = TTOT
               HORASIG$ = TRIM$(HORAMINU$(TT1%))
               If Left$(HORASIG$, 2) = "0:" Then HORASIG$ = "  " + Mid$(HORASIG$, 3)
               TETABU1$(4, CAITAB1%) = HORASIG$
               TETABU1$(5, CAITAB1%) = FECHATEX$(FEXDES)
               TETABU1$(6, CAITAB1%) = HORATEXNOR$(HDES%)
               TETABU1$(9, CAITAB1%) = FECHATEX$(FEXHAS)
                 If TETABU1$(9, CAITAB1%) = TETABU1$(5, CAITAB1%) Then TETABU1$(9, CAITAB1%) = "---id---"
               TETABU1$(7, CAITAB1%) = HORATEXNOR$(HHAS%)
               TETABU1$(8, CAITAB1%) = Mid$(XX$, 225, 6)
                 TT1% = 32767: If TPREP < 32767 Then TT1% = TPREP
               TETABU1$(10, CAITAB1%) = TRIM$(HORAMINU$(TT1%))
               '
             End If
           Next JJ&
           '
           Call PRINTDOC("PLACARMA")   ' planilla de carga
           '
        End If
590  Next i%
     '
     Screen.MousePointer = 1
     Return
     '
End Sub
'

Sub PLACARMA_VIE()        ' PLANILLA DE CARGA DE MAQUINA
     '                   ' DISCONTINUADA AL 7 SET 2006
     Dim MAQUISEL$(512)
     Dim EQSEL$(512), PUNOP$(512)
     '
     FORSELEQU.Show 1
     
     
     CAMAQUI% = 0
     HO% = HOY(HOS$)
     LU$ = LUDAT$
     '
     Call GENEOPERFAP
     '
     TITU$ = "INFORMES DE PRODUCCION - GENERALES"
     '
     FEX = HO%
95   DIADOS% = DIPOST(FEX)
     FEX = DIADOS%
     If DIASEM%(FEX) > 5 Or DIASEM%(FEX) < 1 Then GoTo 95
     '     Call BLANARCH("OPERFAQ")
100  VDEF$ = MKI$(HO%) + "Todos " + MKI$(DIADOS%)
101  DIAEQUI$ = OBJPLA$("DIAEQUI", VDEF$)
     If DIAEQUI$ = "" Then
         REGQ& = 0
         'GoTo 190
         Exit Sub
     End If
     '
     Screen.MousePointer = 11
     DIA% = CVI(Left$(DIAEQUI$, 2))
     FEX = DIA%
110  DIADOS% = CVI(Mid$(DIAEQUI$, 9, 2))
     If DIADOS% < DIA% Then Exit Sub
     MAQUIX$ = Mid$(DIAEQUI$, 3, 6)
     Modo% = 1: If MAQUIX$ = "Todos " Then Modo% = 2
     '
     Screen.MousePointer = 11
     FEX = DIA%
     CORRESP$ = "Corresponde a: " + FECHATEX$(FEX)
     FEX = DIADOS%
     CORRESP$ = CORRESP$ + " - " + FECHATEX$(FEX)
     '
     CAEQSEL% = 0
     For REGI& = 1 To ULTREG&("OPERFAP")
         XX$ = REGLEIDO$("OPERFAP", REGI&)
         DES% = CVI(Mid$(XX$, 135, 2))
         HAS% = CVI(Mid$(XX$, 139, 2))
         If DES% >= DIA% Then
            If HAS% <= DIADOS% Then
               CAPROD# = CVD(Mid$(XX$, 79, 8))
               CAREA# = CVD(Mid$(XX$, 87, 8))
               SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
               Call REPLA(XX$, MKD$(SALDO#), 95, 8)
               Call GRAREG("OPERFAP", XX$, REGI&)
               '
               MAQUIY$ = Mid$(XX$, 73, 6)
               For j% = 1 To CAEQSEL%
                  If EQSEL$(j%) = MAQUIY$ Then
                     PUNOP$(j%) = PUNOP$(j%) + MKS$(REGI&)
                     GoTo 120
                  End If
               Next j%
               '
               If Modo% = 2 Or MAQUIY$ = MAQUIX$ Then
                  CAEQSEL% = CAEQSEL% + 1
                  EQSEL$(CAEQSEL%) = MAQUIY$
                  PUNOP$(CAEQSEL%) = MKS$(REGI&)
               End If
               '
            End If
         End If
120   Next REGI&
      '
      For i% = 1 To CAEQSEL%
         For j% = i% + 1 To CAEQSEL%
            If EQSEL$(i%) > EQSEL$(j%) Then
               EQRES$ = EQSEL$(j%)
               PURES$ = PUNOP$(j%)
               EQSEL$(j%) = EQSEL$(i%)
               PUNOP$(j%) = PUNOP$(i%)
               EQSEL$(i%) = EQRES$
               PUNOP$(i%) = PURES$
            End If
         Next j%
      Next i%
      '
130   FORIPRE$ = TRIM$(CONTROL$("", "PLACARMA"))
      If FORIPRE$ <> "" Then
         GoSub 500
         GoTo 199
      End If
      '
      REGQ& = 0
      Call HEADERS("OPERFAQ", "")
      Screen.MousePointer = 11
      For i% = 1 To CAEQSEL%
        If Len(PUNOP$(i%)) >= 4 Then
          MAQUIX$ = EQSEL$(i%)
          DEQUIX$ = ECOARCH$("PADMAQ", MAQUIX$)
          If REGQ& > 0 Then
            REGQ& = REGQ& + 1
            XX$ = REGBLAN$("OPERFAQ")
            Call GRAREG("OPERFAQ", XX$, REGQ&)
            REGQ& = REGQ& + 1
            Call GRAREG("OPERFAQ", String$(256, "="), REGQ&)
          End If
          REGQ& = REGQ& + 1
          XX$ = REGBLAN$("OPERFAQ")
          Call GRAREG("OPERFAQ", XX$, REGQ&)
          REGQ& = REGQ& + 1
          XX$ = REGBLAN$("OPERFAQ")
          Call REPLA(XX$, "@Equipo: " + TRIM$(MAQUIX$) + " - " + DEQUIX$, 17, 48)
          Call GRAREG("OPERFAQ", XX$, REGQ&)
          REGQ& = REGQ& + 1
          XX$ = REGBLAN$("OPERFAQ")
          Call REPLA(XX$, "@******* ", 17, 48)
          Call GRAREG("OPERFAQ", XX$, REGQ&)
          REGQ& = REGQ& + 1
          XX$ = REGBLAN$("OPERFAQ")
          Call GRAREG("OPERFAQ", XX$, REGQ&)
          '
          For JJ& = 1 To Len(PUNOP$(i%)) Step 4
            REGI& = CVS(Mid$(PUNOP$(i%), JJ&, 4))
            If REGI& > 0 Then
              XX$ = REGLEIDO$("OPERFAP", REGI&)
              DES% = CVI(Mid$(XX$, 135, 2))
              HAS% = CVI(Mid$(XX$, 139, 2))
              CAPROD# = CVD(Mid$(XX$, 79, 8))
              CAREA# = CVD(Mid$(XX$, 87, 8))
              SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
              Call REPLA(XX$, MKD$(SALDO#), 95, 8)
              Call GRAREG("OPERFAP", XX$, REGI&)
              '
              'pasar tiempo de minutos a segundos
              For KK1% = 103 To 131 Step 4
                 TIESEGUN = 60 * CVS(Mid$(XX$, KK1%, 4))
                 Call REPLA(XX$, MKS$(TIESEGUN), KK1%, 4)
              Next KK1%
              REGQ& = REGQ& + 1
              Call GRAREG("OPERFAQ", XX$, REGQ&)
            End If
          Next JJ&
        End If
     Next i%
     '
     Screen.MousePointer = 1
     If REGQ& < 1 Then
        Call MENSERR(24, "No hay Operaciones Programadas\para esa Fecha y ese Equipo.")
        GoTo 100
     End If
     '
190  Call SETULTREG("OPERFAQ", REGQ&)
     '
     Call CIERRARCH("OPERFAQ")
     If ULTREG&("OPERFAQ") > 0 Then
        Call HEADERS("OPERFAQ", "")
        Call HEADERS("OPERFAQ", CORRESP$)
        Call CIERRARCH("*.*")
        Call CONECRUN("*OPERFAQ", "Planilla de Carga de Maquina")
     End If
     '
199  If Modo% = 1 Then GoTo 100
     Exit Sub
     '
500  Screen.MousePointer = 11
     HOII% = HOY(HOS$)
     FECHDOC$ = HOS$
     NUMEDOC$ = ""
     TIPODOC$ = "Planilla de Carga de Máquina"
     '
     CODPARAM$(1) = "EQUIPO1"
     CAPARDOC% = 1
     '
     CODTABU1$(1) = "NOP-OPER"
     CODTABU1$(2) = "DES-OPER"
     CODTABU1$(3) = "CANTIDAD"
     CODTABU1$(4) = "TIE-TOTA"
     CODTABU1$(5) = "FECH-VEN"
     CODTABU1$(6) = "TIE-PREP"
     CODTABU1$(9) = "FECH-REC"
     CODTABU1$(7) = "TIE-PROD"
     CODTABU1$(8) = "OBS-OPER"
     CODTABU1$(10) = "TOT-PREP"
     CACOLTAB1% = 10
     '
     For i% = 1 To CAEQSEL%
        If Len(PUNOP$(i%)) >= 4 Then
           MAQUIX$ = EQSEL$(i%)
           DEQUIX$ = ECOARCH$("PADMAQ", MAQUIX$)
           DOCPARAM$(1) = DEQUIX$
           '
           CAITAB1% = 0
           For JJ& = 1 To Len(PUNOP$(i%)) Step 4
             REGI& = CVS(Mid$(PUNOP$(i%), JJ&, 4))
             If REGI& > 0 Then
               XX$ = REGLEIDO$("OPERFAP", REGI&)
               NOFAX$ = Left$(XX$, 12)
               CIXI% = CVI(Mid$(XX$, 13, 2))
               CAPROD# = CVD(Mid$(XX$, 79, 8))
               CAREA# = CVD(Mid$(XX$, 87, 8))
               SALDOI = CAPROD# - CAREA#: If SALDOI < 0 Then SALDOI = 0
               TPREP = CVS(Mid$(XX$, 103, 4))
               TTOT = CVS(Mid$(XX$, 131, 4))
               FEXDES = CVI(Mid$(XX$, 135, 2))
               HDES% = CVI(Mid$(XX$, 137, 2))
               FEXHAS = CVI(Mid$(XX$, 139, 2))
               HHAS% = CVI(Mid$(XX$, 141, 2))
               '
               CAITAB1% = CAITAB1% + 1
               TETABU1$(1, CAITAB1%) = NOFAX$
               TETABU1$(2, CAITAB1%) = TRIM$(CODEXT$(CIXI%)) + " - " + DESCRIT0$(CIXI%)
               TETABU1$(3, CAITAB1%) = Str$(SALDOI)
                 TT1% = 32767: If TTOT < 32767 Then TT1% = TTOT
               HORASIG$ = TRIM$(HORAMINU$(TT1%))
               If Left$(HORASIG$, 2) = "0:" Then HORASIG$ = "  " + Mid$(HORASIG$, 3)
               TETABU1$(4, CAITAB1%) = HORASIG$
               TETABU1$(5, CAITAB1%) = FECHATEX$(FEXDES)
               TETABU1$(6, CAITAB1%) = HORATEXNOR$(HDES%)
               TETABU1$(9, CAITAB1%) = FECHATEX$(FEXHAS)
                 If TETABU1$(9, CAITAB1%) = TETABU1$(5, CAITAB1%) Then TETABU1$(9, CAITAB1%) = "---id---"
               TETABU1$(7, CAITAB1%) = HORATEXNOR$(HHAS%)
               TETABU1$(8, CAITAB1%) = Mid$(XX$, 225, 6)
                 TT1% = 32767: If TPREP < 32767 Then TT1% = TPREP
               TETABU1$(10, CAITAB1%) = TRIM$(HORAMINU$(TT1%))
               '
             End If
           Next JJ&
           '
           Call PRINTDOC("PLACARMA")   ' planilla de carga
           '
        End If
     Next i%
     '
     Screen.MousePointer = 1
     Return
     '
End Sub
'



