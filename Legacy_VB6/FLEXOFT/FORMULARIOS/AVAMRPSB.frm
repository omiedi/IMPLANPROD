VERSION 5.00
Begin VB.Form AVAMRPSB 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFF0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cálculo Necesidad de Componentes según Programa de Entregas"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5940
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
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
      Height          =   740
      Left            =   4920
      Picture         =   "AVAMRPSB.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Abandona Proceso de Cálculo"
      Top             =   315
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFF0FF&
      Caption         =   "Ultimo Re-Calculo Stocks I.Temp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   4515
      Begin VB.CommandButton Command3 
         Caption         =   "Re-Calc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3240
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   400
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   820
      Left            =   4920
      Picture         =   "AVAMRPSB.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Inicia Proceso de Cálculo"
      Top             =   1270
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4935
      ScaleHeight     =   75
      ScaleWidth      =   810
      TabIndex        =   9
      Top             =   3885
      Width           =   870
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   820
      Left            =   4920
      Picture         =   "AVAMRPSB.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Abandona Proceso de Cálculo"
      Top             =   2160
      Width           =   855
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFF0FF&
      Caption         =   "Avance de Proceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3585
      Left            =   210
      TabIndex        =   7
      Top             =   1200
      Width           =   4515
      Begin VB.CheckBox Check8 
         BackColor       =   &H00FFF0FF&
         Caption         =   "Generando Tabla Anual Necesidades"
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
         Left            =   240
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3075
         Width           =   4250
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00FFF0FF&
         Caption         =   "Cálculo de Necesidades  por Tipificadas"
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
         Left            =   240
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2760
         Width           =   4250
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00FFF0FF&
         Caption         =   "Armando Tabla Programa de Procuración"
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
         Left            =   240
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2220
         Width           =   4250
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00FFF0FF&
         Caption         =   "Neteando Programa por Existencias en Stock"
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
         Left            =   240
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1905
         Width           =   4250
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFF0FF&
         Caption         =   "Generación Tablas Programa Bruto"
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
         Left            =   240
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   645
         Width           =   3405
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFF0FF&
         Caption         =   "Neteo Programas por Entregas Realizadas"
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
         Left            =   240
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   960
         Width           =   4250
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFF0FF&
         Caption         =   "Generando Programa Neto de Entregas"
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
         Left            =   240
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1275
         Width           =   4250
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFF0FF&
         Caption         =   "Armando Tabla Programa Neto Entregas"
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
         Left            =   240
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1590
         Width           =   4250
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   3840
         Picture         =   "AVAMRPSB.frx":075E
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4365
      Picture         =   "AVAMRPSB.frx":0A68
      Top             =   4250
      Width           =   2010
   End
End
Attribute VB_Name = "AVAMRPSB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CODIPRO%(16384), PROGENT$(16384)
Dim FEPLA%(24), DEPERI$(24)
Dim TIENEPROGEN%(32767)
Dim HAYPROPIE%(32767)

Private Sub command1_Click()
   Command1.Enabled = False
   If COMALTER%("Realmente Desea Cancelar el Cálculo\\No, Continuar\Si, Cancelar") = 2 Then
     Call CIERRARCH("*.*")
     Command1.Enabled = True
     Hide
   End If
   Command1.Enabled = True
End Sub

'
Private Sub Command2_Click()
   Command2.Enabled = False
      DoEvents
   HOII% = HOY(HOS$)
   Timer1.Enabled = True
   '
' 0.- Arma tabla de meses
   '
   For KU% = 0 To 24
     TTYY$ = TRIM$(Mid$(APROGMRP.Label5, 9 * KU% + 1, 9))
     If KU% > 0 Then
         TTYY$ = MESANO$(TTYY$, DES%, HAS%)
         FEPLA%(KU%) = HAS%
         DEPERI$(KU%) = MESANO$(TRIM$(Mid$(APROGMRP.Label5, 9 * KU% + 1, 9)), DES%, HAS%)
       Else
         FEPLA%(KU%) = DIANTE(FECHANUM("01/01/" + Mid$(HOS$, 7, 2)))
         DEPERI$(KU%) = TRIM$(Mid$(APROGMRP.Label5, 9 * KU% + 1, 9))
     End If
   Next KU%
   '
' GoTo 800
   '
' 1.- Transforma programa de entregas bruto en tablas anuales
   '
   Check1.Value = 2
   For KKI% = 0 To NUMAS%
     PROGENT$(KKI%) = String$(100, 0)
   Next KKI%
   '
   KK& = 0
   For IPF& = 1 To ULTREG("PROFABRI")
     IPX% = IPF&
     RFF$ = RECFILT$("PROGENT", 10, MKI$(IPX%))
     FIN& = Len(RFF$)
     '
     For U& = 1 To FIN& Step 4
       Call TRACE(20, U&, FIN&)
       I& = CVS(Mid$(RFF$, U&, 4))
       P00$ = REGLEIDO$("PROGENT", I&)
       CI% = CVI(Left$(P00$, 2))
       CANPRO = CVS(Mid$(P00$, 7, 4))
       '
       If CANPRO > 0.05 Then
         If CI% > 0 Then
           If CI% <= NUMAS% Then
             FEXI% = CVI(Mid$(P00$, 3, 2))
             If FEXI% <= FEPLA%(0) Then
                COLUX% = 0
                GoTo 20
              Else
                For KKJ% = 24 To 1 Step -1
                  If FEXI% >= FEPLA%(KKJ%) Then
                    COLUX% = KKJ%
                    GoTo 20
                  End If
                Next KKJ%
             End If
             GoTo 39
             '
20           For KKI% = 1 To CACODIP%
               If CODIPRO%(KKI%) = CI% Then GoTo 30
             Next KKI%
             CACODIP% = CACODIP% + 1
             CODIPRO%(CACODIP%) = CI%
             KKI% = CACODIP%
             '
30           If COLUX% >= 0 Then
               If COLUX% <= 24 Then
                 CANPRO = CANPRO + CVS(Mid$(PROGENT$(KKI%), 1 + 4 * COLUX%, 4))
                 Mid$(PROGENT$(KKI%), 1 + 4 * COLUX%, 4) = MKS$(CANPRO)
               End If
             End If
           End If
         End If
       End If
       '
39   Next U&
     '
     REBLA$ = REGBLAN$("PRGBRUTA")
     REBLB$ = REGBLAN$("PRGBRUTB")
     '
     For KKI% = 1 To CACODIP%
       '
       CIKKI% = CODIPRO%(KKI%)
       TTXA$ = MKI$(CIKKI%) + MKI$(IPX%)
       TTXB$ = MKI$(CIKKI%) + MKI$(IPX%)
       '
       SUMA = 0
       For JJ& = 0 To 12
         XX$ = Mid$(PROGENT$(KKI%), 4 * JJ& + 1, 4)
         SUMA = SUMA + CVS(XX$)
         TTXA$ = TTXA$ + XX$
       Next JJ&
       TTXA$ = TTXA$ + MKS$(SUMA) + String$(64, 0)
       '
       TTXB$ = TTXB$ + MKS$(SUMA)
       For JJ& = 13 To 24
         XX$ = Mid$(PROGENT$(KKI%), 4 * JJ& + 1, 4)
         SUMA = SUMA + CVS(XX$)
         TTXB$ = TTXB$ + XX$
       Next JJ&
       TTXB$ = TTXB$ + MKS$(SUMA) + String$(64, 0)
       '
       KK& = KK& + 1
       Call GRAREG("PRGBRUTA", TTXA$, KK&)
       Call GRAREG("PRGBRUTB", TTXB$, KK&)
     Next KKI%
     '
     For KKI% = 0 To CACODIP%
       CODIPRO%(KKI%) = 0
       PROGENT$(KKI%) = String$(100, 0)
     Next KKI%
     CACODIP% = 0
     '
   Next IPF&
   '
   Call SETULTREG("PRGBRUTA", KK&)
   Call SETULTREG("PRGBRUTB", KK&)
   Check1.Value = 1
   '
   '
   '
' 2.- Netea Programa con entregas efectuadas
   '
' 2.1.- Marca los clientes que tienen programa de entregas
   Check2.Value = 2
   For U& = 1 To 32766
     TIENEPROGEN%(U&) = 0
     HAYPROPIE%(U&) = 0
   Next U&
   '
   Screen.MousePointer = 11
   For U& = 1 To ULTREG("PROFABRI")
     XX$ = REGLEIDO$("PROFABRI", U&)
     NCLIE% = CVI(Mid$(XX$, 53, 2))
     If NCLIE% > 0 Then
       If NCLIE% <= 32767 Then
         TIENEPROGEN%(NCLIE%) = 1
       End If
     End If
   Next U&
   '
   REBLA$ = REGBLAN$("PRONET")
   HOII% = HOY(HOS$)
   FECHIN% = FECHANUM("01/01/" + Mid$(HOS$, 7, 2))
   JJ& = 0
   REBLA$ = REGBLAN$("PRONET")
   '
   FIN& = ULTREG("PROGENT")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("PROGENT", U&)
     CI% = CVI(Left$(XX$, 2))
     CAPROG = CVS(Mid$(XX$, 7, 4))
     If CI% > 0 Then
       If CAPROG > 0.05 Then
         HAYPROPIE%(CI%) = 1
         NPRG% = CVI(Mid$(XX$, 23, 2))
         FF% = CVI(Mid$(XX$, 3, 2))
         ZZ$ = REBLA$
         Call REPLA(ZZ$, MKI$(NPRG%), 1, 2)
         Call REPLA(ZZ$, MKI$(CI%), 3, 2)
         Call REPLA(ZZ$, MKI$(FF%), 5, 2)
         ' EN CAMPO COMPROBANTE PONER TEXTO MES Y AÑO
         For KKU% = 0 To 24
           If FEPLA%(KKU%) >= FF% Then
             Call REPLA(ZZ$, DEPERI$(KKU%), 7, 12)
             Exit For
           End If
         Next KKU%
         ' COMPLETAR CLIENTE
         NCLIE% = CVI(Mid$(XX$, 5, 2))
         If NCLIE% < 1 Then
             RACLI$ = "Clientes Varios"
           Else
             RACLI$ = RASOCLI$(NCLIE%)
         End If
         Call REPLA(ZZ$, RACLI$, 19, 30)
         Call REPLA(ZZ$, MKS$(CAPROG), 49, 4)
         JJ& = JJ& + 1
         Call GRAREG("PRONET", ZZ$, JJ&)
       End If
     End If
   Next U&
   '
   For U& = 1 To ULTREG("PROFABRI")
     NPRG% = U&
     XX$ = REGLEIDO$("PROFABRI", U&)
     NCLIE% = CVI(Mid$(XX$, 53, 2))
     '
     FIN& = ULTREG&("MOVISTO")
     For WW& = 1 To FIN&
       Call TRACE(20, WW&, FIN&)
       XX$ = REGLEIDO$("MOVISTO", WW&)
       CMOV$ = TRIM$(UCase$(Mid$(XX$, 21, 2)))
       DOCUPRI$ = TRIM$(Mid$(XX$, 23, 10))
       NCLI1% = CVI(Mid$(XX$, 84, 2))
       If NCLI1% > 0 Then ' solo movimientos de clientes
         If NCLI1% = NCLIE% Or (NCLIE% = 0 And TIENEPROGEN%(NCLI1%) = 0) Then
           FEXI% = CVI(Left$(XX$, 2))
           If FEXI% >= FECHIN% Then
             CI% = CVI(Mid$(XX$, 3, 2))
             If HAYPROPIE(CI%) > 0 Then
               REPFI$ = REBLA$
               CAEN = CVD(Mid$(XX$, 96, 8)) - CVD(Mid$(XX$, 88, 8))
               Call REPLA(REPFI$, MKI$(NPRG%), 1, 2)
               Call REPLA(REPFI$, MKI$(CI%), 3, 2)
               Call REPLA(REPFI$, MKI$(FEXI%), 5, 2)
               Call REPLA(REPFI$, DOCUPRI$, 7, 12)
               Call REPLA(REPFI$, RASOCLI$(NCLI1%), 19, 30)
               If CAEN > 0 Then
                   Call REPLA(REPFI$, MKS$(CAEN), 53, 4)
                 Else
                   Call REPLA(REPFI$, MKS$(Abs(CAEN)), 49, 4)
               End If
               JJ& = JJ& + 1
               Call GRAREG("PRONET", REPFI$, JJ&)
             End If
           End If
         End If
       End If
     Next WW&
   Next U&
   '
   Call SETULTREG("PRONET", JJ&)
   Call CIERRARCH("PRONET")
   Check2.Value = 1
   '
' 3.- ARMANDO ARCHIVO DE PROGRAMA NETO DE ENTREGAS EFECTUADAS
   '
   Check3.Value = 2
   Screen.MousePointer = 11
   Call COPYSTRU("PROGENET", "INTERPRO")
   JK0& = 0
   For NPF& = 1 To ULTREG&("PROFABRI")
      NPRG% = NPF&
      Call FILTERFILE("PRONET", "INTERPRO", 1, MKI$(NPRG%))
      '
      Call FILESORT("INTERPRO", "", 1, 4)
      FIN& = ULTREG&("INTERPRO")
      For KKL& = 1 To FIN&
        XX$ = REGLEIDO$("INTERPRO", KKL&)
        Call GRAREG("!PROGENET", XX$, KKL&)
      Next KKL&
      Call SETULTREG("!PROGENET", FIN&)
      '
      CIAN% = 0: ACUMU = 0: RINICO& = 0: UENTRA& = 0: UENTRB& = 0
      FIN& = ULTREG&("!PROGENET")
      For U& = 1 To FIN&
        XX$ = REGLEIDO$("!PROGENET", U&)
        CI% = CVI(Mid$(XX$, 3, 2))
        If CI% <> CIAN% Then
          ACUMU = 0
          If CIAN% > 0 Then
            GoSub 150
            RINICO& = U&
          End If
          CIAN% = CI%
          XX$ = REGLEIDO$("!PROGENET", U&)
        End If
        If CVS(Mid$(XX$, 53, 4)) > 0 Then UENTRA& = U&
        ACUMU = ACUMU + CVS(Mid$(XX$, 49, 4)) - CVS(Mid$(XX$, 53, 4))
        Call REPLA(XX$, MKS$(ACUMU), 57, 4)
        Call GRAREG("!PROGENET", XX$, U&)
      Next U&
      ACUMU = 0
      GoSub 150
      CIAN% = 0
      Call CIERRARCH("!PROGENET")
      '
      FIN& = ULTREG&("!PROGENET")
      For KKL& = 1 To FIN&
        XX$ = REGLEIDO$("!PROGENET", KKL&)
        JK0& = JK0& + 1
        Call GRAREG("PROGENET", XX$, JK0&)
      Next KKL&
   Next NPF&
   '
   Call SETULTREG("PROGENET", JK0&)
   Call CIERRARCH("PROGENET")
   Call BAJALDISCO
   Check3.Value = 1
   '
' 4.- GENERA PROGRAMAS NETOS EN FORMA DE TABLAS
   '
   Check4.Value = 2
   KK& = 0
   For IPF& = 1 To ULTREG("PROFABRI")
     IPX% = IPF&
     '
     For KKI% = 0 To 16384
       PROGENT$(KKI%) = String$(100, 0)
       CODIPRO%(KKI%) = 0
     Next KKI%
     '
     RFF$ = RECFILT$("PROGENET", 1, MKI$(IPX%))
     FIN& = Len(RFF$)
     '
     For U& = 1 To FIN& Step 4
       Call TRACE(20, U&, FIN&)
       I& = CVS(Mid$(RFF$, U&, 4))
       P00$ = REGLEIDO$("PROGENET", I&)
       CI% = CVI(Mid$(P00$, 3, 2))
       CANPRO = CVS(Mid$(P00$, 57, 4))
       '
       If CANPRO > 0.05 Then
         If CI% > 0 Then
           If CI% <= NUMAS% Then
             FEXI% = CVI(Mid$(P00$, 5, 2))
             If FEXI% <= FEPLA%(0) Then
                COLUX% = 0
                GoTo 120
              Else
                For KKJ% = 24 To 1 Step -1
                  If FEXI% >= FEPLA%(KKJ%) Then
                    COLUX% = KKJ%
                    GoTo 120
                  End If
                Next KKJ%
             End If
             GoTo 139
             '
120          For KKI% = 1 To CACODIP%
               If CODIPRO%(KKI%) = CI% Then GoTo 130
             Next KKI%
             CACODIP% = CACODIP% + 1
             CODIPRO%(CACODIP%) = CI%
             KKI% = CACODIP%
             '
130          If COLUX% >= 0 Then
               If COLUX% <= 24 Then
                 CANPRO = CANPRO + CVS(Mid$(PROGENT$(KKI%), 1 + 4 * COLUX%, 4))
                 Mid$(PROGENT$(KKI%), 1 + 4 * COLUX%, 4) = MKS$(CANPRO)
               End If
             End If
           End If
         End If
       End If
       '
139  Next U&
     '
     REBLA$ = REGBLAN$("PROGNETA")
     REBLB$ = REGBLAN$("PROGNETB")
     '
     For KKI% = 1 To CACODIP%
       '
       CIKKI% = CODIPRO%(KKI%)
       TTXA$ = MKI$(CIKKI%) + MKI$(IPX%)
       '
       SUMA = 0
       For JJ& = 0 To 12
         XX$ = Mid$(PROGENT$(KKI%), 4 * JJ& + 1, 4)
         SUMA = SUMA + CVS(XX$)
         TTXA$ = TTXA$ + XX$
       Next JJ&
       TTXA$ = TTXA$ + MKS$(SUMA) + String$(64, 0)
       '
       TTXB$ = MKI$(CIKKI%) + MKI$(IPX%) + MKS$(SUMA)
       For JJ& = 13 To 24
         XX$ = Mid$(PROGENT$(KKI%), 4 * JJ& + 1, 4)
         SUMA = SUMA + CVS(XX$)
         TTXB$ = TTXB$ + XX$
       Next JJ&
       TTXB$ = TTXB$ + MKS$(SUMA) + String$(128, 0)
       '
       KK& = KK& + 1
       Call GRAREG("PROGNETA", TTXA$, KK&)
       Call GRAREG("PROGNETB", TTXB$, KK&)
     Next KKI%
   Next IPF&
   '
   Call SETULTREG("PROGNETA", KK&)
   Call SETULTREG("PROGNETB", KK&)
   Check4.Value = 1
   '
   '
   '
' 5.- ARMANDO ARCHIVO DE PROGRAMA NETO DE ENTREGAS Y STOCK
   '
   Check5.Value = 2
   Screen.MousePointer = 11
   JK0& = 0
   HOII% = HOY(HOS$)
   FF00% = DIANTE(FECHANUM("01/01/" + Mid$(HOS$, 7, 2)))
   '
   For NPF& = 1 To ULTREG&("PROFABRI")
      NPRG% = NPF&
      For KKU% = 1 To NUMAS%
        HAYPROPIE%(KKU%) = 0
      Next KKU%
      '
      JJ& = 0
      Call FILTERFILE("PROGENET", "INTERPRO", 1, MKI$(NPRG%))
      For U& = 1 To ULTREG("INTERPRO")
        XX$ = REGLEIDO$("INTERPRO", U&)
        CI% = CVI(Mid$(XX$, 3, 2))
        If CVS(Mid$(XX$, 57, 4)) >= 0.05 Then
          CASAEN = CVS(Mid$(XX$, 57, 4))
          Call REPLA(XX$, MKS$(CASAEN), 49, 4)
          Call REPLA(XX$, MKS$(0), 53, 4)
          Call REPLA(XX$, MKS$(0), 57, 4)
          JJ& = JJ& + 1
          Call GRAREG("INTERPRO", XX$, JJ&)
          HAYPROPIE%(CI%) = 1
        End If
      Next U&
      '
      For KKI% = 1 To NUMAS%
        If HAYPROPIE%(KKI%) > 0 Then
          ZZ$ = REGBLAN$("INTERPRO")
          Call REPLA(ZZ$, MKI$(NPRG%), 1, 2)
          Call REPLA(ZZ$, MKI$(KKI%), 3, 2)
          Call REPLA(ZZ$, MKI$(FF00%), 5, 2)
          Call REPLA(ZZ$, "Stk." + HOS$, 7, 12)
          Call REPLA(ZZ$, MKS$(STOCKACT(KKI%)), 53, 4)
          JJ& = JJ& + 1
          Call GRAREG("INTERPRO", ZZ$, JJ&)
        End If
      Next KKI%
      Call SETULTREG("INTERPRO", JJ&)
      '
      Call FILESORT("INTERPRO", "", 1, 4)
      FIN& = ULTREG&("INTERPRO")
      For KKL& = 1 To FIN&
        XX$ = REGLEIDO$("INTERPRO", KKL&)
        Call GRAREG("!PROGENET", XX$, KKL&)
      Next KKL&
      Call SETULTREG("!PROGENET", FIN&)
      '
      CIAN% = 0: ACUMU = 0: RINICO& = 0: UENTRA& = 0: UENTRB& = 0
      FIN& = ULTREG&("!PROGENET")
      For U& = 1 To FIN&
        XX$ = REGLEIDO$("!PROGENET", U&)
        CI% = CVI(Mid$(XX$, 3, 2))
        If CI% <> CIAN% Then
          ACUMU = 0
          If CIAN% > 0 Then
            GoSub 150
            RINICO& = U&
          End If
          CIAN% = CI%
          XX$ = REGLEIDO$("!PROGENET", U&)
        End If
        If CVS(Mid$(XX$, 53, 4)) > 0 Then UENTRA& = U&
        ACUMU = ACUMU + CVS(Mid$(XX$, 49, 4)) - CVS(Mid$(XX$, 53, 4))
        Call REPLA(XX$, MKS$(ACUMU), 57, 4)
        Call GRAREG("!PROGENET", XX$, U&)
      Next U&
      ACUMU = 0
      GoSub 150
      CIAN% = 0
      Call CIERRARCH("!PROGENET")
      '
      FIN& = ULTREG&("!PROGENET")
      For KKL& = 1 To FIN&
        XX$ = REGLEIDO$("!PROGENET", KKL&)
        JK0& = JK0& + 1
        Call GRAREG("PRGNETST", XX$, JK0&)
      Next KKL&
      '
   Next NPF&
   Call SETULTREG("PRGNETST", JK0&)
   Call CIERRARCH("PRGNETST")
   Call BAJALDISCO
   '
   Screen.MousePointer = 1
   Check5.Value = 1
   '
   '
' 6.- GENERA PROGRAMAS NETOS DE STOCK EN FORMA DE TABLAS
   '
   Check6.Value = 2
   KK& = 0
   For IPF& = 1 To ULTREG("PROFABRI")
     IPX% = IPF&
     '
     For KKI% = 0 To 16384
       PROGENT$(KKI%) = String$(100, 0)
       CODIPRO%(KKI%) = 0
     Next KKI%
     CACODIP% = 0
     '
     RFF$ = RECFILT$("PRGNETST", 1, MKI$(IPX%))
     FIN& = Len(RFF$)
     '
     For U& = 1 To FIN& Step 4
       Call TRACE(20, U&, FIN&)
       I& = CVS(Mid$(RFF$, U&, 4))
       P00$ = REGLEIDO$("PRGNETST", I&)
       CI% = CVI(Mid$(P00$, 3, 2))
       CANPRO = CVS(Mid$(P00$, 57, 4))
       '
       If CANPRO > 0.05 Then
         If CI% > 0 Then
           If CI% <= NUMAS% Then
             FEXI% = CVI(Mid$(P00$, 5, 2))
             If FEXI% <= FEPLA%(0) Then
                COLUX% = 0
                GoTo 220
              Else
                For KKJ% = 24 To 1 Step -1
                  If FEXI% >= FEPLA%(KKJ%) Then
                    COLUX% = KKJ%
                    GoTo 220
                  End If
                Next KKJ%
             End If
             GoTo 239
             '
220          For KKI% = 1 To CACODIP%
               If CODIPRO%(KKI%) = CI% Then GoTo 230
             Next KKI%
             CACODIP% = CACODIP% + 1
             CODIPRO%(CACODIP%) = CI%
             KKI% = CACODIP%
             '
230          If COLUX% >= 0 Then
               If COLUX% <= 24 Then
                 CANPRO = CANPRO + CVS(Mid$(PROGENT$(KKI%), 1 + 4 * COLUX%, 4))
                 Mid$(PROGENT$(KKI%), 1 + 4 * COLUX%, 4) = MKS$(CANPRO)
               End If
             End If
           End If
         End If
       End If
       '
239  Next U&
     '
     REBLA$ = REGBLAN$("PROGSTKA")
     REBLB$ = REGBLAN$("PROGSTKB")
     '
     For KKI% = 1 To CACODIP%
       '
       CIKKI% = CODIPRO%(KKI%)
       TTXA$ = MKI$(CIKKI%) + MKI$(IPX%)
       '
       SUMA = 0
       For JJ& = 0 To 12
         XX$ = Mid$(PROGENT$(KKI%), 4 * JJ& + 1, 4)
         SUMA = SUMA + CVS(XX$)
         TTXA$ = TTXA$ + XX$
       Next JJ&
       TTXA$ = TTXA$ + MKS$(SUMA) + String$(64, 0)
       '
       TTXB$ = MKI$(CIKKI%) + MKI$(IPX%) + MKS$(SUMA)
       For JJ& = 13 To 24
         XX$ = Mid$(PROGENT$(KKI%), 4 * JJ& + 1, 4)
         SUMA = SUMA + CVS(XX$)
         TTXB$ = TTXB$ + XX$
       Next JJ&
       TTXB$ = TTXB$ + MKS$(SUMA) + String$(128, 0)
       '
       KK& = KK& + 1
       Call GRAREG("PROGSTKA", TTXA$, KK&)
       Call GRAREG("PROGSTKB", TTXB$, KK&)
     Next KKI%
   Next IPF&
   '
   Call SETULTREG("PROGSTKA", KK&)
   Call SETULTREG("PROGSTKB", KK&)
   Check6.Value = 1
   '
   '
' 7.- GENERA NECESIDADES BRUTAS DE MATERIALES POR TIPIFICADAS
   '
   
800 Check7.Value = 2
   Screen.MousePointer = 11
   FIN& = ULTREG("PROGENET")
   NMS% = NUMAS%
   REBLA$ = REGBLAN$("NECALPRO")
   JJ& = 0
   '
   For I& = 1 To FIN&
     Call TRACE(20, I&, FIN&)
     P00$ = REGLEIDO$("PROGENET", I&)
     NPROG% = CVI(Left$(P00$, 2))
     CIXI% = CVI(Mid$(P00$, 3, 2))
     CANPRO = CVS(Mid$(P00$, 57, 4))
     '
     If CANPRO > 0.05 Then
       If CIXI% > 0 Then
         If CIXI% <= NUMAS% Then
           FEXI% = CVI(Mid$(P00$, 5, 2))
           '
           YY0$ = REBLA$
           Call REPLA(YY0$, MKI$(CIXI%), 33, 2)
           Call REPLA(YY0$, MKI$(FEXI%), 35, 2)
Call REPLA(YY0$, MKI$(NCLIE%), 37, 2)
           Call REPLA(YY0$, MKI$(NPROG%), 39, 2)
           Call REPLA(YY0$, MKS$(CANPRO), 41, 4)
           '
           RFF$ = RECFILT$("TIPIFIC", 1, MKI$(CIXI%))
           If Len(RFF$) > 0 Then
             For UK& = 1 To Len(RFF$) Step 4
               RETI& = CVS(Mid$(RFF$, UK&, 4))
               ZZ$ = REGLEIDO$("TIPIFIC", RETI&)
               CICO% = CVI(Mid$(ZZ$, 49, 2))
               If CICO% > 0 Then
                 If CICO% <= NMS% Then
                   UNICO$ = Mid$(ZZ$, 51, 4)
                   If TRIM$(UNICO$) = "" Then
                     UNICO$ = UNIMED$(CICO%)
                   End If
                   USNETO# = CVD(Mid$(ZZ$, 55, 8))
                   PORMER = CVS(Mid$(ZZ$, 71, 4))
                   PORPER = CVS(Mid$(ZZ$, 83, 4))
                   USBRUT# = CVD(Mid$(ZZ$, 87, 8))
                   '
                   YY$ = YY0$
                   ILOTEC& = ILOTEC& + 1
                   Call REPLA(YY$, MKS$(ILOTEC&), 65, 4)
                   Call REPLA(YY$, MKI$(CICO%), 69, 2)
                   Call REPLA(YY$, UNICO$, 71, 4)
                   Call REPLA(YY$, MKD$(USNETO#), 75, 8)
                   Call REPLA(YY$, MKS$(PORMER), 83, 4)
                   Call REPLA(YY$, MKS$(PORPER), 87, 4)
                   Call REPLA(YY$, MKD$(USBRUT#), 91, 8)
                   Call REPLA(YY$, MKD$(USBRUT# * CANPRO), 99, 8)
                   '
                   Call REPLA(YY$, MKI$(CICO%), 1, 2)
                   Call REPLA(YY$, UNICO$, 3, 4)
                   Call REPLA(YY$, MKI$(FEXI%), 7, 2)
                   '
                   JJ& = JJ& + 1
                   Call GRAREG("NECALPRO", YY$, JJ&)
                   '
                 End If
               End If
             Next UK&
           End If
         End If
       End If
     End If
   Next I&
   '
   Call SETULTREG("NECALPRO", JJ&)
   Call FILESORT("NECALPRO", "", 1, 7)
   '
   JJ1& = 0: SUMAD# = 0
   REBLA = REGBLAN$("NECALPR3")
   FIN& = ULTREG&("NECALPRO")
   For U& = 1 To FIN&
     XX$ = REGLEIDO$("NECALPRO", U&)
     SUMAD# = SUMAD# + CVD(Mid$(XX$, 99, 8))
     YY$ = REGLEIDO$("NECALPRO", U& + 1)
     If Left$(YY$, 8) <> Left$(XX$, 8) Then
        ZZ$ = REBLA$
        Call REPLA(ZZ$, Left$(XX$, 8), 1, 8)
        Call REPLA(ZZ$, MKD$(SUMAD#), 33, 8)
        JJ1& = JJ1& + 1
        Call GRAREG("NECALPR3", ZZ$, JJ1&)
        SUMAD# = 0
     End If
   Next U&
   Call SETULTREG("NECALPR3", JJ1&)
   '
   Check7.Value = 1
   '
' 8.- GENERA NECESIDADES BRUTAS DE STOCK EN FORMA DE TABLAS
   '
   Check8.Value = 2
   For KKI% = 0 To 16384
     PROGENT$(KKI%) = String$(100, 0)
     CODIPRO%(KKI%) = 0
   Next KKI%
   CACODIP% = 0
   '
   KK& = 0
   '
   FIN& = ULTREG("NECALPR3")
   For U& = 1 To FIN&
     P00$ = REGLEIDO$("NECALPR3", U&)
     CI% = CVI(Left$(P00$, 2))
     CANPRO = CVD(Mid$(P00$, 33, 8))
     '
     If CANPRO > 0.0005 Then
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           FEXI% = CVI(Mid$(P00$, 7, 2))
           If FEXI% <= FEPLA%(0) Then
              COLUX% = 0
              GoTo 320
            Else
              For KKJ% = 24 To 1 Step -1
                If FEXI% >= FEPLA%(KKJ%) Then
                  COLUX% = KKJ%
                  GoTo 320
                End If
              Next KKJ%
           End If
           GoTo 339
           '
320        For KKI% = 1 To CACODIP%
             If CODIPRO%(KKI%) = CI% Then GoTo 330
           Next KKI%
           CACODIP% = CACODIP% + 1
           CODIPRO%(CACODIP%) = CI%
           KKI% = CACODIP%
           '
330        If COLUX% >= 0 Then
             If COLUX% <= 24 Then
               CANPRO = CANPRO + CVS(Mid$(PROGENT$(KKI%), 1 + 4 * COLUX%, 4))
               Mid$(PROGENT$(KKI%), 1 + 4 * COLUX%, 4) = MKS$(CANPRO)
             End If
           End If
         End If
       End If
     End If
     '
339 Next U&
   '
   REBLA$ = REGBLAN$("NETIBRA")
   REBLB$ = REGBLAN$("NETIBRB")
   KK& = 0
   '
   For KKI% = 1 To CACODIP%
     '
     CIKKI% = CODIPRO%(KKI%)
     TTXA$ = MKI$(CIKKI%) + MKI$(PROVHABI%(CIKKI%))
     '
     SUMA = 0
     For JJ& = 0 To 12
       XX$ = Mid$(PROGENT$(KKI%), 4 * JJ& + 1, 4)
       SUMA = SUMA + CVS(XX$)
       TTXA$ = TTXA$ + XX$
     Next JJ&
     TTXA$ = TTXA$ + MKS$(SUMA) + String$(64, 0)
     '
     TTXB$ = Left$(TTXA$, 4) + MKS$(SUMA)
     For JJ& = 13 To 24
       XX$ = Mid$(PROGENT$(KKI%), 4 * JJ& + 1, 4)
       SUMA = SUMA + CVS(XX$)
       TTXB$ = TTXB$ + XX$
     Next JJ&
     TTXB$ = TTXB$ + MKS$(SUMA) + String$(128, 0)
     '
     KK& = KK& + 1
     Call GRAREG("NETIBRA", TTXA$, KK&)
     Call GRAREG("NETIBRB", TTXB$, KK&)
   Next KKI%
   '
   Call SETULTREG("NETIBRA", KK&)
   Call SETULTREG("NETIBRB", KK&)
   Check8.Value = 1
   '
' 9.- GENERA NECESIDADES DE MATERIALES POR TIPIFICADAS NETAS DE STOCK
   '
   JJ1& = 0
   HOII% = HOY(HOS$)
   REBLA$ = REGBLAN$("NECALPR4")
   FIN& = ULTREG&("NECALPR3")
   For U& = 1 To FIN&
     XX$ = REGLEIDO$("NECALPR3", U&)
     Call REPLA(XX$, Space$(24), 9, 24)
     If CVI(Mid$(XX$, 7, 2)) < HOII% Then
       Call REPLA(XX$, MKI$(HOII%), 7, 2)
     End If
     If JJ1& > 0 Then
       YY$ = REGLEIDO$("NECALPR4", JJ1&)
       If Left$(YY$, 8) = Left$(XX$, 8) Then
         CATONE# = CVD(Mid$(XX$, 33, 8)) + CVD(Mid$(YY$, 33, 8))
         Call REPLA(YY$, MKD$(CATONE#), 33, 8)
         Call GRAREG("NECALPR4", YY$, JJ1&)
         GoTo 349
       End If
     End If
     JJ1& = JJ1& + 1
     Call GRAREG("NECALPR4", XX$, JJ1&)
349 Next U&
   '
   FIN& = ULTREG&("SALIMPT")
   For U& = 1 To FIN&
     YY$ = REGLEIDO$("SALIMPT", U&)
     CIXI% = CVI(Mid$(YY$, 9, 2))
     FEXI% = CVI(Mid$(YY$, 57, 2))
     PERIT$ = Mid$(YY$, 33, 24)
     SALIT# = CVD(Mid$(YY$, 97, 8))
     '
     XX$ = REBLA$
     Call REPLA(XX$, MKI$(CIXI%), 1, 2)
     Call REPLA(XX$, UNIMED$(CIXI%), 3, 4)
     Call REPLA(XX$, MKI$(FEXI%), 7, 2)
     Call REPLA(XX$, PERIT$, 9, 24)
     Call REPLA(XX$, MKD$(SALIT#), 41, 8)
     JJ1& = JJ1& + 1
     Call GRAREG("NECALPR4", XX$, JJ1&)
   Next U&
   Call SETULTREG("NECALPR4", JJ1&)
   Call CIERRARCH("NECALPR4")
   '
   Call FILESORT("NECALPR4", "", 1, 4)
   CIAN% = 0: ACUMUL# = 0: RINICO& = 0: UENTRA& = 0: UENTRB& = 0
   FIN& = ULTREG&("NECALPR4")
   For U& = 1 To FIN&
     XX$ = REGLEIDO$("NECALPR4", U&)
     CI% = CVI(Mid$(XX$, 1, 2))
     If CI% <> CIAN% Then
       ACUMUL# = 0
       If CIAN% > 0 Then
         GoSub 250
         RINICO& = U&
       End If
       CIAN% = CI%
       XX$ = REGLEIDO$("NECALPR4", U&)
     End If
     If CVD(Mid$(XX$, 41, 8)) > 0 Then UENTRA& = U&
     ACUMUL# = ACUMUL# + CVD(Mid$(XX$, 33, 8)) - CVD(Mid$(XX$, 41, 8))
     Call REPLA(XX$, MKD$(ACUMUL#), 49, 8)
     Call GRAREG("NECALPR4", XX$, U&)
   Next U&
   ACUMUL# = 0
   GoSub 250
   CIAN% = 0
   Call CIERRARCH("NECALPR4")
   '
' 10.- GENERA NECESIDADES TIPIFICADAS NETAS DE STOCK EN FORMA DE TABLAS
   '
   'Check9.Value = 2
   For KKI% = 0 To 16384
     PROGENT$(KKI%) = String$(100, 0)
     CODIPRO%(KKI%) = 0
   Next KKI%
   CACODIP% = 0
   '
   KK& = 0
   '
   FIN& = ULTREG("NECALPR4")
   For U& = 1 To FIN&
     P00$ = REGLEIDO$("NECALPR4", U&)
     CI% = CVI(Left$(P00$, 2))
     CANPRO = CVD(Mid$(P00$, 49, 8))
     '
     If CANPRO > 0.0005 Then
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           FEXI% = CVI(Mid$(P00$, 7, 2))
           If FEXI% <= FEPLA%(0) Then
              COLUX% = 0
              GoTo 420
            Else
              For KKJ% = 24 To 1 Step -1
                If FEXI% >= FEPLA%(KKJ%) Then
                  COLUX% = KKJ%
                  GoTo 420
                End If
              Next KKJ%
           End If
           GoTo 439
           '
420        For KKI% = 1 To CACODIP%
             If CODIPRO%(KKI%) = CI% Then GoTo 430
           Next KKI%
           CACODIP% = CACODIP% + 1
           CODIPRO%(CACODIP%) = CI%
           KKI% = CACODIP%
           '
430        If COLUX% >= 0 Then
             If COLUX% <= 24 Then
               CANPRO = CANPRO + CVS(Mid$(PROGENT$(KKI%), 1 + 4 * COLUX%, 4))
               Mid$(PROGENT$(KKI%), 1 + 4 * COLUX%, 4) = MKS$(CANPRO)
             End If
           End If
         End If
       End If
     End If
     '
439 Next U&
   '
   REBLA$ = REGBLAN$("NETINTA")
   REBLB$ = REGBLAN$("NETINTB")
   KK& = 0
   '
   For KKI% = 1 To CACODIP%
     '
     CIKKI% = CODIPRO%(KKI%)
     TTXA$ = MKI$(CIKKI%) + MKI$(PROVHABI%(CIKKI%))
     '
     SUMA = 0
     For JJ& = 0 To 12
       XX$ = Mid$(PROGENT$(KKI%), 4 * JJ& + 1, 4)
       SUMA = SUMA + CVS(XX$)
       TTXA$ = TTXA$ + XX$
     Next JJ&
     TTXA$ = TTXA$ + MKS$(SUMA) + String$(64, 0)
     '
     TTXB$ = Left$(TTXA$, 4) + MKS$(SUMA)
     For JJ& = 13 To 24
       XX$ = Mid$(PROGENT$(KKI%), 4 * JJ& + 1, 4)
       SUMA = SUMA + CVS(XX$)
       TTXB$ = TTXB$ + XX$
     Next JJ&
     TTXB$ = TTXB$ + MKS$(SUMA) + String$(128, 0)
     '
     KK& = KK& + 1
     Call GRAREG("NETINTA", TTXA$, KK&)
     Call GRAREG("NETINTB", TTXB$, KK&)
   Next KKI%
   '
   Call SETULTREG("NETINTA", KK&)
   Call SETULTREG("NETINTB", KK&)
   'Check8.Value = 1
   '
   '
   '
99 Screen.MousePointer = 1
   Timer1.Enabled = False
   Call COMUNI("Calculo completo")
   Command2.Enabled = True
   
   On Error Resume Next
   Hide
   On Error GoTo 0
   GoTo 999
   '
   '
   '
' **********************************************************************
150  If RINICO& < 1 Then RINICO& = 1
     If UENTRA& < RINICO& - 1 Then UENTRA& = RINICO& - 1
     '
     For KLU& = RINICO& To UENTRA& - 1
       XX$ = REGLEIDO$("!PROGENET", KLU&)
       Call REPLA(XX$, MKS$(0), 57, 4)
       Call GRAREG("!PROGENET", XX$, KLU&)
     Next KLU&
     '
     For KLU& = UENTRA& To U& - 1
        XX$ = REGLEIDO$("!PROGENET", KLU&)
        CAPE = CVS(Mid$(XX$, 57, 4))
        If CAPE < 0 Then
            Call REPLA(XX$, MKS$(0), 57, 4)
            Call GRAREG("!PROGENET", XX$, KLU&)
          Else
            UENTRB& = KLU&
            GoTo 155
        End If
     Next KLU&
     GoTo 159
     '
155  For KLU& = UENTRB& + 1 To U& - 1
        XX$ = REGLEIDO$("!PROGENET", KLU&)
        CAPE = CVS(Mid$(XX$, 49, 4))
        Call REPLA(XX$, MKS$(CAPE), 57, 4)
        Call GRAREG("!PROGENET", XX$, KLU&)
     Next KLU&
     '
159  Return
' **********************************************************************
   '
250  If RINICO& < 1 Then RINICO& = 1
     If UENTRA& < RINICO& - 1 Then UENTRA& = RINICO& - 1
     '
     For KLU& = RINICO& To UENTRA& - 1
       XX$ = REGLEIDO$("NECALPR4", KLU&)
       Call REPLA(XX$, MKD$(0), 49, 8)
       Call GRAREG("NECALPR4", XX$, KLU&)
     Next KLU&
     '
     For KLU& = UENTRA& To U& - 1
        XX$ = REGLEIDO$("NECALPR4", KLU&)
        CAPEL# = CVD(Mid$(XX$, 49, 8))
        If CAPEL# < 0 Then
            Call REPLA(XX$, MKD$(0), 49, 8)
            Call GRAREG("NECALPR4", XX$, KLU&)
          Else
            UENTRB& = KLU&
            GoTo 255
        End If
     Next KLU&
     GoTo 259
     '
255  For KLU& = UENTRB& + 1 To U& - 1
        XX$ = REGLEIDO$("NECALPR4", KLU&)
        CAPEL# = CVD(Mid$(XX$, 33, 8))
        Call REPLA(XX$, MKD$(CAPEL#), 49, 8)
        Call GRAREG("NECALPR4", XX$, KLU&)
     Next KLU&
     '
259  Return
     '
' **********************************************************************
   '
999 End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   STIMTEMP.Show 1
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   If COMALTER%("Realmente Desea Cancelar el Cálculo\\No, Continuar\Si, Cancelar") = 2 Then
     Call CIERRARCH("*.*")
     Command1.Enabled = True
     Hide
   End If
   Command4.Enabled = True
End Sub

   '
Private Sub Form_Load()
   Label1 = TRIM$(CONTROL$("CALCONIT", "FULTICAL"))
End Sub

Private Sub Timer1_Timer()
   '
   If Image1.Visible = True Then
       Image1.Visible = False
     Else
       Image1.Visible = True
   End If
   DoEvents
   '
End Sub
