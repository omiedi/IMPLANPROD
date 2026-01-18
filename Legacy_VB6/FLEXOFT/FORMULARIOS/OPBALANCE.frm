VERSION 5.00
Begin VB.Form OPBALANCE 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Balance General - Sumas y Saldos"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9630
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5160
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9675
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   8610
         ScaleHeight     =   75
         ScaleWidth      =   795
         TabIndex        =   21
         Top             =   2625
         Width           =   855
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton command1 
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
         Left            =   8640
         Picture         =   "OPBALANCE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   420
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   8610
         Picture         =   "OPBALANCE.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Emite Reporte Solicitado"
         Top             =   2940
         Width           =   855
      End
      Begin VB.CommandButton Command7 
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
         Left            =   8610
         Picture         =   "OPBALANCE.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   1365
         Width           =   855
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Tipo de Balance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1800
         Left            =   210
         TabIndex        =   10
         Top             =   1850
         Width           =   4005
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Balance de Sumas y Saldos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   735
            TabIndex        =   16
            Top             =   1260
            Value           =   -1  'True
            Width           =   3060
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Resultados Mensuales"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   735
            TabIndex        =   12
            Top             =   420
            Width           =   2850
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Saldos - Estado de Cuentas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   735
            TabIndex        =   11
            Top             =   840
            Width           =   3060
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Formato de Salida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   4410
         TabIndex        =   9
         Top             =   315
         Width           =   4005
         Begin VB.OptionButton Option5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "Nivel 1"
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
            Left            =   2480
            TabIndex        =   20
            Top             =   1470
            Width           =   1000
         End
         Begin VB.OptionButton Option4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "Nivel 2"
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
            Left            =   2480
            TabIndex        =   19
            Top             =   1155
            Width           =   1000
         End
         Begin VB.OptionButton Option3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "Condensado Nivel 3"
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
            Left            =   1365
            TabIndex        =   18
            Top             =   840
            Width           =   2115
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Formato Completo"
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
            Left            =   420
            TabIndex        =   17
            Top             =   420
            Value           =   -1  'True
            Width           =   2535
         End
      End
      Begin VB.ListBox List1 
         Height          =   540
         Left            =   8610
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   2415
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Mes y Año"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   210
         TabIndex        =   4
         Top             =   315
         Width           =   4005
         Begin VB.CommandButton Command6 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3570
            TabIndex        =   6
            Top             =   630
            Width           =   175
         End
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1365
            TabIndex        =   5
            Top             =   630
            Width           =   2220
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Período:"
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
            Left            =   -1050
            TabIndex        =   7
            Top             =   735
            Width           =   2325
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Selección"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1120
         Left            =   4410
         TabIndex        =   1
         Top             =   2520
         Width           =   4005
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Todas las Cuentas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   2415
            TabIndex        =   3
            Top             =   525
            Width           =   1380
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Cuentas con Movimiento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   3
            Left            =   420
            TabIndex        =   2
            Top             =   525
            Value           =   -1  'True
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "OPBALANCE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Unload OPBALANCE
End Sub
'

Private Sub Command2_Click()
   Command2.Enabled = False
   Call BALSUSAL
   Command2.Enabled = True
End Sub

Private Sub Command6_Click()
   Call FRESHALL
   Call CARSELMESA
10 Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   Text6.TEXT = MESASEL$
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
    Call CIERRARCH("*.*")
End Sub

'
Sub BALSUSAL()  ' balance de sumas y saldos
   '
   MMXX0$ = TRIM$(Text6.TEXT)
   MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
   If MMXX1$ <> "" Then GoTo 20
   '
   If Option1(4).Value = True Then
5     VDEF$ = MKI$(COMEJ%) + MKI$(FINEJ%)
      OBXX$ = OBJPLA$("DESHASFECHA", VDEF$)
      If OBXX$ = "" Then Exit Sub
      DES% = CVI(Left$(OBXX$, 2))
      HAS% = CVI(Mid$(OBXX$, 3, 2))
      If HAS% < DES% Then GoTo 5
      If DES% < COMEJ% Or HAS% > FINEJ% Then GoTo 5
      FEX = DES%: MMXX1$ = FECHATEX$(FEX)
      FEX = HAS%: MMXX1$ = MMXX1$ + " - " + FECHATEX$(FEX)
      GoTo 20
   End If
   '
   Call FRESHALL
   Call CARSELMESA
10 Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) = "" Then
      Exit Sub
   End If
   '
   MMXX0$ = TRIM$(MESASEL$)
   MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
   If TRIM$(MMXX1$) = "" Then
      Exit Sub
   End If
   '
20 MESASEL$ = MMXX1$
   If DES% < COMEJ% Or HAS% > FINEJ% Then
      Call MENSERR(24, "Período Fuera de Límites\del Ejercicio Activo.")
      GoTo 10
   End If
   '
   Text6.TEXT = MESASEL$
   Text6.Refresh
   '
   TODAS% = 0: If Option1(0).Value = True Then TODAS% = 1
   NIVCOND% = 4
   If Option3.Value = True Then
         NIVCOND% = 3
      ElseIf Option4.Value = True Then
         NIVCOND% = 2
      ElseIf Option5.Value = True Then
         NIVCOND% = 1
   End If
   TIBA% = 1                           'BALANCE MENSUAL
   TIBALAN$ = "Balance Mensual de " + MESASEL$
   If Option1(2) = True Then           'ESTADO DE CUENTAS
      TIBA% = 2
      TIBALAN$ = "Estado de Cuentas a " + MESASEL$
   End If
   If Option1(4) = True Then
      TIBA% = 3 'SUMAS Y SALDOS
      TIBALAN$ = "Balance de Sumas y Saldos de " + MESASEL$
   End If
   '
   Screen.MousePointer = 11
   Dim SALINI#(16384)
   Dim SUDEBE#(16384)
   Dim SUHABE#(16384)
   '
   ' CARGAR MOVIMIENTOS DE DIARIO
   FIN& = MAXDIARIO%
   For KK% = 1 To MAXDIARIO%
      UU& = KK%
      Call TRACE(20, UU&, FIN&)
      Call LEEDIARIO(KK%)
      CUE% = CVI(E11$)
      NUAS = CVS(E12$)
      FEAS% = CVI(E15$)
      IPAS# = CVD(E16$)
      '
      If CUE% > 0 Then
         If CUE% <= 16384 Then
            If FEAS% < DES% Or (FEAS% = DES% And NUAS = 1) Then
                 SALINI#(CUE%) = SALINI#(CUE%) + (Int(100 * IPAS# + 0.5)) / 100
               Else
                 If FEAS% <= HAS% Then
                    If IPAS# > 0 Then
                        SUDEBE#(CUE%) = SUDEBE#(CUE%) + (Int(100 * IPAS# + 0.5)) / 100
                       Else
                        SUHABE#(CUE%) = SUHABE#(CUE%) - (Int(100 * IPAS# + 0.5)) / 100
                    End If
                 End If
            End If
         End If
      End If
   Next KK%
   '
   ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   FIN& = MAXDIA&
   For KKL& = 1 To MAXDIA&
      UU& = KKL&
      Call TRACE(20, UU&, FIN&)
      Call LEESUCAJA(KKL&)
      If Abs(CVS(D19$)) < 0.1 Then
         CUE% = CVI(D11$)
         FEAS% = CVI(D15$)
         IPAS# = CVD(D16$)
        If CUE% > 0 Then
            If CUE% <= 16384 Then
               If FEAS% < DES% Then
                    SALINI#(CUE%) = SALINI#(CUE%) + (Int(100 * IPAS# + 0.5)) / 100
                  Else
                    If FEAS% <= HAS% Then
                       If IPAS# > 0 Then
                           SUDEBE#(CUE%) = SUDEBE#(CUE%) + (Int(100 * IPAS# + 0.5)) / 100
                          Else
                           SUHABE#(CUE%) = SUHABE#(CUE%) - (Int(100 * IPAS# + 0.5)) / 100
                       End If
                    End If
               End If
            End If
         End If
      End If
   Next KKL&
   '
   On TIBA% GoTo 100, 100, 300
   Call CIERRARCH("*.*")
   Exit Sub
   '
100 ' BALANCE MENSUAL *******************************
   '
   If NXP% < 1 Then
      NXP% = FILEOPEN%(LUDAT$ + "PLAGRAL" + "." + EXTE$, 0, 4)
   End If
   '
   JJ& = 0: KLJ& = 0
   Dim IVP As String * 4
   Dim SUNIVEL#(5), DENIVEL$(5)
   '
   FIN& = LOF(NXP%) / 4
   For KK% = 1 To LOF(NXP%) / 4
      UU& = KK%
      Call TRACE(20, UU&, FIN&)
      Get #NXP%, KK%, IVP$
      CUE% = CVI(Mid$(IVP$, 1, 2))
      NIVEI% = CVI(Mid$(IVP$, 3, 2))
      Get #NXP%, KK% + 1, IVP$
      PRONI% = CVI(Mid$(IVP$, 3, 2))
      If CUE% > 0 Then
         If CUE% <= 16384 Then
            IMPO# = SUDEBE#(CUE%) - SUHABE#(CUE%)
            If TIBA% = 2 Then IMPO# = IMPO# + SALINI#(CUE%)
            If NIVEI% > NIVANT% Then
                 If NIVEI% <= NIVCOND% Then
                    ZZ$ = Space$(128)
                    JJ& = JJ& + 1
                    Call GRAREG("BALANCE", ZZ$, JJ&)
                 End If
               ElseIf NIVEI% < NIVANT% Then
                 For KQ% = NIVANT% - 1 To NIVEI% Step -1
                   If KQ% <= NIVCOND% Then
                      If TODAS% = 1 Or Abs(SUNIVEL#(KQ%)) > 0.005 Then
                         ZZ$ = Space$(128)
                         JJ& = JJ& + 1
                         Call GRAREG("BALANCE", ZZ$, JJ&)
                         ZZ$ = "." + Space$(127)
                         QQ$ = Space$(88) + String$(40, 0)
                         TXTX$ = "TOTAL " + DENIVEL$(KQ%)
                         Call REPLA(QQ$, TXTX$, 15, 74)
                         If Abs(SUNIVEL#(KQ%)) >= 0.005 Then TXTX$ = TXTX$ + " " + String$(128, ".")
                         ITXT$ = CSTRING$(MKD$(SUNIVEL#(KQ%)), 4, 12, 2, 2) + Space$(128)
                         Call REPLA(ZZ$, TXTX$, 1 + 5 * KQ%, 128)
                         Call REPLA(ZZ$, ITXT$, 117 - 10 * KQ%, 128)
                         JJ& = JJ& + 1
                         Call GRAREG("BALANCE", ZZ$, JJ&)
                         '
                         Call REPLA(QQ$, MKD$(SUNIVEL#(KQ%)), 121 - 8 * KQ%, 8)
                         KLJ& = KLJ& + 1
                         Call GRAREG("BALANCF", QQ$, KLJ&)
                      End If
                   End If
                   SUNIVEL#(KQ%) = 0
                   DENIVEL$(KQ%) = ""
                 Next KQ%
                 If NIVEI% <= NIVCOND% Then
                    ZZ$ = Space$(128)
                    JJ& = JJ& + 1
                    Call GRAREG("BALANCE", ZZ$, JJ&)
                 End If
            End If
            '
            If NIVEI% <= NIVCOND% Then
               If TODAS% = 1 Or Abs(IMPO#) >= 0.005 Or PRONI% > NIVEI% Then
                  ZZ$ = "." + Space$(127)
                  QQ$ = Space$(88) + String$(40, 0)
                  TXTX$ = TRIM$(CODCUE$(CUE%)) + " " + TRIM$(DENOCUE$(CUE%))
                  Call REPLA(QQ$, CODCUE$(CUE%), 1, 14)
                  Call REPLA(QQ$, DENOCUE$(CUE%), 15, 74)
                  If Abs(IMPO#) >= 0.005 Then TXTX$ = TXTX$ + " " + String$(128, ".")
                  ITXT$ = CSTRING$(MKD$(IMPO#), 4, 12, 2, 2) + Space$(128)
                  Call REPLA(ZZ$, TXTX$, 1 + 5 * NIVEI%, 128)
                  Call REPLA(ZZ$, ITXT$, 117 - 10 * NIVEI%, 128)
                  JJ& = JJ& + 1
                  Call GRAREG("BALANCE", ZZ$, JJ&)
                  '
                  Call REPLA(QQ$, MKD$(IMPO#), 121 - 8 * NIVEI%, 8)
                  KLJ& = KLJ& + 1
                  Call GRAREG("BALANCF", QQ$, KLJ&)
               End If
            End If
            '
            For KQ% = 0 To NIVEI% - 1
               SUNIVEL#(KQ%) = SUNIVEL#(KQ%) + IMPO#
            Next KQ%
            SUNIVEL#(NIVEI%) = IMPO#
            DENIVEL$(NIVEI%) = TRIM$(DENOCUE$(CUE%))
            NIVANT% = NIVEI%
            '
         End If
      End If
   Next KK%
   '
   For KQ% = NIVANT% - 1 To 0 Step -1
      If KQ% <= NIVCOND% Then
         If TODAS% = 1 Or Abs(SUNIVEL#(KQ%)) >= 0.005 Then
            ZZ$ = Space$(128)
            JJ& = JJ& + 1
            Call GRAREG("BALANCE", ZZ$, JJ&)
            ZZ$ = "." + Space$(127)
            QQ$ = Space$(88) + String$(40, 0)
            TXTX$ = "TOTAL " + DENIVEL$(KQ%)
            Call REPLA(QQ$, TXTX$, 15, 74)
            If Abs(SUNIVEL#(KQ%)) >= 0.005 Then TXTX$ = TXTX$ + " " + String$(128, ".")
            ITXT$ = CSTRING$(MKD$(SUNIVEL#(KQ%)), 4, 12, 2, 2) + Space$(128)
            Call REPLA(ZZ$, TXTX$, 1 + 5 * KQ%, 128)
            Call REPLA(ZZ$, ITXT$, 117 - 10 * KQ%, 128)
            JJ& = JJ& + 1
            Call GRAREG("BALANCE", ZZ$, JJ&)
            '
            Call REPLA(QQ$, MKD$(SUNIVEL#(KQ%)), 121 - 8 * KQ%, 8)
            KLJ& = KLJ& + 1
            Call GRAREG("BALANCF", QQ$, KLJ&)
         End If
      End If
      SUNIVEL#(KQ%) = 0
      DENIVEL$(KQ%) = ""
   Next KQ%
   ZZ$ = Space$(128)
   JJ& = JJ& + 1
   Call GRAREG("BALANCE", ZZ$, JJ&)
   '
   Call SETULTREG("BALANCE", JJ&)
   Call CIERRARCH("BALANCE")
   '
   Call SETULTREG("BALANCF", KLJ&)
   Call CIERRARCH("BALANCF")
   '
   Call HEADERS("BALANCE", "")
   Call HEADERS("BALANCE", "Corresponde a: " + MESASEL$)
   '
   Call HEADERS("BALANCF", "")
   Call HEADERS("BALANCF", "Corresponde a: " + MESASEL$)
   '
   Close #NXP%: NXP% = 0
   Screen.MousePointer = 1
   '
   OPX1% = COMALTER%("Elija Formato de Salida:\\Planilla Plana\Exportable a Excel")
   If OPX1% = 1 Then
       Call CONECRUN("*BALANCE", "Balance General")
     ElseIf OPX1% = 2 Then
       Call CONECRUN("*BALANCF", "Balance General")
   End If
   Exit Sub
   '
300 ' BALANCE DE SUMAS Y SALDOS *********************
   '
   If NX3% < 1 Then
      NX3% = FILEOPEN%(LUDAT$ + "INVECUE" + "." + EXTE$, 0, 16)
   End If
   '
   JJ& = 0
   Dim IVX As String * 16
   FIN& = LOF(NX3%) / 16
   For KK% = 1 To LOF(NX3%) / 16
      UU& = KK%
      Call TRACE(20, UU&, FIN&)
      Get #NX3%, KK%, IVX$
      CUE% = CVI(Mid$(IVX$, 13, 2))
      If CUE% > 0 Then
         If CUE% <= 16384 Then
            If TODAS% = 1 Or Abs(SALINI#(CUE%)) + Abs(SUDEBE#(CUE%)) + Abs(SUHABE#(CUE%)) >= 0.005 Then
               ZZ$ = REGBLAN$("BASUSAL")
               Call REPLA(ZZ$, CODCUE$(CUE%), 1, 12)
               Call REPLA(ZZ$, DENOCUE$(CUE%), 13, 52)
               Call REPLA(ZZ$, MKD$(SALINI#(CUE%)), 65, 8)
               Call REPLA(ZZ$, MKD$(SUDEBE#(CUE%)), 73, 8)
               Call REPLA(ZZ$, MKD$(SUHABE#(CUE%)), 81, 8)
               Call REPLA(ZZ$, MKD$(SUDEBE#(CUE%) - SUHABE#(CUE%)), 89, 8)
               Call REPLA(ZZ$, MKD$(SALINI#(CUE%) + SUDEBE#(CUE%) - SUHABE#(CUE%)), 97, 8)
               JJ& = JJ& + 1
               Call GRAREG("BASUSAL", ZZ$, JJ&)
            End If
         End If
      End If
   Next KK%
   '
   Call SETULTREG("BASUSAL", JJ&)
   Call CIERRARCH("BASUSAL")
   '
   Call HEADERS("BASUSAL", "")
   Call HEADERS("BASUSAL", "Corresponde a: " + MESASEL$)
   '
   Screen.MousePointer = 1
   '
   Call CONECRUN("*BASUSAL", "Balance de Sumas y Saldos")
   '
End Sub

Private Sub Text6_DblClick()
   Call FRESHALL
   Call CARSELMESA
10 Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   Text6.TEXT = MESASEL$
End Sub
