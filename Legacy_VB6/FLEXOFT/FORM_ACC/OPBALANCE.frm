VERSION 5.00
Begin VB.Form OPBALANCE 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Balance General - Sumas y Saldos"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9555
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
   ScaleHeight     =   3510
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0D0D0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6720
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9915
      Begin VB.CommandButton Command1 
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
         Height          =   855
         Left            =   8580
         Picture         =   "OPBALANCE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   2500
         Width           =   855
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0D0D0&
         Caption         =   "Tipo de Balance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   3360
         TabIndex        =   9
         Top             =   240
         Width           =   3165
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0D0D0&
            Caption         =   "Sumas y Saldos"
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
            Left            =   240
            TabIndex        =   12
            Top             =   1080
            Value           =   -1  'True
            Width           =   1860
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0D0D0&
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
            Left            =   240
            TabIndex        =   11
            Top             =   480
            Width           =   2370
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0D0D0&
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
            Left            =   240
            TabIndex        =   10
            Top             =   780
            Width           =   2700
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0D0D0&
         Caption         =   "Formato de Salida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   6600
         TabIndex        =   8
         Top             =   240
         Width           =   2805
         Begin VB.OptionButton Option5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0D0D0&
            Caption         =   "Nv-1"
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
            Left            =   1755
            TabIndex        =   16
            Top             =   1080
            Width           =   765
         End
         Begin VB.OptionButton Option4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0D0D0&
            Caption         =   "Nv-2"
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
            Left            =   1755
            TabIndex        =   15
            Top             =   840
            Width           =   765
         End
         Begin VB.OptionButton Option3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0D0D0&
            Caption         =   "Condensado Nv-3"
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
            Left            =   645
            TabIndex        =   14
            Top             =   600
            Width           =   1875
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0D0D0&
            Caption         =   "Completo"
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
            Left            =   360
            TabIndex        =   13
            Top             =   300
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0D0D0&
         Caption         =   "Periodo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3165
         Begin VB.CommandButton Command6 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2640
            TabIndex        =   7
            Top             =   720
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
            Height          =   300
            Left            =   480
            TabIndex        =   6
            Top             =   720
            Width           =   2220
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0D0D0&
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
         Height          =   1550
         Left            =   120
         TabIndex        =   2
         Top             =   1800
         Width           =   8325
         Begin VB.TextBox Text10 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3135
            TabIndex        =   20
            Top             =   480
            Width           =   1485
         End
         Begin VB.CommandButton Command9 
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
            Height          =   300
            Left            =   4605
            TabIndex        =   19
            Top             =   480
            Width           =   175
         End
         Begin VB.TextBox Text8 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6375
            TabIndex        =   18
            Text            =   " "
            Top             =   465
            Width           =   1485
         End
         Begin VB.CommandButton Command8 
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
            Height          =   300
            Left            =   7845
            TabIndex        =   17
            Top             =   465
            Width           =   175
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0D0D0&
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
            Left            =   360
            TabIndex        =   4
            Top             =   960
            Width           =   1260
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0D0D0&
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
            Left            =   360
            TabIndex        =   3
            Top             =   480
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2400
            TabIndex        =   24
            Top             =   870
            Width           =   2415
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5640
            TabIndex        =   23
            Top             =   870
            Width           =   2415
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
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
            Left            =   720
            TabIndex        =   22
            Top             =   585
            Width           =   2325
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta :"
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
            Left            =   3960
            TabIndex        =   21
            Top             =   570
            Width           =   2325
         End
      End
      Begin VB.Image Image2 
         Height          =   510
         Left            =   8000
         Picture         =   "OPBALANCE.frx":030A
         Top             =   1920
         Width           =   2010
      End
   End
End
Attribute VB_Name = "OPBALANCE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
   Command1.Enabled = False
   Call BALSUSAL
   Command1.Enabled = True
End Sub



Private Sub Command6_Click()
   Call FRESHALL
   Call CARSELMESA
10 Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   Text6.TEXT = MESASEL$
End Sub

Private Sub Command8_Click()
    Call SELECHO("ECUECOT")
    PRICUEX$ = VALACT1$("ECUECOT")
    If PRICUEX$ <> "" Then
        Text8.TEXT = PRICUEX$
    End If
End Sub

Private Sub Command9_Click()
    Call SELECHO("ECUECOT")
    PRICUEX$ = VALACT1$("ECUECOT")
    If PRICUEX$ <> "" Then
        Text10.TEXT = PRICUEX$
    End If
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Text10 = Left$(REGLEIDO$("ECUECOT", 1), 12)
    Text8 = Left$(REGLEIDO$("ECUECOT", ULTREG&("ECUECOT")), 12)
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
9 Call SELECHO("SELMESA")
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
      GoTo 9
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
   Dim SALINI1#(16384)
   Dim SUDEBE1#(16384)
   Dim SUHABE1#(16384)
   '
   ' CARGAR MOVIMIENTOS DE DIARIO
   SUCUI% = 0
   SQLX$ = "SELECT FECHASIEN, NUIDASIEN, REFEPASE, CUECOINT, IDEBEPASE, IHABEPASE "
   SQLX$ = SQLX$ + " FROM DETAASIEN WHERE "
   SQLX$ = SQLX$ + "(CDBL(FECHASIEN) >= " & CDbl(CVDAT(COMEJ%)) & " "
   SQLX$ = SQLX$ + "AND CDBL(FECHASIEN) <= " & CDbl(CVDAT(HAS%)) & " "
   SQLX$ = SQLX$ + "AND SUCUASIen = " & SUCUI% & " "
   SQLX$ = SQLX$ + "AND STATPASE > 0); "
   Set MODIASEL = Contable.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   '
   CARESE& = 0
   On Error Resume Next
   MODIASEL.MoveLast
   CARESE& = MODIASEL.RecordCount
   If CARESE& < 1 Then
      Call MENSERR(24, "Sin Movimientos")
      Exit Sub
   End If
   '
   JJ& = 0: FIN& = CARESE&
   '
   With MODIASEL
     .MoveFirst
10   JJ& = JJ& + 1
     Call TRACE(20, JJ&, FIN&)
     '
     CUE% = !CUECOINT
     FEAS% = FECHANUM(Format$(!FECHASIEN, "DD/MM/YY"))
     IPAS# = !IDebePase - !IHabePase
     '
     If CUE% > 0 Then
       If CUE% <= 16384 Then
         If FEAS% < DES% Then ' Or (FEAS% = DES% And NUAS = 1) Then
             SALINI1#(CUE%) = SALINI1#(CUE%) + (Int(100 * IPAS# + 0.5)) / 100
           Else
             If FEAS% <= HAS% Then
               If IPAS# > 0 Then
                   SUDEBE1#(CUE%) = SUDEBE1#(CUE%) + (Int(100 * IPAS# + 0.5)) / 100
                 Else
                   SUHABE1#(CUE%) = SUHABE1#(CUE%) - (Int(100 * IPAS# + 0.5)) / 100
               End If
             End If
           End If
         End If
      End If
      '
      If JJ& < CARESE& Then
        MODIASEL.MoveNext
        GoTo 10
      End If
   End With
   '
   MODIASEL.Close
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
            IMPO# = SUDEBE1#(CUE%) - SUHABE1#(CUE%)
            If TIBA% = 2 Then IMPO# = IMPO# + SALINI1#(CUE%)
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
   For KQ% = NIVANT% - 1 To NIVEI% Step -1
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
   On Error Resume Next
   Contable.Execute "DROP TABLE BALASUSA"
   Contable.Execute "SELECT codexcue, denocue, codincue INTO Balasusa FROM Plancuen;"
   On Error GoTo 0
   '
   SQLX$ = "SELECT CODEXCUE, DENOCUE, CODINCUE, SALINI, SUDEBE, SUHABE, RESULT, SALFIN "
   SQLX$ = SQLX$ + " FROM PLANCUEN "
   Set BALAPRIM = Contable.OpenRecordset(SQLX$)
   '
   CACUEN& = 0
   On Error Resume Next
   With BALAPRIM
     .MoveLast
     FIN& = .RecordCount
     For U& = 1 To FIN&
       If U& = 1 Then
            .MoveFirst
          Else
            .MoveNext
       End If
       CUEI% = .CODINCUE
       .Edit
       !SALINI = SALINI1#(CUEI%)
       !SUDEBE = SUDEBE1#(CUEI%)
       !SUHABE = SUHABE1#(CUEI%)
       !RESULT = SUDEBE1#(CUEI%) - SUHABE1#(CUEI%)
       !SALFIN = SALINI1#(CUEI%) + SUDEBE1#(CUEI%) - SUHABE1#(CUEI%)
       .Update
     Next U&
   End With
   '
   SQLX$ = "SELECT CODEXCUE, DENOCUE, CODINCUE, SALINI, SUDEBE, SUHABE, RESULT, SALFIN "
   SQLX$ = SQLX$ + " FROM PLANCUEN "
   ACONDI% = 0
   If Option1(3) = True Then
     SQLX$ = SQLX$ + "WHERE (SALINI > 0 OR SUDEBE > 0 OR SUHABE > 0 "
     SQLX$ = SQLX$ + "OR RESULT > 0 OR SALFIN > 0 ) "
     ACONDI% = 1
   End If
   '
   If TRIM$(Text10) <> "" Then
     If ACONDI% > 0 Then
         SQLX$ = SQLX$ + "AND "
       Else
         SQLX$ = SQLX$ + "WHERE "
     End If
     SQLX$ = SQLX$ + "CODEXCUE >= '" & TRIM$(Text10) & "' "
     ACONDI% = 1
   End If
   '
   If TRIM$(Text8) <> "" Then
     If ACONDI% > 0 Then
         SQLX$ = SQLX$ + "AND "
       Else
         SQLX$ = SQLX$ + "WHERE "
     End If
     SQLX$ = SQLX$ + "CODEXCUE <= '" & TRIM$(Text8) & "' "
   End If
   '
   'MsgBox SQLX$
   Set BALASEC = Contable.OpenRecordset(SQLX$)
   BALASEC.MoveLast
   If BALASEC.RecordCount < 1 Then
     Call MENSERR(24, "Sin Registros Activos")
     Exit Sub
   End If
   '
   Set BASUMSAL.Data1.Recordset = BALASEC ' Contable.OpenRecordset(SQLX$) ', dbOpenSnapshot, dbReadOnly)
   BASUMSAL.Data1.Refresh
   BASUMSAL.DBGrid1.Refresh
   '
   Set COLUI = BASUMSAL.DBGrid1.Columns
   COLUI(0).Caption = "Cuenta"
   COLUI(1).Caption = "Descripcion"
   COLUI(3).Caption = "S.Inicial"
   COLUI(4).Caption = "Debe"
   COLUI(5).Caption = "Haber"
   COLUI(6).Caption = "Resultado"
   COLUI(7).Caption = "Saldo Final"
   For COLUJ% = 3 To 7
     COLUI(COLUJ%).NumberFormat = "########0.00"
     COLUI(COLUJ).Width = 1200
   Next COLUJ%
   COLUI(2).Visible = 0
   '
   COLUI(0).Width = 1470
   COLUI(1).Width = 3645
   COLUI(2).Width = 1020
   COLUI(3).Width = 1200
   COLUI(4).Width = 1200
   COLUI(5).Width = 1200
   COLUI(6).Width = 1200
   COLUI(7).Width = 1200
   Screen.MousePointer = 1
   BASUMSAL.Show 1
   Exit Sub
   '
   On Error Resume Next
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
            If TODAS% = 1 Or Abs(SALINI1#(CUE%)) + Abs(SUDEBE1#(CUE%)) + Abs(SUHABE1#(CUE%)) >= 0.005 Then
               ZZ$ = REGBLAN$("BASUSAL")
               Call REPLA(ZZ$, CODCUE$(CUE%), 1, 12)
               Call REPLA(ZZ$, DENOCUE$(CUE%), 13, 52)
               Call REPLA(ZZ$, MKD$(SALINI1#(CUE%)), 65, 8)
               Call REPLA(ZZ$, MKD$(SUDEBE1#(CUE%)), 73, 8)
               Call REPLA(ZZ$, MKD$(SUHABE1#(CUE%)), 81, 8)
               Call REPLA(ZZ$, MKD$(SUDEBE1#(CUE%) - SUHABE1#(CUE%)), 89, 8)
               Call REPLA(ZZ$, MKD$(SALINI1#(CUE%) + SUDEBE1#(CUE%) - SUHABE1#(CUE%)), 97, 8)
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

Private Sub Text10_Change()
    Label2 = ECOARCH("ECUECOT", Text10.TEXT)
End Sub

Private Sub Text6_DblClick()
   Call FRESHALL
   Call CARSELMESA
10 Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   Text6.TEXT = MESASEL$
End Sub

Sub BALSUSAL1()  ' balance de sumas y saldos
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
9 Call SELECHO("SELMESA")
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
      GoTo 9
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
   Dim SALINI1#(16384)
   Dim SUDEBE1#(16384)
   Dim SUHABE1#(16384)
   '
   ' CARGAR MOVIMIENTOS DE DIARIO
   SUCUI% = 0
   SQLX$ = "SELECT FECHASIEN, NUIDASIEN, REFEPASE, CUECOINT, IDEBEPASE, IHABEPASE "
   SQLX$ = SQLX$ + " FROM DETAASIEN WHERE "
   SQLX$ = SQLX$ + "(CDBL(FECHASIEN) >= " & CDbl(CVDAT(COMEJ%)) & " "
   SQLX$ = SQLX$ + "AND CDBL(FECHASIEN) <= " & CDbl(CVDAT(HAS%)) & " "
   SQLX$ = SQLX$ + "AND SUCUASIen = " & SUCUI% & " "
   SQLX$ = SQLX$ + "AND STATPASE > 0); "
   Set MODIASEL = Contable.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   '
   CARESE& = 0
   On Error Resume Next
   MODIASEL.MoveLast
   CARESE& = MODIASEL.RecordCount
   If CARESE& < 1 Then
      Call MENSERR(24, "Sin Movimientos")
      Exit Sub
   End If
   '
   JJ& = 0: FIN& = CARESE&
   '
   With MODIASEL
     .MoveFirst
10   JJ& = JJ& + 1
     Call TRACE(20, JJ&, FIN&)
     '
     CUE% = !CUECOINT
     FEAS% = FECHANUM(Format$(!FECHASIEN, "DD/MM/YY"))
     IPAS# = !IDebePase - !IHabePase
     '
     If CUE% > 0 Then
       If CUE% <= 16384 Then
         If FEAS% < DES% Then ' Or (FEAS% = DES% And NUAS = 1) Then
             SALINI1#(CUE%) = SALINI1#(CUE%) + (Int(100 * IPAS# + 0.5)) / 100
           Else
             If FEAS% <= HAS% Then
               If IPAS# > 0 Then
                   SUDEBE1#(CUE%) = SUDEBE1#(CUE%) + (Int(100 * IPAS# + 0.5)) / 100
                 Else
                   SUHABE1#(CUE%) = SUHABE1#(CUE%) - (Int(100 * IPAS# + 0.5)) / 100
               End If
             End If
           End If
         End If
      End If
      '
      If JJ& < CARESE& Then
        MODIASEL.MoveNext
        GoTo 10
      End If
   End With
   '
   MODIASEL.Close
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
            IMPO# = SUDEBE1#(CUE%) - SUHABE1#(CUE%)
            If TIBA% = 2 Then IMPO# = IMPO# + SALINI1#(CUE%)
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
   For KQ% = NIVANT% - 1 To NIVEI% Step -1
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
   SQLX$ = "SELECT CODEXCUE, DENOCUE, CODINCUE, SALINI, SUDEBE, SUHABE, RESULT, SALFIN "
   SQLX$ = SQLX$ + " FROM PLANCUEN " 'WHERE "
   'SQLX$ = SQLX$ '+ "(CDBL(FECHASIEN) >= " & DESYY & " "
   'SQLX$ = SQLX$ '+ "AND CDBL(FECHASIEN) <= " & HASYY & " "
   'SQLX$ = SQLX$ '+ "AND SUCUASIen = " & SUCUI% & " "
   'SQLX$ = SQLX$ '+ "AND STATPASE > 0) "
   'SQLX$ = SQLX$ '+ "ORDER BY CODEXCUE;"
   Set BASUMSAL.Data1.Recordset = Contable.OpenRecordset(SQLX$)
   '
   BASUMSAL.Data1.Refresh
   BASUMSAL.DBGrid1.Refresh
   With BASUMSAL.Data1.Recordset
     .MoveLast
     FIN& = .RecordCount
     For U& = 1 To FIN&
       If U& = 1 Then
            .MoveFirst
          Else
            .MoveNext
       End If
       CUEI% = .CODINCUE
       .Edit
       .SALINI = SALINI1#(CUEI%)
       .SUDEBE = SUDEBE1#(CUEI%)
       .SUHABE = SUHABE1#(CUEI%)
       .RESULT = SUDEBE1#(CUEI%) - SUHABE1#(CUEI%)
       .SALFIN = SALINI1#(CUEI%) + SUDEBE1#(CUEI%) - SUHABE1#(CUEI%)
       .Update
     Next U&
   End With
   '
   BASUMSAL.Data1.Refresh
   BASUMSAL.DBGrid1.Refresh
   '
   Set COLUI = BASUMSAL.DBGrid1.Columns
   COLUI(0).Caption = "Cuenta"
   COLUI(1).Caption = "Descripcion"
   COLUI(3).Caption = "S.Inicial"
   COLUI(4).Caption = "Debe"
   COLUI(5).Caption = "Haber"
   COLUI(6).Caption = "Resultado"
   COLUI(7).Caption = "Saldo Final"
   For COLUJ% = 3 To 7
     COLUI(COLUJ%).NumberFormat = "########0.00"
     COLUI(COLUJ).Width = 1200
   Next COLUJ%
   COLUI(2).Visible = 0
   COLUI(1).Width = 3500
   Screen.MousePointer = 1
   BASUMSAL.Show 1
   Exit Sub
   '
   On Error Resume Next
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
            If TODAS% = 1 Or Abs(SALINI1#(CUE%)) + Abs(SUDEBE1#(CUE%)) + Abs(SUHABE1#(CUE%)) >= 0.005 Then
               ZZ$ = REGBLAN$("BASUSAL")
               Call REPLA(ZZ$, CODCUE$(CUE%), 1, 12)
               Call REPLA(ZZ$, DENOCUE$(CUE%), 13, 52)
               Call REPLA(ZZ$, MKD$(SALINI1#(CUE%)), 65, 8)
               Call REPLA(ZZ$, MKD$(SUDEBE1#(CUE%)), 73, 8)
               Call REPLA(ZZ$, MKD$(SUHABE1#(CUE%)), 81, 8)
               Call REPLA(ZZ$, MKD$(SUDEBE1#(CUE%) - SUHABE1#(CUE%)), 89, 8)
               Call REPLA(ZZ$, MKD$(SALINI1#(CUE%) + SUDEBE1#(CUE%) - SUHABE1#(CUE%)), 97, 8)
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


Private Sub Text8_Change()
    Label1 = ECOARCH("ECUECOT", Text8.TEXT)
End Sub
