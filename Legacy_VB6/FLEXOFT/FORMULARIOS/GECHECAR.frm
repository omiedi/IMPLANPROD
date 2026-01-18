VERSION 5.00
Begin VB.Form GECHECAR 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gestion de Cartera de Valores"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5610
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
   ScaleHeight     =   6885
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   4540
      Picture         =   "GECHECAR.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1200
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Actualización"
      Height          =   2535
      Left            =   210
      TabIndex        =   4
      Top             =   210
      Width           =   4110
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1575
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   5
         Top             =   1680
         Width           =   2010
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
            Picture         =   "GECHECAR.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Cheques Rechazados"
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
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1260
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   8
         Top             =   1260
         Width           =   2010
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
            Left            =   -20
            Picture         =   "GECHECAR.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Entrega  de Valores"
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
            Top             =   0
            Width           =   960
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   11
         Top             =   840
         Width           =   2010
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
            Left            =   -20
            Picture         =   "GECHECAR.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Carga Manual de Cartera"
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
            Top             =   0
            Width           =   1380
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   16
         Top             =   420
         Width           =   2010
         Begin VB.CommandButton Command8 
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
            Picture         =   "GECHECAR.frx":0C28
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Depósito Bancario"
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
            Top             =   0
            Width           =   1380
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Consultas - Listados"
      Height          =   3795
      Left            =   210
      TabIndex        =   1
      Top             =   2940
      Width           =   5160
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   1215
         Left            =   120
         TabIndex        =   24
         Top             =   2520
         Width           =   4935
         Begin VB.Frame Frame6 
            BackColor       =   &H00C0E0FF&
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
            Left            =   120
            TabIndex        =   28
            Top             =   120
            Width           =   3690
            Begin VB.OptionButton Option4 
               BackColor       =   &H00C0E0FF&
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
               TabIndex        =   30
               Top             =   250
               Width           =   1275
            End
            Begin VB.OptionButton Option5 
               BackColor       =   &H00C0E0FF&
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
               TabIndex        =   29
               Top             =   250
               Value           =   -1  'True
               Width           =   1380
            End
         End
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
            Left            =   4005
            Picture         =   "GECHECAR.frx":0F32
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   195
            Width           =   855
         End
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   120
            ScaleHeight     =   75
            ScaleWidth      =   4695
            TabIndex        =   25
            Top             =   960
            Width           =   4750
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   26
               Top             =   0
               Width           =   12000
            End
         End
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cheques Recibidos p/Rango Fechas"
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
         TabIndex        =   23
         Top             =   600
         Width           =   3585
      End
      Begin VB.CommandButton Command9 
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
         Height          =   240
         Left            =   4095
         TabIndex        =   22
         Top             =   1950
         Width           =   855
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Conciliación Bancaria por Fecha Registración"
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
         TabIndex        =   21
         Top             =   2205
         Width           =   4530
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Posición Diferida"
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
         TabIndex        =   20
         ToolTipText     =   "Doble-Click para Conciliar"
         Top             =   1890
         Width           =   3900
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cheques Propios Entregados por Debitar"
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
         TabIndex        =   19
         Top             =   1575
         Width           =   3900
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
         Height          =   525
         Left            =   4095
         Picture         =   "GECHECAR.frx":123C
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   315
         Width           =   855
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Búsqueda de Cheque por Número"
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
         Top             =   315
         Value           =   -1  'True
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Detalle de Valores en Cartera"
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
         Top             =   1280
         Width           =   3900
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         BorderWidth     =   2
         X1              =   0
         X2              =   5145
         Y1              =   1050
         Y2              =   1050
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
      Left            =   4540
      Picture         =   "GECHECAR.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   330
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3990
      Picture         =   "GECHECAR.frx":1690
      Top             =   2220
      Width           =   2010
   End
End
Attribute VB_Name = "GECHECAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Unload GECHECAR
End Sub

Private Sub Command2_Click()
   If Option1.Value = True Then
        Call LICHECAR
      ElseIf Option2.Value = True Then
        Call LICHEQEM
      ElseIf Option3.Value = True Then
        Call PODIFER
      ElseIf Option7.Value = True Then
        Call CONCIBANC
   End If
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call CARGACHECAR
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call ENTREGACHE
    Command4.Enabled = True
End Sub


Private Sub Command5_Click()
   Command5.Enabled = False
   Call CHERECHAZA
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Call HELPONLINE("GECHECAR")
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   If Option6.Value = True Then
       Call BUSCACHEQ
     ElseIf Option8.Value = True Then
       Call LICHEREC
   End If
   Command7.Enabled = True
End Sub
'
Sub LICHEREC()
   '
   HOII% = HOY(HOS$)
   FEX = FINIPDIF%
   HAS% = HOII%
   DES% = FECHANUM("01" + Mid$(HOS$, 3))
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   '
5  OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then Exit Sub
   '
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   If HAS% < DES% Then Exit Sub
   '
   If EXISTE%(LUDAT$ + "LICHEREC.RFS") < 1 Then
     Call COPYSTRU("BADACHE", "LICHEREC")
   End If
   '
   FIN& = ULTREG&("BADACHE")
   JJ& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("BADACHE", U&)
     FEMII% = CVI(Mid$(XX$, 61, 2))
     If FEMII% >= DES% Then
       If FEMII% <= HAS% Then
         JJ& = JJ& + 1
         Call GRAREG("LICHEREC", XX$, JJ&)
       End If
     End If
   Next U&
   Call SETULTREG("LICHEREC", JJ&)
   Call CIERRARCH("LICHEREC")
   '
   Call HEADERS("LICHEREC", "")
   Call HEADERS("LICHEREC", "Desde - Hasta: " + RANFECHA$(DES%, HAS%))
   Call CONECRUN("*SOCHEREC", "Listado de Cheques Recibidos")
   '
End Sub
'
Private Sub Command8_Click()
   Command8.Enabled = False
   'Call BLOQARCH("ASICONT")
   TIASI% = 4
   ASICAJA.Show 1
   Unload ASICAJA
   Command8.Enabled = True
End Sub

Sub PODIFER()

   Static CAMP%, CUEMP%(2048), MEPA%(2048)
   '
   Dim REFE$(2048)
   Dim FECHAMO%(2048)
   Dim IPODIF#(2048)
   '
   If EXTE$ <> "DAT" Then
      Call MENSERR(24, "Para Consultar Posición Diferida\Debe Pasar al Ejercicio ACTUAL")
      Exit Sub
   End If
   '
   If CAMP% <= 0 Then
      For I& = 1 To ULTREG&("MEPAGO")
         F0$ = REGLEIDO$("MEPAGO", I&)
         CAMP% = I&
         MEPA%(CAMP%) = CVI(Left$(F0$, 2))
         CUEX$ = Mid$(F0$, 45, 12)
         CUEMP%(CAMP%) = CUEINT%(CUEX$)
         If TRIM$(CODCUE$(CUEMP%(CAMP%))) <> TRIM$(CUEX$) Then
            Call MENSERR(24, "Cuenta Contable Incorrecta\para Medio de Pago '" + TRIM$(Mid$(F0$, 3, 12)) + "'\-\Resultados Imprevisibles.")
            Call FINAL("")
         End If
     Next I&
   End If
   '
   FINIPDIF% = FECHANUM(CONTROL$("", "FINIPDIF"))
   HOII% = HOY(HOS$)
   FEX = FINIPDIF%
   DES% = DIPOST(FEX)
   VDEF$ = MKI$(DES%) + MKI$(HOII%)
   '
5  OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then Exit Sub
   '
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   If HAS% < DES% Then Exit Sub
   '
   If DES% <= FINIPDIF% Then
      Call MENSERR(24, "Fecha no Válida")
      GoTo 5
   End If
   '
   'If HAS% > FINEJ% Then
   '   Call MENSERR(24, "Fecha 'Hasta' es Posterior\al Fin del Ejercicio.")
   '   GoTo 5
   'End If
   '
10 'Call SELECHO("ECUECON")
   'CUEX$ = VALACT1$("ECUECON")
   'If CUEX$ = "" Then Exit Sub
   '
   Screen.MousePointer = 11
   KLM& = 0
   '
   For cuebal& = 1 To ULTREG&("CUEBANC")
   CBB$ = REGLEIDO$("CUEBANC", cuebal&)
   DCBB$ = Left$(CBB$, 32)
   CUEX$ = Left$(CBB$, 12)
   SALINI# = CVD(Mid$(CBB$, 63, 8))
   CUE% = CUEINT%(CUEX$)
   If CUE% < 1 Then
      Call MENSERR(24, "Cuenta Inexistente")
      GoTo 80
   End If
   '
   Frame1.Refresh
   Call FRESHALL
   CUEBIN$ = MKI$(CUE%)
   CODENOCUE$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%))
   JJ& = 0
   'SALINI# = 0
   '
   'If DES% >= FECHANUM("01/10/01") Then
   '   CCCC$ = TRIM$(CODCUE$(CUE%))
   '   If CCCC$ = "111401" Then
   '      SALINI# = 10886.85   ' 11409.33
   '      SALINI# = SALINI# + 114.95 ' para tener en cuenta cheque duplicado flama
   '   End If
   '   If CCCC$ = "111502" Then SALINI# = 11791.1
   '   If CCCC$ = "111403" Then SALINI# = 13039.69   ' 2414.3
   '   If CCCC$ = "111406" Then SALINI# = 30026.22   ' 29769.7
   '   If CCCC$ = "111506" Then SALINI# = 381.86
   'End If
   '
   ' CARGAR MOVIMIENTOS DE DIARIO
   EXTEA$ = EXTE$
   If EXTEA$ = "DAT" Then
     Close #NX9%: NX9% = 0
     Close #NX10%: NX10% = 0
     Close #NX11%: NX11% = 0
     EXTE$ = "OLD"
   End If
   '
12 PMC$ = PUNMODIAR$(CUE%)
   FIN& = Len(PMC$)
   For KK% = 1 To Len(PMC$) Step 2
      UU& = KK% + 1
      Call TRACE(20, UU&, FIN&)
      REGDIA% = CVI(Mid$(PMC$, KK%, 2))
      Call LEEDIARIO(REGDIA%)
      If CVS(E12$) < 1.05 Then GoTo 19 ' SALTEA ASIENTO APERTURA
      If Asc(E18$) <> 1 Then  ' SALTEA ASIENTOS RESUMEN DE CAJA
         '
         Call LEEPASDIA(REGDIA%, 2)
         DETAPAS$ = TRIM$(DETAPAS$)
         If DETAPAS$ <> "" Then
            DETAPAS$ = TRIM$(DETAPAS$) + " - "
         End If
         DETAPAS$ = DETAPAS$ + REFASIEN$
         '
         ACIERR% = 0
         If InStr(UCase$(DETAPAS$), "(CIERR") > 0 Or InStr(UCase$(DETAPAS$), "(*") > 0 Then
           ACIERR% = 1
         End If
         MARCADIA$ = "*"
         FEMOREA% = CVI(E15$)
         IMPO# = CVD(E16$)
         If FEMOREA% < DES% Then
             If ACIERR% = 1 Then IMPO# = 0
             'If FEMOREA% > FECHANUM("31/09/01") Then GoSub 90 'TOTALIZAR INICIAL
             If FEMOREA% > FINIPDIF% Then GoSub 90 'TOTALIZAR INICIAL
          ElseIf FEMOREA% <= HAS% Then
            FEMORE% = FEMOREA%
            FEREGI% = FEMOREA%
            GoSub 95
         End If
      End If
19 Next KK%
   '
   If EXTE$ = "OLD" Then
     If EXTEA$ = "DAT" Then
       Close #NX9%: NX9% = 0
       Close #NX10%: NX10% = 0
       Close #NX11%: NX11% = 0
       EXTE$ = "DAT"
       EXTEA$ = ""
       GoTo 12
     End If
   End If
   '
   '
20 ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   EXTEA$ = EXTE$
   If EXTEA$ = "DAT" Then
     Close #NX5%: NX5% = 0
     Close #NX6%: NX6% = 0
     Close #NX7%: NX7% = 0
     EXTE$ = "OLD"
   End If
   '
22 PMC$ = PUNMOCAJA$(CUE%)
   FIN& = Len(PMC$)
   For KK% = 1 To Len(PMC$) Step 4
      UU& = KK% + 1
      Call TRACE(20, UU&, FIN&)
      REGDIAL& = CVS(Mid$(PMC$, KK%, 4))
      Call LEEPASCAJ(REGDIAL&, 1)
      FEMOVI% = CVI(D15$)
      IMPO# = CVD(D16$)
      '
      If TRIM$(DETAPAS$) = "" Then
         Call LEEPASCAJ(REGDIAL&, 2)
         DETAPAS$ = REFASIEN$
      End If
      MARCADIA$ = ""
      If InStr(DETAPAS$, "- OP") > 0 Then GoTo 29
      If Left$(DETAPAS$, 8) = "DEPOSITO" Then
         FEMORE% = FECHACRE%(FEMOVI%, Val(Mid$(DETAPAS$, 10, 3)))
         If FEMORE% < DES% Then
            'If FEMORE% > FECHANUM("31/09/01") Then GoSub 90
            If FEMORE% > FINIPDIF% Then GoSub 90
         ElseIf FEMORE% <= HAS% Then
            GoSub 95
         End If
         GoTo 29
      End If
      If FEMOVI% < DES% Then
        'If FEMOVI% > FECHANUM("31/09/01") Then GoSub 90
        If FEMOVI% > FINIPDIF% Then GoSub 90
        GoTo 29
      End If
      If FEMOVI% >= DES% Then
        If FEMOVI% <= HAS% Then
          FEMORE% = FEMOVI%
          GoSub 95
        End If
      End If
      '
29 Next KK%
   '
   If EXTE$ = "OLD" Then
     If EXTEA$ = "DAT" Then
       Close #NX5%: NX5% = 0
       Close #NX6%: NX6% = 0
       Close #NX7%: NX7% = 0
       EXTE$ = "DAT"
       EXTEA$ = ""
       GoTo 22
     End If
   End If
   '
   For U& = 1 To ULTREG&("CHEQUEMI")
     X$ = REGLEIDO$("CHEQUEMI", U&)
      PROV% = CVI(Left$(X$, 2))
      NUMERO$ = Mid$(X$, 3, 12)
      BANCO$ = Mid$(X$, 15, 16)
      EMIS% = CVI(Mid$(X$, 47, 2)) ' toma fecha de O/Pago
      IMPO# = (-1) * CVD(Mid$(X$, 33, 8))
      NOPAGO& = CVS(Mid$(X$, 41, 4))
      VENCI% = CVI(Mid$(X$, 49, 2))
      '
      If VENCI% < EMIS% Then  ' modificacion del 28/01/02. Si
         VENCI% = EMIS%       ' la fecha del valor es anterior
      End If                  ' a la OP, toma la de la OP
      '
      QUE% = CVI(Mid$(X$, 45, 2))
            For UX% = 1 To CAMP%
               If MEPA%(UX%) = QUE% Then GoTo 38
            Next UX%
            Call MENSERR(24, "Medio de Pago" + Str$(QUE%) + " no Definido")
            Call FINAL("")
            '
38    CUEH% = CUEMP%(UX%)
      If CUEH% = CUE% Then
        FEMORE% = VENCI%
'If EMIS% <= FECHANUM("31/09/01") And VENCI% > FECHANUM("31/09/01") Then
If EMIS% <= FINIPDIF% And VENCI% > FINIPDIF% Then
SALINICA# = SALINICA# + IMPO#
AAA = BBB
End If
        If FEMORE% < DES% Then
          'If FEMORE% > FECHANUM("31/09/01") Then GoSub 90
          If FEMORE% > FINIPDIF% Then GoSub 90
        End If
        If FEMORE% >= DES% Then
          If FEMORE% <= HAS% Then
            FEX = EMIS%
            DETAPAS$ = "Ch." + TRIM$(NUMERO$)
            DETAPAS$ = DETAPAS$ + " - OP." + TRIM$(Str$(NOPAGO&))
            DETAPAS$ = DETAPAS$ + " - " + TRIM$(RASOCLI$((-1) * PROV%))
            GoSub 95
          End If
        End If
      End If
   Next U&
   '
   FEX = DES%
   FECHAMO%(0) = DIANTE(FEX)
   FEX = FECHAMO(0)
   REFE$(0) = "Saldo al " + FECHATEX$(FEX)
   IPODIF#(0) = SALINI#
   '
   For KKL& = 0 To JJ&
     For KKL1& = KKL& + 1 To JJ&
       If FECHAMO%(KKL1&) < FECHAMO%(KKL&) Then
         FX1% = FECHAMO%(KKL&)
         FX2$ = REFE$(KKL&)
         FX3# = IPODIF#(KKL&)
         FECHAMO%(KKL&) = FECHAMO%(KKL1&)
         REFE$(KKL&) = REFE$(KKL1&)
         IPODIF#(KKL&) = IPODIF#(KKL1&)
         FECHAMO%(KKL1&) = FX1%
         REFE$(KKL1&) = FX2$
         IPODIF#(KKL1&) = FX3#
        End If
     Next KKL1&
     GoSub 96
   Next KKL&
AAA = SALINICA#
ABC$ = CCCC$
SALINICA# = 0
80 Next cuebal&
   '
   DCBB$ = "Cartera de Valores"
   JJ& = (-1)
   For U& = 1 To ULTREG&("BADACHE")
     X$ = REGLEIDO$("BADACHE", U&)
     If CVI(Mid$(X$, 133, 2)) < 1 Then  ' aun en cartera
       FEMORE% = CVI(Mid$(X$, 59, 2))
       If FEMORE% <= HAS% Then
         DETAPAS$ = Left$(X$, 40)
         IMPO# = CVD(Mid$(X$, 49, 8))
         GoSub 95
       End If
     End If
   Next U&
   '
   For KKL& = 0 To JJ&
     For KKL1& = KKL& + 1 To JJ&
       If FECHAMO%(KKL1&) < FECHAMO%(KKL&) Then
         FX1% = FECHAMO%(KKL&)
         FX2$ = REFE$(KKL&)
         FX3# = IPODIF#(KKL&)
         FECHAMO%(KKL&) = FECHAMO%(KKL1&)
         REFE$(KKL&) = REFE$(KKL1&)
         IPODIF#(KKL&) = IPODIF#(KKL1&)
         FECHAMO%(KKL1&) = FX1%
         REFE$(KKL1&) = FX2$
         IPODIF#(KKL1&) = FX3#
        End If
     Next KKL1&
     GoSub 96
   Next KKL&
   '
   Call SETULTREG("PODIFER", KLM&)
   Call CIERRARCH("PODIFER")
   Call CIERRARCH("*.*")
   '
   Screen.MousePointer = 1
   Call FINAL("*PODIFER")
   '
   Exit Sub
   
90 SALINI# = SALINI# + IMPO#
   Return
   '
95 If Abs(IMPO#) >= 0.005 Then
      JJ& = JJ& + 1
      REFE$(JJ&) = DETAPAS$
      FECHAMO%(JJ&) = FEMORE%
      IPODIF#(JJ&) = IMPO#
   End If
   Return
   '
96 ZZZ$ = REGBLAN$("PODIFER")
   If DCBB$ <> DCBBA$ Then
     SARRAS# = 0
     DCBBA$ = DCBB$
   End If
   Call REPLA(ZZZ$, DCBB$, 1, 32)
   Call REPLA(ZZZ$, MKI$(FECHAMO%(KKL&)), 33, 2)
   Call REPLA(ZZZ$, REFE$(KKL&), 35, 30)
   ACIERR% = 0
   If InStr(UCase$(REFE$(KKL&)), "(CIERR") > 0 Or InStr(REFE$(KKL&), "(*") > 0 Then
      ACIERR% = 1
   End If
   SARRAS# = SARRAS# + IPODIF#(KKL&)
   If ACIERR% > 0 Then SARRAS# = SARRAS# - IPODIF#(KKL&)
   If IPODIF#(KKL&) > 0 Then
      Call REPLA(ZZZ$, MKD$(IPODIF#(KKL&)), 65, 8)
      If ACIERR% > 0 Then Call REPLA(ZZZ$, MKD$(Abs(IPODIF#(KKL&))), 73, 8)
     Else
      Call REPLA(ZZZ$, MKD$(Abs(IPODIF#(KKL&))), 73, 8)
      If ACIERR% > 0 Then Call REPLA(ZZZ$, MKD$(Abs(IPODIF#(KKL&))), 65, 8)
   End If
   Call REPLA(ZZZ$, MKD$(SARRAS#), 81, 8)
   KLM& = KLM& + 1
   Call GRAREG("PODIFER", ZZZ$, KLM&)
   Return
   '
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   If DERACCE%("SALIPDIF", "Saldos Iniciales Posición Diferida") = 2 Then
      FEXI% = FECHANUM(CONTROL$("", "FINIPDIF"))
      VDEF$ = MKI$(FEXI%)
      OBX$ = OBJPLA$("FINIPDIF", VDEF$)
      If OBX$ <> "" Then
        FEX = CVI(OBX$)
        If FEX > 0 Then
          FEXI% = FEX
          Call TRALOCAL("CUEBANC", "")
          For U& = 1 To ULTREG&("!CUEBANC")
            XX$ = REGLEIDO$("!CUEBANC", U&)
            Call REPLA(XX$, MKI$(FEXI%), 61, 2)
            Call GRAREG("!CUEBANC", XX$, U&)
          Next U&
          Call CIERRARCH("!CUEBANC")
          VTB% = VENTABU%("SINIPDIF/Saldos Pos.Diferida al " + OBX$)
          If VTB% >= 0 Then
            Call TRACENTRAL("CUEBANC", "")
            Call GRACONTROL("", "FINIPDIF", FECHATEX$(FEX))
          End If
        End If
      End If
   End If
   Command9.Enabled = True
End Sub

Private Sub Option3_DblClick()
   Call INDEXCAJA(MAXDIA&)
   '
   DES% = COMEJ%
   HAS% = FINEJ%
   '
10 Call SELECHO("ECUECON")
   CUEX$ = VALACT1$("ECUECON")
   If CUEX$ = "" Then Exit Sub
   '
   CUE% = CUEINT%(CUEX$)
   If CUE% < 1 Then
      Call MENSERR(24, "Cuenta Inexistente")
      GoTo 10
   End If
   '
   Call FRESHALL
   CUEBIN$ = MKI$(CUE%)
   CODENOCUE$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%))
   JJ& = 0: TRANSPORTE# = 0
   '
   ' CARGAR MOVIMIENTOS DE DIARIO
   PMC$ = PUNMODIAR$(CUE%)
   FIN& = Len(PMC$)
   For KK% = 1 To Len(PMC$) Step 2
      UU& = KK% + 1
      Call TRACE(20, UU&, FIN&)
      REGDIA% = CVI(Mid$(PMC$, KK%, 2))
      Call LEEDIARIO(REGDIA%)
      If Asc(E18$) <> 1 Then  ' SALTEA ASIENTOS RESUMEN DE CAJA
         Call LEEPASDIA(REGDIA%, 1)
         If TRIM$(DETAPAS$) = "" Then
            Call LEEPASDIA(REGDIA%, 2)
            DETAPAS$ = REFASIEN$
         End If
         MARCADIA$ = "*"
         If FECHASIEN% < DES% Then
             TRANSPORTE# = TRANSPORTE# + IMPOPAS#
             GoTo 19
           ElseIf FECHASIEN% <= HAS% Then
             GoSub 90
         End If
      End If
19 Next KK%
   '
   ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   PMC$ = PUNMOCAJA$(CUE%)
   FIN& = Len(PMC$)
   For KK% = 1 To Len(PMC$) Step 4
      UU& = KK% + 1
      Call TRACE(20, UU&, FIN&)
      REGDIAL& = CVS(Mid$(PMC$, KK%, 4))
      Call LEEPASCAJ(REGDIAL&, 1)
      If TRIM$(DETAPAS$) = "" Then
         Call LEEPASCAJ(REGDIAL&, 2)
         DETAPAS$ = REFASIEN$
      End If
      MARCADIA$ = ""
      If FECHASIEN% < DES% Then
          TRANSPORTE# = TRANSPORTE# + IMPOPAS#
          GoTo 29
        ElseIf FECHASIEN% <= HAS% Then
          GoSub 90
      End If
29 Next KK%
   '
   If Abs(TRANSPORTE#) >= 0.005 Then
      FEX = DES%
      FECHASIEN% = DIANTE(FEX)
      NROASIEN = 0
      NROASDEF% = 0
      MARCADIA$ = " "
      DETAPAS$ = "SALDO ANTERIOR"
      IMPOPAS# = TRANSPORTE#
      GoSub 90
   End If
   '
   Call SETULTREG("RESUCON", JJ&)
   Call CIERRARCH("RESUCON")
   '
   ' hasta aqui genero resume de cuenta de C&B en RESUCON.WKF
   '
   '
   Call GEPODIF(CUE%)
   '
   'hasta aqui genero posicion diferida en PODIFER.WKF
   '
   '
   UPO& = ULTREG&("PODIFER")
   FIN& = ULTREG&("RESUCON")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("RESUCON", U&)
     IDERE# = CVD(Mid$(XX$, 105, 8))
     IHARE# = CVD(Mid$(XX$, 113, 8))
     '
     For KKU& = 1 To UPO&
       YY$ = REGLEIDO$("PODIFER", KKU&)
       If Abs(CVD(Mid$(YY$, 65, 8)) - IDERE#) < 0.005 Then
         If Abs(CVD(Mid$(YY$, 73, 8)) - IHARE#) < 0.005 Then
           Call GRAREG("PODIFER", String$(96, 0), KKU&)
           Call GRAREG("RESUCON", String$(128, 0), U&)
           GoTo 59
         End If
       End If
     Next KKU&
     '
59 Next U&
   '
   JJ& = 0
   For U& = 1 To ULTREG&("PODIFER")
     XX$ = REGLEIDO$("PODIFER", U&)
     If XX$ <> String$(96, 0) Then
       JJ& = JJ& + 1
       Call GRAREG("PODIFER", XX$, JJ&)
     End If
   Next U&
   Call SETULTREG("PODIFER", JJ&)
   Call CIERRARCH("PODIFER")
   '
   JJ& = 0
   For U& = 1 To ULTREG&("RESUCON")
     XX$ = REGLEIDO$("RESUCON", U&)
     If XX$ <> String$(128, 0) Then
       JJ& = JJ& + 1
       Call GRAREG("RESUCON", XX$, JJ&)
     End If
   Next U&
   Call SETULTREG("RESUCON", JJ&)
   Call CIERRARCH("RESUCON")
   '
   Call CIERRARCH("*.*")
80 OPX% = COMALTER%("Conciliacion Completa\\Pos.Diferida\Res.Cuenta\Terminar")
   If OPX% = 1 Then
       Call FINAL("*CPODIFE")
       GoTo 80
     ElseIf OPX% = 2 Then
       Call FINAL("*CRESUCO")
       GoTo 80
   End If
   Exit Sub
   '
   '
90 XX$ = REGBLAN$("RESUCON")
   Call REPLA(XX$, CODENOCUE$, 1, 64)
   Call REPLA(XX$, MKI$(FECHASIEN%), 65, 2)
   Call REPLA(XX$, MKS$(NROASIEN), 67, 4)
   Call REPLA(XX$, MKI$(NROASDEF%), 71, 2)
   Call REPLA(XX$, MARCADIA$, 73, 2)
   Call REPLA(XX$, DETAPAS$, 75, 30)
   If IMPOPAS# >= 0 Then
         Call REPLA(XX$, MKD$(IMPOPAS#), 105, 8)
      Else
         Call REPLA(XX$, MKD$(Abs(IMPOPAS#)), 113, 8)
   End If
   JJ& = JJ& + 1
   Call GRAREG("RESUCON", XX$, JJ&)
   '
Return
'
End Sub

Private Sub Form_Load()
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
   End If
   If InStr(UCase$(EPAC$), "DOSIVA") < 1 Then
     Frame3.Top = Frame3.Top - 600
     Frame1.Height = Frame1.Height - 600
     Height = Height - 600
     Command9.Visible = False
     DoEvents
   End If
   Option6.Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
End Sub

Sub BUSCACHEQ()
   '
30 NCHE$ = OBJPLA$("CARNUCHE", NCHE$)
   If NCHE$ = "" Then
      Exit Sub
   End If
   '
   NCHE$ = TRIM$(NCHE$)
   Screen.MousePointer = 11
   Call FRESHECHO("CHESELEC")
   JJ& = 0
   FIN& = ULTREG&("BADACHE")
   '
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("BADACHE", I&)
      NCHX$ = Mid$(X$, 19, 12)
      '
      If NCHE$ = "" Then
         If TRIM$(NCHX$) = "" Then
            Call REPLA(X$, MKS$(I&), 45, 4)
            JJ& = JJ& + 1
            Call GRAREG("CHESELEC", X$, JJ&)
         End If
         GoTo 39
      End If
      '
      If InStr(NCHX$, NCHE$) > 0 Then
         Call REPLA(X$, MKS$(I&), 45, 4)
         JJ& = JJ& + 1
         Call GRAREG("CHESELEC", X$, JJ&)
      End If
39 Next I&
   '
   Call SETULTREG("CHESELEC", JJ&)
   Call CIERRARCH("CHESELEC")
   Screen.MousePointer = 1
   '
   CASEL& = ULTREG&("CHESELEC")
   If CASEL& < 1 Then
         Call COMUNI("Número no Encontrado")
         GoTo 30
       ElseIf CASEL& = 1 Then
         Y$ = REGLEIDO$("CHESELEC", 1&)
         GoTo 50
   End If
   '
   Y$ = REGSEL$("CHESELEC")
   If Len(Y$) <= 4 Then GoTo 30
   '
50 X$ = OBJPLA$("MUESTRACHEQUE", Y$)
   If X$ <> "" Then
       '
       If X$ <> Y$ Then
          Call MENSERR(24, "Referencias no Modificables")
          GoTo 50
       End If
       '
       REG& = CVS(Mid$(X$, 45, 4))
       Call REPLA(X$, String$(4, 0), 45, 4)
       If REG& > 0 Then
          If REG& <= ULTREG&("BADACHE") Then
             Call GRAREG("BADACHE", X$, REG&)
          End If
       End If
   End If
   GoTo 30
   '
End Sub
'

Sub CARGACHECAR()
   '
   HOII% = HOY(HOS$)
   '
10 CHENUE$ = OBJPLA$("CARGACHEQUE", CHENUE$)
   If CHENUE$ = "" Then Exit Sub
   '
   REFE$ = TRIM$(Left$(CHENUE$, 30))
   If REFE$ <> "" Then
     If CVD(Mid$(CHENUE$, 49, 8)) > 0.01 Then
       If CVI(Mid$(CHENUE$, 57, 2)) > 0 Then
         If CVI(Mid$(CHENUE$, 59, 2)) > 0 Then
           If CVI(Mid$(CHENUE$, 61, 2)) > 0 Then  ' FECHA DE RECEPCION
             If CVI(Mid$(CHENUE$, 61, 2)) > HOII% Then
               Call MENSERR(24, "Fecha de Recepción no Puede ser\Posterior a Fecha Actual.")
               GoTo 10
             End If
             If TRIM$(Mid$(CHENUE$, 63, 10)) <> "" Then
               If TRIM$(Mid$(CHENUE$, 73, 18)) <> "" Then
                 RECIDE$ = TRIM$(Mid$(CHENUE$, 73, 17)) + " (Ing.Manual)"
                 Call REPLA(CHENUE$, RECIDE$, 73, 30)
                 FEX = CVI(Mid$(CHENUE$, 59, 2))
                 FEVE$ = FECHATEX$(FEX)
                 Call REPLA(CHENUE$, FEVE$, 33, 8)
                 Call REGAPP("BADACHE", CHENUE$)
                 Call CIERRARCH("BADACHE")
                 CHENUE$ = ""
                 Call COMUNI("Carga Completa")
                 GoTo 10
               End If
             End If
           End If
         End If
       End If
     End If
   End If
   '
   Call MENSERR(24, "Datos Incompletos !!!\  \Corrija o Agregue los Faltantes.")
   GoTo 10
   '
End Sub
'

Sub ENTREGACHE()
   '
30 NCHE$ = OBJPLA$("CARNUCHE", NCHE$)
   If NCHE$ = "" Then
      Exit Sub
   End If
   '
   NCHE$ = TRIM$(NCHE$)
   Screen.MousePointer = 11
   Call FRESHECHO("CHESELEC")
   JJ& = 0
   FIN& = ULTREG&("BADACHE")
   '
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("BADACHE", I&)
      NCHX$ = Mid$(X$, 19, 12)
      If CVI(Mid$(X$, 133, 2)) < 1 Then  'aun en cartera
         '
         If NCHE$ = "" Then
            If TRIM$(NCHX$) = "" Then
               Call REPLA(X$, MKS$(I&), 45, 4)
               JJ& = JJ& + 1
               Call GRAREG("CHESELEC", X$, JJ&)
            End If
            GoTo 39
         End If
         '
         If InStr(NCHX$, NCHE$) > 0 Then
            Call REPLA(X$, MKS$(I&), 45, 4)
            JJ& = JJ& + 1
            Call GRAREG("CHESELEC", X$, JJ&)
         End If
         '
      End If
39 Next I&
   '
   Call SETULTREG("CHESELEC", JJ&)
   Call CIERRARCH("CHESELEC")
   Screen.MousePointer = 1
   '
   CASEL& = ULTREG&("CHESELEC")
   If CASEL& < 1 Then
         Call COMUNI("Número no Encontrado en Cartera Actual.")
         GoTo 30
       ElseIf CASEL& = 1 Then
         Y$ = REGLEIDO$("CHESELEC", 1&)
         GoTo 50
   End If
   '
   Y$ = REGSEL$("CHESELEC")
   If Len(Y$) <= 4 Then GoTo 30
   '
50 HOII% = HOY(HOS$)
   REFESAL$ = MKI$(HOII%) + Space$(120)
   Call REPLA(REFESAL$, "Salida Manual de Cartera.", 3, 40)
   Call REPLA(REFESAL$, HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 43, 40)
   Call REPLA(REFESAL$, "Usuario: " + TRIM$(USERNAME$), 83, 40)
   Call REPLA(Y$, REFESAL$, 133, 122)
   '
60 X$ = OBJPLA$("MUESTRACHEQUE", Y$)
   If X$ <> "" Then
      If TRIM$(Mid$(X$, 103, 30)) = "" Then
         Call MENSERR(24, "Falta Indicar Destino del Valor.")
         GoTo 60
      End If
      '
      If Mid$(X$, 135, 120) <> Mid$(REFESAL$, 3, 120) Then
         Call MENSERR(24, "Referencias no Modificables")
         GoTo 60
      End If
      '
      If CVI(Mid$(X$, 133, 2)) > 0 Then
         REG& = CVS(Mid$(X$, 45, 4))
         Call REPLA(X$, String$(4, 0), 45, 4)
         If REG& > 0 Then
           If REG& <= ULTREG&("BADACHE") Then
             Call GRAREG("BADACHE", X$, REG&)
           End If
         End If
      End If
   End If
   GoTo 30
   '
End Sub
'

Sub LICHECAR()
   '
   CUECHECA$ = CONTROL$("", "CUECHECA")
   CUEICHEQ% = CUEINT%(CUECHECA$)
   '
   OPX% = 1
   If Option4.Value = True Then OPX% = 2
   '
   HASI% = HOY(HOS$)
   FEX = HAS%
   HASS$ = FECHATEX$(FEX)
   DESS$ = "01/01/80"
   DESI% = FECHANUM(DESS$)
   VDEF$ = MKI$(DESI%) + MKI$(HASI%)
   '
10 OBX$ = OBJPLA$("DESHESFECHREC", VDEF$)
   If OBX$ = "" Then
      Exit Sub
   End If
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   
   Screen.MousePointer = 11
   ASALID$ = "CHESELEC"
   If OPX% = 1 Then ASALID$ = "!" + ASALID$
   JJ& = 0
   FIN& = ULTREG&("BADACHE")
   FORMALTER.LINOSORT.Clear
   '
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("BADACHE", I&)
      If CVI(Mid$(X$, 133, 2)) < 1 Then
        FEVENCHE% = CVI(Mid$(X$, 59, 2))
        Z$ = CSTRING$(MKI$(FEVENCHE%), 1, 6, 0, 2) + CSTRING$(MKS$(I&), 4, 8, 0, 2)
        For JLI% = 1 To FORMALTER.LINOSORT.ListCount
          If FORMALTER.LINOSORT.List(JLI% - 1) > Z$ Then
            FORMALTER.LINOSORT.AddItem Z$, JLI% - 1
            GoTo 19
          End If
        Next JLI%
        FORMALTER.LINOSORT.AddItem Z$
      End If
19 Next I&
   '
   SUMACHE# = 0
   FIN& = FORMALTER.LINOSORT.ListCount
   For JI& = 1 To FIN&
      I& = Val(Mid$(FORMALTER.LINOSORT.List(JI& - 1), 7, 8))
      X$ = REGLEIDO$("BADACHE", I&)
      If CVI(Mid$(X$, 133, 2)) < 1 Then
        FERECEP% = CVI(Mid$(X$, 61, 2))
        If FERECEP% >= DES% Then
          If FERECEP% <= HAS% Then
            SUMACHE# = SUMACHE# + CVD(Mid$(X$, 49, 8))
            FEX = CVI(Mid$(X$, 59, 2))
            MESA$ = Mid$(FECHATEX$(FEX), 4, 5)
            Call REPLA(X$, String$(154, 0), 103, 154)
            Call REPLA(X$, MESA$, 103, 6)
            Call REPLA(X$, MKI$(1), 117, 2)
            JJ& = JJ& + 1
            Call GRAREG(ASALID$, X$, JJ&)
          End If
        End If
      End If
   Next JI&
   '
   Call SETULTREG(ASALID$, JJ&)
   Call CIERRARCH(ASALID$)
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DOSIVA") > 0 Then
     If EJER% > 0 Then
       If DES% = DESI% Then
         If HAS% = HASI% Then
           SARRACU# = SARRASCUE#(CUEICHEQ%)
           DIFESA# = Abs(SARRACU# - SUMACHE#)
           If DIFESA# >= 0.005 Then
             Call MENSERR(24, "Hay una diferencia de " + TRIM$(FORMATNUM$(DIFESA#, "F12.2")) + " entre\la B.D.Cheques y la Cuenta 'Valores a Depositar'.\Investigue !!!")
           End If
         End If
       End If
     End If
   End If
   '
   Screen.MousePointer = 1
   '
   If OPX% = 1 Then
        VTB% = VENTABU%("CHESELEC")
      Else
        Call FINAL("*SOCHECAR")
   End If
   '
End Sub

Sub LICHEQEM()
   '
   OPX% = 2
   'If Option5.Value = True Then
   '   If HOY(HOS$) < FECHANUM("08/11/01") Then
   '     NOPANT& = 0
   '     For UU& = 1 To ULTREG&("CHEQUEMI")
   '       XX$ = REGLEIDO$("CHEQUEMI", UU&)
   '       NOPA& = CVS(Mid$(XX$, 41, 4))
   '       If NOPA& < NOPANT& Then
   '         Call CONECRUN("*SORCHEM", "Ordenando por Numero de O.P.")
   '         Exit Sub
   '       End If
   '       NOPANT& = NOPA&
   '     Next UU&
   '   End If
   '   Call CONECRUN("*COCHEMI", "Consulta Cheques Emitidos")
   '   Exit Sub
   'End If
   '
   Screen.MousePointer = 11
   Call COPYSTRU("CHEQUEMI", "CHEQUEMJ")
   ASALID$ = "LICHEQEM"
   If OPX% = 1 Then ASALID$ = "!" + ASALID$
   HOII% = HOY(HOS$)
   JJ& = 0
   FIN& = ULTREG&("CHEQUEMI")
   ' FLIMEM% = HOII% - 32 * 13   ' CHEQUES VENCIDOS MAS DE UN AÑO
   '
   Call SELECHO("SELFECHA/Emitidos al")
   FF11$ = VALACT1$("SELFECHA")
   If FF11$ = "" Then GoTo 99
   DESEM% = FECHANUM(FF11$)
   '
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("CHEQUEMI", I&)
      PROV% = CVI(Left$(X$, 2))
      NUMERO$ = Mid$(X$, 3, 12)
      BANCO$ = Mid$(X$, 15, 16)
      EMIS% = CVI(Mid$(X$, 47, 2))   ' EMISION=FECHA O/PAGO
      If EMIS% < 1 Then
        EMIS% = CVI(Mid$(X$, 31, 2))
      End If
      IMPO# = CVD(Mid$(X$, 33, 8))
      VENCI% = CVI(Mid$(X$, 49, 2))
      '
      If CVI(Mid$(X$, 45, 2)) <> 999 Then ' excluye cheques terceros
        If VENCI% > DESEM% Then
          If EMIS% >= FLIMEM% Then
            If EMIS% <= DESEM% Then
              If TRIM$(BANCO$) <> "" Then
                If InStr(BANCO$, "EFECT") < 1 Then
                  If InStr(UCase$(BANCO$), "EFECT") < 1 Then
                    If InStr(UCase$(BANCO$), "DOLAR") < 1 Then
                      If InStr(UCase$(BANCO$), "PATACON") < 1 Then
                        Y$ = REGBLAN$("LICHEQEM")
                        Call REPLA(Y$, MKI$(VENCI%), 1, 2)
                        Call REPLA(Y$, BANCO$, 3, 16)
                        Call REPLA(Y$, NUMERO$, 19, 12)
                        Call REPLA(Y$, MKD$(IMPO#), 31, 8)
                        Call REPLA(Y$, MKI$(EMIS%), 39, 2)
                        Call REPLA(Y$, MKI$(PROV%), 41, 2)
                        JJ& = JJ& + 1
                        Call GRAREG(ASALID$, Y$, JJ&)
                        Call GRAREG("CHEQUEMJ", X$, JJ&)
                      End If
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
   Next I&
   '
   Call SETULTREG(ASALID$, JJ&)
   Call CIERRARCH(ASALID$)
   Call SETULTREG("CHEQUEMJ", JJ&)
   Call CIERRARCH("CHEQUEMJ")
   '
   If Option5.Value = True Then
       Call CONECRUN("*COCHEMI", "Consulta Cheques Emitidos")
     Else
       Call CONECRUN("*SOCHEMI", "Listado de Cheques Emitidos")
   End If
99 Screen.MousePointer = 1
   '
End Sub

Sub CHERECHAZA()
   '
   Dim TOFAC#(10), IMV#(10), F%(10), GRU%(10), IPAR#(10), DATCLI$(32)
   Dim CTA%(128), IMPO#(128), DETA$(128)

10 NCHE$ = OBJPLA$("CARNUCHE", NCHE$)
   If NCHE$ = "" Then Exit Sub
   NCHE$ = TRIM$(NCHE$)
   '
   Screen.MousePointer = 11
   '
   Call BLANARCH("!CHESELEC")
   FIN& = ULTREG&("BADACHE")
   '
   For I& = 1 To FIN&
     Call TRACE(20, I&, FIN&)
     X$ = REGLEIDO$("BADACHE", I&)
     If CVI(Mid$(X$, 255, 2)) < 1 Then  ' SALTEA RECHAZADOS
       NCHX$ = Mid$(X$, 19, 12)
        '
        If NCHE$ = "" Then
          If TRIM$(NCHX$) = "" Then
            Call REPLA(X$, MKS$(I&), 45, 4)
            Call REGAPP("!CHESELEC", X$)
          End If
          GoTo 19
        End If
        '
        If InStr(NCHX$, NCHE$) > 0 Then
          Call REPLA(X$, MKS$(I&), 45, 4)
          Call REGAPP("!CHESELEC", X$)
        End If
      End If
19 Next I&
   '
   Call CIERRARCH("!CHESELEC")
   Screen.MousePointer = 1
   '
   CASEL& = ULTREG&("!CHESELEC")
   If CASEL& < 1 Then
         Call COMUNI("Número no Encontrado")
         Exit Sub
       ElseIf CASEL& = 1 Then
         Y$ = REGLEIDO$("!CHESELEC", 1&)
         GoTo 25
   End If
   '
   Call FRESHECHO("!CHESELEC")
   Y$ = REGSEL$("!CHESELEC")
   If Len(Y$) <= 4 Then GoTo 10
   '
25 X$ = OBJPLA$("MUESTRACHEQUE", Y$)
   If X$ = "" Then GoTo 10
   '
   If CVI(Mid$(Y$, 133, 2)) < 1 Then
       Call COMUNI("Cheque Aún en Cartera")
       GoTo 10
   End If
   '
   ICHE# = CVD(Mid$(Y$, 49, 8))
   REFECHE$ = TRIM$(TRIM$(Left$(Y$, 16)) + " " + TRIM$(Mid$(Y$, 19, 12)))
   NCHX$ = Mid$(Y$, 19, 12)
   REBADA& = CVS(Mid$(Y$, 45, 4))
   '
30 PROVE$ = Mid$(Y$, 103, 30)
   For U& = 1 To ULTREG&("PROVED1")
      Z$ = REGLEIDO$("PROVED1", U&)
      If Mid$(Z$, 3) = PROVE$ Then
         PROV% = CVI(Left$(Z$, 2))
         GoTo 31
      End If
   Next U&
   PROV% = 0
   '
31 CLIEN$ = Mid$(Y$, 73, 30)
   For U& = 1 To ULTREG&("DEUDOR1")
      Z$ = REGLEIDO$("DEUDOR1", U&)
      If Mid$(Z$, 3) = CLIEN$ Then
         CLI% = CVI(Left$(Z$, 2))
         GoTo 33
      End If
   Next U&
   CLI% = 0
   '
33 MODO% = COMALTER%("Indique el Origen del Rechazo:\\Proveedor\Banco\Abandonar")
   On MODO% GoTo 64, 70
   GoTo 10
   '
64 HO% = HOY(HOS$)
   VDEF$ = MKI$(PROV%) + MKI$(CLI%) + MKI$(HO%)
   Z$ = OBJPLA$("DEVUCHEPRO", VDEF$)
   If Z$ = "" Then GoTo 10
   '
   PROV% = CVI(Left$(Z$, 2)): If PROV% < 1 Then GoTo 64
   CLI% = CVI(Mid$(Z$, 3, 2)): If CLI% < 1 Then GoTo 64
   FF% = CVI(Mid$(Z$, 5, 2))
   If FF% < COMEJ% Or FF% > HO% Then GoTo 64
   '
   RGXX& = REGICLI&(CLI%)
   If RGXX& < 1 Or RGXX& > ULTREG&("DEUDOR1") Then
      Call MENSERR(24, "Imposible Registrar Rechazo.\Cliente no Encontrado.")
      GoTo 10
   End If
   '
   RGYY& = REGICLI&((-1) * PROV%)
   If RGYY& < 1 Or RGYY& > ULTREG&("PROVED1") Then
      Call MENSERR(24, "Imposible Registrar Rechazo.\Proveedor no Encontrado.")
      GoTo 10
   End If
   '
   Call BLOQARCH("FACTUR")
   Call BLOQARCH("SVD202")
   Call BLOQARCH("DEUDOR2")
   Call BLOQARCH("CUECORR")
   Call BLOQARCH("PROVED2")
   '
   OKEY% = 0
   If CONTROL$("", "CHERECHA") = "NOTADEBI" Then
        CLIX% = CLI%
        Call COMUNI("Recuerde Hacer la NOTA DE DEBITO POR RECHAZO\en la Cuenta " + TRIM$(Str$(CLIX%)) + " - " + TRIM$(RASOCLI$(CLIX%)))
        OKEY% = 1
      Else
        Call GRAJUCLI(CLI%, FF%, ICHE#, OKEY%)
   End If
   Call GRAJUPRO(PROV%, FF%, ICHE#, NCHX$, OKEY%)
   If OKEY% = 2 Then
        Call COMUNI("Rechazo Registrado")
      Else
        Call COMUNI("Rechazo mal Registrado.\Consulte !!!")
   End If
   If REBADA& > 0 Then
     If REBADA& <= ULTREG&("BADACHE") Then
       X$ = REGLEIDO$("BADACHE", REBADA&)
       Call REPLA(X$, MKI$(FF%), 255, 2)
       Call GRAREG("BADACHE", X$, REBADA&)
     End If
   End If
   Call CIERRARCH("BADACHE")
   '
   GoTo 10
   '
70 HO% = HOY(HOS$)
   CLIEN$ = Mid$(Y$, 73, 30)
   For U& = 1 To ULTREG&("DEUDOR1")
      Z$ = REGLEIDO$("DEUDOR1", U&)
      If Mid$(Z$, 3) = CLIEN$ Then
         CLI% = CVI(Left$(Z$, 2))
         GoTo 71
      End If
   Next U&
   CLI% = 0
   '
71 VDEF$ = Space$(24) + MKI$(HO%) + MKI$(CLI%)
   CUECHERA$ = AJUSTI$(CONTROL$("", "CUECHERA"), 12)
   Call REPLA(VDEF$, CUECHERA$, 13, 12)
72 Z$ = OBJPLA$("RECHACHEBA", VDEF$)
   VDEF$ = Z$
   If Z$ = "" Then Exit Sub
   '
   CLI% = CVI(Mid$(Z$, 27, 2)): If CLI% < 1 Then GoTo 72
   FF% = CVI(Mid$(Z$, 25, 2)): If FF% < 1 Then GoTo 72
   '
   For II% = 1 To 128
       CTA%(II%) = 0
       IMPO#(II%) = 0
       DETA$(II%) = ""
   Next II%
   '
   CUEBANCO$ = Left$(Z$, 12): CTA%(2) = 0
   CUECHERA$ = Mid$(Z$, 13, 12): CTA%(1) = 0
   '
   For U& = 1 To ULTREG&("ECUECON")
      X$ = REGLEIDO$("ECUECON", U&)
      If Left$(X$, 12) = CUEBANCO$ Then CTA%(2) = CVI(Mid$(X$, 41, 2))
      If Left$(X$, 12) = CUECHERA$ Then CTA%(1) = CVI(Mid$(X$, 41, 2))
   Next U&
   '
   FF% = CVI(Mid$(Z$, 25, 2)): If FF% < 1 Then Exit Sub
   If CTA%(1) < 1 Then GoTo 72
   If CTA%(2) < 1 Then GoTo 72
   Call GRACONTROL("", "CUECHERA", CUECHERA$)
   '
   IMPO#(1) = ICHE#
   IMPO#(2) = (-1) * ICHE#
   DETA$(1) = REFECHE$
   DETA$(2) = "Cheque Rechazado"
   REFE$ = "Ch.Rech." + REFECHE$
   CONTA% = 2
   '
   FEX = FF%
   If FF% < COMEJ% Then
      Call MENSERR(24, "Fecha Anterior al Comienzo del Ejercicio")
      VDEF$ = Z$
      GoTo 72
   End If
   If FF% > FINEJ% Then
      Call MENSERR(24, "Fecha Posterior al Cierre del Ejercicio")
      VDEF$ = Z$
      GoTo 72
   End If
   '
   If FF% >= FULTAS% Then
      NROAS = Int(ULTAS + 1)
      GoTo 90
   End If
   '
   Screen.MousePointer = 11
   For UI% = MAXENC% To 2 Step -1
      Call LEENCAJA(UI%)
      NAS2 = CVS(D1$): FAS2% = CVI(D3$)
      If FAS2% < FF% Then GoTo 80
      Call LEENCAJA(UI% - 1)
      NAS1 = CVS(D1$): FAS1% = CVI(D3$)
      If FAS1% <= FF% Then
        If FAS2% >= FF% Then
          If (Int(10 * NAS2 + 0.5) - Int(10 * NAS1 + 0.5)) >= 2 Then
            NROAS = (1 + Int(10 * NAS1 + 0.5)) / 10
            Screen.MousePointer = 1
            GoTo 90
          End If
        End If
      End If
   Next UI%
   '
80 Screen.MousePointer = 1
   Call MENSERR(24, "Imposible Registrar Rechazo\en la Fecha Elegida.\  \Fecha Propuesta " + FECHATEX$(FEX))
   VDEF$ = Z$
   Call REPLA(Z$, MKI$(FULTAS%), 25, 2)
   GoTo 72
   '
90 Call GRASICAJA(NROAS, FF%, REFE$, CONTA%, CTA%(), IMPO#(), DETA$())
   If CONTROL$("", "CHERECHA") = "NOTADEBI" Then
        CLIX% = CLI%
        Call COMUNI("Recuerde Hacer la NOTA DE DEBITO POR RECHAZO\en la Cuenta " + TRIM$(Str$(CLIX%)) + " - " + TRIM$(RASOCLI$(CLIX%)))
      Else
        Call GRAJUCLI(CLI%, FF%, ICHE#, OKEY%)
   End If
   If REBADA& > 0 Then
     If REBADA& <= ULTREG&("BADACHE") Then
       X$ = REGLEIDO$("BADACHE", REBADA&)
       Call REPLA(X$, MKI$(FF%), 255, 2)
       Call GRAREG("BADACHE", X$, REBADA&)
     End If
   End If
   Call CIERRARCH("BADACHE")
   '
   Call COMUNI("Rechazo Cheque " + REFECHE$ + "\Registrado y Contabilizado.")
   GoTo 10
   '
End Sub

Sub GRAJUCLI(CLI%, FF%, ICHE#, OKEY%)
   '
   RGXX& = REGICLI&(CLI%)
   If RGXX& < 1 Or RGXX& > ULTREG&("DEUDOR1") Then
      Call MENSERR(24, "Imposible Registrar Rechazo.\Cliente no Encontrado.")
      Exit Sub
   End If
   '
   NAJU& = 1 + ULTREG&("FACTUR")
   '
   Z$ = REGBLAN$("FACTUR")
   Call REPLA(Z$, MKI$(CLI%), 1, 2)
   Call REPLA(Z$, MKS$(NAJU&), 3, 4)
   Call REPLA(Z$, MKI$(FF%), 7, 2)
   Call REPLA(Z$, MKI$(FF%), 9, 2)
   Call REPLA(Z$, MKD$(ICHE#), 15, 8)
   Call REPLA(Z$, MKD$(ICHE#), 39, 8)
   Call REGAPP("FACTUR", Z$)
   '
   Z$ = REGBLAN$("SVD202")
   Call REPLA(Z$, MKI$(CLI%), 1, 2)
   Call REPLA(Z$, MKS$(NAJU&), 5, 4)
   Call REPLA(Z$, MKI$(FF%), 9, 2)
   Call REPLA(Z$, MKD$(ICHE#), 11, 8)
   Call REPLA(Z$, MKI$(FF%), 19, 2)
   Call REPLA(Z$, MKD$(ICHE#), 21, 8)
   Call REGAPP("SVD202", Z$)
   '
   TTXS$ = REGLEIDO$("DEUDOR2", RGXX&)
   SALNUE# = (Int(100 * ICHE# + 0.5)) / 100 + (Int(100 * CVD(Mid$(TTXS$, 109, 8)) + 0.5)) / 100
   Call REPLA(TTXS$, MKD$(SALNUE#), 109, 8)
   Call GRAREG("DEUDOR2", TTXS$, RGXX&)
   '
   Call CIERRARCH("FACTUR")
   Call CIERRARCH("SVD202")
   Call CIERRARCH("DEUDOR2")
   '
   OKEY% = OKEY% + 1
   '
End Sub

Sub GRAJUPRO(PROV%, FF%, ICHE#, NCHE$, OKEY%)
   '
   RGYY& = REGICLI&((-1) * PROV%)
   If RGYY& < 1 Or RGYY& > ULTREG&("PROVED1") Then
      Call MENSERR(24, "Imposible Registrar Rechazo.\Proveedor no Encontrado.")
      Exit Sub
   End If
   '
   NCHEQUE$ = Right$("00000000" + TRIM$(NCHE$), 8)
   TINUVA$ = "AS-X-ReCh-" + NCHEQUE$
   '
39 RECORD$ = MKI$(FF%)
   RECORD$ = RECORD$ + MKI$(0) + MKI$(PROV%)
   RECORD$ = RECORD$ + TINUVA$
   '
   REFE$ = Left$("Ch.Rech." + NCHE$ + Space$(24), 24)
   RECORD$ = RECORD$ + REFE$
   '
   MO$ = "$ "
   RECORD$ = RECORD$ + MO$

   TOTAL$ = MKD$(ICHE#) + String$(56, 0)
   RECORD$ = RECORD$ + TOTAL$
   RECORD$ = RECORD$ + String$(8, 0) + MKD$(ICHE#) + String$(8, 0)

   VENCI$ = MKI$(FF%) + MKD$(ICHE#) + String$(50, 0)
   RECORD$ = RECORD$ + VENCI$
   '
   IMPU$ = String$(80, 9)
   RECORD$ = RECORD$ + IMPU$
   '
   RECORD$ = RECORD$ + String$(11, 0)
   HOO = HOY(HO$)
   While HOO > 32767: HOO = HOO - 10000:   Wend
   HOI% = HOO
   RECORD$ = RECORD$ + MKI$(HOI%)
   '
   HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
   While HP > 32767: HP = HP - 10000: Wend
   HPI% = HP
   RECORD$ = RECORD$ + MKI$(HPI%)
   RECORD$ = RECORD$ + Chr$(USER%)
   RECORD$ = RECORD$ + String$(6, 0)
   '
   Call REGAPP("CUECORR", RECORD$)
   '
   TTXS$ = REGLEIDO$("PROVED2", RGYY&)
   SALNUE# = (Int(100 * CVD(Mid$(TTXS$, 109, 8)) + 0.5) - (Int(100 * ICHE# + 0.5))) / 100
   Call REPLA(TTXS$, MKD$(SALNUE#), 109, 8)
   Call GRAREG("PROVED2", TTXS$, RGYY&)
   '
   Call CIERRARCH("CUECORR")
   Call LIBARCH("CUECORR")
   Call CIERRARCH("PROVED2")
   Call LIBARCH("PROVED2")
   '
   OKEY% = OKEY% + 1
   '
End Sub

Sub CONCIBANC()

   Static CAMP%, CUEMP%(2048), MEPA%(2048)
   '
   Dim REFE$(2048)
   Dim FEREGIS%(2048)
   Dim FECHAMO%(2048)
   Dim IPODIF#(2048)
   '
   If CAMP% <= 0 Then
      For I& = 1 To ULTREG&("MEPAGO")
         F0$ = REGLEIDO$("MEPAGO", I&)
         CAMP% = I&
         MEPA%(CAMP%) = CVI(Left$(F0$, 2))
         CUEX$ = Mid$(F0$, 45, 12)
         CUEMP%(CAMP%) = CUEINT%(CUEX$)
         If TRIM$(CODCUE$(CUEMP%(CAMP%))) <> TRIM$(CUEX$) Then
            Call MENSERR(24, "Cuenta Contable Incorrecta\para Medio de Pago '" + TRIM$(Mid$(F0$, 3, 12)) + "'\-\Resultados Imprevisibles.")
            Call FINAL("")
         End If
     Next I&
   End If
   '
   'CMJJ% = COMEJ%
   'If CMJJ% < FECHANUM("30/09/01") Then CMJJ% = FECHANUM("30/09/01")
   CMJJ% = FECHANUM("30/09/01")
   '
   HOII% = HOY(HOS$)
   DES% = FECHANUM("01" + Mid$(HOS$, 3))
   VDEF$ = MKI$(DES%) + MKI$(HOII%)
   '
5  OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then Exit Sub
   '
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   If HAS% < DES% Then Exit Sub
   '
   If DES% < FECHANUM("01/10/01") Then
      Call MENSERR(24, "Fecha no Válida")
      GoTo 5
   End If
   '
   If DES% < COMEJ% Then
      Call MENSERR(24, "Fecha 'Desde' no Válida\Anterior a Comienzo de Ejercicio")
      GoTo 5
   End If
   '
   If HAS% > FINEJ% Then
      Call MENSERR(24, "Fecha 'Hasta' es Posterior\al Fin del Ejercicio.")
      GoTo 5
   End If
   '
  'For CUEBAL& = 1 To ULTREG&("CUEBANC")
10 CBB$ = REGSEL$("CUEBANC")
   If Len(CBB$) <= 4 Then Exit Sub
   DCBB$ = Left$(CBB$, 32)
   CUEX$ = Left$(CBB$, 12)
   CUE% = CUEINT%(CUEX$)
   If CUE% < 1 Then
      Call MENSERR(24, "Cuenta Inexistente")
      GoTo 10
   End If
   '
   Screen.MousePointer = 11
   KLM& = 0
   '
   Frame1.Refresh
   Call FRESHALL
   CUEBIN$ = MKI$(CUE%)
   CODENOCUE$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%))
   JJ& = 0
   SUMAX# = 0
   SALINI# = 0
   SARRAS# = 0
   '
   '
   ' CARGAR MOVIMIENTOS DE DIARIO
   EXTEA$ = EXTE$
   'If EXTEA$ = "DAT" Then
   '  Close #NX9%: NX9% = 0
   '  Close #NX10%: NX10% = 0
   '  Close #NX11%: NX11% = 0
   '  EXTE$ = "OLD"
   'End If
   '
12 PMC$ = PUNMODIAR$(CUE%)
   FIN& = Len(PMC$)
   For KK% = 1 To Len(PMC$) Step 2
      UU& = KK% + 1
      Call TRACE(20, UU&, FIN&)
      REGDIA% = CVI(Mid$(PMC$, KK%, 2))
      Call LEEDIARIO(REGDIA%)
      ASIAPER% = 0
      If CVS(E12$) < 1.05 Then
        If EXTE$ = "DAT" Then
          ASIAPER% = 1
          ' GoTo 19 ' SALTEA ASIENTO APERTURA
        End If
      End If
      If Asc(E18$) <> 1 Then  ' SALTEA ASIENTOS RESUMEN DE CAJA
         Call LEEPASDIA(REGDIA%, 1)
         If TRIM$(DETAPAS$) = "" Then
            Call LEEPASDIA(REGDIA%, 2)
            DETAPAS$ = REFASIEN$
         End If
         MARCADIA$ = "*"
         FEMOREA% = CVI(E15$)
         IMPO# = CVD(E16$)
         '
         'ACIERR% = 0
         'If InStr(UCase$(DETAPAS$), "(CIERR") > 0 Or InStr(UCase$(DETAPAS$), "(*") > 0 Then
         '  ACIERR% = 1
         'End If
         'If ACIERR% > 0 Then IMPO# = 0
         '
         If FEMOREA% < DES% Or ASIAPER% = 1 Then
            SUMAX# = SUMAX# + IMPO#
            If FEMOREA% >= CMJJ% Then GoSub 90 'TOTALIZAR INICIAL
          ElseIf FEMOREA% <= HAS% Then
            FEMORE% = FEMOREA%
            FEREGI% = FEMOREA%
            GoSub 95
         End If
      End If
19 Next KK%
   '
   'If EXTE$ = "OLD" Then
   '  If EXTEA$ = "DAT" Then
   '    Close #NX9%: NX9% = 0
   '    Close #NX10%: NX10% = 0
   '    Close #NX11%: NX11% = 0
   '    EXTE$ = "DAT"
   '    EXTEA$ = ""
   '    GoTo 12
   '  End If
   'End If
   '
   '
   ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   'EXTEA$ = EXTE$
   'If EXTEA$ = "DAT" Then
   '  Close #NX5%: NX5% = 0
   '  Close #NX6%: NX6% = 0
   '  Close #NX7%: NX7% = 0
   '  EXTE$ = "OLD"
   'End If
   '
22 PMC$ = PUNMOCAJA$(CUE%)
   FIN& = Len(PMC$)
   For KK% = 1 To Len(PMC$) Step 4
      UU& = KK% + 1
      Call TRACE(20, UU&, FIN&)
      REGDIAL& = CVS(Mid$(PMC$, KK%, 4))
      Call LEEPASCAJ(REGDIAL&, 1)
      FEMOVI% = CVI(D15$)
      IMPO# = CVD(D16$)
      If FEMOVI% < DES% Then SUMAX# = SUMAX# + IMPO#
      '
      If TRIM$(DETAPAS$) = "" Then
         Call LEEPASCAJ(REGDIAL&, 2)
         DETAPAS$ = REFASIEN$
      End If
      MARCADIA$ = ""
      If InStr(DETAPAS$, "- OP") > 0 Then GoTo 29
      If Left$(DETAPAS$, 8) = "DEPOSITO" Then
         FEREGI% = FEMOVI%
         FEMORE% = FECHACRE%(FEMOVI%, Val(Mid$(DETAPAS$, 10, 3)))
         If FEREGI% < DES% Then
            If FEREGI% >= CMJJ% Then GoSub 90
          ElseIf FEREGI% <= HAS% Then
            GoSub 95
         End If
         GoTo 29
      End If
      If FEMOVI% < DES% Then
        GoSub 90
        GoTo 29
      End If
      If FEMOVI% >= DES% Then
        If FEMOVI% <= HAS% Then
          FEREGI% = FEMOVI%
          FEMORE% = FEMOVI%
          GoSub 95
        End If
      End If
      '
29 Next KK%
   '
   'If EXTE$ = "OLD" Then
   '  If EXTEA$ = "DAT" Then
   '    Close #NX5%: NX5% = 0
   '    Close #NX6%: NX6% = 0
   '    Close #NX7%: NX7% = 0
   '    EXTE$ = "DAT"
   '    EXTEA$ = ""
   '    GoTo 22
   '  End If
   'End If
   '
   'AAA = SUMAX#
   '
   For U& = 1 To ULTREG&("CHEQUEMI")
     X$ = REGLEIDO$("CHEQUEMI", U&)
      PROV% = CVI(Left$(X$, 2))
      NUMERO$ = Mid$(X$, 3, 12)
      BANCO$ = Mid$(X$, 15, 16)
      EMIS% = CVI(Mid$(X$, 47, 2)) ' toma fecha de O/Pago
      IMPO# = (-1) * CVD(Mid$(X$, 33, 8))
      NOPAGO& = CVS(Mid$(X$, 41, 4))
      If NOPAGO& < 1 Then GoTo 39 ' INCLUIDO POR DOSIVAC - SCARINCI DEL 26/12/04
      If Abs(IMPO#) < 0.005 Then GoTo 39 ' PARA SUPRIMIR O/P's ANULADAS
      '
      VENCI% = CVI(Mid$(X$, 49, 2))
      QUE% = CVI(Mid$(X$, 45, 2))
            For UX% = 1 To CAMP%
               If MEPA%(UX%) = QUE% Then GoTo 38
            Next UX%
            Call MENSERR(24, "Medio de Pago" + Str$(QUE%) + " no Definido\en Orden de Pago Número " + TRIM$(Str$(NOPAGO&)) + ".")
            'Call FINAL("")
            '
38    CUEH% = CUEMP%(UX%)
      If CUEH% = CUE% Then
        FEREGI% = EMIS%
        If FEREGI% < DES% Then
           If FEREGI% >= CMJJ% Then GoSub 90
        End If
        FEMORE% = VENCI%
        If FEREGI% >= DES% Then
          If FEREGI% <= HAS% Then
            FEX = EMIS
            DETAPAS$ = "Ch." + TRIM$(NUMERO$)
            DETAPAS$ = DETAPAS$ + " - OP." + TRIM$(Str$(NOPAGO&))
            DETAPAS$ = DETAPAS$ + " - " + TRIM$(RASOCLI$((-1) * PROV%))
            GoSub 95
          End If
        End If
      End If
39 Next U&
   '
   SALINI# = SUMAX#
   FEX = DES%
   FECHAMO%(0) = DIANTE(FEX)
   FEREGIS%(0) = FECHAMO%(0)
   FEX = FECHAMO%(0)
   REFE$(0) = "Saldo al " + FECHATEX$(FEX)
   IPODIF#(0) = SALINI#
   '
   For KKL& = 0 To JJ&
     For KKL1& = KKL& + 1 To JJ&
       If FEREGIS%(KKL1&) < FEREGIS%(KKL&) Then
         FX1% = FECHAMO%(KKL&)
         FX4% = FEREGIS%(KKL&)
         FX2$ = REFE$(KKL&)
         FX3# = IPODIF#(KKL&)
         FECHAMO%(KKL&) = FECHAMO%(KKL1&)
         FEREGIS%(KKL&) = FEREGIS%(KKL1&)
         REFE$(KKL&) = REFE$(KKL1&)
         IPODIF#(KKL&) = IPODIF#(KKL1&)
         FECHAMO%(KKL1&) = FX1%
         FEREGIS%(KKL1&) = FX4%
         REFE$(KKL1&) = FX2$
         IPODIF#(KKL1&) = FX3#
        End If
     Next KKL1&
     GoSub 96
   Next KKL&
   '
   Call SETULTREG("CONCIBAN", KLM&)
   Call CIERRARCH("CONCIBAN")
   Call CIERRARCH("*.*")
   '
   Screen.MousePointer = 1
   If Option5.Value = True Then
        Call CONECRUN("*CONCIBA/" + AJUSTI$(DCBB$, 30), "Conciliacion Bancaria (Pantalla)")
      Else
        Call HEADERS("CONCIBAN", "")
        Call HEADERS("CONCIBAN", "Banco: " + DCBB$)
        Call CONECRUN("*LCONBAN", "Conciliacion Bancaria (Listado)")
   End If
   GoTo 10
   '
   Exit Sub
   
90 SALINI# = SALINI# + IMPO#
   Return
   '
95 If Abs(IMPO#) >= 0.005 Then
      JJ& = JJ& + 1
      REFE$(JJ&) = DETAPAS$
      FECHAMO%(JJ&) = FEMORE%
      FEREGIS%(JJ&) = FEREGI%
      IPODIF#(JJ&) = IMPO#
   End If
   Return
   '
96 ZZZ$ = REGBLAN$("CONCIBAN")
   If DCBB$ <> DCBBA$ Then
     SARRAS# = 0
     DCBBA$ = DCBB$
   End If
   Call REPLA(ZZZ$, DCBB$, 1, 30)
   Call REPLA(ZZZ$, MKI$(FEREGIS%(KKL&)), 31, 2)
   Call REPLA(ZZZ$, MKI$(FECHAMO%(KKL&)), 33, 2)
   Call REPLA(ZZZ$, REFE$(KKL&), 35, 30)
   SARRAS# = SARRAS# + IPODIF#(KKL&)
   If IPODIF#(KKL&) > 0 Then
      Call REPLA(ZZZ$, MKD$(IPODIF#(KKL&)), 65, 8)
     Else
      Call REPLA(ZZZ$, MKD$(Abs(IPODIF#(KKL&))), 73, 8)
   End If
   Call REPLA(ZZZ$, MKD$(SARRAS#), 81, 8)
   KLM& = KLM& + 1
   Call GRAREG("CONCIBAN", ZZZ$, KLM&)
   Return
   '
End Sub

Sub GEPODIF(CUEKK%)
'
   Static CAMP%, CUEMP%(2048), MEPA%(2048)
   '
   Dim REFE$(2048)
   Dim FECHAMO%(2048)
   Dim IPODIF#(2048)
   '
   If CAMP% <= 0 Then
      For I& = 1 To ULTREG&("MEPAGO")
         F0$ = REGLEIDO$("MEPAGO", I&)
         CAMP% = I&
         MEPA%(CAMP%) = CVI(Left$(F0$, 2))
         CUEX$ = Mid$(F0$, 45, 12)
         CUEMP%(CAMP%) = CUEINT%(CUEX$)
         If TRIM$(CODCUE$(CUEMP%(CAMP%))) <> TRIM$(CUEX$) Then
            Call MENSERR(24, "Cuenta Contable Incorrecta\para Medio de Pago '" + TRIM$(Mid$(F0$, 3, 12)) + "'\-\Resultados Imprevisibles.")
            Call FINAL("")
         End If
     Next I&
   End If
   '
   HOII% = HOY(HOS$)
   DES% = COMEJ%
   HAS% = FINEJ%
   '
10 Screen.MousePointer = 11
   KLM& = 0
   '
   For cuebal& = 1 To ULTREG&("CUEBANC")
     CBB$ = REGLEIDO$("CUEBANC", cuebal&)
     DCBB$ = Left$(CBB$, 32)
     CUEX$ = Left$(CBB$, 12)
     CUE% = CUEINT%(CUEX$)
     If CUE% = CUEKK% Then GoTo 11
   Next cuebal&
   Call MENSERR(24, "Cuenta Inexistente")
   Call FINAL("")
   '
11 CUEBIN$ = MKI$(CUE%)
   CODENOCUE$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%))
   JJ& = 0
   SALINI# = 0
   '
   If DES% >= FECHANUM("01/10/01") Then
      CCCC$ = TRIM$(CODCUE$(CUE%))
      If CCCC$ = "111401" Then
         SALINI# = 10886.85   ' 11409.33
         SALINI# = SALINI# + 114.95 ' para tener en cuenta cheque duplicado flama
      End If
      If CCCC$ = "111502" Then SALINI# = 11791.1
      If CCCC$ = "111403" Then SALINI# = 13039.69   ' 2414.3
      If CCCC$ = "111406" Then SALINI# = 30026.22   ' 29769.7
      If CCCC$ = "111506" Then SALINI# = 381.86
   End If
   '
   ' CARGAR MOVIMIENTOS DE DIARIO
   EXTEA$ = EXTE$
   If EXTEA$ = "DAT" Then
     Close #NX9%: NX9% = 0
     Close #NX10%: NX10% = 0
     Close #NX11%: NX11% = 0
     EXTE$ = "OLD"
   End If
   '
12 PMC$ = PUNMODIAR$(CUE%)
   FIN& = Len(PMC$)
   For KK% = 1 To Len(PMC$) Step 2
      UU& = KK% + 1
      Call TRACE(20, UU&, FIN&)
      REGDIA% = CVI(Mid$(PMC$, KK%, 2))
      Call LEEDIARIO(REGDIA%)
      If CVS(E12$) < 1.05 Then GoTo 19 ' SALTEA ASIENTO APERTURA
      If Asc(E18$) <> 1 Then  ' SALTEA ASIENTOS RESUMEN DE CAJA
         '
         Call LEEPASDIA(REGDIA%, 2)
         DETAPAS$ = TRIM$(DETAPAS$)
         If DETAPAS$ <> "" Then
            DETAPAS$ = TRIM$(DETAPAS$) + " - "
         End If
         DETAPAS$ = DETAPAS$ + REFASIEN$
         '
         ACIERR% = 0
         If InStr(UCase$(DETAPAS$), "(CIERR") > 0 Or InStr(UCase$(DETAPAS$), "(*") > 0 Then
           ACIERR% = 1
         End If
         MARCADIA$ = "*"
         FEMOREA% = CVI(E15$)
         IMPO# = CVD(E16$)
         If FEMOREA% < DES% Then
             If ACIERR% = 1 Then IMPO# = 0
             If FEMOREA% > FECHANUM("31/09/01") Then GoSub 90 'TOTALIZAR INICIAL
          ElseIf FEMOREA% <= HAS% Then
            FEMORE% = FEMOREA%
            FEREGI% = FEMOREA%
            GoSub 95
         End If
      End If
19 Next KK%
   '
   If EXTE$ = "OLD" Then
     If EXTEA$ = "DAT" Then
       Close #NX9%: NX9% = 0
       Close #NX10%: NX10% = 0
       Close #NX11%: NX11% = 0
       EXTE$ = "DAT"
       EXTEA$ = ""
       GoTo 12
     End If
   End If
   '
   '
20 ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   EXTEA$ = EXTE$
   If EXTEA$ = "DAT" Then
     Close #NX5%: NX5% = 0
     Close #NX6%: NX6% = 0
     Close #NX7%: NX7% = 0
     EXTE$ = "OLD"
   End If
   '
22 PMC$ = PUNMOCAJA$(CUE%)
   FIN& = Len(PMC$)
   For KK% = 1 To Len(PMC$) Step 4
      UU& = KK% + 1
      Call TRACE(20, UU&, FIN&)
      REGDIAL& = CVS(Mid$(PMC$, KK%, 4))
      Call LEEPASCAJ(REGDIAL&, 1)
      FEMOVI% = CVI(D15$)
      IMPO# = CVD(D16$)
      '
      If TRIM$(DETAPAS$) = "" Then
         Call LEEPASCAJ(REGDIAL&, 2)
         DETAPAS$ = REFASIEN$
      End If
      MARCADIA$ = ""
      If InStr(DETAPAS$, "- OP") > 0 Then GoTo 29
      If Left$(DETAPAS$, 8) = "DEPOSITO" Then
         FEMORE% = FECHACRE%(FEMOVI%, Val(Mid$(DETAPAS$, 10, 3)))
         If FEMORE% < DES% Then
            If FEMORE% > FECHANUM("31/09/01") Then GoSub 90
          ElseIf FEMORE% <= HAS% Then
            GoSub 95
         End If
         GoTo 29
      End If
      If FEMOVI% < DES% Then
        If FEMOVI% > FECHANUM("31/09/01") Then GoSub 90
        GoTo 29
      End If
      If FEMOVI% >= DES% Then
        If FEMOVI% <= HAS% Then
          FEMORE% = FEMOVI%
          GoSub 95
        End If
      End If
      '
29 Next KK%
   '
   If EXTE$ = "OLD" Then
     If EXTEA$ = "DAT" Then
       Close #NX5%: NX5% = 0
       Close #NX6%: NX6% = 0
       Close #NX7%: NX7% = 0
       EXTE$ = "DAT"
       EXTEA$ = ""
       GoTo 22
     End If
   End If
   '
   For U& = 1 To ULTREG&("CHEQUEMI")
     X$ = REGLEIDO$("CHEQUEMI", U&)
      PROV% = CVI(Left$(X$, 2))
      NUMERO$ = Mid$(X$, 3, 12)
      BANCO$ = Mid$(X$, 15, 16)
      EMIS% = CVI(Mid$(X$, 47, 2)) ' toma fecha de O/Pago
      IMPO# = (-1) * CVD(Mid$(X$, 33, 8))
      NOPAGO& = CVS(Mid$(X$, 41, 4))
      VENCI% = CVI(Mid$(X$, 49, 2))
      '
      If VENCI% < EMIS% Then  ' modificacion del 28/01/02. Si
         VENCI% = EMIS%       ' la fecha del valor es anterior
      End If                  ' a la OP, toma la de la OP
      '
      QUE% = CVI(Mid$(X$, 45, 2))
            For UX% = 1 To CAMP%
               If MEPA%(UX%) = QUE% Then GoTo 38
            Next UX%
            Call MENSERR(24, "Medio de Pago" + Str$(QUE%) + " no Definido")
            Call FINAL("")
            '
38    CUEH% = CUEMP%(UX%)
      If CUEH% = CUE% Then
        FEMORE% = VENCI%
If EMIS% <= FECHANUM("31/09/01") And VENCI% > FECHANUM("31/09/01") Then
SALINICA# = SALINICA# + IMPO#
AAA = BBB
End If
        If FEMORE% < DES% Then
          If FEMORE% > FECHANUM("31/09/01") Then GoSub 90
        End If
        If FEMORE% >= DES% Then
          If FEMORE% <= HAS% Then
            FEX = EMIS%
            DETAPAS$ = "Ch." + TRIM$(NUMERO$)
            DETAPAS$ = DETAPAS$ + " - OP." + TRIM$(Str$(NOPAGO&))
            DETAPAS$ = DETAPAS$ + " - " + TRIM$(RASOCLI$((-1) * PROV%))
            GoSub 95
          End If
        End If
      End If
   Next U&
   '
   FEX = DES%
   FECHAMO%(0) = DIANTE(FEX)
   FEX = FECHAMO(0)
   REFE$(0) = "Saldo al " + FECHATEX$(FEX)
   IPODIF#(0) = SALINI#
   '
   For KKL& = 0 To JJ&
     For KKL1& = KKL& + 1 To JJ&
       If FECHAMO%(KKL1&) < FECHAMO%(KKL&) Then
         FX1% = FECHAMO%(KKL&)
         FX2$ = REFE$(KKL&)
         FX3# = IPODIF#(KKL&)
         FECHAMO%(KKL&) = FECHAMO%(KKL1&)
         REFE$(KKL&) = REFE$(KKL1&)
         IPODIF#(KKL&) = IPODIF#(KKL1&)
         FECHAMO%(KKL1&) = FX1%
         REFE$(KKL1&) = FX2$
         IPODIF#(KKL1&) = FX3#
        End If
     Next KKL1&
     GoSub 96
   Next KKL&
80 'Next cuebal&
   '
   Call SETULTREG("PODIFER", KLM&)
   Call CIERRARCH("PODIFER")
   Call CIERRARCH("*.*")
   '
   Screen.MousePointer = 1
   '
   Exit Sub
   
90 SALINI# = SALINI# + IMPO#
   Return
   '
95 If Abs(IMPO#) >= 0.005 Then
      JJ& = JJ& + 1
      REFE$(JJ&) = DETAPAS$
      FECHAMO%(JJ&) = FEMORE%
      IPODIF#(JJ&) = IMPO#
   End If
   Return
   '
96 ZZZ$ = REGBLAN$("PODIFER")
   If DCBB$ <> DCBBA$ Then
     SARRAS# = 0
     DCBBA$ = DCBB$
   End If
   Call REPLA(ZZZ$, DCBB$, 1, 32)
   Call REPLA(ZZZ$, MKI$(FECHAMO%(KKL&)), 33, 2)
   Call REPLA(ZZZ$, REFE$(KKL&), 35, 30)
   ACIERR% = 0
   If InStr(UCase$(REFE$(KKL&)), "(CIERR") > 0 Or InStr(REFE$(KKL&), "(*") > 0 Then
      ACIERR% = 1
   End If
   SARRAS# = SARRAS# + IPODIF#(KKL&)
   If ACIERR% > 0 Then SARRAS# = SARRAS# - IPODIF#(KKL&)
   If IPODIF#(KKL&) > 0 Then
      Call REPLA(ZZZ$, MKD$(IPODIF#(KKL&)), 65, 8)
      If ACIERR% > 0 Then Call REPLA(ZZZ$, MKD$(Abs(IPODIF#(KKL&))), 73, 8)
     Else
      Call REPLA(ZZZ$, MKD$(Abs(IPODIF#(KKL&))), 73, 8)
      If ACIERR% > 0 Then Call REPLA(ZZZ$, MKD$(Abs(IPODIF#(KKL&))), 65, 8)
   End If
   Call REPLA(ZZZ$, MKD$(SARRAS#), 81, 8)
   KLM& = KLM& + 1
   Call GRAREG("PODIFER", ZZZ$, KLM&)
   Return
   '

End Sub

