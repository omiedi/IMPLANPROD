VERSION 5.00
Begin VB.Form ACIERRE04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asiento de Cierre"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   6045
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
   ScaleHeight     =   5460
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0D0D0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9915
      Begin VB.PictureBox Picture16 
         Height          =   5500
         Left            =   5040
         ScaleHeight     =   5445
         ScaleWidth      =   1035
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   1100
         Begin VB.Image Image3 
            Height          =   5600
            Left            =   0
            Picture         =   "ACIERRE04.frx":0000
            Stretch         =   -1  'True
            Top             =   -120
            Width           =   1095
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00004000&
         FillColor       =   &H00C0C0FF&
         Height          =   6735
         Left            =   5100
         ScaleHeight     =   6675
         ScaleWidth      =   1950
         TabIndex        =   16
         Top             =   0
         Width           =   2010
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
            Height          =   675
            Left            =   105
            Picture         =   "ACIERRE04.frx":2398
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Inicia Proceso de Generación del Asiento de Cierre"
            Top             =   4620
            Width           =   650
         End
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   120
            ScaleHeight     =   75
            ScaleWidth      =   585
            TabIndex        =   19
            Top             =   3675
            Width           =   650
         End
         Begin VB.CommandButton Command18 
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
            Height          =   675
            Left            =   120
            Picture         =   "ACIERRE04.frx":26A2
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   840
            Width           =   650
         End
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
            Height          =   650
            Left            =   120
            Picture         =   "ACIERRE04.frx":29AC
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Terminar - Salir de la Aplicación"
            Top             =   105
            Width           =   650
         End
         Begin VB.Image Image2 
            Height          =   390
            Left            =   -300
            Picture         =   "ACIERRE04.frx":2AF6
            Top             =   3990
            Width           =   1515
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0D0D0&
         Caption         =   "Filtro por Tipo de Cuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   210
         TabIndex        =   3
         ToolTipText     =   "Doble-Click para Incluir Todos los Tipos de Cuenta"
         Top             =   3045
         Width           =   4650
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C0D0D0&
            Caption         =   "Haber"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1260
            TabIndex        =   13
            Top             =   425
            Width           =   990
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H00C0D0D0&
            Caption         =   "Ingresos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2310
            TabIndex        =   12
            Top             =   425
            Width           =   1125
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H00C0D0D0&
            Caption         =   "Gastos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3570
            TabIndex        =   11
            Top             =   425
            Width           =   990
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0D0D0&
            Caption         =   "Debe"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   210
            TabIndex        =   10
            Top             =   425
            Width           =   990
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0D0D0&
         Caption         =   "Cuenta de Contra-Partida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   210
         TabIndex        =   2
         Top             =   4095
         Width           =   4650
         Begin VB.TextBox Text1 
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
            Left            =   2730
            TabIndex        =   26
            Top             =   315
            Width           =   1485
         End
         Begin VB.CommandButton Command2 
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
            Height          =   285
            Left            =   4230
            TabIndex        =   14
            Top             =   320
            Width           =   175
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta:"
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
            Left            =   300
            TabIndex        =   21
            Top             =   420
            Width           =   2325
         End
         Begin VB.Label Label7 
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
            Height          =   300
            Left            =   210
            TabIndex        =   15
            Top             =   735
            Width           =   4200
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0D0D0&
         Caption         =   "Rango de Cuentas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2280
         Left            =   200
         TabIndex        =   1
         Top             =   630
         Width           =   4650
         Begin VB.TextBox Text2 
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
            Left            =   2730
            TabIndex        =   25
            Top             =   1365
            Width           =   1485
         End
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
            Left            =   2730
            TabIndex        =   22
            Top             =   375
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
            Height          =   330
            Left            =   4230
            TabIndex        =   5
            Top             =   375
            Width           =   175
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
            Height          =   285
            Left            =   4200
            TabIndex        =   4
            Top             =   1365
            Width           =   175
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
            Height          =   300
            Left            =   210
            TabIndex        =   24
            Top             =   1785
            Width           =   4200
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
            Height          =   300
            Left            =   210
            TabIndex        =   23
            Top             =   800
            Width           =   4200
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Desde Cuenta:"
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
            Left            =   300
            TabIndex        =   7
            Top             =   450
            Width           =   2325
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta Cuenta:"
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
            Left            =   300
            TabIndex        =   6
            Top             =   1440
            Width           =   2325
         End
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
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
         Height          =   300
         Left            =   3800
         TabIndex        =   9
         Top             =   210
         Width           =   1050
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha del Asiento :"
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
         Left            =   1365
         TabIndex        =   8
         Top             =   315
         Width           =   2325
      End
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu aprobar 
         Caption         =   "Aprobar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "ACIERRE04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub CALCUCIE()
   '
   '\\\OT-07-0065-RG-21/02/07///
   '
   Screen.MousePointer = 11
   HAS% = FECHANUM(Label4)   ' FIN DEL EJERCICIO
   Dim SALFIN1#(16384)
   CONTRAPA% = CUEINT%(Text1)
   TIASI% = 9            ' ASIENTO DE CIERRE
   '
   'Acumuladores para suma total de columnas y muestra
   SUMAINI# = 0
   SUMADEB# = 0
   SUMAHAB# = 0
   RESU# = 0
   SALD# = 0
   '
   ' CARGAR MOVIMIENTOS DE DIARIO
   '
   ' ***** OJO QUE SI SE MODIFICA EL CRITERIO ACÁ, TAMBIÉN EN ****
   ' ***** LA RUTINA DE GENERACIÓN DEL ASIENTO DE CIERRE !!!! ****
   ' ***** EL CRITERIO ES QUE EL ASIENTO DE CIERRRE REFLEJE ******
   ' ***** EXACTAMENTE LO MISMO QUE EL BALANCE DE SUMAS Y ********
   ' ***** SALDOS, AÚN CUANDO AQUEL ESTÉ MAL PORQUE NO SE ********
   ' ***** HIZO EL CIERRE DEL EJERCICIO ANTERIOR *****************
   SUCUI% = 0
   SQLX$ = "SELECT FECHASIEN, NUIDASIEN, REFEPASE, CUECOINT, IDEBEPASE, IHABEPASE "
   SQLX$ = SQLX$ + " FROM DETAASIEN WHERE "
   'SQLX$ = SQLX$ + "CDBL(FECHASIEN) >= " & CDbl(CVDAT(COMEJ%)) & " "
   SQLX$ = SQLX$ + "CDBL(FECHASIEN) <= " & CDbl(CVDAT(HAS%)) & " "
   SQLX$ = SQLX$ + "AND SUCUASIen = " & SUCUI% & " "
   SQLX$ = SQLX$ + "AND STATPASE > 0; "
   Set MODIASEL = Contable.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   '
   CARESE& = 0
   On Error Resume Next
   MODIASEL.MoveLast
   CARESE& = MODIASEL.RecordCount
   JJ& = 0: FIN& = CARESE&
   '
   With MODIASEL
     .MoveFirst
10   JJ& = JJ& + 1
     Call TRACE(20, JJ&, FIN&)
     '
     CUE% = !CUECOINT
     FEAS% = FECHANUM(Format$(!FECHASIEN, "DD/MM/YY"))
     IPAS# = ROUND#(!IDEBEPASE - !IHABEPASE)
     '
     If CUE% > 0 Then
       If CUE% <= 16384 Then
         If FEAS% <= HAS% Then
            SALFIN1#(CUE%) = SALFIN1#(CUE%) + IPAS# ' (Int(100 * IPAS# + 0.5)) / 100
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

   ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   Call ABRECABAN
   On Error GoTo 0
   ' ***** OJO QUE SI SE MODIFICA EL CRITERIO ACÁ, TAMBIÉN EN ****
   ' ***** LA RUTINA DE GENERACIÓN DEL BALANCE DE SALDOS !!!! ****
   ' ***** EL CRITERIO ES QUE EL ASIENTO DE CIERRRE REFLEJE ******
   ' ***** EXACTAMENTE LO MISMO QUE EL BALANCE DE SUMAS Y ********
   ' ***** SALDOS, AÚN CUANDO AQUEL ESTÉ MAL PORQUE NO SE ********
   ' ***** HIZO EL CIERRE DEL EJERCICIO ANTERIOR *****************
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + " WHERE (NuAsidiar = '' or isnull(NuAsiDiar) = true OR IsEmpty(NuAsiDiar) = true) "
   'SQLX$ = SQLX$ + " AND CDBL(FechConta) >= " & CDbl(CVDAT(COMEJ%))
   'SQLX$ = SQLX$ + " AND CDBL(FechConta) <= " & CDbl(CVDAT(FINEJ%))
   SQLX$ = SQLX$ + " ORDER BY CodiCom_Lar ASC;"
   '
   Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   ULTICOMPRO$ = ""
   With CABATEMP
     On Error Resume Next
     .MoveLast
     If Err < 1 Then
       FIN& = .RecordCount
       JJ& = 0
       .MoveFirst
       Do While Not .EOF
         JJ& = JJ& + 1
         Call TRACE(20, JJ&, FIN&)
         '
         CUE% = !CueContaI
         FEAS% = CVINT(!FechConta)
         IPAS# = ROUND(!IDEBECOMP - !IHABECOMP)
         '
25       If CUE% > 0 Then
           If CUE% <= 16384 Then
             If FEAS% <= HAS% Then
               SALFIN1#(CUE%) = SALFIN1#(CUE%) + IPAS#
             End If
           End If
         End If
         '
       .MoveNext
       Loop
       '
     End If
   End With
   '
   '
   ASIDIAR.Ltext1 = Label4
   ASIDIAR.Text4 = "ASIENTO DE CIERRE"
   CONTA% = 0: SUMA# = 0
   ASIDIAR.List1.Clear
   ASIDIAR.List2.Clear
   FIN& = 16384
   '
   SUMA# = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     IPASE# = (-1) * SALFIN1#(U&)
     If Abs(IPASE#) >= 0.005 Then
        CUEI% = U&
        TPCX$ = UCase$(Left$(TIPCUE$(CUEI%), 1))
        If Check1.Value <> 1 And TPCX$ = "D" Then GoTo 35
        If Check2.Value <> 1 And TPCX$ = "H" Then GoTo 35
        If Check3.Value <> 1 And TPCX$ = "I" Then GoTo 35
        If Check4.Value <> 1 And TPCX$ = "G" Then GoTo 35
        '
        COCUE$ = CODCUE$(CUEI%)
        If COCUE$ >= Text10 And COCUE$ <= Text2 Then
             IPASE# = SALFIN1#(U&)
             If Abs(IPASE#) >= 0.005 Then
               SUMA# = SUMA# + ROUND#(IPASE#)
             End If
           Else
35           SALFIN1#(U&) = 0
        End If
     End If
   Next U&
   SALFIN1#(CONTRAPA%) = SALFIN1#(CONTRAPA%) - SUMA#
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     IPASE# = (-1) * SALFIN1#(U&)
     If Abs(IPASE#) >= 0.005 Then
       CUEI% = U&
       CONTA% = CONTA% + 1
       SUMA# = SUMA# + ROUND#(IPASE#)
       '
       Z$ = Space$(128)
       Call REPLA(Z$, CODCUE$(CUEI%), 1, 12)
       DEPASE$ = DENOCUE$(CUEI%)
       Call REPLA(Z$, DEPASE$, 14, 36)
       If IPASE# > 0 Then
            Call REPLA(Z$, CSTRING$(MKD$(IPASE#), 3, 12, 2, 2), 52, 13)
          Else
            Call REPLA(Z$, CSTRING$(MKD$(Abs(IPASE#)), 3, 12, 2, 2), 66, 13)
       End If
       ASIDIAR.List2.AddItem "A" + Z$
     End If
   Next U&
   '
   For U& = 1 To ASIDIAR.List2.ListCount
      Z$ = Mid$(ASIDIAR.List2.List(U& - 1), 2)
      ASIDIAR.List1.AddItem Z$
   Next U&
   '
   If ASIDIAR.List1.ListCount < 1 Then
      Call MENSERR(24, "Asiento Vacío.")
      GoTo 90
   End If
   '
   Call ASIDIAR.SUMASIEN
   ASIDIAR.Label3 = "(Asiento de Cierre)"
   NASI$ = ASIDIAR.NUPROA$
   ASIDIAR.Ltext2 = NASI$
   Screen.MousePointer = 1
   '
   On Error Resume Next
   ASIDIAR.Show 1
   If Err Then
     On Error GoTo 0
     ASIDIAR.Visible = True
   End If
   '
90 Unload Me
   '
End Sub

Private Sub aprobar_Click()
   Call Command1_Click
End Sub

Private Sub Command1_Click()
   '\\\OT-07-0065-RG-21/02/07///
   Command1.Enabled = False
   '
   If XVALO(Text10) <= 0 Then
      Call MENSERR(24, "Falta Ingresar Cuenta")
      Text10.SetFocus
      GoTo 99
   End If
   '
   If XVALO(Text2) <= 0 Then
      Call MENSERR(24, "Falta Ingresar Cuenta")
      Text2.SetFocus
      GoTo 99
   End If
   '
   If CUEINT%(Text10) = 0 Then
      Call MENSERR(24, "Número de Cuenta No Válido")
      Text10 = ""
      Text10.SetFocus
      GoTo 99
   End If
   '
   If CUEINT%(Text2) = 0 Then
      Call MENSERR(24, "Número de Cuenta No Válido")
      Text2 = ""
      Text2.SetFocus
      GoTo 99
   End If
   '
   If XVALO(Text2) < XVALO(Text10) Then
      Call MENSERR(24, "Rango de Cuentas No Válido")
      Text2 = ""
      Text2.SetFocus
      GoTo 99
   End If
   '
   If Check1.Value = 0 And Check2.Value = 0 And Check3.Value = 0 And Check4.Value = 0 Then
      Call MENSERR(24, "Falta Seleccionar Alguna Opción de Filtro")
      GoTo 99
   End If
   '
   If XVALO(Text1) <= 0 Then
      Call MENSERR(24, "Falta Ingresar Cuenta de Contra-Partida")
      Text1.SetFocus
      GoTo 99
   End If
   '
   If CUEINT%(Text1) = 0 Then
      Call MENSERR(24, "Número de Cuenta No Válido")
      Text1 = ""
      Text1.SetFocus
      GoTo 99
   End If
   '
   Call CALCUCIE
   '
99 Screen.MousePointer = 1
   Command1.Enabled = True
   '\\\OT-07-0065-RG-FIN///
End Sub



Private Sub Command2_Click()
    '\\\OT-07-0065-RG-20/02/07///
    Call SELECHO("ECUECOT")
    PRICUEX$ = VALACT1$("ECUECOT")
    If PRICUEX$ <> "" Then
        Text1.TEXT = PRICUEX$
    End If
    '\\\OT-07-0065-RG-FIN///
End Sub

Private Sub Command3_Click()
    '\\\OT-07-0065-RG-20/02/07///
    Call CIERRARCH("*.*")
    Unload Me
    '\\\OT-07-0065-RG-FIN///
End Sub

Private Sub Command8_Click()
    Call SELECHO("ECUECOT")
    PRICUEX$ = VALACT1$("ECUECOT")
    If PRICUEX$ <> "" Then
        Text2.TEXT = PRICUEX$
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
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6
      Me.Check1.BackColor = &HFCD7B6
      Me.Check2.BackColor = &HFCD7B6
      Me.Check3.BackColor = &HFCD7B6
      Me.Check4.BackColor = &HFCD7B6
      
      MARCOBARRA.Top = 5200
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
    End If

    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Text10 = Left$(REGLEIDO$("ECUECOT", 1), 12)
    Text2 = Left$(REGLEIDO$("ECUECOT", ULTREG&("ECUECOT")), 12)
    '\\\OT-07-0065-RG-21/02/07///
    Label4.Caption = FECHATEX$(FINEJ%)
    '\\\OT-07-0065-RG-FIN///
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
End Sub


Private Sub Frame4_DblClick()
   Check1.Value = 1
   Check2.Value = 1
   Check3.Value = 1
   Check4.Value = 1
End Sub

Private Sub salir_Click()
   Call Command3_Click
End Sub

Private Sub Text1_Change()
    '\\\OT-07-0065-RG-20/02/07///
    Label7 = ECOARCH("ECUECOT", Text1.TEXT)
    '\\\OT-07-0065-RG-FIN///
End Sub

Private Sub Text1_DblClick()
    '\\\OT-07-0065-RG-20/02/07///
    Call Command2_Click
    '\\\OT-07-0065-RG-FIN///
End Sub

Private Sub Text10_Change()
    Label2 = ECOARCH("ECUECOT", Text10.TEXT)
End Sub


Private Sub Text10_DblClick()
   '\\\OT-07-0065-RG-20/02/07///
   Call Command9_Click
   '\\\OT-07-0065-RG-FIN///
End Sub

Private Sub Text2_Change()
  Label1 = ECOARCH("ECUECOT", Text2.TEXT)
End Sub

Private Sub Text2_DblClick()
   '\\\OT-07-0065-RG-20/02/07///
  Call Command8_Click
  '\\\OT-07-0065-RG-FIN///
End Sub
