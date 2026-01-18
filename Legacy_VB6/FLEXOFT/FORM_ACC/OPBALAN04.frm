VERSION 5.00
Begin VB.Form OPBALAN04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Balance General - Sumas y Saldos"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9420
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
   ScaleWidth      =   9420
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
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   8580
         ScaleHeight     =   75
         ScaleWidth      =   690
         TabIndex        =   25
         Top             =   2415
         Width           =   750
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
         Left            =   8580
         Picture         =   "OPBALAN04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   2670
         Width           =   750
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
         Left            =   3255
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
         Left            =   6510
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
         Width           =   3060
         Begin VB.CheckBox Check2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0D0D0&
            Caption         =   "Asiento Apertura"
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
            Left            =   1000
            TabIndex        =   28
            Top             =   880
            Width           =   1770
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0D0D0&
            Caption         =   "Asientos de Cierre"
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
            Left            =   840
            TabIndex        =   27
            Top             =   1155
            Width           =   1935
         End
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
            Left            =   2535
            TabIndex        =   7
            Top             =   405
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
            Left            =   420
            TabIndex        =   6
            Top             =   405
            Width           =   2220
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Excluir:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   105
            TabIndex        =   29
            Top             =   880
            Width           =   800
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
            Left            =   3765
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
            Left            =   5235
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
            Left            =   630
            TabIndex        =   4
            Top             =   945
            Width           =   2100
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
            Left            =   630
            TabIndex        =   3
            Top             =   525
            Value           =   -1  'True
            Width           =   2220
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
            Left            =   3030
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
            Left            =   3030
            TabIndex        =   22
            Top             =   585
            Width           =   645
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
   End
End
Attribute VB_Name = "OPBALAN04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Command1.Enabled = False
   Call BALSUSAL
   Command1.Enabled = True
End Sub

Private Sub Command2_Click()
   Call APERBASDAT("CABAN")
   SQLX$ = "SELECT  CODICOM_LAR,  SUM(IDEBECOMP - IHABECOMP) AS SUMA FROM CAJABANC "
   SQLX$ = SQLX$ + " WHERE (NuAsidiar = '' or isnull(NuAsiDiar) = true OR IsEmpty(NuAsiDiar) = true) "
   SQLX$ = SQLX$ + " GROUP BY CODICOM_LAR"
   '
   Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   CABATEMP.MoveFirst
   Do While Not CABATEMP.EOF
      If Abs(CABATEMP!SUMA) >= 0.005 Then
         AA1 = CABATEMP!FechEmisi
         AA2 = CABATEMP!CodiCom_Lar
         AA3 = BB3
      End If
   CABATEMP.MoveNext
   Loop
   '
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
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6

      Me.Option1(0).BackColor = &HFCD7B6
      Me.Option1(1).BackColor = &HFCD7B6
      Me.Option1(2).BackColor = &HFCD7B6
      Me.Option1(3).BackColor = &HFCD7B6
      Me.Option1(4).BackColor = &HFCD7B6
      Me.Option2.BackColor = &HFCD7B6
      Me.Option3.BackColor = &HFCD7B6
      Me.Option4.BackColor = &HFCD7B6
      Me.Option5.BackColor = &HFCD7B6

      Me.Check1.BackColor = &HFCD7B6
      Me.Check2.BackColor = &HFCD7B6
      '
    End If


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
   Dim NUMECO$(), SUMASIT#(), FECO%()
   ReDim NUMECO$(1024), SUMASIT#(1024), FECO%(1024)
   CANCOM& = 0
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
   SUCUI% = 0
   CMJ& = CVDAT(COMEJ%)
   SQLX$ = "SELECT DETAASIEN.FECHASIEN, DETAASIEN.NUIDASIEN, DETAASIEN.REFEPASE, DETAASIEN.CUECOINT, DETAASIEN.IDEBEPASE, DETAASIEN.IHABEPASE, ENCABEASIEN.TIPOASIEN "
   SQLX$ = SQLX$ + " FROM DETAASIEN inner join ENCABEASIEN  on DETAASIEN.NUIDASIEN = ENCABEASIEN.NUIDASIEN WHERE "
   'SQLX$ = SQLX$ + "(CDBL(FECHASIEN) >= " & CDbl(CVDAT(COMEJ%)) & " "
   SQLX$ = SQLX$ + "CDBL(DETAASIEN.FECHASIEN) <= " & CDbl(CVDAT(HAS%)) & " "
   SQLX$ = SQLX$ + "AND DETAASIEN.SUCUASIEN = " & SUCUI% & " "
   SQLX$ = SQLX$ + "AND ENCABEASIEN.STATASIEN > 0 "
   If Check1.Value = 1 Then
        SQLX$ = SQLX$ + "AND (ENCABEASIEN.TIPOASIEN <> 9 OR DETAASIEN.FECHASIEN <= " & CMJ& & ") "
   End If
   SQLX$ = SQLX$ + "AND DETAASIEN.STATPASE > 0; "
   
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
         'If FEAS% < DES% Then ' Or (FEAS% = DES% And NUAS = 1) Then
         If FEAS% < DES% Or (Check2.Value = 1 And !TIPOASIEN = 8) Then ' Or (FEAS% = DES% And NUAS = 1) Then
             SALINI1#(CUE%) = SALINI1#(CUE%) + IPAS# '(Int(100 * IPAS# + 0.5)) / 100
             SUMAINI# = SUMAINI# + IPAS#
           Else
             If FEAS% <= HAS% Then
               If IPAS# > 0 Then
                   SUDEBE1#(CUE%) = SUDEBE1#(CUE%) + IPAS# ' (Int(100 * IPAS# + 0.5)) / 100
                   SUMADEB# = SUMADEB# + IPAS#
                 Else
                   SUHABE1#(CUE%) = SUHABE1#(CUE%) - IPAS# ' (Int(100 * IPAS# + 0.5)) / 100
                   SUMAHAB# = SUMAHAB# - IPAS#
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
   
   ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   Call ABRECABAN
   On Error GoTo 0
   '
   ' ***** OJO QUE SI SE MODIFICA EL CRITERIO ACÁ, TAMBIÉN EN ****
   ' ***** LA RUTINA DE GENERACIÓN DEL ASIENTO DE CIERRE !!!! ****
   SQLX$ = "SELECT * FROM CAJABANC "
   'SQLX$ = SQLX$ + "where (CueContaI = " & CUEII% & " "
   SQLX$ = SQLX$ + " WHERE (NuAsidiar = '' or isnull(NuAsiDiar) = true OR IsEmpty(NuAsiDiar) = true) "
   'SQLX$ = SQLX$ + " AND CDBL(FechConta) >= " & FEINI
   'SQLX$ = SQLX$ + " AND CDBL(FechConta) <= " & FEFIN & ") "
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
         NUMECOL$ = !CodiCom_Lar
         '
         ' ESTO ES PARA VERIFICAC BALANCE DE COMPROBANTES
         ' VERIFICA UNICAMENTE LOS DEL EJERCICIO
         If FEAS% < DES% Then
              ' ESTO PARA VERIFICAR BALANCE INICIAL
              SUMASIT#(0) = SUMASIT#(0) + IPAS#   '(Int(100 * IPAS# + 0.5)) / 100
            ElseIf FEAS% >= DES% Then
              If FEAS% <= HAS% Then
                If NUMECOL$ <> ULTICOMPRO$ Then GoTo 22
                For KKU& = CANCOM& To 1 Step -1
                  If NUMECO$(KKU&) = NUMECOL$ Then
                    SUMASIT#(KKU&) = SUMASIT#(KKU&) + IPAS#   '(Int(100 * IPAS# + 0.5)) / 100
                    GoTo 25
                  End If
                Next KKU&
                '
22              ULTICOMPRO$ = NUMECOL$
                CANCOM& = CANCOM& + 1
                If UBound(NUMECO$) < CANCOM& Then
                  UBA& = UBound(NUMECO$) + 1024
                  ReDim Preserve NUMECO$(UBA&), FECO%(UBA&), SUMASIT#(UBA)
                End If
                NUMECO$(CANCOM&) = NUMECOL$
                FECO%(CANCOM&) = FEAS%
                SUMASIT#(CANCOM&) = IPAS#    ' (Int(100 * IPAS# + 0.5)) / 100
              End If
         End If
         ' FIN VERIFICA BALANCE DE COMPROBANTES
         '
25       If CUE% > 0 Then
           If CUE% <= 16384 Then
             If FEAS% < DES% Then ' Or (FEAS% = DES% And NUAS = 1) Then
                 SALINI1#(CUE%) = SALINI1#(CUE%) + IPAS# '(Int(100 * IPAS# + 0.5)) / 100
                 SUMAINI# = SUMAINI# + IPAS#
               Else
                 If FEAS% <= HAS% Then
                   If IPAS# > 0 Then
                       SUDEBE1#(CUE%) = SUDEBE1#(CUE%) + IPAS# ' (Int(100 * IPAS# + 0.5)) / 100
                       'SUDEBE1#(CUE%) = SUDEBE1#(CUE%) + (Int(100 * IPAS# + 0.5)) / 100
                       SUMADEB# = SUMADEB# + IPAS#
                     Else
                       SUHABE1#(CUE%) = SUHABE1#(CUE%) - IPAS# ' - (Int(100 * IPAS# + 0.5)) / 100
                       'SUHABE1#(CUE%) = SUHABE1#(CUE%) - (Int(100 * IPAS# + 0.5)) / 100
                       SUMAHAB# = SUMAHAB# - IPAS#
                   End If
                 End If
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
   If Abs(SUMASIT#(0)) >= 0.005 Then
      TTXX$ = "Se ha Detectado un Des-Balance Inicial\Anterior al Comienzo del Ejercicio Actual."
      TTXX$ = TTXX$ + "\  \Causa Probable: Error de Migración de Datos\Anteriores a la Puesta en Marcha del Sistema."
      TTXX$ = TTXX$ + "\  \Puede Recuperarse Mediante un Pase de Ajuste."
      If COMALTER%(TTXX$ + "\\Continuar\Realizar Ajuste") = 2 Then
        DIFEREC# = SUMASIT#(0)
        GoSub RECUBAL
        Exit Sub
      End If
   End If
   '
   For KKU& = 1 To CANCOM&
     Call TRACE(20, KKU&, CANCOM&)
     If Abs(SUMASIT#(KKU&)) >= 0.005 Then
       FEX = FECO%(KKU&)
       '
       ASIAN$ = NUMECO$(KKU&)
       If Asc(Left$(ASIAN$ + " ", 1)) > 57 Then ' SOLO ASIENTOS DE C&B
         Call RECUMOCAJ(ASIAN$, CVDAT(FECO%(KKU&)), OKEYY%)
         If OKEYY% > 0 Then
           HUBORECU% = 1
           GoTo 89
         End If
       End If
       '
       Call MENSERR(24, "Comprobante '" + NUMECO$(KKU&) + "' del " + FECHATEX$(FEX) + " No Balancea.\Este Comprobante Debe Anularse\y Cargarse Nuevamente.")
     End If
89 Next KKU&
   '
   If HUBORECU% > 0 Then
     Call COMUNI("Movimientos No Balanceados Recuperados.\  \Debe Volver a Pedir el Reporte.")
     Exit Sub
   End If
   '
   On TIBA% GoTo 100, 100, 300
   Call CIERRARCH("*.*")
   Exit Sub
   '
100 ' BALANCE MENSUAL *******************************
   '
   On Error Resume Next
   Contable.Execute "DROP TABLE BALASUSA"
   Contable.Execute "SELECT codexcue, denocue, codincue INTO Balasusa FROM Plancuen;"
   On Error GoTo 0
   '
   SQLX$ = "SELECT NIVAGRU, CODEXCUE, DENOCUE, CODINCUE, SALINI, SUDEBE, SUHABE, RESULT, SALFIN, NuOrdIt "
   '\\\OT-5-549-FG-29/07/05
   SQLX$ = SQLX$ + " FROM PLANCUEN ORDER BY NuOrdIt ASC"
   '\\\OT-5-549-FG-FIN
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
   Call ABREPLANCUEN
   SQLX$ = "SELECT NIVAGRU, CODEXCUE, DENOCUE, CODINCUE, SALINI, SUDEBE, SUHABE, RESULT, SALFIN "
   '\\\OT-5-549-FG-29/07/05
   SQLX$ = SQLX$ + " FROM PLANCUEN "
   '\\\OT-5-549-FG-FIN
   ACONDI% = 0
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
   SQLX$ = SQLX$ + "ORDER BY NuOrdIt ASC"
   '
   Set Balasec = Contable.OpenRecordset(SQLX$)
   Balasec.MoveLast
   If Balasec.RecordCount < 1 Then
     Call MENSERR(24, "Sin Registros Activos")
     Screen.MousePointer = 1
     Exit Sub
   End If
   '
   'If NXP% < 1 Then
   '   NXP% = FILEOPEN%(LUDAT$ + "PLAGRAL" + "." + EXTE$, 0, 4)
   'End If
   '
   JJ& = 0: KLJ& = 0
   Dim IVP As String * 4
   Dim SUNIVEL#(5), DENIVEL$(5)
   '
   With Balasec
     FIN& = .RecordCount
     .MoveFirst
     UU& = 1
     Do While Not .EOF
       Call TRACE(20, UU&, FIN&)
       CUE% = !CODINCUE         'CVI(Mid$(IVP$, 1, 2))
       NIVEI% = !NIVAGRU        'CVI(Mid$(IVP$, 3, 2))
       PRONI% = NIVEI%
       If Not .EOF Then
         .MoveNext
         PRONI% = !NIVAGRU
         .MovePrevious
       End If
       '
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
       '
       .MoveNext
       UU& = UU& + 1
     Loop
   End With
   '
   For KQ% = NIVANT% - 1 To 0 Step -1   'NIVEI% Step -1
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
   '\\\OT-5-549-FG-29/07/05
   SQLX$ = SQLX$ + " FROM PLANCUEN ORDER BY NuOrdIt ASC"
   '\\\OT-5-549-FG-FIN
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
   SQLX$ = "SELECT CODEXCUE, DENOCUE, CODINCUE, SALINI, SUDEBE, SUHABE, RESULT, SALFIN ,NuOrdIt"
   '\\\OT-5-549-FG-29/07/05
   SQLX$ = SQLX$ + " FROM PLANCUEN "
   '\\\OT-5-549-FG-FIN
   ACONDI% = 0
   If Option1(3) = True Then
     SQLX$ = SQLX$ + "WHERE (SALINI <> 0 OR SUDEBE <> 0 OR SUHABE <> 0 "
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
   SQLX$ = SQLX$ + " ORDER BY NuOrdIt ASC "
   '
   Set Balasec = Contable.OpenRecordset(SQLX$, dbOpenSnapshot)
   Balasec.MoveLast
   If Balasec.RecordCount < 1 Then
     Call MENSERR(24, "Sin Registros Activos")
     Screen.MousePointer = 1
     mouspointer = 1
     Exit Sub
   End If
   '
   Set BASUMSAL.Data1.Recordset = Balasec ' Contable.OpenRecordset(SQLX$) ', dbOpenSnapshot, dbReadOnly)
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
   '
   '
   '
   JJ& = 0
   With Data1.Recordset
     For U& = 1 To FIN&
       If U& = 1 Then
            .MoveFirst
          Else
            .MoveNext
       End If
       '
       ZZ$ = REGBLAN$("BASUSAL")
       Call REPLA(ZZ$, .Fields(0).Value, 1, 12)
       Call REPLA(ZZ$, .Fields(1).Value, 13, 52)
       Call REPLA(ZZ$, MKD$(.Fields(3).Value), 65, 8)
       Call REPLA(ZZ$, MKD$(.Fields(4).Value), 73, 8)
       Call REPLA(ZZ$, MKD$(.Fields(5).Value), 81, 8)
       Call REPLA(ZZ$, MKD$(.Fields(6).Value), 89, 8)
       Call REPLA(ZZ$, MKD$(.Fields(7).Value), 97, 8)
       JJ& = JJ& + 1
       Call GRAREG("BASUSAL", ZZ$, JJ&)
     Next U&
   End With
   '
   Call SETULTREG("BASUSAL", JJ&)
   Call CIERRARCH("BASUSAL")
   '
   Screen.MousePointer = 1
   '
   BASUMSAL.Label2 = Text6
   'acá agregaría valores de suma a las 3 etiquetas
   BASUMSAL.Label3.Caption = FORMATNUM$(SUMAINI#, "F12.2")
   BASUMSAL.Label4.Caption = FORMATNUM$(SUMADEB#, "F12.2")
   BASUMSAL.Label5.Caption = FORMATNUM$(SUMAHAB#, "F12.2")
   RESU# = SUMADEB# - SUMAHAB#
   SALD# = SUMAINI# + RESU#
   BASUMSAL.Label6.Caption = FORMATNUM$(RESU#, "F12.2")
   BASUMSAL.Label7.Caption = FORMATNUM$(SALD#, "f12.2")
   
   BASUMSAL.Show 1
   Exit Sub
   '
RECUBAL:
   Call ABRECABAN
   With CajaBanc
     .FindFirst "CodiCom_Lar = 'AJUS-INI' "
     If .NoMatch = False Then
       If CVINT(!FechConta) >= DES% Then GoTo 980
       NUEVAL# = !IDEBECOMP - !IHABECOMP - DIFEREC#
       .Edit
       !ValAbsComp = Abs(NUEVAL#)
       If NUEVAL# > 0 Then
          !IDEBECOMP = NUEVAL#
          !IHABECOMP = 0
        Else
          !IDEBECOMP = 0
          !IHABECOMP = Abs(NUEVAL#)
       End If
       .Update
       Screen.MousePointer = 1
       Call COMUNI("Diferencia Recuperada.\Vuelva a Pedir el Listado.")
       Exit Sub
     End If
     '
980  Screen.MousePointer = 1
     Call COMUNI("Imposible Recuperar Diferencia - Consulte.")
   End With
Return
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

Private Sub Text8_Change()
    Label1 = ECOARCH("ECUECOT", Text8.TEXT)
End Sub

