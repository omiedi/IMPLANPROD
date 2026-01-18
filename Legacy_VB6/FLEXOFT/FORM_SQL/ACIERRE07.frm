VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form ACIERRE07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asiento de Cierre"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5805
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
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H0095C5E1&
      BorderStyle     =   0  'None
      Height          =   6720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9915
      Begin VB.PictureBox Picture999 
         Height          =   330
         Index           =   0
         Left            =   600
         ScaleHeight     =   270
         ScaleWidth      =   1110
         TabIndex        =   26
         Top             =   120
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H002C3D46&
         FillColor       =   &H00C0C0FF&
         ForeColor       =   &H80000008&
         Height          =   6735
         Left            =   4920
         ScaleHeight     =   6705
         ScaleWidth      =   1980
         TabIndex        =   16
         Top             =   0
         Width           =   2010
         Begin VB.PictureBox MARCOBARRA 
            Height          =   195
            Left            =   105
            ScaleHeight     =   135
            ScaleWidth      =   585
            TabIndex        =   28
            Top             =   4200
            Width           =   650
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H0048CC9D&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   -120
               TabIndex        =   29
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
            Left            =   105
            Picture         =   "ACIERRE07.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Inicia Proceso de Generación del Asiento de Cierre"
            Top             =   4620
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
            Picture         =   "ACIERRE07.frx":030A
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
            Picture         =   "ACIERRE07.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Terminar - Salir de la Aplicación"
            Top             =   105
            Width           =   650
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H0095C5E1&
         Caption         =   "FILTRO POR TIPO DE CUENTA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Doble-Click para Incluir Todos los Tipos de Cuenta"
         Top             =   3045
         Width           =   4650
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H0095C5E1&
            Caption         =   "Haber"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1260
            TabIndex        =   13
            Top             =   425
            Width           =   990
         End
         Begin VB.CheckBox Check3 
            Appearance      =   0  'Flat
            BackColor       =   &H0095C5E1&
            Caption         =   "Ingresos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2310
            TabIndex        =   12
            Top             =   425
            Width           =   1125
         End
         Begin VB.CheckBox Check4 
            Appearance      =   0  'Flat
            BackColor       =   &H0095C5E1&
            Caption         =   "Gastos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3570
            TabIndex        =   11
            Top             =   425
            Width           =   990
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H0095C5E1&
            Caption         =   "Debe"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   210
            TabIndex        =   10
            Top             =   425
            Width           =   990
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H0095C5E1&
         Caption         =   "CUENTA CONTRA-PARTIDA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   2
         Top             =   4095
         Width           =   4650
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2730
            TabIndex        =   25
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
            Height          =   300
            Left            =   4230
            TabIndex        =   14
            Top             =   320
            Width           =   175
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   300
            TabIndex        =   20
            Top             =   420
            Width           =   2325
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   210
            TabIndex        =   15
            Top             =   720
            Width           =   4200
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0095C5E1&
         Caption         =   "RANGO DE CUENTAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2280
         Left            =   120
         TabIndex        =   1
         Top             =   630
         Width           =   4650
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2730
            TabIndex        =   24
            Top             =   1365
            Width           =   1485
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2730
            TabIndex        =   21
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
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   210
            TabIndex        =   23
            Top             =   1785
            Width           =   4200
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   210
            TabIndex        =   22
            Top             =   800
            Width           =   4200
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Desde Cuenta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   300
            TabIndex        =   7
            Top             =   480
            Width           =   2325
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta Cuenta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   300
            TabIndex        =   6
            Top             =   1470
            Width           =   2325
         End
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   2625
         OleObjectBlob   =   "ACIERRE07.frx":075E
         Top             =   0
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
         Height          =   225
         Index           =   0
         Left            =   240
         OleObjectBlob   =   "ACIERRE07.frx":0992
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3675
         TabIndex        =   9
         Top             =   210
         Width           =   1050
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha del Asiento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1245
         TabIndex        =   8
         Top             =   315
         Width           =   2325
      End
   End
End
Attribute VB_Name = "ACIERRE07"
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
   SUMAINI# = 0
   SUMADEB# = 0
   SUMAHAB# = 0
   RESU# = 0
   SALD# = 0
   '
   ' CARGAR MOVIMIENTOS DE DIARIO
   '
10  SQLX$ = "SELECT cuecoint, SUM(IDEBEPASE-IHABEPASE) AS SALDOFI FROM DETAASIEN "
    SQLX$ = SQLX$ + " WHERE FECHASIEN  <= '" & FETEXCOR1$(HAS%) & "' "
    ' SQLX$ = SQLX$ + "AND SUCUASIen = " & SUCUI% & " "
    SQLX$ = SQLX$ + "AND STATPASE > 0 "
    SQLX$ = SQLX$ + "GROUP BY CUECOINT "
    '
    Dim MODIASEL As ADODB.Recordset
    Set MODIASEL = New ADODB.Recordset
    On Error Resume Next
    MODIASEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then
       On Error GoTo 0
       GoTo 15
    End If
    On Error GoTo 0
    '
    With MODIASEL
       If Not .BOF And Not .EOF Then
         Do While Not .EOF
           NCLIE% = !CUECOINT
           If NCLIE% > 0 Then
               SALFIN1#(NCLIE%) = SALFIN1#(NCLIE%) + ROUND(!SALDOFI)
           End If
         .MoveNext
         Loop
       End If
    End With
    MODIASEL.Close
    '
    ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
15  Call ABRECABAN
    SQLX$ = "SELECT CueContaI, SUM(IDEBECOMP - IHABECOMP) AS SALDOFI FROM CAJABANC "
    SQLX$ = SQLX$ + " WHERE FechConta  <= '" & FETEXCOR1$(HAS%) & "' "
    SQLX$ = SQLX$ + " AND (NuAsidiar = '' or NuAsiDiar IS NULL)  "
    
    SQLX$ = SQLX$ + "GROUP BY CueContaI "
    '
    Dim CABTEMP As ADODB.Recordset
    Set CABTEMP = New ADODB.Recordset
    On Error Resume Next
    CABTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then
       On Error GoTo 0
       GoTo 25
    End If
    On Error GoTo 0
    '
    With CABTEMP
       If Not .BOF And Not .EOF Then
         Do While Not .EOF
           CUE% = !CUECONTAI
           If CUE% > 0 Then
               SALFIN1#(CUE%) = SALFIN1#(CUE%) + ROUND(!SALDOFI)
           End If
         .MoveNext
         Loop
       End If
    End With
    CABTEMP.Close
   '
25 ASIDIAR07.Ltext1 = Label4
   ASIDIAR07.Text4 = "ASIENTO DE CIERRE"
   CONTA% = 0: SUMA# = 0
   ASIDIAR07.List1.Clear
   ASIDIAR07.List2.Clear
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
       ASIDIAR07.List2.AddItem "A" + Z$
     End If
   Next U&
   '
   For U& = 1 To ASIDIAR07.List2.ListCount
      Z$ = Mid$(ASIDIAR07.List2.LIST(U& - 1), 2)
      ASIDIAR07.List1.AddItem Z$
   Next U&
   '
   If ASIDIAR07.List1.ListCount < 1 Then
      Call MENSERR(24, "Asiento Vacío.")
      GoTo 90
   End If
   '
   Call ASIDIAR07.SUMASIEN
   ASIDIAR07.Label3 = "(Asiento de Cierre)"
   ASIDIAR07.Ltext1 = Label4
   NASI$ = ASIDIAR07.NUPROA$
   ASIDIAR07.Ltext2 = NASI$
   Screen.MousePointer = 1
   '
   On Error Resume Next
   ASIDIAR07.Show 1
   If Err Then
     On Error GoTo 0
     ASIDIAR07.Visible = True
   End If
   '
90 Unload Me
   '
End Sub

Private Sub Command1_Click()
   '\\\OT-07-0065-RG-21/02/07///
   Command1.Enabled = False
   '
   If TRIM$(Text10) = "" Then
      Call MENSERR(24, "Falta Ingresar Cuenta")
      Text10.SetFocus
      GoTo 99
   End If
   '
   If TRIM$(Text2) = "" Then
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
   If Text2 < Text10 Then
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
   If TRIM$(Text1) = "" Then
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
'    EPAC$ = TRIM$(EMPREAC$)
'    If EPAC$ <> "" Then
'        Caption = Caption + "  -  " + EPAC$
'    End If
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    
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
