VERSION 5.00
Begin VB.Form DIARGEN 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Libro Diario General"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   3300
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
   ScaleHeight     =   3585
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture16 
      Height          =   3600
      Left            =   2400
      ScaleHeight     =   3540
      ScaleWidth      =   945
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   1000
      Begin VB.Image Image3 
         Height          =   3700
         Left            =   0
         Picture         =   "DIARGEN.frx":0000
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   1000
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00004040&
      Height          =   8600
      Left            =   2415
      ScaleHeight     =   8535
      ScaleWidth      =   1215
      TabIndex        =   22
      Top             =   -105
      Width           =   1275
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   555
         TabIndex        =   26
         Top             =   2280
         Width           =   615
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   12000
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
         Height          =   675
         Left            =   120
         Picture         =   "DIARGEN.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   380
         Width           =   615
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
         Picture         =   "DIARGEN.frx":24E2
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1065
         Width           =   615
      End
      Begin VB.CommandButton Command15 
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
         Left            =   120
         Picture         =   "DIARGEN.frx":27EC
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Genera Reporte Solicitado"
         Top             =   2835
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0F0&
      Caption         =   "Selección"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1100
      Left            =   600
      TabIndex        =   19
      Top             =   4800
      Width           =   2445
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0F0&
         Caption         =   "Cuentas c/Movim."
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
         Index           =   3
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0F0&
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
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   1980
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0E0F0&
      Caption         =   "Rango de Fechas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Doble-Click para Editar Rango de Fechas"
      Top             =   210
      Width           =   2115
      Begin VB.CommandButton Command21 
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
         Left            =   1740
         TabIndex        =   14
         Top             =   705
         Width           =   175
      End
      Begin VB.CommandButton Command4 
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
         Left            =   1740
         TabIndex        =   13
         Top             =   1440
         Width           =   175
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
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
         Left            =   30
         TabIndex        =   18
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label3 
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
         Left            =   -1440
         TabIndex        =   17
         Top             =   465
         Width           =   2325
      End
      Begin VB.Label label5 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   630
         TabIndex        =   16
         Top             =   720
         Width           =   1110
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   630
         TabIndex        =   15
         Top             =   1455
         Width           =   1110
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0F0&
      Caption         =   "Rango de Cuentas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2150
      Left            =   0
      TabIndex        =   3
      Top             =   5775
      Width           =   4005
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
         Left            =   3675
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   360
         Width           =   175
      End
      Begin VB.TextBox Text1 
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
         Left            =   1890
         TabIndex        =   6
         Top             =   360
         Width           =   1800
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
         Height          =   285
         Left            =   3675
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1320
         Width           =   175
      End
      Begin VB.TextBox Text2 
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
         Left            =   1890
         TabIndex        =   4
         Top             =   1320
         Width           =   1800
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
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   3600
      End
      Begin VB.Label Label7 
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
         Left            =   840
         TabIndex        =   10
         Top             =   420
         Width           =   960
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
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   3600
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
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
         Left            =   840
         TabIndex        =   8
         Top             =   1380
         Width           =   960
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0F0&
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
      Height          =   1050
      Left            =   120
      TabIndex        =   0
      Top             =   2415
      Width           =   2115
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0F0&
         Caption         =   "Impresora"
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
         Left            =   420
         TabIndex        =   2
         Top             =   680
         Width           =   1260
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0F0&
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
         Height          =   240
         Index           =   1
         Left            =   420
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1140
      End
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu aprobar 
         Caption         =   "Aprobar"
         Begin VB.Menu pantalla 
            Caption         =   "Pantalla "
         End
         Begin VB.Menu impresora 
            Caption         =   "Impresora"
         End
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "DIARGEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Unload Me
End Sub

Private Sub Command25_Click()
   '
   Command15.Enabled = False
   If FECHANUM(Label5) > FECHANUM(Label6) Or FECHANUM(Label5) < COMEJ% Or FECHANUM(Label6) > FINEJ% Then
     Call MENSERR(24, "Rango de Fechas Imposible")
     GoTo 99
   End If
   DESYY = CDbl(CVDAT(FECHANUM(Label5)))
   HASYY = CDbl(CVDAT(FECHANUM(Label6)))
   '
   CUEDES$ = Text1
   If CUEINT%(CUEDES$) < 1 Then
     Call MENSERR(24, "Cuenta 'Desde' no Válida")
     GoTo 99
   End If
   '
   CUEHAS$ = Text2
   If CUEINT%(CUEHAS$) < 1 Then
     Call MENSERR(24, "Cuenta 'Hasta' no Válida")
     GoTo 99
   End If
   '
   If Option1(1).Value = True Then
   On Error GoTo 0
   SUCUI% = 0
   JJ& = 0
     SQLX$ = "SELECT FECHASIEN, NUIDASIEN, REFEPASE, CUECOEXT, IDEBEPASE, IHABEPASE "
       SQLX$ = SQLX$ + " FROM DETAASIEN WHERE "
       SQLX$ = SQLX$ + "(CDBL(FECHASIEN) >= " & DESYY & " "
       SQLX$ = SQLX$ + "AND CDBL(FECHASIEN) <= " & HASYY & " "
       SQLX$ = SQLX$ + "AND CUECOEXT >= '" & CUEDES$ & "' "
       SQLX$ = SQLX$ + "AND CUECOEXT <= '" & CUEHAS$ & "' "
       SQLX$ = SQLX$ + "AND SUCUASIen = " & SUCUI% & " "
       SQLX$ = SQLX$ + "AND STATPASE > 0) "
       SQLX$ = SQLX$ + "ORDER BY NUIDASIEN, ORNUPASE;"
       Set DIARPAN.Data1.Recordset = Contable.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
       On Error Resume Next
       '
   ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   Call ABRECABAN
   On Error GoTo 0
   SQLX$ = "SELECT * FROM CAJABANC "
   'SQLX$ = SQLX$ + "where (CueContaI = " & CUEII% & " "
   SQLX$ = SQLX$ + " WHERE (NuAsidiar = '' or isnull(NuAsiDiar) = true OR IsEmpty(NuAsiDiar) = true) "
   'SQLX$ = SQLX$ + " AND CDBL(FechConta) >= " & FEINI
   'SQLX$ = SQLX$ + " AND CDBL(FechConta) <= " & FEFIN & ") "
   'SQLX$ = SQLX$ + " ORDER BY FECHCONTA ASC, FEHOREAL ASC;"
   '
   Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
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
         IPAS# = !IDEBECOMP - !IHABECOMP
         '
'DIARPAN.Data1.Recordset.AddNew
'DIARPAN.Data1.Recordset.Fields(0).Value = CVDAT(FEAS%)
'DIARPAN.Data1.Recordset.Fields(1).Value = !CODICOM_COR
'DIARPAN.Data1.Recordset.Update
'On Error GoTo 0
'AA1 = DIARPAN.Data1.Recordset.RecordCount
'DIARPAN.Data1.Recordset.Fields(0).Value = CVDAT(FEAS%)
'DIARPAN.Data1.Recordset.Fields(0).Value = CVDAT(FEAS%)
'DIARPAN.Data1.Recordset.Fields(0).Value = CVDAT(FEAS%)

'         If CUE% > 0 Then
'           If CUE% <= 16384 Then
'             If FEAS% < DES% Then ' Or (FEAS% = DES% And NUAS = 1) Then
'                 SALINI1#(CUE%) = SALINI1#(CUE%) + (Int(100 * IPAS# + 0.5)) / 100
'               Else
'                 If FEAS% <= HAS% Then
'                   If IPAS# > 0 Then
'                       SUDEBE1#(CUE%) = SUDEBE1#(CUE%) + (Int(100 * IPAS# + 0.5)) / 100
'                     Else
'                       SUHABE1#(CUE%) = SUHABE1#(CUE%) - (Int(100 * IPAS# + 0.5)) / 100
'                   End If
'                 End If
'             End If
'           End If
'         End If
         '
       .MoveNext
       Loop
       '
     End If
   End With
'select * from cajabanc

       DIARPAN.Data1.Refresh
       DIARPAN.DBGrid1.Refresh
       '
       On Error GoTo 0
       Set COLUI = DIARPAN.DBGrid1.Columns
       COLUI(0).Caption = "Fecha"
       COLUI(1).Caption = "Nro-Id"
       COLUI(2).Caption = "Referencia"
       COLUI(3).Caption = "Cuenta"
       COLUI(4).Caption = "Debe"
       COLUI(5).Caption = "Haber"
       COLUI(4).NumberFormat = "########0.00"
       COLUI(5).NumberFormat = "########0.00"
       COLUI(0).Width = 1110
       COLUI(1).Width = 1095
       COLUI(2).Width = 4425
       COLUI(3).Width = 1320
       COLUI(4).Width = 1580
       COLUI(5).Width = 1580
       DIARPAN.Show 1
     Else
       Call DIARIOPRIN
   End If
   '
99 Command15.Enabled = True
   
End Sub
'
Sub DIARIOPRIN()
   '
   Call FRESHALL
   '
10 DES% = FECHANUM(Label5): HAS% = FECHANUM(Label6)
   CORRESP$ = DENOEJ$ + " (" + Label5 + " - " + Label6 + ")"
   '
   DESYY = CDbl(CVDAT(FECHANUM(Label5)))
   HASYY = CDbl(CVDAT(FECHANUM(Label6)))
   '
   CUEDES$ = Text1
   CUEHAS$ = Text2
   '
   'MODOXX2% = COMALTER%("Puede Incluir los Asientos de Caja y Bancos\o Dejar Solo los Asientos Resumen\\Asientos de C & B\Asientos Resumen")
   'If MODOXX2% < 1 Or MODOXX2% > 2 Then GoTo 99
   'Call GRACONTROL("CONTAGEN", "INCLUCAJ", Str$(MODOXX2% Mod 2))
   'MODOXX3% = COMALTER%("Puede Re-Numerar Correlativamente los Asientos\o Expresar los Números Provisorios de Asientos\\Re-Numerar Asientos\Números Provisorios")
   'If MODOXX3% < 1 Or MODOXX3% > 2 Then GoTo 99
   'Call GRACONTROL("CONTAGEN", "RENUMERA", Str$(MODOXX3% Mod 2))
   '
   Screen.MousePointer = 11
   SUDEBE# = 0
   SUHABE# = 0
   '
   ' 1.- DETERMINAR TRANSPORTE
   SUCUI% = 0
   DESYY = CDbl(CVDAT(DES%))
   SQLX$ = "SELECT IDEBEPASE, IHABEPASE FROM DETAASIEN "
   'SQLX$ = SQLX$ + "(CDBL(FECHASIEN) >= " & DESYY & " " COMIENZO DE EJERCICIO
       SQLX$ = SQLX$ + "WHERE (CDBL(FECHASIEN) < " & DESYY & " "
       SQLX$ = SQLX$ + "AND CUECOEXT >= '" & CUEDES$ & "' "
       SQLX$ = SQLX$ + "AND CUECOEXT <= '" & CUEHAS$ & "' "
       SQLX$ = SQLX$ + "AND SUCUASIen = " & SUCUI% & " "
       SQLX$ = SQLX$ + "AND STATPASE > 0) "
       Set DIARPAN.Data1.Recordset = Contable.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   
   SCG_PRIN.SORTLIST.Clear
   '
   FIN& = MAXDIARIO%
   For KK% = 1 To MAXDIARIO%
      UU& = KK%
      Call TRACE(20, UU&, FIN&)
      Call LEEPASDIA(KK%, 1)
      FEAS% = CVI(E15$)
      If MODOXX2% = 1 Then
        If Asc(E18$) = 1 Then GoTo 29 ' SALTEA ASIENTOS RESUMEN DE CAJA
      End If
      RESOR$ = Left$(TRIM$(Str$(10000 + FEAS%)) + Space$(10), 10)
      RESOR$ = RESOR$ + "DI" + TRIM$(Str$(100000 + KK%))
      SCG_PRIN.SORTLIST.AddItem RESOR$
29 Next KK%
   '
   If MODOXX2% = 1 Then
     FIN& = MAXDIA&
     For KKL& = 1 To MAXDIA&
       UU& = KKL&
       Call TRACE(20, UU&, FIN&)
       Call LEEPASCAJ(KKL&, 1)
       FEAS% = CVI(D15$)
       RESOR$ = Left$(TRIM$(Str$(10000 + FEAS%)) + Space$(10), 10)
       RESOR$ = RESOR$ + "SU" + TRIM$(Str$(100000 + KKL&))
       SCG_PRIN.SORTLIST.AddItem RESOR$
     Next KKL&
   End If
   '
   ' CARGAR MOVIMIENTOS DE DIARIO
   FIN& = SCG_PRIN.SORTLIST.ListCount    ' MAXDIARIO%
   YATRA% = 0: JJ& = 0: ASIANT = 0: ASIANTX = 0
   '
   NUASDEF& = 0
   If MODOXX3% = 1 Then
      Screen.MousePointer = 1
      NUASDEF& = Val(InputBox("Re-Numera a Partir de", "Numeración Asientos Libro Diario", "1.0")) - 1
      If NUASDEF& < 0 Then Exit Sub
      Screen.MousePointer = 11
   End If
   '
   For KK% = 1 To FIN&     'MAXDIARIO%
      UU& = KK%
      Call TRACE(20, UU&, FIN&)
      RESOR$ = SCG_PRIN.SORTLIST.List(KK% - 1)
      REGII& = Val(Mid$(RESOR$, 13)) - 100000
      GRUXI$ = TRIM$(Mid$(RESOR$, 11, 2))
      SGX1% = 1
      FEASI% = Val(TRIM$(Left$(RESOR$, 10))) - 10000
      MODOLE% = 1
      If FEASI% >= DES% And FEASI% <= HAS% Then MODOLE% = 2
      If GRUXI$ = "DI" Then
         REGIJ% = REGII&
         Call LEEPASDIA(REGIJ%, MODOLE%)
         FEAS% = CVI(E15$)
         IPAS# = CVD(E16$)
         CUE% = CVI(E11$)
         NUAS = CVS(E12$)
         REFEL$ = "- " + TRIM$(DETAPAS$)
         REFE$ = REPLACAR$(REFEL$, "@", "A")
         If TRIM$(REFE$) = "-" Then
           REFE$ = "- " + TRIM$(DENOCUE$(CUE%))
         End If
         DENOMI$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%))
        Else
         SGX1% = (-1)
         Call LEEPASCAJ(REGII&, MODOLE%)
         FEAS% = CVI(D15$)
         IPAS# = CVD(D16$)
         CUE% = CVI(D11$)
         NUAS = CVS(D12$)
         REFEL$ = "- " + TRIM$(DETAPAS$)
         REFE$ = REPLACAR$(REFEL$, "@", "A")
         If TRIM$(REFE$) = "-" Then
           REFE$ = "- " + TRIM$(DENOCUE$(CUE%))
         End If
         DENOMI$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%))
      End If
      '
      If NUAS * SGX1% <> ASIANT Then
        NUASDEF& = NUASDEF& + 1
        ASIANT = NUAS * SGX1%
      End If
      '
      NUADE% = NUASDEF& Mod 32700
      If SGX1% > 0 Then
          Get #NX10%, REGII&, E10$
          EXXX$ = E10$
          Call REPLA(EXXX$, MKI$(NUADE%), 7, 2)
          E10$ = EXXX$
          Put #NX10%, REGII&, E10$
        Else
          Get #NX6%, REGII&, D10$
          EXXX$ = D10$
          Call REPLA(EXXX$, MKI$(NUADE%), 7, 2)
          D10$ = EXXX$
          Put #NX6%, REGII&, D10$
      End If
      '
      If FEAS% < DES% Then
         If IPAS# > 0 Then
              SUDEBE# = SUDEBE# + IPAS#
            Else
              SUHABE# = SUHABE# - IPAS#
         End If
         GoTo 90
      End If
      '
      If YATRA% = 0 Then
         Z$ = REGBLAN$("DIARGEN")
         FEX = DES%
         DIA0% = DIANTE(FEX)
         Call REPLA(Z$, MKI$(DIA0%), 1, 2)
         Call REPLA(Z$, MKS$(99999999999#), 3, 4)
         Call REPLA(Z$, "Transporte", 7, 53)
         Call REPLA(Z$, MKD$(SUDEBE#), 105, 8)
         Call REPLA(Z$, MKD$(SUHABE#), 113, 8)
         JJ& = JJ& + 1
         Call GRAREG("DIARGEN", Z$, JJ&)
         Z$ = REGBLAN$("DIARGEN")
         JJ& = JJ& + 1
         Call GRAREG("DIARGEN", Z$, JJ&)
         YATRA% = 1
      End If
      '
      If FEAS% <= HAS% Then
        If NUAS * SGX1% <> ASIANTX Then
           ASIANTX = NUAS * SGX1%
           Z$ = REGBLAN$("DIARGEN")
           Call REPLA(Z$, MKI$(FEAS%), 1, 2)
           Call REPLA(Z$, MKS$(NUAS), 3, 4)
           If MODOXX3% = 1 Then
             Call REPLA(Z$, MKS$(NUASDEF&), 3, 4)
           End If
           Call REPLA(Z$, REFASIEN$, 7, 53)
           JJ& = JJ& + 1
           Call GRAREG("DIARGEN", Z$, JJ&)
         End If
         '
         Z$ = REGBLAN$("DIARGEN")
         Call REPLA(Z$, MKI$(FEAS%), 1, 2)
         Call REPLA(Z$, MKS$(NUAS), 3, 4)
         If MODOXX3% = 1 Then
            Call REPLA(Z$, MKS$(NUASDEF&), 3, 4)
         End If
         Call REPLA(Z$, REFE$, 7, 53)
         Call REPLA(Z$, DENOMI$, 60, 45)
         If IPAS# > 0 Then
              Call REPLA(Z$, MKD$(IPAS#), 105, 8)
            Else
              Call REPLA(Z$, MKD$((-1) * IPAS#), 113, 8)
         End If
         JJ& = JJ& + 1
         Call GRAREG("DIARGEN", Z$, JJ&)
      End If
90 Next KK%
   '
   Call SETULTREG("DIARGEN", JJ&)
   Call CIERRARCH("DIARGEN")
   '
   Call HEADERS("DIARGEN", "")
   Call HEADERS("DIARGEN", "Corresponde a: " + CORRESP$)
   '
   Screen.MousePointer = 1
   Call CONECRUN("*SUBDIAR", "Libro Diario General")
   '
99 End Sub

Private Sub Command15_Click()
   ' LA VERSION ANTERIOR ESTA EN COMMAND25_CLICK
   Command15.Enabled = False
   If FECHANUM(Label5) > FECHANUM(Label6) Or FECHANUM(Label5) < COMEJ% Or FECHANUM(Label6) > FINEJ% Then
     Call MENSERR(24, "Rango de Fechas Imposible")
     GoTo 99
   End If
   DESYY = CDbl(CVDAT(FECHANUM(Label5)))
   HASYY = CDbl(CVDAT(FECHANUM(Label6)))
   '
   On Error GoTo 0
   FEINI = Int(CDbl(CVDAT(COMEJ%)))
   FEFIN = Int(CDbl(CVDAT(FINEJ%)))
   '
   Call BLOQARCH("DIARPRIM")
   Screen.MousePointer = 11
   SUCUI% = 0
   JJ& = 0
   REBLA$ = REGBLAN$("DIARPRIM")
   Call BLANARCH("DIARPRIM/NOERR")
   '
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM DetaAsien "
   'SQLX$ = SQLX$ + "WHERE (CueCoInt=" & CUEII%
   SQLX$ = SQLX$ + " WHERE SucuAsien = " & SUCUI%
   SQLX$ = SQLX$ + " AND CDBL(Fechasien) >= " & FEINI
   SQLX$ = SQLX$ + " AND CDBL(Fechasien) <= " & FEFIN
   SQLX$ = SQLX$ + " AND statpase >0  "
   SQLX$ = SQLX$ + "ORDER BY FECHASIEN, ORFEASIEN, OrNupase ASC;"
   Set PASEASIEN = Contable.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   '
   JJ& = 0
   TRANSPOR# = 0: TRANDEBE# = 0: TRANHABE# = 0
10 With PASEASIEN
     On Error Resume Next
     .MoveLast
     If Err < 1 Then
       On Error GoTo 0
       FIN& = .RecordCount
       U& = 0
       .MoveFirst
       Do While Not .EOF
         U& = U& + 1
         Call TRACE(20, U&, FIN&)
         IDEBE# = !IDEBEPASE
         IHABE# = !IHABEPASE
         If Int(CDbl(!FECHASIEN)) < DESYY Then   ' FEINI Then
             TRANDEBE# = TRANDEBE# + IDEBE#
             TRANHABE# = TRANHABE# + IHABE#
           ElseIf Int(CDbl(!FECHASIEN)) <= HASYY Then   ' FEFIN Then
             Z$ = REBLA$
             Call REPLA(Z$, MKI$(CVINT(!FECHASIEN)), 1, 2)
             Call REPLA(Z$, !NUIDASIEN, 3, 16)
               ' AQUI FALTA NUMERO DE ASIENTO DEFINITIVO
                 REFEX$ = REFGENAS(!NUIDASIEN)
             Call REPLA(Z$, REFEX$, 23, 64)
               ' AQUI FALTA REFERENCIA GENERAL DEL ASIENTO
             Call REPLA(Z$, CODCUE$(!CUECOINT), 87, 16)
               DEPASE$ = !REFEPASE
             Call REPLA(Z$, DEPASE$, 103, 64)
             Call REPLA(Z$, MKD$(IDEBE#), 167, 8)
             Call REPLA(Z$, MKD$(IHABE#), 175, 8)
             JJ& = JJ& + 1
             Call GRAREG("DIARPRIM", Z$, JJ&)
         End If
       .MoveNext
       Loop
     End If
   End With
   On Error GoTo 0
   '
   Call ABRECABAN
   SQLX$ = "SELECT * FROM CAJABANC "
   'SQLX$ = SQLX$ + "where (CueContaI = " & CUEII% & " "
   SQLX$ = SQLX$ + " WHERE (NuAsidiar = '' or isnull(NuAsiDiar) = true OR IsEmpty(NuAsiDiar) = true) "
   SQLX$ = SQLX$ + " AND CDBL(FechConta) >= " & FEINI
   'SQLX$ = SQLX$ + " AND CDBL(FechConta) <= " & FEFIN & ") "
   SQLX$ = SQLX$ + " ORDER BY FECHCONTA ASC, FEHOREAL ASC;"
   '
   Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   With CABATEMP
     On Error Resume Next
     .MoveLast
     If Err < 1 Then
       On Error GoTo 0
       FIN& = .RecordCount
       U& = 0
       .MoveFirst
       '
       Do While Not .EOF
         U& = U& + 1
         Call TRACE(20, U&, FIN&)
         IDEBE# = !IDEBECOMP
         IHABE# = !IHABECOMP
         If Int(CDbl(!FechConta)) < DESYY Then   ' FEINI Then
             TRANDEBE# = TRANDEBE# + IDEBE#
             TRANHABE# = TRANHABE# + IHABE#
           ElseIf Int(CDbl(!FechConta)) <= HASYY Then   ' FEFIN Then
             Z$ = REBLA$
             Call REPLA(Z$, MKI$(CVINT(!FechConta)), 1, 2)
             Call REPLA(Z$, !CodiCom_Cor, 3, 12)
             ' AQUI FALTA NUMERO DE ASIENTO DEFINITIVO
             ' AQUI FALTA REFERENCIA GENERAL DEL ASIENTO
             Call REPLA(Z$, CODCUE$(!CueContaI), 87, 16)
               DEPASE$ = !RefeCom
             Call REPLA(Z$, DEPASE$, 103, 64)
             Call REPLA(Z$, MKD$(IDEBE#), 167, 8)
             Call REPLA(Z$, MKD$(IHABE#), 175, 8)
               DEMOVI$ = !DescriMov
             Call REPLA(Z$, DEMOVI$, 191, 40)
             JJ& = JJ& + 1
             Call GRAREG("DIARPRIM", Z$, JJ&)
         End If
       .MoveNext
       Loop
     End If
   End With
   Screen.MousePointer = 1
   '
   Call SETULTREG("DIARPRIM", JJ&)
   Call FILESORT("DIARPRIM", "", 1, 2, "ASC")
   TRANSPOR# = TRANDEBE# - TRANHABE#
   '
   DESX% = FECHANUM(Label5)
   HASX% = FECHANUM(Label6)
   FIN& = ULTREG&("DIARPRIM")
   SUMASI# = 0: ASIAN$ = ""
   HUBORECU% = 0
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("DIARPRIM", U&)
     FEAS% = CVI(Left$(X$, 2))
     '
     If FEAS% < DESX% Then
       TRANDEBE# = TRANDEBE# + CVD(Mid$(X$, 167, 8))
       TRANHABE# = TRANHABE# + CVD(Mid$(X$, 175, 8))
       TRANSPOR# = TRANDEBE# - TRANHABE#
     End If
     '
     If FEAS% >= DESX% Then
       If FEAS% <= HASX% Then
         If DIARPAN.List1.ListCount = 0 Then
           TTXX$ = Space$(128)
           Call REPLA(TTXX$, FECHATEX$(DIANTE(DESX%)), 1, 8)
           Call REPLA(TTXX$, "Transporte ...............................", 47, 32)
           Call REPLA(TTXX$, FORMATNUM$(TRANDEBE#, "F14.2"), 80, 14)
           Call REPLA(TTXX$, FORMATNUM$(TRANHABE#, "F14.2"), 94, 14)
           DIARPAN.List1.AddItem TTXX$
           '
           TTXX$ = Space$(128)
           Call REPLA(TTXX$, " ............ ", 80, 14)
           Call REPLA(TTXX$, "............. ", 94, 14)
           DIARPAN.List1.AddItem TTXX$
         End If
         TTXX$ = Space$(128 + 40)
         Call REPLA(TTXX$, FECHATEX$(FEAS%), 1, 8)
           NIDAS$ = TRIM$(Mid$(X$, 3, 16))
         Call REPLA(TTXX$, NIDAS$, 11, 12)
           If NIDAS$ <> ASIAN$ Then
             If Abs(SUMASI#) >= 0.005 Then
               '
               If Asc(Left$(ASIAN$ + " ", 1)) > 57 Then ' SOLO ASIENTOS DE C&B
                 Call RECUMOCAJ(ASIAN$, CVDAT(FEASAN%), OKEYY%)
                 If OKEYY% > 0 Then
                   HUBORECU% = 1
                   GoTo 89
                 End If
               End If
               '
               'Call MENSERR(24, "Comprobante " + ASIAN$ + " No Balanceado.")
               If COMALTER%("Comprobante " + ASIAN$ + " No Balanceado.\\Continuar\Verificar y Recuperar") = 2 Then
                 Call RECUNOBAL(ASIAN$, FEASAN%)
                 Call COMUNI("Debe Ahora Volver a Pedir el Reporte.")
                 GoTo 99
               End If
             End If
89           SUMASI# = 0
           End If
           ASIAN$ = NIDAS$
           FEASAN% = FEAS%
           NASDEF& = CVS(Mid$(X$, 19, 4))
         Call REPLA(TTXX$, FORMATNUM$(NASDEF, "F7.0"), 24, 7)
           COCUE$ = TRIM$(Mid$(X$, 87, 16))
         Call REPLA(TTXX$, COCUE$, 33, 12)
           REFPASE$ = TRIM$(Mid$(X$, 103, 64))
         Call REPLA(TTXX$, REFPASE$, 47, 32)
'atencion
           IDEBE# = ROUND#(CVD(Mid$(X$, 167, 8)), 2)
           IHABE# = ROUND#(CVD(Mid$(X$, 175, 8)), 2)
           If IDEBE# < 0 Then IHABE# = IHABE# + Abs(IDEBE#): IDEBE# = 0
           If IHABE# < 0 Then IDEBE# = IDEBE# + Abs(IHABE#): IHABE# = 0
           If COCUE$ <> "" Then
                SUMASI# = SUMASI# + IDEBE# - IHABE#
           End If
         Call REPLA(TTXX$, FORMATNUM$(IDEBE#, "F14.2"), 80, 14)
         Call REPLA(TTXX$, FORMATNUM$(IHABE#, "F14.2"), 94, 14)
           DEMOVI$ = Mid$(X$, 191, 40)
         Call REPLA(TTXX$, DEMOVI$, 129, 40)
         DIARPAN.List1.AddItem TTXX$
       End If
     End If
   Next U&
   '
   If HUBORECU% > 0 Then
     Call COMUNI("Movimientos No Balanceados Recuperados.\  \Debe Volver a Pedir el Reporte.")
     GoTo 99
   End If
   '
   If Option1(1).Value = True Then
       DIARPAN.Show 1
     Else
       Call DIARPAN.LIDIARPRIN
   End If
   '
99 Call CIERRARCH("DIARPRIM")
   DIARPAN.List1.Clear
   Command15.Enabled = True
   Screen.MousePointer = 1
   
End Sub

Sub RECUNOBAL(ASIAN$, FEASAN%)
   ' RECUPERACION COMPROBANTES NO BALANCEADOS
   If DERACCE%("RECUNOBA", "Recuperación Comprobantes no Balanceados") < 2 Then
     Exit Sub
   End If
   '
   FEASEL = CDbl(CVDAT(FEASAN%))
   Call ABRECABAN
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + " WHERE CodiCom_Cor = '" & TRIM$(ASIAN$) & "' "
   SQLX$ = SQLX$ + " AND CDBL(FechConta) = " & FEASEL & " "
   SQLX$ = SQLX$ + " ORDER BY FECHCONTA ASC, FEHOREAL ASC;"
   '
   Set CABASELE = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   RECUNOBA.List1.Clear
   With CABASELE
     On Error Resume Next
     .MoveLast
     If Err > 0 Then
       On Error GoTo 0
       Call COMUNI("Imposible Recuperar Des-Balance - Consulte")
       Exit Sub
     End If
     '
     NORI% = 0
     .MoveFirst
     Do While Not .EOF
       NORI% = NORI% + 1
       .Edit
       !ValAbsComp = NORI%
       .Update
       TTXX$ = Space$(80)
       Call REPLA(TTXX$, FORMATNUM$(NORI%, "I4"), 1, 4)
       Call REPLA(TTXX$, CODCUE$(!CueContaI), 7, 16)
       Call REPLA(TTXX$, !RefeCom, 25, 32)
         IDEBE# = !IDEBECOMP
         IHABE# = !IHABECOMP
       Call REPLA(TTXX$, FORMATNUM(IDEBE#, "F14.2"), 57, 14)
       Call REPLA(TTXX$, FORMATNUM(IHABE#, "F14.2"), 71, 14)
       RECUNOBA.List1.AddItem TTXX$
     .MoveNext
     Loop
     '
     RECUNOBA.List2.Clear
     RECUNOBA.Label6 = ""
     FEX = FEASAN%
     RECUNOBA.Caption = "Recuperación Comprobante No Balanceado " + TRIM$(ASIAN$) + " - " + FECHATEX$(FEX)
     Call RECUNOBA.TOTALI
     '
     Screen.MousePointer = 1
     RECUNOBA.Show 1
     '
     If RECUNOBA.Label6 = "OK" Then
       For KU& = 1 To RECUNOBA.List2.ListCount
         TTXX$ = RECUNOBA.List2.List(KU& - 1)
         NORI% = Val(Mid$(TTXX$, 1, 4))
105      .FindFirst "ValAbsComp = " & NORI%
         If .NoMatch = False Then
           .Delete
         End If
       Next KU&
     End If
     '
     .MoveFirst
     Do While Not .EOF
       .Edit
       !ValAbsComp = Abs(!IDEBECOMP - !IHABECOMP)
       .Update
     .MoveNext
     Loop
     '
   End With
   '
End Sub

Private Sub Frame6_DblClick()
   DES1% = FECHANUM(Label5)
   HAS1% = FECHANUM(Label6)
   Call DESHASFECHA(DES1%, HAS1%, PERIO1$)
   If PERIO1$ <> "" Then
      Label5 = FECHATEX$(DES1%)
      Label6 = FECHATEX$(HAS1%)
   End If
End Sub

Private Sub impresora_Click()
    Option1(2).Value = True
    Call Command15_Click

End Sub

'
Private Sub label5_DblClick()
  Call Command21_Click
End Sub

Private Sub Label6_DblClick()
   Call Command4_Click
End Sub

Private Sub pantalla_Click()
    Option1(1).Value = True
    Call Command15_Click
End Sub

Private Sub salir_Click()
  Call Command1_Click
End Sub

Private Sub Text1_Change()
   Label1 = DENOCUE$(CUEINT%(Text1))
End Sub

Private Sub Text1_DblClick()
   Call Command8_Click
End Sub

Private Sub Text2_Change()
   Label2 = DENOCUE$(CUEINT%(Text2))
End Sub

Private Sub Command8_Click()
    Call SELECHO("ECUECON")
    PRICUEX$ = VALACT1$("ECUECON")
    If PRICUEX$ <> "" Then
        Text1 = PRICUEX$
    End If
End Sub

Private Sub Command9_Click()
    Call SELECHO("ECUECON")
    ULTCUEX$ = VALACT1$("ECUECON")
    If ULTCUEX$ <> "" Then
        Text2 = ULTCUEX$
    End If
End Sub

Private Sub Command21_Click()
   Call SELECHO("SELFECHA")
   FESELEC$ = VALACT1$("SELFECHA")
   '
   If FESELEC$ <> "" Then
     FFII% = FECHANUM(FESELEC$)
     If FFII% >= COMEJ% Then
       If FFII% <= FINEJ% Then
         Label5 = FECHATEX$(FFII%)
         If FECHANUM(Label6) < FFII% Then
           Label6 = Label5
         End If
       End If
     End If
   End If
End Sub

Private Sub Command4_Click()
   Call SELECHO("SELFECHA")
   FESELEC$ = VALACT1$("SELFECHA")
   '
   If FESELEC$ <> "" Then
     FFII% = FECHANUM(FESELEC$)
     If FFII% >= COMEJ% Then
       If FFII% <= FINEJ% Then
         Label6 = FECHATEX$(FFII%)
         If FECHANUM(Label5) > FFII% Then
           Label5 = Label6
         End If
       End If
     End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
End Sub


Private Sub Text10_Change()
    Text9.TEXT = ECOARCH("ECUECON", Text10.TEXT)
End Sub

Private Sub Text5_DblClick()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text5.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text5_GotFocus()
   Text5.SelStart = 0
   Text5.SelLength = Len(Text5.TEXT)
End Sub

Private Sub Text6_GotFocus()
   Text6.SelStart = 0
   Text6.SelLength = Len(Text6.TEXT)
End Sub

Private Sub Text8_Change()
    Text7.TEXT = ECOARCH("ECUECON", Text8.TEXT)
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("ECUECON")
    PRICUEX$ = VALACT1$("ECUECON")
    If PRICUEX$ <> "" Then
        Text10.TEXT = PRICUEX$
    End If
End Sub

Private Sub Text6_DblClick()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text6.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text7_CLICK()
    Text8.SetFocus
End Sub

Private Sub Text8_DblClick()
    Call SELECHO("ECUECON")
    ULTCUEX$ = VALACT1$("ECUECON")
    If ULTCUEX$ <> "" Then
        Text8.TEXT = ULTCUEX$
    End If
End Sub

Private Sub Text9_CLICK()
    Text10.SetFocus
End Sub

Sub LIMAYORPAN()
   '
   DES% = FECHANUM(Text6.TEXT)
   If DES% < COMEJ% Then
      DES% = COMEJ%
      FEX = DES%
      Text6.TEXT = FECHATEX$(FEX)
      Text6.Refresh
      Exit Sub
   End If
   '
   HAS% = FECHANUM(Text5.TEXT)
   FULTIAS% = FULTAS%
   If FULTASDIA% > FULTIAS% Then FULTIAS% = FULTASDIA%
   If HAS% > FULTIAS% Then
      HAS% = FULTIAS%
      FEX = HAS%
      Text5.TEXT = FECHATEX$(FEX)
      Text5.Refresh
      Exit Sub
   End If
   '
   If DES% > HAS% Then
      Call MENSERR(24, "Rango de Fechas Imposible")
      Exit Sub
   End If
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
   Frame1.Refresh
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
      Call LEEPASDIA(REGDIA%, 1)
      If TRIM$(DETAPAS$) = "" Then
         Call LEEPASDIA(REGDIA%, 2)
         DETAPAS$ = REFASIEN$
      End If
      MARCADIA$ = ""
      If FECHASIEN% < DES% Then
          TRANSPORTE# = TRANSPORTE# + IMPOPAS#
          GoTo 19
        ElseIf FECHASIEN% <= HAS% Then
          GoSub 90
      End If
19 Next KK%
   '
   ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   PMC$ = PUNMOCAJA$(CUE%)
   FIN& = Len(PMC$)
   ASICAPEN# = 0
   For KKL& = 1 To Len(PMC$) Step 4
      UU& = KKL& + 1
      Call TRACE(20, UU&, FIN&)
      REGDIAL& = CVS(Mid$(PMC$, KKL&, 4))
      Call LEESUCAJA(REGDIAL&)
      If Abs(CVS(D19$)) < 0.1 Then
         ASICAPEN# = ASICAPEN# + CVD(D16$)
      End If
29 Next KKL&
   '
   If Abs(ASICAPEN#) > 0.005 Then
      FECHASIEN% = HOY(HOS$)
      NROASIEN = 0
      NROASDEF% = 0
      MARCADIA$ = " "
      DETAPAS$ = "Asientos de Caja Pendientes"
      IMPOPAS# = ASICAPEN#
      If FECHASIEN% < DES% Then
          TRANSPORTE# = TRANSPORTE# + IMPOPAS#
          GoTo 39
        ElseIf FECHASIEN% <= HAS% Then
          GoSub 90
      End If
   End If
   '
39 If Abs(TRANSPORTE#) >= 0.005 Then
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
   If JJ& < 1 Then
        Call MENSERR(24, "Cuenta " + CODENOCUE$ + "\sin Movimientos en el Ejercicio.")
      Else
         Call CONECRUN("*MAYOCON/" + CODENOCUE$, "Libro Mayor por Cuenta")
   End If
   GoTo 10
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

Sub LIMAYORPRIN()
   '
   Dim SUDEBIN#(16384)
   Dim SUHABIN#(16384)
   Dim SUDEBE#(16384)
   Dim SUHABE#(16384)
   Dim ASICAPE#(16384)
   Dim PUNTE$(16384)
   '
   DES% = FECHANUM(Text6.TEXT)
   If DES% < COMEJ% Then
      DES% = COMEJ%
      FEX = DES%
      Text6.TEXT = FECHATEX$(FEX)
      Text6.Refresh
   End If
   '
   HAS% = FECHANUM(Text5.TEXT)
   If HAS% > FULTASDIA% Then
      HAS% = FULTASDIA%
      FEX = HAS%
      Text5.TEXT = FECHATEX$(FEX)
      Text5.Refresh
   End If
   '
   If DES% > HAS% Then
      Call MENSERR(24, "Rango de Fechas Imposible")
      Exit Sub
   End If
   '
   DESCUE$ = Text10.TEXT
   If TRIM$(DESCUE$) = "" Then
      Call MENSERR(24, "Falta Rango de Cuentas")
   End If
   '
   HASCUE$ = Text8.TEXT
   If TRIM$(HASCUE$) = "" Then
      Call MENSERR(24, "Falta Rango de Cuentas")
   End If
   '
   If DESCUE$ > HASCUE$ Then
      Call MENSERR(24, "Rango de Cuentas Imposible")
      Exit Sub
   End If
   '
   INCLUCAJ% = Val(CONTROL$("CONTAGEN", "INCLUCAJ"))
   If INCLUCAJ% <> 1 Then INCLUCAJ% = 0
   RENUMERA% = Val(CONTROL$("CONTAGEN", "RENUMERA"))
   If RENUMERA% <> 1 Then RENUMERA% = 0
   '
   Screen.MousePointer = 11
   TODAS% = 0
   If Option1(0).Value = True Then TODAS% = 1
   '
   If INCLUCAJ% < 1 Then
     FIN& = MAXDIA&
     For KKL& = 1 To MAXDIA&
       UU& = KKL&
       Call TRACE(20, UU&, FIN&)
       Call LEESUCAJA(KKL&)
       '
       CUE% = CVI(D11$)
       FEAS% = CVI(D15$)
       IPAS# = CVD(D16$)
       If CUE% > 0 Then
         If CUE% <= 16384 Then
           If FEAS% <= HAS% Then
             If Abs(CVS(D19$)) < 0.1 Then
               ASICAPE#(CUE%) = ASICAPE#(CUE%) + CVD(D16$)
             End If
           End If
         End If
       End If
     Next KKL&
   End If
   '
   Screen.MousePointer = 11
   FIN& = MAXDIARIO%
   For KK% = 1 To MAXDIARIO%
      UU& = KK%
      Call TRACE(20, UU&, FIN&)
      Call LEEDIARIO(KK%)
      '
      If INCLUCAJ% > 0 Then
        If Asc(E18$) = 1 Then GoTo 29 ' SALTEA ASIENTOS RESUMEN DE CAJA
      End If
      '
      CUE% = CVI(E11$)
      FEAS% = CVI(E15$)
      IPAS# = CVD(E16$)
      If CUE% > 0 Then
         If CUE% <= 16384 Then
            If FEAS% < DES% Then
                  If IPAS# > 0 Then
                      SUDEBIN#(CUE%) = SUDEBIN#(CUE%) + (Int(100 * IPAS# + 0.5)) / 100
                     Else
                      SUHABIN#(CUE%) = SUHABIN#(CUE%) - (Int(100 * IPAS# + 0.5)) / 100
                  End If
                Else
                  If FEAS% <= HAS% Then
                     If IPAS# > 0 Then
                         SUDEBE#(CUE%) = SUDEBE#(CUE%) + (Int(100 * IPAS# + 0.5)) / 100
                        Else
                         SUHABE#(CUE%) = SUHABE#(CUE%) - (Int(100 * IPAS# + 0.5)) / 100
                     End If
                     KKL& = KK%
                     PUNTE$(CUE%) = PUNTE$(CUE%) + MKS$(KKL&)
                  End If
            End If
         End If
      End If
      '
29 Next KK%
   '
   If INCLUCAJ% = 1 Then
     FIN& = MAXDIA&
     For KKL& = 1 To MAXDIA&
       UU& = KKL&
       Call TRACE(20, UU&, FIN&)
       Call LEESUCAJA(KKL&)
       '
       CUE% = CVI(D11$)
       FEAS% = CVI(D15$)
       IPAS# = CVD(D16$)
       If CUE% > 0 Then
         If CUE% <= 16384 Then
            If FEAS% < DES% Then
                  If IPAS# > 0 Then
                      SUDEBIN#(CUE%) = SUDEBIN#(CUE%) + (Int(100 * IPAS# + 0.5)) / 100
                     Else
                      SUHABIN#(CUE%) = SUHABIN#(CUE%) - (Int(100 * IPAS# + 0.5)) / 100
                  End If
                Else
                  If FEAS% <= HAS% Then
                     If IPAS# > 0 Then
                         SUDEBE#(CUE%) = SUDEBE#(CUE%) + (Int(100 * IPAS# + 0.5)) / 100
                        Else
                         SUHABE#(CUE%) = SUHABE#(CUE%) - (Int(100 * IPAS# + 0.5)) / 100
                     End If
                     PUNTE$(CUE%) = PUNTE$(CUE%) + MKS$((-1) * KKL&)
                  End If
            End If
         End If
       End If
       '
     Next KKL&
   End If
   '
   If NX3% < 1 Then
      NX3% = FILEOPEN%(LUDAT$ + "INVECUE" + "." + EXTE$, 0, 16)
   End If
   '
   Screen.MousePointer = 11
   JJ& = 0
   Dim IVX As String * 16
   FIN& = LOF(NX3%) / 16
   For KK% = 1 To LOF(NX3%) / 16
      UU& = KK%
      Call TRACE(20, UU&, FIN&)
      Get #NX3%, KK%, IVX$
      COCUE$ = TRIM$(Left$(IVX$, 12))
      '
      If COCUE$ >= DESCUE$ Then
        If COCUE$ <= HASCUE$ Then
          CUE% = CVI(Mid$(IVX$, 13, 2))
          If CUE% > 0 Then
            If CUE% <= 16384 Then
              If TODAS% = 1 Or Abs(SUDEBIN#(CUE%)) + Abs(SUHABIN#(CUE%)) + Abs(SUDEBE#(CUE%)) + Abs(SUHABE#(CUE%)) + Abs(ASICAPE#(CUE%)) >= 0.005 Then
                '
                List1.Clear
                SUTRA# = 0
                TTTT = Len(PUNTE$(CUE%))
                For KL% = 1 To Len(PUNTE$(CUE%)) Step 4
                  KML& = CVS(Mid$(PUNTE$(CUE%), KL%, 4))
                  If KML& > 0 Then
                     KM% = KML&
                     Call LEEDIARIO(Abs(KM%))
                     VALOSOR% = CVI(E15$)
                     If RENUMERA% = 1 Then VALOSOR% = CVI(E13$)
                     ZZ$ = "A" + CSTRING$(MKI$(VALOSOR%), 1, 6, 0, 2) + Space$(9)
                     Call REPLA(ZZ$, "A", 8, 1)
                     Call REPLA(ZZ$, Str$(KM%), 9, 8)
                     List1.AddItem ZZ$
                   ElseIf KML& < 0 Then
                     Call LEESUCAJA(Abs(KML&))
                     VALOSOR% = CVI(D15$)
                     If RENUMERA% = 1 Then VALOSOR% = CVI(D13$)
                     ZZ$ = "A" + CSTRING$(MKI$(VALOSOR%), 1, 6, 0, 2) + Space$(9)
                     Call REPLA(ZZ$, "A", 8, 1)
                     Call REPLA(ZZ$, Str$(KML&), 9, 8)
                     List1.AddItem ZZ$
                 End If
                Next KL%
                '
                If JJ& > 0 Then
                  XX$ = REGBLAN$("LIMAYOL"): GoSub 95
                  XX$ = String$(128, "="): GoSub 95
                  XX$ = REGBLAN$(LIMAYOL$): GoSub 95
                End If
                '
                FECHASIEN% = 0: NROASIEN = 0: NROASDEF = 0: MARCADIA$ = "": REFE$ = "": IDEBE# = 0: IHABE# = 0
                REFE$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%)): GoSub 90
                REFE$ = String$(Len(REFE$), "*"): GoSub 90
                XX$ = REGBLAN$("LIMAYOL"): GoSub 95
                '
                FEX = DES
                FECHASIEN% = DIANTE(FEX)
                NROASIEN = 0
                NROASDEF = 0
                MARCADIA$ = " "
                REFE$ = "Transporte"
                IDEBE# = SUDEBIN#(CUE%)
                IHABE# = SUHABIN#(CUE%)
                GoSub 90
                '
                XX$ = REGBLAN$("LIMAYOL"): GoSub 95
                '
                For KL% = 1 To List1.ListCount
                  KML& = Val(Mid$(List1.List(KL% - 1), 9))
                  If KML& > 0 Then
                     KM% = KML&
                     Call LEEPASDIA(KM%, 2)
                     FECHASIEN% = CVI(E15$)
                     NROASIEN = CVS(E12$)
                     NROASDEF% = CVI(E13$)
                     MARCADIA$ = "  "
                     REFE$ = TRIM$(DETAPAS$)
                     If REFE$ <> "" Then REFE$ = REFE$ + " - "
                     REFE$ = REFE$ + TRIM$(REFASIEN$)
                     IPAS# = CVD(E16$)
                     If IPAS# > 0 Then
                         IDEBE# = IPAS#: IHABE# = 0
                       Else
                         IDEBE# = 0: IHABE# = Abs(IPAS#)
                     End If
                     If Abs(IDEBE#) + Abs(IHABE#) > 0.005 Then GoSub 90
                   ElseIf KML& < 0 Then
                     Call LEEPASCAJ((Abs(KML&)), 2)
                     FECHASIEN% = CVI(D15$)
                     NROASIEN = CVS(D12$)
                     NROASDEF% = CVI(D13$)
                     MARCADIA$ = "  "
                     REFE$ = TRIM$(DETAPAS$)
                     If REFE$ <> "" Then REFE$ = REFE$ + " - "
                     REFE$ = REFE$ + TRIM$(REFASIEN$)
                     IPAS# = CVD(D16$)
                     If IPAS# > 0 Then
                         IDEBE# = IPAS#: IHABE# = 0
                       Else
                         IDEBE# = 0: IHABE# = Abs(IPAS#)
                     End If
                     If Abs(IDEBE#) + Abs(IHABE#) > 0.005 Then GoSub 90
                  End If
                Next KL%
                '
                If Abs(ASICAPE#(CUE%)) >= 0.005 Then
                   XX$ = REGBLAN$("LIMAYOL"): GoSub 95
                   FECHASIEN% = HAS%
                   NROASIEN = 0
                   NROASDEF% = 0
                   MARCADIA$ = "* "
                   REFE$ = "Asientos de Caja Pendientes"
                   IPAS# = ASICAPE#(CUE%)
                   If IPAS# > 0 Then
                        IDEBE# = IPAS#: IHABE# = 0
                       Else
                        IDEBE# = 0: IHABE# = Abs(IPAS#)
                    End If
                    GoSub 90
                End If
                '
              End If
            End If
          End If
        End If
      End If
      '
   Next KK%
   '
   Close #NX3%: NX3% = 0
   Call SETULTREG("LIMAYOL", JJ&)
   Call CIERRARCH("LIMAYOL")
   '
   Call HEADERS("LIMAYOL", "")
   Call HEADERS("LIMAYOL", "Corresponde a: " + EJACT$)
   '
   Screen.MousePointer = 1
   Call CONECRUN("*LIMAYOL", "Libro Mayor por Cuenta")
   '
   Exit Sub

90 XX$ = REGBLAN$("LIMAYOL")
   FFAX$ = Space$(8)
   If FECHASIEN% > 0 Then
      FEX = FECHASIEN%: FFAX$ = FECHATEX$(FEX)
   End If
   Call REPLA(XX$, FFAX$, 1, 8)
   If RENUMERA% = 1 Then
     NROASIEN = NROASDEF%
   End If
   Call REPLA(XX$, MKS$(NROASIEN), 9, 4)
   'Call REPLA(XX$, MKI$(NROASDEF%), 13, 2)
   Call REPLA(XX$, MARCADIA$, 15, 2)
   Call REPLA(XX$, REFE$, 17, 56)
   Call REPLA(XX$, MKD$(IDEBE#), 73, 8)
   Call REPLA(XX$, MKD$(IHABE#), 81, 8)
   SUTRA# = SUTRA# + IDEBE# - IHABE#
   Call REPLA(XX$, MKD$(SUTRA#), 89, 8)
   '
95 JJ& = JJ& + 1
   Call GRAREG("LIMAYOL", XX$, JJ&)
   Return
   '
End Sub

Private Sub Form_Load()
    '
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6
      Me.Option1(1).BackColor = &HFCD7B6
      Me.Option1(2).BackColor = &HFCD7B6
      MARCOBARRA.Top = 3300
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
    End If
    
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Screen.MousePointer = 1
    '
End Sub


Private Sub Text2_DblClick()
   Call Command9_Click
End Sub
