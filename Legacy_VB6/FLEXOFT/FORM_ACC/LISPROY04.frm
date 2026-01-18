VERSION 5.00
Begin VB.Form LISPROY04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administrador de Proyectos  - Consultas y Listados"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   5235
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Gestión de Proyectos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   115
      TabIndex        =   14
      Top             =   105
      Width           =   4005
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1050
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   15
         Top             =   1050
         Width           =   2325
         Begin VB.CommandButton Command14 
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
            Picture         =   "LISPROY04.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Anulación de Proyectos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Anulación no Reversible"
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
            TabIndex        =   17
            Top             =   105
            Width           =   1275
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   18
         Top             =   525
         Width           =   2325
         Begin VB.CommandButton Command46 
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
            Picture         =   "LISPROY04.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Cierre de Proyectos"
            Top             =   0
            Width           =   530
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cierre de Proyectos"
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
            TabIndex        =   20
            Top             =   60
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Consultas - Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2745
      Left            =   115
      TabIndex        =   0
      Top             =   2100
      Width           =   4005
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1050
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   1
         Top             =   1050
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
            Left            =   -20
            Picture         =   "LISPROY04.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Resumen de Cargos por Proyecto "
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Resumen de Cargos por Proyecto "
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
            Left            =   550
            TabIndex        =   3
            Top             =   50
            Width           =   1800
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   7
         Top             =   525
         Width           =   2325
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
            Left            =   -20
            Picture         =   "LISPROY04.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Detalle de Cargos por Proyecto "
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Detalle de Cargos por Proyecto "
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
            Left            =   550
            TabIndex        =   9
            Top             =   50
            Width           =   1590
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0C0&
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
         Left            =   420
         TabIndex        =   4
         Top             =   1890
         Width           =   3165
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0E0C0&
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
            Left            =   1785
            TabIndex        =   6
            Top             =   250
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0C0&
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
            Index           =   0
            Left            =   315
            TabIndex        =   5
            Top             =   250
            Width           =   1380
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   6000
      Left            =   4305
      ScaleHeight     =   5940
      ScaleWidth      =   1005
      TabIndex        =   10
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command4 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "LISPROY04.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Ayuda en Línea"
         Top             =   900
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   12
         Top             =   3885
         Width           =   645
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -105
            TabIndex        =   13
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
         Height          =   690
         Left            =   105
         Picture         =   "LISPROY04.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   210
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -225
         Picture         =   "LISPROY04.frx":107C
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   1410
      End
   End
End
Attribute VB_Name = "LISPROY04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command14_Click()
  '
  Command14.Enabled = False
    Call COMUNI("No Disponible en la Presente Versión")
  Command14.Enabled = True
  '
End Sub

 Sub Command3_Click()
  '
  Command3.Enabled = False
  Call LICARCUM
  Command3.Enabled = True
  '
End Sub

Private Sub Command4_Click()
  '
  Command4.Enabled = False
    Call HELPONLINE("AMAPROY")
  Command4.Enabled = True
  '
End Sub

Private Sub Command46_Click()
  Command46.Enabled = False
  '
  If DERACCE%("CIERPROYE", "Cierre de Proyecto") < 2 Then
    Exit Sub
  End If
  '
  REBLA1$ = REGBLAN$("INDIPROY")
  JJ& = 0
  HOII% = HOY(HO$)
  Call COPYSTRU("INDIPROY", "SELPROY")
  SQLX$ = "SELECT * FROM TAPROYEC"
  SQLX$ = SQLX$ + " WHERE STAT_PROY < 3 "
  Call ABREPLANCUEN
  Set CIETMP = Contable.OpenRecordset(SQLX$, 4)
  With CIETMP
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      Call MENSERR(24, "No Existen Proyectos que \Puedan Cerrarse")
      GoTo 99
    End If
    On Error GoTo 0
    Do While .EOF = False
     XZ$ = REBLA1$
     Call REPLA(ZX$, !Cod_Proy, 1, 8)
     Call REPLA(ZX$, TRIM$(FANTCLI$(!Clien_Proy)) & " - " & !Deno_Proy, 9, 48)
     JJ& = JJ& + 1
     Call GRAREG("!SELPROY", ZX$, JJ&)
     .MoveNext
    Loop
  End With
  Call SETULTREG("!SELPROY", JJ&)
  '
  Call FRESHECHO("!SELPROY")
  Call SELECHO("!SELPROY")
  NUOPROY$ = TRIM$(VALACT1$("!SELPROY"))
  If TRIM$(NUOPROY$) = "" Then
    GoTo 99
  End If
  OPX1% = COMALTER%("¿ Confirma Cerrar Proyecto:  " + TRIM$(NUOPROY$) + " ?\\No, Cancelar\Si, Cerrar")
  If OPX1% = 2 Then
    Call ABREPLANCUEN
    SQLX$ = "SELECT * FROM TAPROYEC"
    Set OTTMP = Contable.OpenRecordset(SQLX$, 2)
    With OTTMP
      .FindFirst "  Cod_Proy = '" & NUOPROY$ & "' "
      If .NoMatch = True Then
         On Error GoTo 0
         Call MENSERR(24, "No se Ha Completado \el Cierre del Proyecto.")
         GoTo 99
      End If
      On Error GoTo 0
      '
      .Edit
         !Stat_Proy = 3 'Cerrado.
         !FeCie_Proy = CVDAT(HOII%) 'Fecha de Cierre Hoy.
      .Update
    End With
    'Marco Fecha Hasta Vigencia en la tabla de Centro de Costos
    'Call FILTERDELETE("TACENCOS", 1, NUOPROY$)
    XZ$ = RECFILT$("TACENCOS", 1, NUOPROY$)
    NR& = CVS(XZ$)
    YY$ = REGLEIDO$("TACENCOS", NR&)
    Call REPLA(YY$, MKI$(HOII%), 47, 2)
    Call GRAREG("TACENCOS", YY$, NR&)
    '
    Call COMUNI("Cierre Completo.")
    '
  End If
  '
99 Command46.Enabled = True
End Sub

Sub Command8_Click()
  Command8.Enabled = False
    Call LICARDET
  Command8.Enabled = True
End Sub
Sub LICARCUM() 'Listado de Cargos por Proyecto (ACUMULADO)
  '
  Call DESHASFECHA(DES%, HAS%, PERIO$)
  If TRIM$(PERIO$) = "" Then Exit Sub
  '
  Screen.MousePointer = 11
  DES1 = Int(CDbl(CVDAT(DES%)))
  HAS1 = Int(CDbl(CVDAT(HAS%)))
  REBLA1$ = REGBLAN$("DCARPRO")
  JJ& = 0
  '
  Call GENINDIPROY
  FIN& = ULTREG&("!SELPROY")
  For J& = 1 To FIN&
    X$ = REGLEIDO$("!SELPROY", J&)
    CCX1$ = Mid$(X$, 1, 8)
    '1.- Cargos por Facturación de Proveedores
    IMPX1# = 0
    IMPX2# = 0
    IMPX3# = 0
    IMPX4# = 0
    IMPXT# = 0
    '
    Call ABRECABAN
    SQLX$ = "SELECT * FROM CAJABANC"
    SQLX$ = SQLX$ + " WHERE CenCosto = '" & CCX1$ & "' "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) >= " & DES1 & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HAS1 & " "
    SQLX$ = SQLX$ + "ORDER BY FechConta "
    Set CABTMP = CueCorri.OpenRecordset(SQLX$, 4)
    With CABTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        GoTo 5
      End If
      On Error GoTo 0
      Do While .EOF = False
        IMPX1# = IMPX1# + !IDEBECOMP - IHABECOMP
        .MoveNext
      Loop
      '
    End With
    '2.- Cargos por Asientos Contables
5   Call ABREPLANCUEN
    SQLX$ = "SELECT * FROM DetaAsien "
    SQLX$ = SQLX$ + " WHERE CenCos = '" & CCX1$ & "' "
    SQLX$ = SQLX$ + " AND StatPase > 0 "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechAsien)) >= " & DES1 & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechAsien)) <= " & HAS1 & " "
    SQLX$ = SQLX$ + "ORDER BY FechAsien "
    Set ASITMP = Contable.OpenRecordset(SQLX$, 4)
    With ASITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        GoTo 6
      End If
      On Error GoTo 0
      Do While .EOF = False
        IMPX2# = IMPX2# + !IDEBEPASE - IHABEPASE
        .MoveNext
      Loop
    End With
    '3.- Cargos por Vales de Consumo de Materiales
6   Call ABRESTOCKS
    CMO$ = TRIM$(CONTROL$("", "CODCONSU"))
    If CMO$ = "" Then CMO$ = "VC"
    '
    SQLX$ = "SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE INT(CDBL( FechMov )) >= " & DES1 & " "
    SQLX$ = SQLX$ + " AND INT(CDBL( FechMov )) <= " & HAS1 & " "
    SQLX$ = SQLX$ + " AND CodiMovi = '" & CMO$ & "'"
    CCXX1$ = "CC-" + CCX1$
    SQLX$ = SQLX$ + " AND  DoSecCor  = '" & CCXX1$ & "' "
    SQLX$ = SQLX$ + "ORDER BY FechMov "
    Set VALETMP = Stocks.OpenRecordset(SQLX$, 4)
    With VALETMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        GoTo 7
      End If
      On Error GoTo 0
      Do While .EOF = False
        IMPX3# = IMPX3# + (!ValoUnit * !CantSalid)
        .MoveNext
      Loop
    End With
    '
    '4.- Por Informes de Proyecto
7   Call ABREORFAB
    SQLX$ = "SELECT * FROM INFOPROY"
    SQLX$ = SQLX$ + " WHERE INT(CDBL( Fecha_Inf )) >= " & DES1 & " "
    SQLX$ = SQLX$ + " AND INT(CDBL( Fecha_Inf )) <= " & HAS1 & " "
    SQLX$ = SQLX$ + " AND Proy_Inf = '" & CCX1$ & "' "
    SQLX$ = SQLX$ + "ORDER BY Fecha_Inf "
    Set INFTMP = Ofabric.OpenRecordset(SQLX$, 4)
    With INFTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
       GoTo 8
     End If
     On Error GoTo 0
     Do While .EOF = False
       IMPX4# = IMPX4# + !Total_Inf
       .MoveNext
     Loop
    End With
    '
8   Z$ = REBLA1$
    Call REPLA(Z$, CCX1$, 1, 8)
    FE1% = 0
    Call REPLA(Z$, MKI(FE1%), 9, 2)
    Call REPLA(Z$, "  ", 11, 18)
    Call REPLA(Z$, "  ", 29, 64)
    Call REPLA(Z$, MKD$(IMPX4#), 111, 8)
    Call REPLA(Z$, MKD$(IMPX3#), 119, 8)
    Call REPLA(Z$, MKD$(IMPX2#), 127, 8)
    Call REPLA(Z$, MKD$(IMPX1#), 135, 8)
    IMPXT# = IMPX1# + IMPX2# + IMPX3# + IMPX4#
    Call REPLA(Z$, MKD$(IMPXT#), 143, 8)
    If IMPXT# > 0.005 Then
      JJ& = JJ& + 1
      Call GRAREG("DCARPRO", Z$, JJ&)
    End If
    '
  Next J&
  '
  Call SETULTREG("DCARPRO", JJ&)
  If JJ& = 0 Then
    Call MENSERR(24, "No Existen Cargos para \ el Rango de Fechas \ Seleccionado")
    Screen.MousePointer = 1
    Exit Sub
  End If
  Call FILESORT("DCARPRO", "DCARPRO", 1, 1) 'PROYECTO
  Call HEADERS("DCARPRO", "")
  Call HEADERS("DCARPRO", "Período: " + PERIO$)
  Call FINAL("*LCARPRO")
  '
End Sub
Sub LICARDET() 'Listado de Cargos por Proyecto ( DETALLE )
  '
  Call DESHASFECHA(DES%, HAS%, PERIO$)
  If TRIM$(PERIO$) = "" Then Exit Sub
  '
  Screen.MousePointer = 11
  DES1 = Int(CDbl(CVDAT(DES%)))
  HAS1 = Int(CDbl(CVDAT(HAS%)))
  REBLA1$ = REGBLAN$("DCARPRO")
  JJ& = 0
  '
  '1.- Cargos por Facturación de Proveedores
  Call ABRECABAN
  SQLX$ = "SELECT * FROM CAJABANC"
  SQLX$ = SQLX$ + " WHERE TRIM ( CenCosto ) <> '' "
  SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) >= " & DES1 & " "
  SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HAS1 & " "
  SQLX$ = SQLX$ + "ORDER BY FechConta "
  Set CABTMP = CueCorri.OpenRecordset(SQLX$, 4)
  With CABTMP
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      GoTo 5
    End If
    On Error GoTo 0
    Do While .EOF = False
      Z$ = REBLA1$
      CECO1$ = !CenCosto
      If ECOARCH$("INDIPROY", CECO1$) = "" Then GoTo 44
      Call REPLA(Z$, CECO1$, 1, 8)
      FE1% = CVINT(!FECHCONTA)
      Call REPLA(Z$, MKI(FE1%), 9, 2)
      Call REPLA(Z$, !CodiCom_Cor, 11, 18)
      Call REPLA(Z$, !DescriMov, 29, 64)
      IMPX1# = !IDEBECOMP - IHABECOMP
      Call REPLA(Z$, MKD$(IMPX1#), 93, 8)
      Call REPLA(Z$, MKD$(IMPX1#), 135, 8)
      Call REPLA(Z$, MKD$(IMPX1#), 143, 8)
      JJ& = JJ& + 1
      Call GRAREG("DCARPRO", Z$, JJ&)
44    .MoveNext
    Loop
    '
  End With
  '2.- Cargos por Asientos Contables
5 Call ABREPLANCUEN
  SQLX$ = "SELECT * FROM DetaAsien "
  SQLX$ = SQLX$ + " WHERE TRIM ( CenCos ) <> ''"
  SQLX$ = SQLX$ + " AND StatPase > 0 "
  SQLX$ = SQLX$ + "AND INT(CDBL(FechAsien)) >= " & DES1 & " "
  SQLX$ = SQLX$ + "AND INT(CDBL(FechAsien)) <= " & HAS1 & " "
  SQLX$ = SQLX$ + "ORDER BY FechAsien "
  Set ASITMP = Contable.OpenRecordset(SQLX$, 4)
  With ASITMP
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      GoTo 6
    End If
    On Error GoTo 0
    Do While .EOF = False
      Z$ = REBLA1$
      CECO1$ = !CenCos
      If ECOARCH$("INDIPROY", CECO1$) = "" Then GoTo 55
      Call REPLA(Z$, CECO1$, 1, 8)
      FE1% = CVINT(!FECHASIEN)
      Call REPLA(Z$, MKI(FE1%), 9, 2)
      Call REPLA(Z$, !NuIdAsien, 11, 18)
      Call REPLA(Z$, !RefePase, 29, 64)
      IMPX1# = !IDEBEPASE - IHABEPASE
      Call REPLA(Z$, MKD$(IMPX1#), 93, 8)
      Call REPLA(Z$, MKD$(IMPX1#), 127, 8)
      Call REPLA(Z$, MKD$(IMPX1#), 143, 8)
      JJ& = JJ& + 1
      Call GRAREG("DCARPRO", Z$, JJ&)
55    .MoveNext
    Loop
  End With
  '3.- Vale de Consumo de Materiales
6 Call ABRESTOCKS
  CMO$ = TRIM$(CONTROL$("", "CODCONSU"))
  If CMO$ = "" Then CMO$ = "VC"
  '
  SQLX$ = "SELECT * FROM MOVISTO "
  SQLX$ = SQLX$ + " WHERE INT(CDBL( FechMov )) >= " & DES1 & " "
  SQLX$ = SQLX$ + " AND INT(CDBL( FechMov )) <= " & HAS1 & " "
  SQLX$ = SQLX$ + " AND CodiMovi = '" & CMO$ & "'"
  SQLX$ = SQLX$ + " AND TRIM( DoSecCor ) <> '' "
  SQLX$ = SQLX$ + "ORDER BY FechMov "
  Set VALETMP = Stocks.OpenRecordset(SQLX$, 4)
  With VALETMP
   On Error Resume Next
   .MoveFirst
   If Err Then
     On Error GoTo 0
     GoTo 7
   End If
   On Error GoTo 0
   Do While .EOF = False
     Z$ = REBLA1$
     CECO1$ = Mid$(!DoSecCor, 4)
     If ECOARCH$("INDIPROY", CECO1$) = "" Then GoTo 66
     Call REPLA(Z$, CECO1$, 1, 8)
     FE1% = CVINT(!FechMov)
     Call REPLA(Z$, MKI(FE1%), 9, 2)
     Call REPLA(Z$, !DoPriCor, 11, 18)
     REFEVAL$ = "(" + CODEXT$(!Cod_Inte) + ")" + !ReferCor
     Call REPLA(Z$, REFEVAL$, 29, 64)
     IMPX1# = (!ValoUnit * !CantSalid)
     Call REPLA(Z$, MKD$(IMPX1#), 93, 8)
     Call REPLA(Z$, MKD$(IMPX1#), 119, 8)
     Call REPLA(Z$, MKD$(IMPX1#), 143, 8)
     JJ& = JJ& + 1
     Call GRAREG("DCARPRO", Z$, JJ&)
66   .MoveNext
   Loop
  End With
  '
  '4.- Por Informes de Proyecto
7 Call ABREORFAB
  SQLX$ = "SELECT * FROM INFOPROY"
  SQLX$ = SQLX$ + " WHERE INT(CDBL( Fecha_Inf )) >= " & DES1 & " "
  SQLX$ = SQLX$ + " AND INT(CDBL( Fecha_Inf )) <= " & HAS1 & " "
  SQLX$ = SQLX$ + "ORDER BY Fecha_Inf "
  Set INFTMP = Ofabric.OpenRecordset(SQLX$, 4)
  With INFTMP
   On Error Resume Next
   .MoveFirst
   If Err Then
     On Error GoTo 0
     GoTo 8
   End If
   On Error GoTo 0
   Do While .EOF = False
     Z$ = REBLA1$
     CECO1$ = !Proy_Inf
     If ECOARCH$("INDIPROY", CECO1$) = "" Then GoTo 77
     Call REPLA(Z$, CECO1$, 1, 8)
     FE1% = CVINT(!Fecha_Inf)
     Call REPLA(Z$, MKI(FE1%), 9, 2)
     Call REPLA(Z$, !Nume_Inf, 11, 18)
     Call REPLA(Z$, !Refer_Inf, 29, 64)
     IMPX1# = !Total_Inf
     Call REPLA(Z$, MKD$(IMPX1#), 93, 8)
     Call REPLA(Z$, MKD$(IMPX1#), 111, 8)
     Call REPLA(Z$, MKD$(IMPX1#), 143, 8)
     JJ& = JJ& + 1
     Call GRAREG("DCARPRO", Z$, JJ&)
77   .MoveNext
   Loop
  End With
  '
8 Call SETULTREG("DCARPRO", JJ&)
  If JJ& = 0 Then
    Call MENSERR(24, "No Existen Cargos para \ el Rango de Fechas \ Seleccionado")
    Screen.MousePointer = 1
    Exit Sub
  End If
  Call FILESORT("DCARPRO", "DCARPRO", 4, 4) 'COMPROBANTE
  Call FILESORT("DCARPRO", "DCARPRO", 3, 3) 'FECHA
  Call FILESORT("DCARPRO", "DCARPRO", 1, 1) 'PROYECTO
  Call HEADERS("DCARPRO", "")
  Call HEADERS("DCARPRO", "Período: " + PERIO$)
  Call FINAL("*DCARPRO")
  '
End Sub
Sub GENINDIPROY()
  '
  REBLA1$ = REGBLAN$("INDIPROY")
  JJ& = 0
  HOII% = HOY(HO$)
  Call COPYSTRU("INDIPROY", "SELPROY")
  SQLX$ = "SELECT * FROM TAPROYEC"
  Call ABREPLANCUEN
  Set CIETMP = Contable.OpenRecordset(SQLX$, 4)
  With CIETMP
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      GoTo 99
    End If
    On Error GoTo 0
    Do While .EOF = False
     XZ$ = REBLA1$
     Call REPLA(ZX$, !Cod_Proy, 1, 8)
     Call REPLA(ZX$, TRIM$(FANTCLI$(!Clien_Proy)) & " - " & !Deno_Proy, 9, 48)
     JJ& = JJ& + 1
     Call GRAREG("!SELPROY", ZX$, JJ&)
     .MoveNext
    Loop
  End With
99 Call SETULTREG("!SELPROY", JJ&)
  '
End Sub
