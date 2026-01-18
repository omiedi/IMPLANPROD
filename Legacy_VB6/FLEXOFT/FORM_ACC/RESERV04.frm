VERSION 5.00
Begin VB.Form RESERV04 
   BackColor       =   &H00EEEEC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reservas de Materiales"
   ClientHeight    =   6645
   ClientLeft      =   2040
   ClientTop       =   1590
   ClientWidth     =   5745
   ControlBox      =   0   'False
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
   ScaleHeight     =   6645
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   6840
      Left            =   4830
      ScaleHeight     =   6780
      ScaleWidth      =   1005
      TabIndex        =   22
      Top             =   0
      Width           =   1065
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
         Height          =   640
         Left            =   120
         Picture         =   "RESERV04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Configuración de Procesos y Formularios"
         Top             =   2625
         Width           =   650
      End
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
         Height          =   640
         Left            =   120
         Picture         =   "RESERV04.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   1785
         Width           =   650
      End
      Begin VB.CommandButton Command6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   120
         Picture         =   "RESERV04.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   950
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
         Height          =   640
         Left            =   120
         Picture         =   "RESERV04.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   300
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   23
         Top             =   5150
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "RESERV04.frx":0BA0
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5470
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EEEEC0&
      Caption         =   "Actualización"
      Height          =   3225
      Left            =   210
      TabIndex        =   8
      Top             =   210
      Width           =   4425
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1785
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   29
         Top             =   2310
         Width           =   2325
         Begin VB.CommandButton Command9 
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
            Picture         =   "RESERV04.frx":2AC2
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Reservas por O/F con Faltantes"
            Top             =   0
            Width           =   600
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Reservas por O/F con Faltantes"
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
            TabIndex        =   30
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1470
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   9
         Top             =   1845
         Width           =   2325
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
            Picture         =   "RESERV04.frx":2DCC
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Depuración de    la Base de Datos"
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
            TabIndex        =   11
            Top             =   0
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1100
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   18
         Top             =   1370
         Width           =   2325
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
            Height          =   600
            Left            =   0
            Picture         =   "RESERV04.frx":30D6
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Control Según O/Fs en Curso"
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
            Top             =   0
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   12
         Top             =   895
         Width           =   2325
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
            Left            =   0
            Picture         =   "RESERV04.frx":33E0
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Cancelación      de Reservas"
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
            TabIndex        =   14
            Top             =   0
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   15
         Top             =   420
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
            Left            =   0
            Picture         =   "RESERV04.frx":36EA
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Generar  Re- servas por O/F"
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
            Top             =   0
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEC0&
      Caption         =   "Consultas - Listados"
      Height          =   2895
      Left            =   210
      TabIndex        =   0
      Top             =   3570
      Width           =   4425
      Begin VB.OptionButton Option6 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Totales por Material Reservado"
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
         TabIndex        =   7
         Top             =   420
         Width           =   3165
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00EEEEC0&
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
         Left            =   210
         TabIndex        =   4
         Top             =   2100
         Width           =   4005
         Begin VB.OptionButton Option4 
            BackColor       =   &H00EEEEC0&
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
            Left            =   1575
            TabIndex        =   5
            Top             =   250
            Width           =   1275
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
            Height          =   490
            Left            =   3340
            Picture         =   "RESERV04.frx":39F4
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   90
            Width           =   645
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00EEEEC0&
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
            TabIndex        =   6
            Top             =   250
            Value           =   -1  'True
            Width           =   1380
         End
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Por Material"
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
         Top             =   1575
         Width           =   2955
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Por Número de O/Fabricación "
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
         Top             =   1260
         Width           =   3060
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Consulta General de Reservas por O/F"
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
         TabIndex        =   1
         Top             =   945
         Value           =   -1  'True
         Width           =   3795
      End
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "RESERV04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Hide
End Sub

Private Sub Command14_Click()
   Command14.Enabled = FALSEÇ
   Call COMUNI("No hay Opciones de Configuración\para esta Aplicación.")
   Command14.Enabled = True
End Sub

 Sub Command2_Click()
    OPCI% = 1
    If Option2.Value = True Then OPCI% = 2
    If Option3.Value = True Then OPCI% = 3
    If Option6.Value = True Then OPCI% = 4
    If OPCI% = 1 Then
        Call GENRESDAT
        Call CONECRUN("*CORESE1", "")
      ElseIf OPCI% = 2 Then
        Call GENRESDAT
        Call CONECRUN("*CORESE2", "")
      ElseIf OPCI% = 3 Then
        Call GENRESDAT
        Call CONECRUN("*CORESE3", "")
      ElseIf OPCI% = 4 Then
        Call GENRESDAT
        Call TOMATRES
      Else
        Hide
    End If
    Screen.MousePointer = 1
    '
End Sub

 Sub Command3_Click()
    Command3.Enabled = False
    Call GENRESER("")
    Command3.Enabled = True
End Sub

 Sub Command4_Click()
    Command4.Enabled = False
    Call ANURESER("")
    Command4.Enabled = True
End Sub

 Sub Command5_Click()
    Command5.Enabled = False
    '
    If DERACCE%("DEPURESE", "Depuración de Reservas de Materiales") < 2 Then GoTo 99
    Screen.MousePointer = 11
    Dim OFCUR$()
    CANOFPE% = 0
    Call APERBASDAT("ORFAB")
    Call APERBASDAT("STOCKS")
    Call APERBASDAT("RESERVA")
    '
    ' ELIMINA RESERVAS DE ORDENES CERRADAS, O ANULADAS
    SQLX$ = "SELECT * FROM ORDEFABR "
    Set OFTMP = Ofabric.OpenRecordset(SQLX$, 4)
    With OFTMP
       On Error Resume Next
       .MoveLast
       If Err Then
         Call MENSERR(24, "No Hay O/F's Pendientes.")
         GoTo 98
       End If
       On Error GoTo 0
       J% = .RecordCount
       ReDim OFCUR$(J%)
       .MoveFirst
       Do While .EOF = False
          NOFAX$ = !IdSubOr
          STATU% = !Statuorf
          If STATU% >= 3 Then
               SQLY$ = "SELECT * FROM RESERVA "
               SQLY$ = SQLY$ + "WHERE IdSubOr = '" & NOFAX$ & "' "
               Set RESETEMP = Stocks.OpenRecordset(SQLY$, dbOpenDynaset)
               With RESETEMP
                 On Error Resume Next
15               .MoveFirst
                 If Err < 1 Then
                   .Delete
                   GoTo 15
                 End If
                 On Error GoTo 0
               End With
             Else
               CANOFPE% = CANOFPE% + 1
               OFCUR$(CANOFPE%) = NOFAX$
          End If
       .MoveNext
       Loop
    End With
    '
    ' PONE EN CERO LOS CONSUMOS DE LAS QUE QUEDAN
    With Reserva
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
          Do While Not .EOF
             .Edit
             !CantCons = 0
             !Saldo_Entre = !CantRes
             .Update
          .MoveNext
          Loop
       End If
    End With
    '
    ' DECLARA LOS CONSUMIDOS
    SQLX$ = "SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + "WHERE CodiMovi = 'CF'"
    Set MOVICON = Stocks.OpenRecordset(SQLX$, 4)
    With MOVICON
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
          Do While .EOF = False
             NOFAX$ = !DoPriCor
             CIXI% = !Cod_Inte
             CACONSU = !Cantsalid
             For KKI% = 1 To CANOFPE%
                If OFCUR$(KKI%) = NOFAX$ Then GoTo 35
             Next KKI%
             GoTo 70
             '
35           With Reserva
                .FindFirst ("IdSubOr = '" & NOFAX$ & "' AND Cod_Inte = " & CIXI%)
37              If .NoMatch = False Then
                   CAPACON = !CantRes - !CantCons
                   If CAPACON > CACONSU Then CAPACON = CACONSU
                   .Edit
                   !CantCons = !CantCons + CAPACON
                   CACONSU = CACONSU - CAPACON
                   !Saldo_Entre = !CantRes - !CantCons
                   .Update
                   If CACONSU >= 0.05 Then
                      .FindNext ("IdSubOr = '" & NOFAX$ & "' AND Cod_Inte = " & CIXI%)
                      GoTo 37
                   End If
                End If
             End With
70        .MoveNext
          Loop
       End If
    End With
    '
98  Screen.MousePointer = 1
99  Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call HELPONLINE("RESERVA")
    Command6.Enabled = True
End Sub

 Sub Command7_Click()
   Command7.Enabled = False
   Call CHECKRES
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
  Command8.Enabled = False
  Call ACCDIR("RESERVA")
  Command8.Enabled = True
End Sub

Private Sub Command9_Click()
  '\\\OT-06-0086-WAR-10/03/06///
  Dim OFCUR$()
  Dim STACT#(32767)
  Dim STDISP#(32767)
  '
  Command9.Enabled = False
  Call ABREORFAB
'1.- CARGAR VECTOR DE OF EN CURSO
  SQLX$ = "SELECT * FROM ORDEFABR "
  SQLX$ = SQLX$ + "WHERE STATUORF <8 "
  Set OFTMP = Ofabric.OpenRecordset(SQLX$, 4)
  With OFTMP
    Screen.MousePointer = 11
    On Error Resume Next
    .MoveLast
    If Err Then
      Call MENSERR(24, "No Hay O/F's Pendientes.")
      Screen.MousePointer = 1
      Command9.Enabled = True
      Exit Sub
    End If
    On Error GoTo 0
    J% = .RecordCount
    ReDim OFCUR$(J%)
    CONTI% = 0
    .MoveFirst
    Do While .EOF = False
      OFCUR$(CONTI%) = !IdSubOr
      CONTI% = CONTI% + 1
      .MoveNext
    Loop
    Screen.MousePointer = 1
  End With
  '
'2.- PARA CADA OF LEER LAS RESERVAS Y CARGAR ARCHIVO HASTA COLU. QT. RESERVADA
  Y$ = REGBLAN$("RESOFALT")
  Call ABRESERVA
  JJ& = 0
  For JX% = 0 To J%
    SQLX$ = "SELECT * FROM RESERVA "
    SQLX$ = SQLX$ + "WHERE IdSubOr = '" & OFCUR$(JX%) & "' "
    SQLX$ = SQLX$ + "AND CantRes-CantCons > 0"
    Set RESTMP = Stocks.OpenRecordset(SQLX$, 4)
    With RESTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        GoTo 22
      End If
      On Error GoTo 0
      Do While .EOF = False
        CARE1# = !CantRes
        CI% = !Cod_Inte
        Call REPLA(Y$, !IdSubOr, 1, 12)
        Call REPLA(Y$, MKI$(CI%), 13, 2)
        Call REPLA(Y$, MKD$(CARE1#), 15, 8)
        JJ& = JJ& + 1
        Call GRAREG("RESOFALT", Y$, JJ&)
        .MoveNext
      Loop
    End With
22 Next JX%
  Call SETULTREG("RESOFALT", JJ&)
  '
'3.- SORT X Nº OF y X CODIGO
  Call FILESORT("RESOFALT", "RESOFALT", 2, 2, "ASC")
  Call FILESORT("RESOFALT", "RESOFALT", 1, 1, "ASC")
  '
'4.- STOCK ACTUAL DE CADA ITEM
  
  For JX% = 1 To NUMAS%
    STACT#(JX%) = STOCKACT(JX%)
    STDISP#(JX%) = STACT#(JX%)
  Next JX%
  Screen.MousePointer = 1
  '
'5.- BARRER EL ARCHIVO
  FIN& = ULTREG&("RESOFALT")
  For JN& = 1 To FIN&
    Call TRACE(24, JN&, FIN&)
    X$ = REGLEIDO$("RESOFALT", JN&)
    CR1# = CVD(Mid$(X$, 15, 8)) 'Q.RESERVADA
    CI% = CVI(Mid$(X$, 13, 2))
'    STD1# = STACT#(CI%) - CR1#
'    If STACT#(CI%) = 0 Then STD1# = 0
    FAL1# = 0
    SOB1# = 0
    If STDISP#(CI%) - CR1# > 0 Then
        SOB1# = STDISP#(CI%) - CR1# 'SOBRANTE
      Else
        FAL1# = Abs(STDISP#(CI%) - CR1#) 'FALTANTE
    End If
    Call REPLA(X$, MKD$(STACT#(CI%)), 23, 8)
    Call REPLA(X$, MKD$(STDISP#(CI%)), 31, 8)
    'AGREGAR FALTANTE Y SOBRANTE AL ARCHIVO
    Call REPLA(X$, MKD$(SOB1#), 39, 8)
    Call REPLA(X$, MKD$(FAL1#), 47, 8)
    Call GRAREG("RESOFALT", X$, JN&)
    STDISP#(CI%) = STDISP#(CI%) - CR1#
    If STDISP#(CI%) < 0 Then STDISP#(CI%) = 0
  Next JN&
  '
  If Option4 Then
    Call FINAL("*RESOFALT")
  ElseIf Option5 Then
    Call CONECRUN("*RESOFAP", "")
  End If
  '
  Command9.Enabled = True
  Screen.MousePointer = 1
  '\\\OT-06-0086-WAR-FIN///
End Sub

Private Sub Form_Load()
   '
   Call TRANSLABELS(Name)
   '
   RESERV04.Option1.Value = True
End Sub
'
Sub TOMATRES()
    '
    Dim CARESE#(32767)
    FIN& = ULTREG&("RESERVA")
    If FIN& < 1 Then
      Call MENSERR(24, "No hay Reservas Registradas")
      Exit Sub
    End If
    '
    Screen.MousePointer = 11
    '
    For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      XX$ = REGLEIDO$("RESERVA", i&)
      CI% = CVI(Left$(XX$, 2))
      If CI% > 0 Then
        If CI% <= NUMAS% Then
          TORE# = CVD(Mid$(XX$, 19, 8))
          CACO# = CVD(Mid$(XX$, 27, 8))
          SIGNO% = Sgn(CVS(Mid$(XX$, 75, 4)))
          SARE# = TORE# - CACO#
          If SARE# * SIGNO% < 0 Then
              SARE# = 0
          End If
          Call REPLA(XX$, MKD$(SARE#), 35, 8)
          Call GRAREG("RESERVA", XX$, i&)
          CARESE#(CI%) = CARESE#(CI%) + SARE#
        End If
      End If
    Next i&
    '
    FIN& = ULTREG&("REPUREPA")
    For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      XX$ = REGLEIDO$("REPUREPA", i&)
      CI% = CVI(Mid$(XX$, 7, 2))
      If CI% > 0 Then
        If CI% <= NUMAS% Then
          TORE# = CVS(Mid$(XX$, 21, 4))
          CACO# = CVS(Mid$(XX$, 25, 4))
          '
          If UCase$(UNIMED$(CIIX%)) = "GK" Then
            TORE# = TORE# / 1000
            CACO# = CACO# / 1000
          End If
          '
          SARE# = TORE# - CACO#
          If SARE# < 0.05 Then GoTo 29
          '
          NOREPARA& = CVS(Left$(XX$, 4))
          REREP& = REGBUS&("PRESREPA", 1, MKS$(NOREPARA&))
          If REREP& > 0 Then
            YY$ = REGLEIDO$("PRESREPA", REREP&)
            If CVI(Mid$(YY$, 123, 2)) > 0 Then GoTo 29 ' O/Rep Cerrada
            If CVS(Mid$(YY$, 125, 4)) < 1 Then GoTo 29 ' falta O/Venta
          End If
          '
          CARESE#(CI%) = CARESE#(CI%) + SARE#
        End If
      End If
29  Next i&
    '
    JJ& = 0
    FIN& = NUMAS%: If FIN& < 1 Then FIN& = 1
    '
    For i& = 1 To NUMAS%
      II$ = REGLEIDO$("INVERT", i& + 1)
      CI% = i&
      If CI% > 0 Then
        If Abs(CARESE#(CI%)) > 0.05 Then
          STOCK# = STOCKACT(CI%)
          XX$ = REGBLAN$("TOMATRES")
          Call REPLA(XX$, MKI$(CI%), 1, 2)
          Call REPLA(XX$, MKD$(CARESE#(CI%)), 3, 8)
          Call REPLA(XX$, MKD$(STOCK#), 11, 8)
          FALTA# = CARESE#(CI%) - STOCK#
          If FALTA# < 0 Then FALTA# = 0
          Call REPLA(XX$, MKD$(FALTA#), 35, 8)
          OBSE$ = "      "
          If FALTA# > 0 Then OBSE$ = "CRIT. "
          If STOCK# - STOMIN(CI%) < 0 Then OBSE$ = "CRIT. "
          If STOCK# < 0 Then OBSE$ = "NEG.  "
          Call REPLA(XX$, OBSE$, 43, 6)
          JJ& = JJ& + 1
          Call GRAREG("TOMATRES", XX$, JJ&)
        End If
      End If
    Next i&
    '
    Call SETULTREG("TOMATRES", JJ&)
    Call CIERRARCH("TOMATRES")
    Call FILESORT("TOMATRES", "", 1, 1, "ASC")
    '
    Screen.MousePointer = 1
    '
    If Option4.Value = True Then
         Call FINAL("*LITOTRE")       ' impresion
       Else
         Call FINAL("*COTOTRE")       ' pantalla
    End If
    '
End Sub
'

Sub GENRESDAT()
    '
    Call ABRESTOCKS
    '
    SQLX$ = " SELECT * FROM RESERVA "
    SQLX$ = SQLX$ + " WHERE CantRes-CantCons > 0 "
    SQLX$ = SQLX$ + " ORDER BY IdSubOr "
    '
    Set RESERTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    With RESERTMP
      On Error Resume Next
      .MoveFirst
      If Err Then Call MENSERR(24, " No Existen Reservas efectuadas. ")
      On Error GoTo 0
      JJ& = 0
      Do While (.EOF = False)
        Y$ = REGBLAN("RESERVA")
        '
        Call REPLA(Y$, MKI$(!Cod_Inte), 1, 2)
        Call REPLA(Y$, !IdSubOr, 3, 12)
        Call REPLA(Y$, MKI$(!Cod_Inte), 15, 2)
        Call REPLA(Y$, MKI$(CVINT(!FechRes)), 17, 2)
        Call REPLA(Y$, MKD$(!CantRes), 19, 8)
        Call REPLA(Y$, MKD$(!CantCons), 27, 8)
        Call REPLA(Y$, MKD$(!Saldo_Entre), 35, 8)
        Call REPLA(Y$, !IdenLote, 43, 16)
        Call REPLA(Y$, Chr$(!Deposito), 60, 1)
        Call REPLA(Y$, MKS$(!UsoUnit), 75, 4)
        '
        JJ& = JJ& + 1
        Call GRAREG("RESERVA", Y$, JJ&)
        '
      .MoveNext
      Loop
    End With
    Call SETULTREG("RESERVA", JJ&)
    Call CIERRARCH("RESERVA")
    '
End Sub

Private Sub salir_Click()
   Call Command1_Click
End Sub
