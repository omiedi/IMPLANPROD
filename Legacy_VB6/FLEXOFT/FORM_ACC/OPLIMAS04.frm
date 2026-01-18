VERSION 5.00
Begin VB.Form OPLIMAS04 
   BackColor       =   &H00C6E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maestro de Items de Stock - Consultas, Listados"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command10 
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
      Left            =   5070
      Picture         =   "OPLIMAS04.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Actualización de Costos"
      Top             =   1980
      Width           =   640
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00004040&
      Height          =   7575
      Left            =   4935
      ScaleHeight     =   7515
      ScaleWidth      =   1005
      TabIndex        =   24
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
         Height          =   600
         Left            =   105
         Picture         =   "OPLIMAS04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Listado de Especificaciones"
         Top             =   3420
         Width           =   640
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Depu-ra"
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
         Left            =   105
         TabIndex        =   34
         Top             =   5370
         Visible         =   0   'False
         Width           =   640
      End
      Begin VB.CommandButton Command12 
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
         Left            =   105
         Picture         =   "OPLIMAS04.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Actualización de Costos por Grupo Contable"
         Top             =   2745
         Width           =   640
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Crystal Reports"
         Height          =   640
         Left            =   105
         TabIndex        =   31
         Top             =   4695
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Command6"
         Height          =   435
         Left            =   105
         TabIndex        =   30
         ToolTipText     =   "Análisis de Costos de Reposición"
         Top             =   4215
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   27
         Top             =   6090
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   28
            Top             =   105
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
         Height          =   650
         Left            =   105
         Picture         =   "OPLIMAS04.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   280
         Width           =   640
      End
      Begin VB.CommandButton Command11 
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
         Left            =   105
         Picture         =   "OPLIMAS04.frx":08A8
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   930
         Width           =   640
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "OPLIMAS04.frx":0BB2
         Top             =   6405
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C6E0E0&
      Caption         =   "Items Inactivos"
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
      Left            =   150
      TabIndex        =   8
      Top             =   5040
      Width           =   4635
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1155
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   9
         Top             =   945
         Width           =   2640
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
            Height          =   600
            Left            =   0
            Picture         =   "OPLIMAS04.frx":2AD4
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Listado de Items Inactivos (Dados de Baja)"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado de Items Inactivos (de Baja)"
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
            Top             =   105
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   12
         Top             =   420
         Width           =   2640
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
            Picture         =   "OPLIMAS04.frx":2DDE
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Listado de Items Inactivos (Dados de Baja)"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Consulta y Restau- ración Items de Baja"
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
            Top             =   45
            Width           =   1905
         End
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C6E0E0&
      Caption         =   "Consultas y Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4740
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   4635
      Begin VB.Frame Frame2 
         BackColor       =   &H00C6E0E0&
         Caption         =   "Costos de Reposición"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1800
         Left            =   210
         TabIndex        =   14
         Top             =   2730
         Width           =   4215
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1050
            ScaleHeight     =   585
            ScaleWidth      =   2580
            TabIndex        =   15
            Top             =   945
            Width           =   2640
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
               Picture         =   "OPLIMAS04.frx":2F28
               Style           =   1  'Graphical
               TabIndex        =   16
               ToolTipText     =   "Listado de Costos de Reposición por Articulo"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Listado por   Código de Artículo"
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
               Left            =   630
               TabIndex        =   17
               Top             =   105
               Width           =   1695
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2580
            TabIndex        =   18
            Top             =   420
            Width           =   2640
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
               Picture         =   "OPLIMAS04.frx":3232
               Style           =   1  'Graphical
               TabIndex        =   19
               ToolTipText     =   "Listado de Costos de Reposición por Proveedor"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Listado por Número de Proveedor"
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
               Top             =   50
               Width           =   1695
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C6E0E0&
         Caption         =   "Listados Generales del Maestro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2325
         Left            =   210
         TabIndex        =   1
         Top             =   315
         Width           =   4215
         Begin VB.PictureBox Picture11 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1050
            ScaleHeight     =   585
            ScaleWidth      =   2580
            TabIndex        =   5
            Top             =   1470
            Width           =   2640
            Begin VB.CommandButton Command15 
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
               Picture         =   "OPLIMAS04.frx":353C
               Style           =   1  'Graphical
               TabIndex        =   6
               ToolTipText     =   "Listado de Captura de Datos para Inventarios"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Listado de Captura para Inventarios"
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
               Left            =   630
               TabIndex        =   7
               Top             =   105
               Width           =   1905
            End
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2580
            TabIndex        =   21
            Top             =   945
            Width           =   2640
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
               Picture         =   "OPLIMAS04.frx":3846
               Style           =   1  'Graphical
               TabIndex        =   22
               ToolTipText     =   "Listado de Atributos Generales de Articulos de Stock"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Stock Mínimo y Lote de Reposición"
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
               TabIndex        =   23
               Top             =   50
               Width           =   1695
            End
         End
         Begin VB.PictureBox Picture12 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2580
            TabIndex        =   2
            Top             =   420
            Width           =   2640
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
               Picture         =   "OPLIMAS04.frx":3C88
               Style           =   1  'Graphical
               TabIndex        =   3
               ToolTipText     =   "Listado de Atributos Generales de Articulos de Stock"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Listado de Atributos Generales"
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
               TabIndex        =   4
               Top             =   50
               Width           =   1695
            End
         End
      End
   End
End
Attribute VB_Name = "OPLIMAS04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub ESPECIFI()

    Screen.MousePointer = 11
    '
    Call COPYSTRU("ESPECIFI", "ESPECIF")
    '
    XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
    XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0) + MKI$(0)
    '
10  RACOD$ = OBJPLA$("DESHASCOD3", VDEF$)
    If RACOD$ = "" Then
        Exit Sub
    End If
    CI1% = CVI(Left$(RACOD$, 2)): COD1$ = CODEXT$(CI1%)
    CI2% = CVI(Mid$(RACOD$, 3, 2)): COD2$ = CODEXT$(CI2%)
    TMT% = Asc(Mid$(RACOD$, 5, 1))
    TIMAT$ = TRIM$(ECOARCH$("TATIMAT", Chr$(TMT%)))
    If TIMAT$ = "" Then
        TMT% = 0
        Call REPLA(RACOD$, Chr$(0), 5, 1)
        TIMAT$ = "Sin Clasificación"
    End If
    GRC% = CVI(Mid$(RACOD$, 6, 2))
    GRUCO$ = TRIM$(ECOARCH$("GRUCOVEN", MKI$(GRC%)))
    If GRUCO$ = "" Then
        GRC% = 0
        Call REPLA(RACOD$, MKI$(0), 6, 2)
        GRUCO$ = "Sin Clasificación"
    End If
    '
    If COD1$ = "" Or COD2$ = "" Or COD2$ < COD1$ Then
        VDEF$ = RACOD$
        GoTo 10
    End If
    '
    JJ& = 0
    '
    SQLX$ = "SELECT * FROM Master "
    SQLX$ = SQLX$ + "WHERE ESPECIFICACION <> '' "
    SQLX$ = SQLX$ + "AND CODIGO>='" & COD1$ & "' "
    SQLX$ = SQLX$ + "AND CODIGO<='" & COD2$ & "' "
    SQLX$ = SQLX$ + "AND (" & TMT% & " =0 OR TIMATE = " & TMT% & ") "
    SQLX$ = SQLX$ + "AND (" & GRC% & " =0 OR GRUCOVEN = " & GRC% & ") "
    SQLX$ = SQLX$ + "ORDER BY CODIGO ASC;"
    '
    Call APERBASDAT("MASTER")
    Set MasterTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With MasterTemp
        On Error Resume Next
        .MoveLast
        'Valida que la tabla no este vacia
        If Err Then
         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
         GoTo 99
        End If
        FIN& = .RecordCount
        '
        On Error GoTo 0
        .MoveFirst
        JJ& = 0
        REBLA$ = REGBLAN$("ESPECIF")
        Call BLOQARCH("ESPECIF")
        Do While (.EOF = False)
           Call TRACE(20, JJ&, FIN&)
           On Error Resume Next
           CIXI% = XVALO(!CODINT)
           ESPE$ = SAFETEXT$(!ESPECIFICACION)
           On Error GoTo 0
           '
           Call FRATEXTO(ESPE$, 62)
           For KKIIX% = 1 To CARETEX%
             ZZ$ = REBLA$
             Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
             Call REPLA(ZZ$, RETEX$(KKIIX%), 3, 62)
             JJ& = JJ& + 1
             Call GRAREG("ESPECIF", ZZ$, JJ&)
           Next KKIIX%
           '
        .MoveNext
        Loop
        '
    End With
    Call SETULTREG("ESPECIF", JJ&)
    Call CIERRARCH("ESPECIF")
    '
    Call HEADERS("ESPECIF", "")
    Call HEADERS("ESPECIF", "Desde Codigo: " + TRIM$(COD1$) + " - " + DESCRIT$(CI1%))
    Call HEADERS("ESPECIF", "Hasta Codigo: " + TRIM$(COD2$) + " - " + DESCRIT$(CI2%))
    '
    Call CONECRUN("*ESPECIF", "Listado Costos de Reposición por Proveedor")
    '
99  Screen.MousePointer = 1
End Sub

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

 Sub Command10_Click()
    '//OT-07-0464//
    Command10.Enabled = False
    VDEF$ = ""
10  OBX$ = OBJPLA$("ACTUCOSTO", VDEF$)
    If OBX$ = "" Then GoTo 991
    '
    VDEF$ = OBX$ 'PARA QUE DEVUELVA LOS DATOS BIEN INGRESADOS SI ENTRO EN ALGUNA VALIDACION
    '
    'VALIDACION DE FECHAS
    DESDE% = CVI(Mid$(OBX$, 3, 2))
    HASTA% = CVI(Mid$(OBX$, 5, 2))
    If DESDE% = 0 Or HASTA% = 0 Or DESDE% > HASTA% Then
      Call MENSERR(24, "Rango de Fechas no Procesable")
      GoTo 10
    End If
    FEC1 = CDbl(CVDAT(DESDE%))
    FEC2 = CDbl(CVDAT(HASTA%))
    '
    barretodo% = 0
    CI% = CVI(Left$(OBX$, 2))
    COSTO# = CVD(Mid$(OBX$, 7, 8))
    If CI% = 0 And COSTO# = (-999) Then
       If COMALTER%("¿ Desea Aplicar los Costos ACTUALES\a TODOS los Artículos\Dentro del Rango de Fechas Elegido ?\\No, Cancelar\Si, Aplicar") = 2 Then
            barretodo% = 1
            GoTo 25
          Else
            VDEF$ = ""
            GoTo 10
       End If
    End If
    '
    'VALIDACION DE CODIGO
    If CI% < 1 Or CI% > NUMAS% Then
      Call MENSERR(24, "Código de Artículo no Válido")
      GoTo 10
    End If
    '
    'VALIDACION DEL COSTO
    If COSTO# < 0.00005 Then
      Call MENSERR(24, "Debe Ingresar el Costo")
      GoTo 10
    End If
    '
    Screen.MousePointer = 11
    '
    'GRABA EL COSTO INGRESADO SEGUN EL
    'RANGO DE FECHA SELECCIONADO Y EL CODIGO ELEGIDO
    '
    Call APERBASDAT("Stocks")
    SQLX$ = "UPDATE MOVISTO "
    SQLX$ = SQLX$ + "SET CostoRep = '" & COSTO# & "' "
    SQLX$ = SQLX$ + " WHERE INT(CDBL(FECHMOV)) >= " & FEC1 & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(FECHMOV)) <= " & FEC2 & " "
    SQLX$ = SQLX$ + " AND Cod_Inte = " & CI% & " ;"
    '
    Stocks.Execute (SQLX$)
    Screen.MousePointer = 1
    GoTo 99
    'FINALIZADA LA ACTUALIZACION
    '
25  Call APERBASDAT("STOCKS")
    SQLX$ = "SELECT * FROM Movisto "
    SQLX$ = SQLX$ + " WHERE INT(CDBL(FECHMOV)) >= " & FEC1 & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(FECHMOV)) <= " & FEC2 & " "
    Set MOVISTMP = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    '
    With MOVISTMP
       On Error Resume Next
       .MoveLast
       If Err Then
          Screen.MousePointer = 1
          Call MENSERR(24, "No Existen Movimientos de Stock\en Rango de Fechas Seleccionado.")
          On Error GoTo 0
          GoTo 991
       End If
       On Error GoTo 0
       FIN& = .RecordCount
       JJ& = 0
       .MoveFirst
       Do While (.EOF = False)
          CIXI% = !Cod_Inte
          COU = COSPESOS(CIXI%)
          .Edit
          !CostoRep = COU
          .Update
          JJ& = JJ& + 1
          Call TRACE(20, JJ&, FIN&)
       .MoveNext
       Loop
    End With
    '
99  Screen.MousePointer = 1
    Call COMUNI("Actualización de Costo Realizada")
991 Command10.Enabled = True
    '//OT-07-0464//
End Sub

Sub Command12_Click()
   Command12.Enabled = False
   OPT% = 0
10 OPT% = ALTERNA%("Actualización de Costos por Grupo Contable\Otros Accesos Directos")
   '
   If OPT% = 1 Then
      Call ACTULISCOSGCONT
       GoTo 10
   ElseIf OPT% = 2 Then
      Call ACCDIR("ANACOS")
       GoTo 10
   End If
   '
   Command12.Enabled = True

End Sub

Private Sub Command13_Click()
    Command13.Enabled = False
    Call ABREMASTER
    
10  Call SELECHO("LISTAS")
    VALOR$ = TRIM$(VALACT1$("LISTAS"))
    If VALOR$ = "" Then GoTo 99
    
    LST% = Asc(Left$(VALOR$, 1))
    DENOL$ = TRIM$(VALACT2$("LISTAS"))
    
    OPX% = COMALTER%("¿ Esta Seguro que desea Anular los Items\de la Lista de Precios '" & DENOL$ & "' ?\\No, Cancelar\Si, Anular")
    If OPX% < 2 Or OPX% > 2 Then GoTo 99

    Do While Len(VALOR$) < 3: VALOR$ = "0" & VALOR$: Loop
    NOMLISTA$ = "LIPRE" & VALOR$
    
    FIN& = ULTREG&(NOMLISTA$)
    For UU& = 1 To ULTREG&(NOMLISTA$)
        Call TRACE(24, UU&, FIN&)
        XX$ = REGLEIDO$(NOMLISTA$, UU&)
        CIXI% = CVI(Mid$(XX$, 1, 2))
        '
        CONDI$ = "CODINT=" & Str$(CIXI%)
        Master.FindFirst CONDI$
        If Master.NoMatch = False Then
            Master.Edit
            Master!Status = -1
            Master.Update
        End If
    Next UU&
    '
    UPDATMASTER% = 1
    Call COMUNI("Proceso Terminado")
    GoTo 10
    '
99  Call GENECOMAS
    '
    Command13.Enabled = True
End Sub

Sub Command14_Click()
   Command14.Enabled = False
   Call ESPECIFI
   Command14.Enabled = True
End Sub


Private Sub Command8_Click()
    '\\\OT-07-0002-RG-03/01/07///
    Command8.Enabled = False
    On Error Resume Next
    CrystalReport1.DataFiles(0) = "X:\BICICLETERIA\FLEXOFT\DATOS\BdAccess\articulo1.mdb"
    CrystalReport1.ReportFileName = "X:\BICICLETERIA\FLEXOFT\DATOS\prueba.rpt"
    CrystalReport1.Action = 1
    On Error GoTo 0
    Command8.Enabled = True
    '\\\OT-07-0002-RG-FIN///
End Sub

 Sub Command9_Click()
    Command9.Enabled = False
    Call LIMASTER(1)
    Command9.Enabled = True
End Sub
'
 Sub Command7_Click()
    Command7.Enabled = False
    Call LIMASTER(3)
    Command7.Enabled = True
End Sub

 Sub COMMAND15_CLICK()
    Command15.Enabled = False
    Call LIMASTER(2)
    Command15.Enabled = True
End Sub
'
 Sub Command2_Click()
   Command2.Enabled = False
   Call COMUNI("Solo Disponible como\Consulta por Pantalla")
   Call Command3_Click
   Command2.Enabled = True
End Sub

 Sub Command3_Click()
   Command3.Enabled = False
   Call APERBASDAT("MASTER")
   '
   Screen.MousePointer = 11
   LITBAJ04.List1.Clear
   '
   Set MasterBaja = Articulos.OpenRecordset("SELECT CODIGO,DESCRIPCION,CODINT FROM Master WHERE STATUS < 0 ORDER BY CODIGO ASC;", dbOpenSnapshot)
   On Error Resume Next
   MasterBaja.MoveLast
   If Err Then
     On Error GoTo 0
     Call MENSERR(24, "No hay Items de Baja\que Puedan Consultarse o Restaurarse.")
     GoTo 99
   End If
   On Error GoTo 0
   '
   CAREC% = MasterBaja.RecordCount
   If CAREC% < 1 Then GoTo 99
   '
   FIN& = CAREC%
   REBLA$ = REGBLAN$("ECOMAST")
   JJ& = 0
   MasterBaja.MoveFirst
   '
10 Call TRACE(20, JJ&, FIN&)
   TTXX$ = Space$(80)
   Call REPLA(TTXX$, Left$(MasterBaja!CODIGO, 15), 1, 16)
   Call REPLA(TTXX$, MasterBaja!DESCRIPCION, 19, 46)
   Call REPLA(TTXX$, Str$(MasterBaja!CODINT), 73, 8)
   LITBAJ04.List1.AddItem TTXX$
   JJ& = JJ& + 1
   '
   If JJ& < CAREC% Then
     MasterBaja.MoveNext
     GoTo 10
   End If
   '
   On Error Resume Next
   MasterBaja.Close
   On Error GoTo 0
   Screen.MousePointer = 1
   LITBAJ04.Show 1
   Call GENECOMAS
   '
99 Screen.MousePointer = 1
   Command3.Enabled = True
End Sub

Sub Command4_Click()
   Command4.Enabled = False
   Call GELICOREP(2)
   Command4.Enabled = True
End Sub

Sub Command5_Click()
   Command5.Enabled = False
   Call GELICOREP(1)
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   '
   Command6.Enabled = False
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DOSIVA") < 1 Then
       Call OPNOIN("Análisis Costos de Reposición")
     Else
       CONTROCOS.Show 1
   End If
   Command6.Enabled = True
   '
End Sub

Sub LIMASTER(modo%)
    '
    HOO% = HOY(HO$)
    If modo% = 0 Then
      modo% = ALTERNA%("Listado de Códigos y Atributos\Código y Descripción para Inventarios\Stock Mínimo y Lote Reposición")
      If modo% < 1 Or modo% > 3 Then
        Exit Sub
      End If
    End If
    '
    XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
    XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0) + MKI$(0)
    '
10  RACOD$ = OBJPLA$("DESHASCOD3", VDEF$)
    If RACOD$ = "" Then
        Exit Sub
    End If
    CI1% = CVI(Left$(RACOD$, 2)): COD1$ = CODEXT$(CI1%)
    CI2% = CVI(Mid$(RACOD$, 3, 2)): COD2$ = CODEXT$(CI2%)
    TMT% = Asc(Mid$(RACOD$, 5, 1))
    TIMAT$ = TRIM$(ECOARCH$("TATIMAT", Chr$(TMT%)))
    If TIMAT$ = "" Then
        TMT% = 0
        Call REPLA(RACOD$, Chr$(0), 5, 1)
        TIMAT$ = "Sin Clasificación"
    End If
    GRC% = CVI(Mid$(RACOD$, 6, 2))
    GRUCO$ = TRIM$(ECOARCH$("GRUCOVEN", MKI$(GRC%)))
    If GRUCO$ = "" Then
        GRC% = 0
        Call REPLA(RACOD$, MKI$(0), 6, 2)
        GRUCO$ = "Sin Clasificación"
    End If
    '
    If COD1$ = "" Or COD2$ = "" Or COD2$ < COD1$ Then
        VDEF$ = RACOD$
        GoTo 10
    End If
    '
    JJ& = 0
    If modo% = 1 Then
        FILEKILL LUDAT$ + "LIMASEL.DAT"
        Call HEADERS("LIMASEL", "")
        Call HEADERS("LIMASEL", "Tipo / Grupo de Material: " + TIMAT$ + " / " + GRUCO$)
        Call HEADERS("LIMASEL", "Desde Codigo: " + TRIM$(COD1$) + " - " + DESCRIT$(CI1%))
        Call HEADERS("LIMASEL", "Hasta Codigo: " + TRIM$(COD2$) + " - " + DESCRIT$(CI2%))
      ElseIf modo% = 2 Then
        FILEKILL LUDAT$ + "LIMAINV.DAT"
        Call HEADERS("LIMAINV", "")
        Call HEADERS("LIMAINV", "Tipo / Grupo de Material: " + TIMAT$ + " / " + GRUCO$)
        Call HEADERS("LIMAINV", "Desde Codigo: " + TRIM$(COD1$) + " - " + DESCRIT$(CI1%))
        Call HEADERS("LIMAINV", "Hasta Codigo: " + TRIM$(COD2$) + " - " + DESCRIT$(CI2%))
        Call HEADERS("LIMAINV", "Stock Nominal al: " + HO$ + " - " + Time$ + " hs.")
      ElseIf modo% = 3 Then
        FILEKILL LUDAT$ + "LIMALOT.DAT"
        Call HEADERS("LIMALOT", "")
        Call HEADERS("LIMALOT", "Material: " + TRIM$(TIMAT$) + " (" + TRIM$(COD1$) + " - " + TRIM$(COD2$) + ")")
      Else
        Screen.MousePointer = 1
        Exit Sub
    End If
    '
    Screen.MousePointer = 11
    'Consulta sql de filtro
    SQLX$ = "SELECT * FROM Master "
    SQLX$ = SQLX$ + "WHERE CODIGO>='" & COD1$ & "' "
    SQLX$ = SQLX$ + "AND CODIGO<='" & COD2$ & "' "
    SQLX$ = SQLX$ + "AND STATUS>=1 "
    SQLX$ = SQLX$ + "AND (" & TMT% & " =0 OR TIMATE = " & TMT% & ") "
    SQLX$ = SQLX$ + "AND (" & GRC% & " =0 OR GRUCOVEN = " & GRC% & ") "
    SQLX$ = SQLX$ + "ORDER BY CODIGO ASC;"
    ' Seteo el recordset
    Screen.MousePointer = 11
    Set MasterTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
    
    J& = 0: FIN& = NUINV%
    With MasterTemp
       On Error Resume Next
       .MoveLast
       'Valida que la tabla no este vacia
       If Err Then
         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
         GoTo 99
       End If
       
       registros& = .RecordCount
       FIN& = registros&
       KU& = 0
       '
       .MoveFirst
       Do While (.EOF = False)
          KU& = KU& + 1
          Call TRACE(20, KU&, FIN&)
          CI% = !CODINT
          If modo% = 1 Then
                'Genera el listado de Codigos y Atributos
                ML$ = REGBLAN$("LIMASEL")
                Call REPLA(ML$, MKI$(CI%), 1, 2)
                Call REPLA(ML$, Chr$(TIMATE%(CI%)), 3, 1)
                Call REPLA(ML$, Chr$(CRITREP%(CI%)), 4, 1)
                Call REPLA(ML$, Chr$(MONECOSTO%(CI%)), 5, 1)
                Call REPLA(ML$, MKI$(STOMIN(CI%)), 15, 2)
                Call REPLA(ML$, MKI$(LOTESTAN(CI%)), 17, 2)
                Call REPLA(ML$, UNIMED$(CI%), 19, 2)
                Call REPLA(ML$, MKS$(COSUNI(CI%)), 21, 4)
                Call REPLA(ML$, MKI$(FECOSTO%(CI%)), 25, 2)
                Call REPLA(ML$, MKI$(PROVHABI(CI%)), 33, 2)
                Call REPLA(ML$, MKI$(GRUCOFAB(CI%)), 35, 2)
                JJ& = JJ& + 1: Call GRAREG("LIMASEL", ML$, JJ&)
             ElseIf modo% = 2 Then
                'Genera Listado de inventario
                ML$ = MKI$(0) + "  " + MKS$(0) + Space$(56)
                'Call REGAPP("LIMAINV", ML$)
                ML$ = REGBLAN$("LIMAINV")
                Call REPLA(ML$, MKI$(CI%), 1, 2)
                Call REPLA(ML$, UNIMED$(CI%), 3, 2)
                Call REPLA(ML$, MKS$(STOCKACT(CI%)), 5, 4)
                JJ& = JJ& + 1: Call GRAREG("LIMAINV", ML$, JJ&)
             ElseIf modo% = 3 Then
                'Genera Listado de maestro por lote
                ML$ = REGBLAN$("LIMALOT")
                Call REPLA(ML$, MKI$(CI%), 1, 2)
                Call REPLA(ML$, MKS$(STOMIN(CI%)), 3, 4)
                Call REPLA(ML$, MKS$(LOTESTAN(CI%)), 7, 4)
                Call REPLA(ML$, MKS$(LOREPOMA(CI%)), 11, 4)
                Call REPLA(ML$, MKS$(COSUNI(CI%)), 15, 4)
                Call REPLA(ML$, MKI$(FECOSTO%(CI%)), 19, 2)
                JJ& = JJ& + 1: Call GRAREG("LIMALOT", ML$, JJ&)
          End If
          '
       .MoveNext
145    Loop
    End With
    '
149 FIECO% = 0
    If modo% = 1 Then
        Call SETULTREG("LIMASEL", JJ&)
        Call CIERRARCH("LIMASEL")
        Call FINAL("*LIMASEL")
      ElseIf modo% = 2 Then
        Call SETULTREG("LIMAINV", JJ&)
        Call CIERRARCH("LIMAINV")
        Call FINAL("*LIMAINV")
      ElseIf modo% = 3 Then
        Call SETULTREG("LIMALOT", JJ&)
        Call CIERRARCH("LIMALOT")
        Call FINAL("*LIMALOT")
    End If
    '
99  Screen.MousePointer = 1
End Sub
'
Sub GELICOREP(modo%)
   '
   ' MODO% = 1 --> Ordenado por Proveedor
   ' MODO% = 2 --> Ordenado por Articulo
   '
   Screen.MousePointer = 11
   'Consulta sql de filtro
   SQLX$ = "SELECT * FROM Master "
   SQLX$ = SQLX$ + "WHERE STATUS>=1 "
   If modo% = 1 Then
      SQLX$ = SQLX$ + "AND PROVHABI > 0 ORDER BY PROVHABI ASC;"
    Else
      SQLX$ = SQLX$ + "ORDER BY CODIGO ASC;"
   End If
   '
   ' Seteo el recordset
   Call APERBASDAT("MASTER")
   Set MasterTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot)
   
   With MasterTemp
       On Error Resume Next
       .MoveLast
       'Valida que la tabla no este vacia
       If Err Then
         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
         GoTo 99
       End If
       FIN& = .RecordCount
       '
       On Error GoTo 0
       .MoveFirst
       JJ& = 0
       REBLA$ = REGBLAN$("LICOREP")
       Call BLOQARCH("LICOREP")
       Do While (.EOF = False)
          Call TRACE(20, JJ&, FIN&)
          On Error Resume Next
          CIXI% = !CODINT
          PHHI% = !PROVHABI
          On Error GoTo 0
          '
          ZZ$ = REBLA$
          Call REPLA(ZZ$, MKI$(PHHI%), 1, 2)
          Call REPLA(ZZ$, MKI$(CIXI%), 3, 2)
          Call REPLA(ZZ$, UNIMED$(CIXI%), 5, 2)
          Call REPLA(ZZ$, MKI$(MONECOSTO%(CIXI%)), 7, 2)
          Call REPLA(ZZ$, MKS$(COSUNI(CIXI%)), 9, 4)
          Call REPLA(ZZ$, MKI$(FECOSTO%(CIXI%)), 13, 2)
          Call REPLA(ZZ$, MKI$(PHHI%), 15, 2)
          JJ& = JJ& + 1
          Call GRAREG("LICOREP", ZZ$, JJ&)
          '
       .MoveNext
       Loop
       '
   End With
   Call SETULTREG("LICOREP", JJ&)
   Call CIERRARCH("LICOREP")
   '
   Call CONECRUN("*LICOREP", "Listado Costos de Reposición por Proveedor")
   '
99 Screen.MousePointer = 1
   End Sub

Sub ACTULISCOSGCONT()

10  Call SELECHO("GRUCOVEN")
    GRCLI% = XVALO(VALACT1$("GRUCOVEN"))
    If GRCLI% > 0 Then
       '
       Call APERBASDAT("MASTER")
       '
       Screen.MousePointer = 11
       Set MasterBaja = Articulos.OpenRecordset("SELECT * FROM Master WHERE Grucoven = " & GRCLI% & " AND Status > 0 ORDER BY CODIGO ASC;", dbOpenSnapshot)
       With MasterBaja
           On Error Resume Next
           MasterBaja.MoveFirst
           If Err Then
              On Error GoTo 0
              Call MENSERR(24, "No hay Registros que Coincidan con el Proveedor Elegido .")
              GoTo 20
           End If
           On Error GoTo 0
           '
20         Call BLANARCH("!ACTUCOS")
           '
           I& = 0
           FIN& = MasterBaja.RecordCount
    
           Do While Not .EOF
              I& = I& + 1
              Call TRACE(20, I&, FIN&)
              CODI% = XVALO(!CODINT)
              If CODI% > 0 Then
                MON% = !MONECOS
                PODES1 = ROUND#(!PorDescCos, 1): If PODES1 > 99 Then PODES1 = 0
                COS1# = !COSUNI / (1 - PODES1 / 100)
                COS2# = !COSUNI
                FECH = !FECOSTO
                '
                Y$ = REGBLAN$("ACTUCOS")
                Call REPLA(Y$, MKI$(CODI%), 1, 2)
                Call REPLA(Y$, MKI$(MON%), 3, 2)
                Call REPLA(Y$, MKD$(COS1#), 5, 8)
                Call REPLA(Y$, MKS$(PODES1), 13, 4)
                Call REPLA(Y$, MKD$(COS2#), 17, 8)
                FCOS% = CVINT%(FECH)
                Call REPLA(Y$, MKI$(FCOS%), 25, 2)
                Call GRAREG("!ACTUCOS", Y$, I&)
              End If
             .MoveNext
             '
           Loop
             '
           Call SETULTREG("!ACTUCOS", I&)
             '
       End With
         '
       Screen.MousePointer = 1
       JUN% = VENTABU%("ACTUCOS/TITUPAN=Costo de Reposición - " + ECOARCH$("GRUCOVEN", MKI$(GRCLI%)))
         '
       If JUN% >= 0 Then
          Screen.MousePointer = 11
          Call ABREMASTER
        
          FIN& = ULTREG&("!ACTUCOS")
          For J& = 1 To ULTREG&("!ACTUCOS")
            Call TRACE(20, J&, FIN&)
            Z$ = REGLEIDO$("!ACTUCOS", J&)
            COD3% = CVI(Mid$(Z$, 1, 2))
              PODES1 = CVS(Mid$(Z$, 13, 4)): If PODES1 < 0 Or PODES1 > 99 Then PODES1 = 0
            COU = (CVD(Mid$(Z$, 5, 8))) * (1 - PODES1 / 100)
            MONEI% = CVI(Mid$(Z$, 3, 2))
            FECOS% = CVI(Mid$(Z$, 25, 2))
            '
            Master.FindFirst ("CODINT =" & COD3%)
            If Master.NoMatch = False Then
               Master.Edit
               Master!PROVHABI = NCX1%
               Master!COSUNI = COU
               Master!PorDescCos = PODES1
               Master!FECOSTO = CVDAT(FECOS%)
               Master!MONECOS = MONEI%
               Master.Update
            End If
            '
          Next J&
         '
       End If
       Screen.MousePointer = 1
       GoTo 10

    End If

End Sub


Private Sub Form_Load()
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "FUINYTER") > 0 Then
        ' PROCESO DE RECUPERACION DE BOLETA DE RECEPCION BR-556
        Command13.Visible = True
   End If

   
  
   
End Sub
