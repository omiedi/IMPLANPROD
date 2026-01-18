VERSION 5.00
Begin VB.Form ESTASTK04 
   BackColor       =   &H00EEEEC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estadísticas de Stock por Tipo de Movimiento"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   8070
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
   ScaleHeight     =   5445
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture6 
      Height          =   5500
      Left            =   7200
      ScaleHeight     =   5445
      ScaleWidth      =   840
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   900
      Begin VB.Image Image3 
         Height          =   5600
         Left            =   0
         Picture         =   "ESTASTK04.frx":0000
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   900
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00EEEEC0&
      Caption         =   "Componentes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   4870
      TabIndex        =   38
      Top             =   1785
      Width           =   2150
      Begin VB.CheckBox Check1 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Incluir Compo- nentes Multinivel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   210
         TabIndex        =   39
         Top             =   315
         Width           =   1905
      End
   End
   Begin VB.CommandButton Command12 
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
      Height          =   690
      Left            =   6360
      Picture         =   "ESTASTK04.frx":2398
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Emite Reporte Solicitado"
      Top             =   4680
      Width           =   650
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00EEEEC0&
      Caption         =   "Tipo de Movimiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   180
      TabIndex        =   29
      Top             =   2700
      Width           =   3405
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Left            =   160
         Style           =   1  'Checkbox
         TabIndex        =   31
         Top             =   270
         Width           =   3075
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEC0&
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
      Height          =   1935
      Left            =   3690
      TabIndex        =   25
      Top             =   2700
      Width           =   3315
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
         Height          =   285
         Left            =   3060
         TabIndex        =   34
         Top             =   1080
         Width           =   175
      End
      Begin VB.TextBox Text6 
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
         Left            =   2640
         TabIndex        =   33
         Top             =   1080
         Width           =   435
      End
      Begin VB.TextBox Text3 
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
         Left            =   2610
         TabIndex        =   27
         Top             =   270
         Width           =   435
      End
      Begin VB.CommandButton Command3 
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
         Left            =   3030
         TabIndex        =   26
         Top             =   270
         Width           =   175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Todos - Sin Clasificación"
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
         Left            =   120
         TabIndex        =   36
         Top             =   1500
         Width           =   3135
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Deposito: "
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
         Left            =   210
         TabIndex        =   35
         Top             =   1170
         Width           =   2325
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Material:"
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
         Left            =   180
         TabIndex        =   30
         Top             =   360
         Width           =   2325
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Todos - Sin Clasificación"
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
         Left            =   90
         TabIndex        =   28
         Top             =   690
         Width           =   3135
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EEEEC0&
      Caption         =   "Rango de Fechas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   180
      TabIndex        =   18
      Top             =   1785
      Width           =   4600
      Begin VB.CommandButton Command5 
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
         Left            =   4215
         TabIndex        =   22
         Top             =   375
         Width           =   175
      End
      Begin VB.TextBox Text5 
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
         Left            =   3225
         TabIndex        =   21
         Text            =   " "
         Top             =   375
         Width           =   1020
      End
      Begin VB.TextBox Text4 
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
         Left            =   1095
         TabIndex        =   20
         Top             =   375
         Width           =   1020
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
         Left            =   2085
         TabIndex        =   19
         Top             =   375
         Width           =   175
      End
      Begin VB.Label Label10 
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
         Left            =   795
         TabIndex        =   24
         Top             =   435
         Width           =   2325
      End
      Begin VB.Label Label9 
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
         Left            =   -1320
         TabIndex        =   23
         Top             =   435
         Width           =   2325
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   7575
      Left            =   7200
      ScaleHeight     =   7515
      ScaleWidth      =   1635
      TabIndex        =   13
      Top             =   0
      Width           =   1695
      Begin VB.CommandButton Command8 
         Caption         =   "Ex"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "ESTASTK04.frx":26A2
         TabIndex        =   40
         ToolTipText     =   "Exportar Movimientos de Stock"
         Top             =   3675
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "ESTASTK04.frx":29AC
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Stock Minimo Recomendado"
         Top             =   2810
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   16
         Top             =   4620
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command11 
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
         Height          =   690
         Left            =   105
         Picture         =   "ESTASTK04.frx":2CB6
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   865
         Width           =   650
      End
      Begin VB.CommandButton Command10 
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
         Picture         =   "ESTASTK04.frx":2FC0
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   180
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "ESTASTK04.frx":310A
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4890
         Width           =   1515
      End
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   9135
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   2205
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00EEEEC0&
      Caption         =   "Rango de Códigos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   180
      TabIndex        =   1
      Top             =   105
      Width           =   6840
      Begin VB.CommandButton Command1 
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
         Left            =   3150
         TabIndex        =   4
         Top             =   460
         Width           =   175
      End
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
         Left            =   1260
         TabIndex        =   5
         Top             =   460
         Width           =   2010
      End
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
         Left            =   4515
         TabIndex        =   3
         Text            =   " "
         Top             =   460
         Width           =   2010
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
         Left            =   6510
         TabIndex        =   2
         Top             =   460
         Width           =   175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   3465
         TabIndex        =   10
         Top             =   900
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   105
         TabIndex        =   9
         Top             =   900
         Width           =   3255
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
         Left            =   -1155
         TabIndex        =   7
         Top             =   525
         Width           =   2325
      End
      Begin VB.Label Label5 
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
         Left            =   2100
         TabIndex        =   6
         Top             =   525
         Width           =   2325
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EEEEC0&
      Caption         =   "Formato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3690
      TabIndex        =   0
      Top             =   4620
      Width           =   2595
      Begin VB.OptionButton Option12 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Acumulado"
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
         Left            =   1250
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   1260
      End
      Begin VB.OptionButton Option11 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Detalle"
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
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   930
      End
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu listado 
      Caption         =   "Listados"
      Begin VB.Menu GenerarReporte 
         Caption         =   "Listar Reporte Segun Criterio Seleccionado"
      End
   End
End
Attribute VB_Name = "ESTASTK04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim COTIMOV$()
Private Sub Command1_Click()
    Call SELECHO("MASTER")
    COD1$ = TRIM$(VALACT1$("MASTER"))
    If COD1$ <> "" Then
       Text1.TEXT = COD1$
    End If
End Sub

Private Sub Command10_Click()
   Call CIERRARCH("*.*")
   Unload Me
End Sub

Private Sub Command12_Click()
    'DECLARACIONES
    Dim CAENT#(), CASAL#(), ACTI%()
    Command12.Enabled = False
    'VALIDACIONES
    If TRIM(Text4) = "" Or TRIM(Text5) = "" Then
      Call COMUNI("Rango de Fecha no valido")
      Command12.Enabled = True
      Exit Sub
    End If
    F1# = CDbl((CVDAT(FECHANUM(TRIM(Text4)))))
    F2# = CDbl((CVDAT(FECHANUM(TRIM(Text5)))))
    If F1# > F2# Then
      Call COMUNI("Rango de Fecha no valido")
      Command12.Enabled = True
      Exit Sub
    End If
    '
    If TRIM(Text1) = "" Or TRIM(Text2) = "" Then
      Call COMUNI("Código no Valido.")
      Command12.Enabled = True
      Exit Sub
    End If
    C1% = CODINT(Text1)
    C2% = CODINT(Text2)
    If C1% <= 0 Or C2% <= 0 Then
      Call COMUNI("Código no Valido.")
      Command12.Enabled = True
      Exit Sub
    End If
    '
    HAYCOMOV% = 0
    ReDim COTIMOV$(List2.ListCount - 1)
    For I% = 0 To List2.ListCount - 1
      COTIMOV$(I%) = ""
    Next I%
    For I% = 0 To List2.ListCount - 1
      If List2.Selected(I%) = True Then
        HAYCOMOV% = 1
        COTIMOV$(I%) = List2.List(I%)
      End If
    Next I%
    If HAYCOMOV% = 0 Then
      Call COMUNI("Código de Movimiento no Valido.")
      Command12.Enabled = True
      Exit Sub
    End If
    'ARMADO DE CONSULTA
    Call ABRESTOCKS
    Screen.MousePointer = 11
    SQLX$ = "SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE CDBL( FechMov ) >= " & F1# & " "
    SQLX$ = SQLX$ + " AND CDBL( FechMov ) <= " & F2# & " "
    M1% = 0
    For I% = 0 To List2.ListCount - 1 '8\\\OT-06-0218-WAR-05/06/06///
      If COTIMOV$(I%) <> "" Then
        If M1% = 0 Then
          COTIMOV1$ = Left(COTIMOV$(I%), 2)
          SQLX$ = SQLX$ + " AND ( CodiMovi = '" & COTIMOV1$ & "'"
          M1% = 1
        Else
          COTIMOV1$ = Left(COTIMOV$(I%), 2)
          SQLX$ = SQLX$ + " OR CodiMovi = '" & COTIMOV1$ & "'  "
        End If
      End If
    Next I%
    If M1% > 0 Then SQLX$ = SQLX$ + " )"
    '
    If XVALO(Text6) > 0 Then 'filtro desposito
      SQLX$ = SQLX$ + "AND DepoMovi = " & XVALO(Text6) & " "
    End If
       '
    'RECORDSET
    Set RST1TMP = Stocks.OpenRecordset(SQLX$, 4)
    With RST1TMP
      On Error Resume Next
      .MoveLast
      If Err Then
        Call COMUNI("No Existen Movimientos para el Criterio Seleccionado.")
        GoTo 99
      End If
      ULT1& = .RecordCount
      .MoveFirst
      On Error GoTo 0
      'INICIALIZACIONES
      JJ& = 0
      CONT& = 0
      ReDim CAENT#(NUMAS%)
      ReDim CASAL#(NUMAS%)
      ReDim ACTI%(NUMAS%)
      For I% = 1 To NUMAS%
        CAENT#(I%) = 0
        CASAL#(I%) = 0
        ACTI%(I%) = 0
      Next I%
      Call COPYSTRU("COKARDEX", "LIMOVSTO")
      REBLA$ = REGBLAN$("COKARDEX")
      CODI1$ = TRIM(Text1)
      CODI2$ = TRIM(Text2)
      '
      FIN& = NUMAS%
      For CIXI% = 1 To NUMAS%
        UK& = CIXI%
        Call TRACE(20, UK&, FIN&)
        CODD$ = CODEXT$(CIXI%)
        If UCase(CODD$) >= UCase(CODI1$) Then
          If UCase(CODD$) <= UCase(CODI2$) Then
            If XVALO(Text3) = 0 Or TIMATE%(CIXI%) = XVALO(Text3) Then 'filtro tipo material
              '
              ACTI%(CIXI%) = 1
              '
              If Check1.Value = 1 Then
                Call EXPLOMUL(CIXI%)
                For KI% = 0 To CONTAMUL%
                  On Error Resume Next
                  ACTI%(CIK%(KI%)) = 1
                  On Error GoTo 0
                Next KI%
              End If
              '
            End If
          End If
        End If
      Next CIXI%
      '
      'BUCLE
      Do While Not .EOF
        CIXI% = !COD_INTE
        If CIXI% < 1 Or CIXI% > NUMAS% Then GoTo 15
        '
        If ACTI%(CIXI%) = 1 Then
          CODD$ = CODEXT$(CIXI%)
          '
          If XVALO(Text6) > 0 Then 'filtro desposito
            If !DEPOMOVI <> XVALO(Text6) Then
              GoTo 15
             End If
          End If
          '
          YY$ = REBLA$
          Call REPLA(YY$, MKI$(CVINT(!FECHMOV)), 1, 2)
          Call REPLA(YY$, MKI$(!COD_INTE), 3, 2)
          Call REPLA(YY$, !IdenLote, 5, 16)
          Call REPLA(YY$, !CodiMovi, 21, 2)
          Call REPLA(YY$, !DoPriCor, 23, 10)
          '/** Por si no existe documento primario, pongo el secundario
          If TRIM$(!DoPriCor) = "" Then
            Call REPLA(YY$, !DoSecCor, 23, 10)
          End If
          '**/
          Call REPLA(YY$, !DoSecCor, 33, 10)
          Call REPLA(YY$, !ReferCor, 43, 30)
          Call REPLA(YY$, MKD$(!VALOUNIT), 73, 8)
            MONEX$ = "$": If !MoneVal = 2 Then MONEX$ = "U$"
          Call REPLA(YY$, MONEX$, 81, 2)
          Call REPLA(YY$, Chr$(!DEPOMOVI), 83, 1)
          Call REPLA(YY$, MKI$(!Nume_Clie), 84, 2)
          Call REPLA(YY$, MKI$(!Nume_Prov), 86, 2)
          Call REPLA(YY$, MKD$(!CANTINGRE), 88, 8)
            CAENT#(CIXI%) = CAENT#(CIXI%) + !CANTINGRE
          Call REPLA(YY$, MKD$(!CANTSALID), 96, 8)
            CASAL#(CIXI%) = CASAL#(CIXI%) + !CANTSALID
          Call REPLA(YY$, MKD$(0), 104, 8)
          Call REPLA(YY$, CVINT(!FeReMovi), 122, 2)
          Call REPLA(YY$, Chr$(!NumUsuar), 126, 1)
''          'PENDIENTE HABLAR CON LUIS ALVAREZ
''          SALDOCI# = (XVALO(!CantIngre) - XVALO(!Cantsalid))
''          Call REPLA(YY$, MKD$(SALDOCI#), 129, 8)
''          SUBTOTCI# = Abs(SALDOCI# * XVALO(!ValoUnit))
''          Call REPLA(YY$, MKD$(SUBTOTCI#), 137, 8)
''          Call REPLA(YY$, MKD$(SUBTOTCI#), 145, 8)
''          '
            JJ& = JJ& + 1
          'Call GRAREG("LIMOVSTO", YY$, JJ&)
          Call GRAREG("#LIMOVSTO", YY$, JJ&)
        End If
        '
15      CONT& = CONT& + 1
        Call TRACE(20, CONT&, ULT1&)
       .MoveNext
      Loop
    End With
    If JJ& = 0 Then
      Call COMUNI("No Existen Movimientos para el Criterio Seleccionado.")
      Command12.Enabled = True
      Exit Sub
    End If
    If Option12 Then 'SI ACUMULADO
      ZZ$ = REGBLAN("LIMOVSTO")
      JJ& = 0
      For I% = 1 To NUMAS%
        If CAENT#(I%) > 0 Or CASAL#(I%) > 0 Then
             Call REPLA(ZZ$, MKI$(I%), 3, 2)
             Call REPLA(ZZ$, MKD$(CAENT(I%)), 88, 8)
             Call REPLA(ZZ$, MKD$(CASAL(I%)), 96, 8)
             JJ& = JJ& + 1
             'Call GRAREG("LIMOVSTO", ZZ$, JJ&)
             Call GRAREG("#LIMOVSTO", ZZ$, JJ&)
        End If
      Next I%
    End If
    'CIERRE, ORDEN Y LLAMDO A ARCHIVO.
    'Call SETULTREG("LIMOVSTO", JJ&)
    Call SETULTREG("#LIMOVSTO", JJ&)
    Call FILEDUMP("LIMOVSTO")
    '
    Call FILESORT("LIMOVSTO", "", 1, 1, "ASC")     ' orden por fecha
    Call FILESORT("LIMOVSTO", "", 2, 2, "ASC")     ' orden por codigo
    If Option12 Then 'SI ACUMULADO
      Call FINAL("*LIMOSTAC")
    Else
      Call FINAL("*LIMOVSTO")
    End If
    Screen.MousePointer = 1
    '
99  On Error GoTo 0
    Command12.Enabled = True
    Screen.MousePointer = 1
End Sub

Private Sub Command2_Click()
    Call SELECHO("MASTER")
    COD2$ = TRIM$(VALACT1$("MASTER"))
    If COD2$ <> "" Then
       Text2.TEXT = COD2$
    End If
End Sub

Private Sub Command3_Click()
    Call SELECHO("TATIMAT")
    VDEV$ = VALACT1$("TATIMAT")
    If VDEV$ <> "" Then
       Text3.TEXT = VDEV$
       If TRIM$(Text3.TEXT) = "" Then
            Label4.Caption = "Todos - Sin Clasificación"
          Else
            Label4.Caption = VALACT2$("TATIMAT")
       End If
    End If
End Sub

Private Sub Command4_Click()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text4.TEXT = VDEV$
    End If
End Sub

Private Sub Command5_Click()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text5.TEXT = VDEV$
    End If
End Sub

Private Sub Command6_Click()
    Call SELECHO("TADEPOSI")
    VDEV$ = VALACT1$("TADEPOSI")
    If VDEV$ <> "" Then
       Text6.TEXT = VDEV$
       If TRIM$(Text6.TEXT) = "" Then
            Label8.Caption = "Todos - Sin Clasificación"
          Else
            Label8.Caption = VALACT2$("TADEPOSI")
       End If
    End If
End Sub

Sub Command7_Click()
  '\\\OT-05-0480-WAR-08/07/08///
  Dim CANTISAL#()
  'ESTE PROCESO ACONSEJA EL STOCK MINIMO QUE SE DEBERIA TENER SEGUN LO VENDIDO
  'SE ELIGE UN RANGO DE FECHAS
  'SE INDICA CUANTO TIEMPO DE COBERTURA SE QUIERE TENER
  'SE CALCULA LA CANTIDAD VENDIDA EN EL RANGO DE FECHA SELECCIONADO
  'SE OBTIENE EL CONSUMO PROMEDIO DIARIO DE LA CANTIDAD
  'SE MULTIPLICA POR LA CANTIDAD DE DIAS QUE SE DESEA CUBRIR Y ESTO DA LA CANTIDAD ACONSEJDADA DEL STOCK MINIMO
  'EJEMPLO EN 90 DIAS VENDI 500 UNIDADES
  'SI QUIERO TENER COBERTURA POR 30 DIAS = (500/90) * 30
  '
  Command7.Enabled = False
  Screen.MousePointer = 11
  Call DESHASFECHA(DES%, HAS%, "")
  DESDE = DES%: HASTA = HAS%
  CANDIA% = DIENTRE%(DESDE, HASTA)
  DES1& = CDbl(CVDAT(DES%))
  HAS1& = CDbl(CVDAT(HAS%))
9 Y2& = XVALO(CONTROL("", "CANDIACO"))
  Y3$ = InputBox("Ingrese la Cantidad de Dias de Cobertura", "Stock Mínimo Recomendado", Y2&)
  Y4 = XVALO(Y3$)
  If Y4 > 0 And Y4 < 30000 Then
      Call GRACONTROL("", "CANDIACO", Y3$)
      CANDIACOB% = Y4 'CANTIDAD DE DIAS DE COBERTURA
    Else
      GoTo 99   ' Call COMUNI("Cantidad de Dias no Valida")
  End If
  '
  ReDim CANTISAL#(NUMAS%)
  For I% = 1 To NUMAS%
    CANTISAL#(I%) = 0
  Next I%
  '
  Call ABRESTOCKS
  SQLX$ = "SELECT * FROM MOVISTO "
  SQLX$ = SQLX$ + " WHERE CodiMovi <> 'AJ' "
  SQLX$ = SQLX$ + " AND CantSalid > 0 "
  SQLX$ = SQLX$ + "AND INT(Cdbl(FechMov)) >= " & DES1& & " "
  SQLX$ = SQLX$ + "AND INT(Cdbl(FechMov)) <= " & HAS1& & " "
  Set SMRTMP = Stocks.OpenRecordset(SQLX$, 4)
  '
  FIN& = SMRTMP.RecordCount
  J2& = 0
  With SMRTMP
    On Error Resume Next
    .MoveFirst
    If Err Then
      Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
      GoTo 99
    End If
    On Error GoTo 0
    '
    Do While Not (.EOF)
      J2& = J2& + 1
      Call TRACE(24, J2&, FIN&)
      CIXI% = !COD_INTE
      On Error Resume Next    ' POR SI CIXI% > NUMAS%
      CASALI = XVALO(!CANTSALID)
      CANTISAL#(CIXI%) = CANTISAL#(CIXI%) + CASALI
      On Error GoTo 0
    .MoveNext
    Loop
  End With
  '
  For I% = 1 To NUMAS% 'DIVIDO POR LA CANTIDAD DE DIAS
    CANTISAL#(I%) = (CANTISAL#(I%) / CANDIA%) * CANDIACOB%
  Next I%
  '
  XX1$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX1$))
  XX2$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX2$))
  VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0)
  '
10 OBX$ = OBJPLA$("DESHASCOD1", VDEF$)
   If Len(OBX$) < 5 Then
       GoTo 99
   End If
   '
   CO1% = CVI(Mid$(OBX$, 1, 2))
   CO2% = CVI(Mid$(OBX$, 3, 2))
   TMT% = Asc(Mid$(OBX$, 5, 1))
   COD1$ = CODEXT$(CO1%)
   COD2$ = CODEXT$(CO2%)
   TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
   If TIMAT$ = "" Then
     TMT% = 0
     TIMAT$ = "Sin Clasificación"
   End If
   '
   Call HEADERS("STMINREC", "")
   Call HEADERS("STMINREC", "Tipo Material: " & TIMAT$)
   REBLA$ = REGBLAN("STMINREC")
   j& = 0
   T1& = XVALO(NUMAS%)
   For CIXI% = 1 To NUMAS%
     CDXX$ = CODEXT$(CIXI%)
     If CDXX$ >= COD1$ Then
       If CDXX$ <= COD2$ Then
         If TMT% = 0 Or TIMATE%(CIXI%) = TMT% Then 'filtro tipo material
            '
            ZZ$ = REBLA$
            Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
            Call REPLA(ZZ$, UNIMED(CIXI%), 3, 2)
            Call REPLA(ZZ$, STOMIN(CIXI%), 5, 4)
            Call REPLA(ZZ$, MKS(CANTISAL#(CIXI%)), 9, 4)
            Call REPLA(ZZ$, MKS(LOTESTAN(CIXI%)), 13, 4)
            If STOMIN(CIXI%) < CANTISAL#(CIXI%) Then
              Call REPLA(ZZ$, "Ajustar St.Minimo", 21, 28)
            End If
            j& = j& + 1
            Call GRAREG("STMINREC", ZZ$, j&)
            '
         End If
       End If
     End If
     J1& = XVALO(CIXI%)
   Call TRACE(24, J1&, T1&)
11 Next CIXI%
   Call SETULTREG("STMINREC", j&)
   Call FINAL("*STMINRE1")
   GoTo 10
   '
99 On Error GoTo 0
   Screen.MousePointer = 1
   Command7.Enabled = True
  '\\\OT-05-0480-WAR-FIN///
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   Call DESHASFECHA(DES%, HAS%, PERIO$)
   If PERIO$ <> "" Then
10   If EXISTE%(LUCOM$ + "STOCKS.MDB") > 0 Then
        Kill LUCOM$ + "STOCKS.MDB"
        GoTo 10
     End If
     '
     Call APERBASDAT("STOCKS")
     Set CREABASTOK = wrkPredeterminado.CreateDatabase(LUCOM$ + "STOCKS.MDB", dbLangGeneral)
     Dim BASTOCK As Database
     Set BASTOCK = OpenDatabase(LUCOM$ + "STOCKS.MDB")
     '
     Dim MOVIST As New TableDef
     
     Set CREATABSTO = BASTOCK.CreateTableDef("MOVISTO")
     NAMX = Movisto.Fields(0).Name
     TIPX = Movisto.Fields(0).Type
     SIZX = Movisto.Fields(0).Size
     CREATABSTO.Fields.Append CREATABSTO.CreateField(NAMX, TIPX, SIZX) ', TIPOVA, LOCAM%)
     BASTOCK.TableDefs.Append CREATABSTO
     Set MOVIST = BASTOCK.TableDefs("MOVISTO")
     '
     FIN& = Movisto.Fields.Count
     For KK1% = 2 To Movisto.Fields.Count
       NAMX = Movisto.Fields(KK1% - 1).Name
       TIPX = Movisto.Fields(KK1% - 1).Type
       SIZX = Movisto.Fields(KK1% - 1).Size
       With MOVIST.Fields
         .Append MOVIST.CreateField(NAMX, TIPX, SIZX)
         TICAM% = TIPX
         NCAMPO$ = NAMX
         If TICAM% = 1 Or TICAM% = 10 Or TICAM% = 12 Then
               MOVIST.Fields(NCAMPO$).AllowZeroLength = True
            ElseIf TICAM% = 1 Then
               MOVIST.Fields(NCAMPO$).DefaultValue = True
            ElseIf TICAM% = 8 Then    ' FECHA
               MOVIST.Fields(NCAMPO$).DefaultValue = CVDAT(0)
            Else
               MOVIST.Fields(NCAMPO$).DefaultValue = 0
         End If
       End With
     Next KK1%
     
     Set MOVISLOC = BASTOCK.OpenRecordset("MOVISTO", dbOpenDynaset)
     '
     F1# = CDbl(CVDAT(DES%))
     F2# = CDbl(CVDAT(HAS%))
     '
     'ARMADO DE CONSULTA
     Call ABRESTOCKS
     Screen.MousePointer = 11
     SQLX$ = "SELECT * FROM MOVISTO "
     SQLX$ = SQLX$ + " WHERE CDBL( FechMov ) >= " & F1# & " "
     SQLX$ = SQLX$ + " AND CDBL( FechMov ) <= " & F2# & " "
     Set RST1TMP = Stocks.OpenRecordset(SQLX$, 4)
     With RST1TMP
       On Error Resume Next
       .MoveLast
       If Err Then
         Call COMUNI("No Existen Movimientos para el Criterio Seleccionado.")
         GoTo 99
       End If
       ULT1& = .RecordCount
       .MoveFirst
       KKII& = 0
       On Error GoTo 0
       '
       Do While Not .EOF
          MOVISLOC.AddNew
          For KKI% = 1 To .Fields.Count
            XVA = .Fields(KKI% - 1).Value
            MOVISLOC.Fields(KKI% - 1).Value = XVA
          Next KKI%
          MOVISLOC.Update
       .MoveNext
       KKII& = KKII& + 1
       Call TRACE(20, KKII&, ULT1&)
       Loop
     End With
     '
   End If
   '
   Call COMUNI("Base de Datos 'STOCKS.MDB'\Creada en 'C:/FLEXOFT'\Con " + TRIM$(Str$(ULT1&)) + " Registros.")
   '
99 Screen.MousePointer = 1
   Command8.Enabled = True
End Sub

Private Sub Form_Load()
'
 If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture6.Visible = True
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6
      Me.Check1.BackColor = &HFCD7B6
      Me.Option11.BackColor = &HFCD7B6
      Me.Option12.BackColor = &HFCD7B6
      MARCOBARRA.Top = 5800
      Set MARCOBARRA.Container = Me.Picture6

 End If

List2.AddItem "AJ - Ajustes"
List2.AddItem "BR - Recepción"
List2.AddItem "CF - Consumo Fábrica"
List2.AddItem "DV - Devolución"
List2.AddItem "OE - Orden de Empaque"
List2.AddItem "OF - Orden de Fabricación"
List2.AddItem "PF - Parte de Fabricación"
List2.AddItem "RC - Remito de Consignación"
List2.AddItem "RS - Remito"
List2.AddItem "RT - Remito de Ventas"
List2.AddItem "TR - Transferencias de Stock"
List2.AddItem "SC - Vale de Scrap"
'\\\OT-8-974-24/10/08
List2.AddItem "VM - Vale de Consumo"
'\\\OT-8-974-24/10/08
'Valores por default
Text4 = FECHATEX$(COMESACT%)
Text5 = HOS$
'
Text1 = TRIM(Left$(REGLEIDO$("ECOMAST", 1), 16))
Text2 = TRIM(Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16))
'
End Sub

Private Sub GenerarReporte_Click()
    Call Command12_Click
End Sub

Private Sub salir_Click()
   Call Command10_Click
End Sub


Private Sub Text1_Change()
    CI1% = CODINT%(Text1.TEXT)
    If CI1% > 0 Then
         Label1.Caption = DESCRIT$(CI1%)
       Else
         Label1.Caption = ""
    End If
End Sub

Private Sub Text2_Change()
    CI2% = CODINT%(Text2.TEXT)
    If CI2% > 0 Then
         Label2.Caption = DESCRIT$(CI2%)
       Else
         Label2.Caption = ""
    End If
End Sub

Private Sub Text3_Change()
    If XVALO(Text3.TEXT) < 256 Then
        VENDE% = XVALO(Text3.TEXT)
        VENDX$ = ECOARCH$("TATIMAT", Chr$(VENDE%))
        If VENDX <> "" Then
            Label4.Caption = VENDX$
            Exit Sub
          Else
            Label4.Caption = "Todos - Sin Clasificación"
        End If
    End If
End Sub

Private Sub Text6_Change()
    If XVALO(Text6.TEXT) < 256 Then
        VENDE% = XVALO(Text6.TEXT)
        VENDX$ = ECOARCH$("TADEPOSI", Chr$(VENDE%))
        If VENDX <> "" Then
            Label8.Caption = VENDX$
            Exit Sub
          Else
            Label8.Caption = "Todos - Sin Clasificación"
        End If
    End If
End Sub
