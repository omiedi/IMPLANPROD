VERSION 5.00
Begin VB.Form OPLIMAS 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maestro de Items de Stock - Complementos, Consultas, Listados"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      BackColor       =   &H80000016&
      Height          =   645
      Left            =   1680
      ScaleHeight     =   585
      ScaleWidth      =   2580
      TabIndex        =   27
      Top             =   2100
      Width           =   2640
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
         Picture         =   "OPLIMAS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Valorización de Lote Mínimo / Económico"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Valorización de Lote Mínimo / Económico"
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
         TabIndex        =   29
         Top             =   105
         Width           =   1905
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Control"
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
      Left            =   4995
      Picture         =   "OPLIMAS.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2205
      Width           =   855
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
      Height          =   855
      Left            =   4995
      Picture         =   "OPLIMAS.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1260
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Items Inactivos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   150
      TabIndex        =   11
      Top             =   5250
      Width           =   4635
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1155
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   12
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
            Picture         =   "OPLIMAS.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   13
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
            TabIndex        =   14
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
         TabIndex        =   15
         Top             =   420
         Width           =   2640
         Begin VB.CommandButton COMMAND3 
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
            Index           =   0
            Left            =   0
            Picture         =   "OPLIMAS.frx":0C28
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Ver y Restaurar Items Inactivos"
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
            TabIndex        =   17
            Top             =   45
            Width           =   1905
         End
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
      Height          =   855
      Left            =   4995
      Picture         =   "OPLIMAS.frx":0D72
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cierra - Abandona la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4995
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   2
      Top             =   6200
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   3
         Top             =   105
         Width           =   12000
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Consultas y Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4950
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   4635
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Costos de Reposición"
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
         Left            =   210
         TabIndex        =   18
         Top             =   2900
         Width           =   4215
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   945
            ScaleHeight     =   585
            ScaleWidth      =   2580
            TabIndex        =   19
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
               Picture         =   "OPLIMAS.frx":0EBC
               Style           =   1  'Graphical
               TabIndex        =   20
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
               TabIndex        =   21
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
            TabIndex        =   22
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
               Picture         =   "OPLIMAS.frx":11C6
               Style           =   1  'Graphical
               TabIndex        =   23
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
               TabIndex        =   24
               Top             =   50
               Width           =   1695
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Listados Generales del Maestro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2300
         Left            =   210
         TabIndex        =   4
         Top             =   420
         Width           =   4215
         Begin VB.PictureBox Picture11 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   945
            ScaleHeight     =   585
            ScaleWidth      =   2580
            TabIndex        =   8
            Top             =   945
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
               Picture         =   "OPLIMAS.frx":14D0
               Style           =   1  'Graphical
               TabIndex        =   9
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
               TabIndex        =   10
               Top             =   80
               Width           =   1905
            End
         End
         Begin VB.PictureBox Picture12 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2580
            TabIndex        =   5
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
               Picture         =   "OPLIMAS.frx":17DA
               Style           =   1  'Graphical
               TabIndex        =   6
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
               TabIndex        =   7
               Top             =   50
               Width           =   1695
            End
         End
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4410
      Picture         =   "OPLIMAS.frx":1924
      Top             =   6530
      Width           =   2010
   End
End
Attribute VB_Name = "OPLIMAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Hide
   '
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call COMUNI("Solo Disponible como\Consulta por Pantalla")
   Call Command3_Click(0)
   Command2.Enabled = True
End Sub

Private Sub Command3_Click(Index As Integer)
   COMMAND3(Index).Enabled = False
   Screen.MousePointer = 11
   FIN& = NUINV%
   LISITBAJ.List1.Clear
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     I0$ = REGLEIDO$("INVERT", U& + 1)
     CIXI% = CVI(Mid$(I0$, 17, 2))
     If CIXI% > 0 Then
       If CIXI% <= NUMAS% Then
         J& = CIXI% + 1
         X$ = REGLEIDO$("MASTER", J&)
         If Asc(Mid$(X$, 92, 1)) = 96 Then
           TTXX$ = Space$(80)
           Call REPLA(TTXX$, Left$(I0$, 15), 1, 15)
           Call REPLA(TTXX$, DESCRIT0$(CIXI%), 19, 46)
           Call REPLA(TTXX$, Str$(U&), 65, 8)
           Call REPLA(TTXX$, Str$(CIXI%), 73, 8)
           LISITBAJ.List1.AddItem TTXX$
         End If
       End If
     End If
   Next U&
   '
   Screen.MousePointer = 1
   LISITBAJ.Show 1
   '
   COMMAND3(Index).Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call GELICOREP(2)
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
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


Private Sub Command8_Click()
    '\\\OT-07-0005-RG-05/01/07///
    Command8.Enabled = False
    Call LIMASTER(4)
    Command8.Enabled = True
    '\\\OT-07-0005-RG-FIN///
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    Call LIMASTER(1)
    Command9.Enabled = True
End Sub
'
Private Sub COMMAND15_CLICK()
    Command9.Enabled = False
    Call LIMASTER(2)
    Command9.Enabled = True
End Sub
'
Sub LIMASTER(MODO%)
    '
    HOO% = HOY(HO$)
    If MODO% = 0 Then
        MODO% = ALTERNA%("Listado de Códigos y Atributos\Código y Descripción para Inventarios\Stock Mínimo y Lote Reposición")
        If MODO% < 1 Or MODO% > 3 Then
          Exit Sub
        End If
      ElseIf MODO% = 1 Then
        MODO% = ALTERNA%("Listado de Códigos y Atributos\Stock Mínimo y Lote Reposición")
        If MODO% < 1 Or MODO% > 2 Then
          Exit Sub
        End If
        If MODO% = 2 Then MODO% = 3
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
    CI1% = CVI(Left$(RACOD$, 2)): cod1$ = CODEXT$(CI1%)
    CI2% = CVI(Mid$(RACOD$, 3, 2)): cod2$ = CODEXT$(CI2%)
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
    If cod1$ = "" Or cod2$ = "" Or cod2$ < cod1$ Then
        VDEF$ = RACOD$
        GoTo 10
    End If
    '
    JJ& = 0
    If MODO% = 1 Then
        'Call BLANARCH("LIMASEL")
        Call HEADERS("LIMASEL", "")
        Call HEADERS("LIMASEL", "Tipo / Grupo de Material: " + TIMAT$ + " / " + GRUCO$)
        Call HEADERS("LIMASEL", "Desde Codigo: " + TRIM$(cod1$) + " - " + DESCRIT$(CI1%))
        Call HEADERS("LIMASEL", "Hasta Codigo: " + TRIM$(cod2$) + " - " + DESCRIT$(CI2%))
      ElseIf MODO% = 2 Then
        'Call BLANARCH("LIMAINV")
        Call HEADERS("LIMAINV", "")
        Call HEADERS("LIMAINV", "Tipo / Grupo de Material: " + TIMAT$ + " / " + GRUCO$)
        Call HEADERS("LIMAINV", "Desde Codigo: " + TRIM$(cod1$) + " - " + DESCRIT$(CI1%))
        Call HEADERS("LIMAINV", "Hasta Codigo: " + TRIM$(cod2$) + " - " + DESCRIT$(CI2%))
        Call HEADERS("LIMAINV", "Stock Nominal al: " + HO$ + " - " + Time$ + " hs.")
      ElseIf MODO% = 3 Then
        'Call BLANARCH("LIMALOT")
        Call HEADERS("LIMALOT", "")
        Call HEADERS("LIMALOT", "Material: " + TRIM$(TIMAT$) + " (" + TRIM$(cod1$) + " - " + TRIM$(cod2$) + ")")
        '
        '\\\OT-07-0005-RG-05/01/07///
      ElseIf MODO% = 4 Then
        Call HEADERS("VAMINEC", "")
        Call HEADERS("VAMINEC", "Tipo / Grupo de Material: " + TIMAT$ + " / " + GRUCO$)
        Call HEADERS("VAMINEC", "Desde Codigo: " + TRIM$(cod1$) + " - " + DESCRIT$(CI1%))
        Call HEADERS("VAMINEC", "Hasta Codigo: " + TRIM$(cod2$) + " - " + DESCRIT$(CI2%))
        '\\\OT-07-0005-RG-FIN///
        '
      Else
        Screen.MousePointer = 1
        Exit Sub
    End If
    '
    Screen.MousePointer = 11
    J& = 0: FIN& = NUINV%
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        IV$ = REGLEIDO$("INVERT", I& + 1)
        If Mid$(IV$, 16, 1) = "`" Then GoTo 145          ' codigo de baja
        '
        COD$ = Left$(IV$, 16)
        If COD$ < cod1$ Then GoTo 145
        If COD$ > cod2$ Then GoTo 149
        '
        CI% = CVI(Mid$(IV$, 17, 2))
        If CI% > 0 Then
            If CI% <= NUMAS% Then
                REGMAS& = CI% + 1
                M$ = REGLEIDO$("MASTER", REGMAS&)
                If TMT% > 0 Then
                    If Asc(Mid$(M$, 93, 1)) <> TMT% Then GoTo 145
                End If
                If GRC% > 0 Then
                    If CVI(Mid$(M$, 125, 2)) <> GRC% Then GoTo 145
                End If
                '
                If MODO% = 1 Then
                   ML$ = Left$(M$, 2) + Mid$(M$, 93, 36)
                   JJ& = JJ& + 1: Call GRAREG("LIMASEL", ML$, JJ&)
                 ElseIf MODO% = 2 Then
                   ML$ = MKI$(0) + "  " + MKS$(0) + Space$(56)
                   Call REGAPP("LIMAINV", ML$)
                   ML$ = Left$(M$, 2) + Mid$(M$, 109, 2)
                   ML$ = ML$ + MKS$(STOCKACT(CI%))
                   ML$ = ML$ + String$(62, ".")
                   JJ& = JJ& + 1: Call GRAREG("LIMAINV", ML$, JJ&)
                 ElseIf MODO% = 3 Then
                   ML$ = REGBLAN$("LIMALOT")
                   Call REPLA(ML$, Left$(M$, 2), 1, 2)
                   Call REPLA(ML$, MKS$(STOMIN(CI%)), 3, 4)
                   Call REPLA(ML$, MKS$(LOTESTAN(CI%)), 7, 4)
                   Call REPLA(ML$, MKS$(LOREPOMA(CI%)), 11, 4)
                   Call REPLA(ML$, MKS$(COSUNI(CI%)), 15, 4)
                   Call REPLA(ML$, MKI$(FECOSTO%(CI%)), 19, 2)
                   JJ& = JJ& + 1: Call GRAREG("LIMALOT", ML$, JJ&)
                 '
                 '\\\OT-07-0005-RG-05/01/07///
                 ElseIf MODO% = 4 Then
                   ML$ = REGBLAN("VAMINEC")
                   Call REPLA(ML$, Left$(M$, 2), 1, 2)
                   Call REPLA(ML$, MKS$(COSUNI(CI%)), 3, 4)
                   Call REPLA(ML$, MKS$(STOMIN(CI%)), 7, 4)
                   PREC# = COSUNI(CI%)
                   STMIN# = STOMIN(CI%)
                   VAL1# = PREC# * STMIN#
                   Call REPLA(ML$, MKS$(VAL1#), 11, 4)
                   Call REPLA(ML$, MKS$(LOTESTAN(CI%)), 15, 4)
                   LECOM# = LOTESTAN(CI%)
                   VAL2# = PREC# * LECOM#
                   Call REPLA(ML$, MKS$(VAL2#), 19, 4)
                   JJ& = JJ& + 1: Call GRAREG("VAMINEC", ML$, JJ&)
                   '\\\OT-07-0005-RG-FIN///
                   '
                End If
            End If
        End If
145 Next I&
    '
149 FIECO% = 0
    If MODO% = 1 Then
        Call SETULTREG("LIMASEL", JJ&)
        Call CIERRARCH("LIMASEL")
        Call FINAL("*LIMASEL")
      ElseIf MODO% = 2 Then
        Call SETULTREG("LIMAINV", JJ&)
        Call CIERRARCH("LIMAINV")
        Call FINAL("*LIMAINV")
      ElseIf MODO% = 3 Then
        Call SETULTREG("LIMALOT", JJ&)
        Call CIERRARCH("LIMALOT")
        Call FINAL("*LIMALOT")
      '
      '\\\OT-07-0005-RG-05/01/07///
      ElseIf MODO% = 4 Then
        Call SETULTREG("VAMINEC", JJ&)
        Call CIERRARCH("VAMINEC")
        Call FINAL("*VAMINEC")
      '\\\OT-07-0005-RG-FIN///
      '
    End If
    '
    Screen.MousePointer = 1
End Sub
'
Sub GELICOREP(MODO%)
   '
   ' MODO% = 1 --> Ordenado por Proveedor
   ' MODO% = 2 --> Ordenado por Articulo
   '
   JJ& = 0
   FIN& = NUINV%
   REBLA$ = REGBLAN$("LICOREP")
   Call BLOQARCH("LICOREP")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     I0$ = REGLEIDO$("INVERT", U& + 1)
     CODE$ = TRIM$(Left$(I0$, 16))
     If Asc(Mid$(I0$, 16, 1)) <> 96 Then
       CIXI% = CVI(Mid$(I0$, 17, 2))
       If CIXI% > 0 Then
         If CIXI% <= NUMAS% Then
           REGMA& = CIXI% + 1
           REMAS$ = REGLEIDO$("MASTER", REGMA&)
           PHHI% = CVI(Mid$(REMAS$, 123, 2))
           If PHHI% > 0 Then
             MONECO% = Asc(Mid$(REMAS$, 95, 1))
             ZZ$ = REBLA$
             Call REPLA(ZZ$, MKI$(PHHI%), 1, 2)
             Call REPLA(ZZ$, MKI$(CIXI%), 3, 2)
             Call REPLA(ZZ$, Mid$(REMAS$, 109, 2), 5, 2)
             Call REPLA(ZZ$, MKI$(MONECO%), 7, 2)
             Call REPLA(ZZ$, Mid$(REMAS$, 111, 4), 9, 4)
             Call REPLA(ZZ$, Mid$(REMAS$, 115, 2), 13, 2)
             Call REPLA(ZZ$, MKI$(PHHI%), 15, 2)
             JJ& = JJ& + 1
             Call GRAREG("LICOREP", ZZ$, JJ&)
           End If
         End If
       End If
     End If
   Next U&
   Call SETULTREG("LICOREP", JJ&)
   Call CIERRARCH("LICOREP")
   '
   If MODO% = 1 Then
       Call CONECRUN("*SOCOREP", "Ordenamiento Costos Reposición s/Proveedor")
     Else
       Call CONECRUN("*LICOREP", "Listado Costos de Reposición por Proveedor")
   End If
   '
End Sub

Private Sub Form_Load()
   If Left$(UCase$(App.EXEName), 6) <> "ANACOS" Then
     Command6.Visible = False
   End If
End Sub

