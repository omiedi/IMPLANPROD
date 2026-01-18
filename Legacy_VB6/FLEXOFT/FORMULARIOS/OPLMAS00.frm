VERSION 5.00
Begin VB.Form OPLMAS00 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maestro de Items de Stock - Complementos, Consultas, Listados"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3945
      Picture         =   "OPLMAS00.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
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
      TabIndex        =   4
      Top             =   2310
      Width           =   3585
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   5
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
            Picture         =   "OPLMAS00.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   6
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
            TabIndex        =   7
            Top             =   105
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   8
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
            Picture         =   "OPLMAS00.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   9
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
            TabIndex        =   10
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
      Left            =   3945
      Picture         =   "OPLMAS00.frx":075E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cierra - Abandona la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   3945
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   2
      Top             =   3255
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
      Height          =   2010
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   3585
      Begin VB.PictureBox Picture11 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   12
         Top             =   1155
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
            Picture         =   "OPLMAS00.frx":08A8
            Style           =   1  'Graphical
            TabIndex        =   13
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
            TabIndex        =   14
            Top             =   105
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture12 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   15
         Top             =   630
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
            Picture         =   "OPLMAS00.frx":0BB2
            Style           =   1  'Graphical
            TabIndex        =   16
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
            TabIndex        =   17
            Top             =   50
            Width           =   1695
         End
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3360
      Picture         =   "OPLMAS00.frx":0CFC
      Top             =   3585
      Width           =   2010
   End
End
Attribute VB_Name = "OPLMAS00"
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

Private Sub Command3_Click()
   Command3.Enabled = False
   Screen.MousePointer = 11
   FIN& = NUINV%
   LISITBAJ.List1.Clear
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     IK0$ = REGLEIDO$("INVERT", U& + 1)
     CIXI% = CVI(Mid$(IK0$, 17, 2))
     If CIXI% > 0 Then
       If CIXI% <= NUMAS% Then
         J& = CIXI% + 1
         X$ = REGLEIDO$("MASTER", J&)
         If Asc(Mid$(X$, 92, 1)) = 96 Then
           TTXX$ = Space$(80)
           Call REPLA(TTXX$, Left$(IK0$, 15), 1, 15)
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
   Command3.Enabled = True
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
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0)
    '
10  RACOD$ = OBJPLA$("DESHASCOD1", VDEF$)
    If RACOD$ = "" Then
        Exit Sub
    End If
    CI1% = CVI(Left$(RACOD$, 2)): COD1$ = CODEXT$(CI1%)
    CI2% = CVI(Mid$(RACOD$, 3, 2)): COD2$ = CODEXT$(CI2%)
    TMT% = Asc(Mid$(RACOD$, 5, 1))
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TIMAT$ = "" Then
        TMT% = 0
        Call REPLA(RACOD$, Chr$(0), 5, 1)
        TIMAT$ = "Sin Clasificación"
    End If
    '
    If COD1$ = "" Or COD2$ = "" Or COD2$ < COD1$ Then
        VDEF$ = RACOD$
        GoTo 10
    End If
    '
    JJ& = 0
    If MODO% = 1 Then
        'Call BLANARCH("LIMASEL")
        Call HEADERS("LIMASEL", "")
        Call HEADERS("LIMASEL", "Tipo de Material: " + TIMAT$)
        Call HEADERS("LIMASEL", "Desde Codigo: " + TRIM$(COD1$) + " - " + DESCRIT$(CI1%))
        Call HEADERS("LIMASEL", "Hasta Codigo: " + TRIM$(COD2$) + " - " + DESCRIT$(CI2%))
      ElseIf MODO% = 2 Then
        'Call BLANARCH("LIMAINV")
        Call HEADERS("LIMAINV", "")
        Call HEADERS("LIMAINV", "Tipo de Material: " + TIMAT$)
        Call HEADERS("LIMAINV", "Desde Codigo: " + TRIM$(COD1$) + " - " + DESCRIT$(CI1%))
        Call HEADERS("LIMAINV", "Hasta Codigo: " + TRIM$(COD2$) + " - " + DESCRIT$(CI2%))
        Call HEADERS("LIMAINV", "Stock Nominal al: " + HO$ + " - " + Time$ + " hs.")
      ElseIf MODO% = 3 Then
        'Call BLANARCH("LIMALOT")
        Call HEADERS("LIMALOT", "")
        Call HEADERS("LIMALOT", "Material: " + TRIM$(TIMAT$) + " (" + TRIM$(COD1$) + " - " + TRIM$(COD2$) + ")")
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
        If COD$ < COD1$ Then GoTo 145
        If COD$ > COD2$ Then GoTo 149
        '
        CI% = CVI(Mid$(IV$, 17, 2))
        If CI% > 0 Then
            If CI% <= NUMAS% Then
                REGMAS& = CI% + 1
                M$ = REGLEIDO$("MASTER", REGMAS&)
                If TMT% > 0 Then
                    If Asc(Mid$(M$, 93, 1)) <> TMT% Then GoTo 145
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
    End If
    '
    Screen.MousePointer = 1
End Sub
'
