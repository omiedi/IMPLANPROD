VERSION 5.00
Begin VB.Form ENSANOR 
   BackColor       =   &H00EEEEDD&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ensayos Normalizados"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   10960
      Picture         =   "ENSANOR.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Alta de Nuevo Código de Ensayo"
      Top             =   945
      Width           =   750
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   10960
      Picture         =   "ENSANOR.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Guarda los Cambios Realizados"
      Top             =   1680
      Width           =   750
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   10900
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   13
      Top             =   7455
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00EEEEDD&
      Caption         =   "Definiciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1800
      Left            =   105
      TabIndex        =   9
      Top             =   105
      Width           =   10620
      Begin VB.CommandButton Command5 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5460
         TabIndex        =   21
         Top             =   420
         Width           =   180
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   7350
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "ENSANOR.frx":0454
         Top             =   420
         Width           =   3060
      End
      Begin VB.Label Label4 
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
         Left            =   210
         TabIndex        =   23
         Top             =   840
         Width           =   5450
      End
      Begin VB.Label Label3 
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
         Left            =   3360
         TabIndex        =   22
         Top             =   420
         Width           =   2115
      End
      Begin VB.Label Label1 
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
         Left            =   3360
         TabIndex        =   17
         Top             =   1260
         Width           =   2295
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Ensayo:"
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
         Index           =   2
         Left            =   1260
         TabIndex        =   16
         Top             =   1350
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
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
         Index           =   0
         Left            =   1260
         TabIndex        =   12
         Top             =   500
         Width           =   1905
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00EEEEDD&
         Caption         =   "Observaciones:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5985
         TabIndex        =   11
         Top             =   420
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEDD&
      Caption         =   "Mediciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   4455
      Left            =   105
      TabIndex        =   2
      Top             =   1995
      Width           =   10620
      Begin VB.Frame Frame6 
         Caption         =   "Frame4"
         Height          =   3800
         Left            =   6825
         TabIndex        =   24
         Top             =   735
         Width           =   15
      End
      Begin VB.Frame Frame5 
         Caption         =   "Frame4"
         Height          =   3800
         Left            =   5480
         TabIndex        =   19
         Top             =   735
         Width           =   15
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame4"
         Height          =   3800
         Left            =   3780
         TabIndex        =   18
         Top             =   735
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   3800
         Left            =   10290
         TabIndex        =   8
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   3800
         Left            =   4200
         TabIndex        =   6
         Top             =   735
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   3800
         Left            =   735
         TabIndex        =   5
         Top             =   630
         Width           =   15
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3660
         Left            =   0
         TabIndex        =   4
         Top             =   735
         Width           =   10575
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   435
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   10515
         TabIndex        =   3
         Top             =   315
         Width           =   10580
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cod.  Variable Larga           UM        Min        Max  Observaciones       "
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   105
            TabIndex        =   7
            Top             =   105
            Width           =   10200
         End
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   10605
         Y1              =   4410
         Y2              =   4410
      End
      Begin VB.Line Line1 
         X1              =   10580
         X2              =   10580
         Y1              =   320
         Y2              =   6225
      End
   End
   Begin VB.CommandButton Command2 
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
      Height          =   700
      Left            =   10960
      Picture         =   "ENSANOR.frx":0456
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   2415
      Width           =   750
   End
   Begin VB.CommandButton Command9 
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
      Height          =   700
      Left            =   10960
      Picture         =   "ENSANOR.frx":0760
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Sale de la Aplicación"
      Top             =   210
      Width           =   750
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   10340
      Picture         =   "ENSANOR.frx":08AA
      Top             =   7710
      Width           =   2010
   End
End
Attribute VB_Name = "ENSANOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Command1.Enabled = False
    '
    Call TRALOCAL("IENSANOR", "")
10
    VTB% = VENTABU("TABENSA")
    If VTB% >= 0 Then
      ' Valida que no hayan codigos de ensayo repetidos
      FIN& = ULTREG&("!IENSANOR")
      For U& = 1 To (FIN& - 1)
        X$ = REGLEIDO$("!IENSANOR", U&)
        For U1& = (U& + 1) To FIN&
          Z$ = REGLEIDO$("!IENSANOR", U1&)
          If UCase(TRIM$(Mid$(X$, 1, 16))) = UCase(TRIM$(Mid$(Z$, 1, 16))) Then
            Call COMUNI("Codigo de Ensayo Repetido")
            GoTo 10
          End If
        Next U1&
      Next U&
      '
      For U& = 1 To ULTREG&("!IENSANOR")
        X$ = REGLEIDO$("!IENSANOR", U&)
        '
        ' Valida que el codigo no sea nulo
        CODENSA$ = TRIM$(Mid$(X$, 1, 16))
        If Len(VARLARGA$) < 0 Then
          Call MENSERR(24, "Codigo de Ensayo Nulo")
          GoTo 10
        End If
      Next U&
      '
      Call TRACENTRAL("IENSANOR", "")
      Call FRESHECHO("IENSANOR")
    End If
    '
99  Command1.Enabled = True
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    '
    X$ = FILTERGETRECORD("IENSANOR", 1, Label3)
    Call REPLA(X$, Text7, 83, 174)
    Call FILTERPUTRECORD("IENSANOR", 1, Label3, X$)
    '
99  Command3.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    '
    X$ = REGSEL$("IENSANOR")
    If X$ = "" Then GoTo 99
    '
    Call BLANFORMU
    Label3 = TRIM$(Mid$(X$, 1, 16))
    X$ = FILTERGETRECORD("IENSANOR", 1, Label3)
    Label4 = TRIM$(Mid$(X$, 17, 64))
    TIPOENSA$ = TRIM$(Mid$(X$, 81, 2))
    Label1 = ECOARCH$("TABTIEN", TIPOENSA$)
    Text7 = TRIM$(Mid$(X$, 83, 64))
    '
    Call FILTERFILE("TADEFEN", "", 1, Label3)
    For U& = 1 To ULTREG&("!TADEFEN")
      X$ = REGLEIDO$("!TADEFEN", U&)
      ZZ$ = " " + AJUSTI$(Mid$(X$, 17, 4), 5) ' Codigo
      ZZ$ = ZZ$ + AJUSTI$(Mid$(X$, 21, 32), 25) ' Variable
      ZZ$ = ZZ$ + AJUSTD$(Mid$(X$, 57, 2), 3) ' Unidad de medida
      ZZ$ = ZZ$ + " " + FORMATNUM(CVD(Mid$(X$, 59, 8)), "F10.3") ' min
      ZZ$ = ZZ$ + " " + FORMATNUM(CVD(Mid$(X$, 67, 8)), "F10.3") ' max
      ZZ$ = ZZ$ + "  " + AJUSTI$(Mid$(X$, 83, 32), 24) ' observaciones
      List1.AddItem ZZ$
    Next U&
    '
99  Command5.Enabled = True
End Sub

Private Sub Command9_Click()
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")
End Sub

Private Sub Label2_Click()
    '
    If TRIM$(Label3) = "" Then Exit Sub
    
    Call FILTERFILE("TADEFEN", "", 1, Label3)
10
    VTB% = VENTABU("TADEFEN")
    If VTB% >= 0 Then
      Call FILTERUPDATE("TADEFEN", "", 1, Label3)
    End If
    '
    List1.Clear
    For U& = 1 To ULTREG&("!TADEFEN")
      X$ = REGLEIDO$("!TADEFEN", U&)
      '
      CODVAR$ = Mid$(X$, 17, 4)
      If TRIM$(CODVAR$) = "" Then
        Call COMUNI("Código de Variable No Válido en Renglón " + CStr(U&))
        GoTo 10
      End If
      ZZ$ = " " + AJUSTI$(CODVAR$, 5) ' Codigo
      '
      ZZ$ = ZZ$ + AJUSTI$(Mid$(X$, 21, 32), 25) ' Variable
      '
      UNIDAD$ = Mid$(X$, 57, 2)
      If TRIM$(UNIDAD$) = "" Then
        Call COMUNI("Falta Unidad de Medidad en Renglón " + CStr(U&))
        GoTo 10
      End If
      ZZ$ = ZZ$ + AJUSTD$(UNIDAD$, 3) ' Unidad de medida
      '
      MINI! = CVD(Mid$(X$, 59, 8))
      MAXI! = CVD(Mid$(X$, 67, 8))
      If MINI! > MAXI! Then
        Call COMUNI("Valor Mínimo mayor a Valor Máximo en Renglón " + CStr(U&))
        GoTo 10
      End If
      ZZ$ = ZZ$ + " " + FORMATNUM(MINI!, "F10.3") ' min
      ZZ$ = ZZ$ + " " + FORMATNUM(MAXI!, "F10.3") ' max
      ZZ$ = ZZ$ + "  " + AJUSTI$(Mid$(X$, 83, 32), 24) ' observaciones
      List1.AddItem ZZ$
    Next U&
    '
End Sub

Sub BLANFORMU()
    '
    Label3 = ""
    Label4 = ""
    Label1 = ""
    Text7 = ""
    List1.Clear
    '
End Sub

Private Sub List1_Click()
    '
    Call Label2_Click
    '
End Sub
