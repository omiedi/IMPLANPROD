VERSION 5.00
Begin VB.Form DISPOSTO 
   BackColor       =   &H00EEEEC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Disponibilidad General de Stocks"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8265
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
   ScaleHeight     =   4650
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   7280
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   30
      Top             =   2940
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Setup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   7245
      Picture         =   "DISPOSTO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1995
      Width           =   855
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
      Height          =   1485
      Left            =   210
      TabIndex        =   16
      Top             =   1940
      Width           =   3585
      Begin VB.CheckBox Check1 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Todos los Tipos"
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
         Left            =   210
         TabIndex        =   21
         Top             =   525
         Value           =   1  'Checked
         Width           =   1695
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
         Left            =   3150
         TabIndex        =   18
         Top             =   525
         Width           =   175
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
         Left            =   2730
         TabIndex        =   17
         Top             =   525
         Width           =   435
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
         Left            =   210
         TabIndex        =   20
         Top             =   945
         Width           =   3135
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
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
         Left            =   1890
         TabIndex        =   19
         Top             =   525
         Width           =   750
      End
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   8610
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   105
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00EEEEC0&
      Caption         =   "Controles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   3885
      TabIndex        =   12
      Top             =   3565
      Width           =   3165
      Begin VB.CheckBox Check2 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Verificar Movimientos y Saldos de Stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   420
         TabIndex        =   29
         Top             =   210
         Width           =   2220
      End
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
      Left            =   210
      TabIndex        =   5
      Top             =   210
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
         TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   525
         Width           =   2325
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EEEEC0&
      Caption         =   "Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   210
      TabIndex        =   4
      Top             =   3565
      Width           =   3585
      Begin VB.OptionButton Option12 
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
         Height          =   240
         Left            =   1890
         TabIndex        =   25
         Top             =   420
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.OptionButton Option11 
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
         Height          =   240
         Left            =   525
         TabIndex        =   24
         Top             =   420
         Width           =   1170
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00EEEEC0&
      Caption         =   "Filtrado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   3885
      TabIndex        =   3
      Top             =   1940
      Width           =   3165
      Begin VB.OptionButton Option4 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Productos Terminados"
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
         Left            =   420
         TabIndex        =   28
         Top             =   1125
         Width           =   2535
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Todos los Artículos"
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
         Left            =   420
         TabIndex        =   26
         Top             =   855
         Width           =   2535
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Artículos con Movimiento"
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
         Left            =   420
         TabIndex        =   23
         Top             =   585
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Solo Artículos Faltantes"
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
         Left            =   420
         TabIndex        =   22
         Top             =   315
         Value           =   -1  'True
         Width           =   2535
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
      Height          =   800
      Left            =   7245
      Picture         =   "DISPOSTO.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1155
      Width           =   855
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
      Height          =   585
      Left            =   7245
      Picture         =   "DISPOSTO.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Emite Reporte Solicitado"
      Top             =   3200
      Width           =   855
   End
   Begin VB.CommandButton command10 
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
      Height          =   800
      Left            =   7245
      Picture         =   "DISPOSTO.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Abandona la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   6680
      Picture         =   "DISPOSTO.frx":0BA0
      Top             =   3960
      Width           =   2010
   End
End
Attribute VB_Name = "DISPOSTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IAT&
'

Private Sub Check1_Click()
   Text3.SetFocus
End Sub

Private Sub Check1_LostFocus()
   If Check1.Value = 1 Then
      Text3.TEXT = ""
      Label4.Caption = "Todos - Sin Clasificación"
   End If
End Sub

Private Sub Command1_Click()
    Call SELECHO("MASTER")
    cod1$ = TRIM$(VALACT1$("MASTER"))
    If cod1$ <> "" Then
       Text1.TEXT = cod1$
    End If
End Sub

Private Sub Command10_Click()
   Call CIERRARCH("*.*")
   Unload DISPOSTO
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   Call STODISP
   Command12.Enabled = True
End Sub

Private Sub Command2_Click()
    Call SELECHO("MASTER")
    cod2$ = TRIM$(VALACT1$("MASTER"))
    If cod2$ <> "" Then
       Text2.TEXT = cod2$
    End If
End Sub

Private Sub Command26_Click()
   Command26.Enabled = False
   If (USNBR% Mod 100) = 1 Then
      On Error Resume Next
      OPDISTOK.Show 1
      On Error GoTo 0
   End If
   Command26.Enabled = True
End Sub

Private Sub Command3_Click()
    Call SELECHO("TATIMAT")
    VDEV$ = VALACT1$("TATIMAT")
    If VDEV$ <> "" Then
       Text3.TEXT = VDEV$
       If TRIM$(Text3.TEXT) = "" Then
            Label4.Caption = "Todos - Sin Clasificación"
            Check1.Value = 1
          Else
            Label4.Caption = VALACT2$("TATIMAT")
            Check1.Value = 0
       End If
    End If
End Sub

Private Sub Command4_Click()
End Sub

Private Sub Command5_Click()
End Sub

Private Sub Command6_Click()
End Sub


Private Sub Form_Load()
    '
    Call TRANSLABELS(Name)
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16)
    Text1.TEXT = XX$
    CI1% = Abs(CODINT%(XX$))
    Label1.Caption = DESCRIT$(CI1%)
    '
    AAA& = ULTREG&("ECOMAST")
10  XX$ = Left$(REGLEIDO$("ECOMAST", AAA&), 16)
    If TRIM$(XX$) = "" Then
      If AAA& > 1 Then
        AAA& = AAA& - 1
        GoTo 10
      End If
    End If
    Text2.TEXT = XX$
    CI2% = Abs(CODINT%(XX$))
    Label2.Caption = DESCRIT$(CI2%)
    '
    Call VERNOVTRA
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
End Sub


Private Sub Text1_DblClick()
    Call SELECHO("MASTER")
    cod1$ = TRIM$(VALACT1$("MASTER"))
    If cod1$ <> "" Then
       Text1.TEXT = cod1$
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

Private Sub Text2_DblClick()
    Call SELECHO("MASTER")
    cod2$ = TRIM$(VALACT1$("MASTER"))
    If cod2$ <> "" Then
       Text2.TEXT = cod2$
    End If
End Sub

Private Sub Text3_DblClick()
    Call SELECHO("TATIMAT")
    VDEV$ = VALACT1$("TATIMAT")
    If VDEV$ <> "" Then
       Text3.TEXT = VDEV$
       If TRIM$(Text3.TEXT) = "" Then
            Label4.Caption = "Todos - Sin Clasificación"
            Check1.Value = 1
          Else
            Label4.Caption = VALACT2$("TATIMAT")
            Check1.Value = 0
       End If
    End If
End Sub

Private Sub Text3_GotFocus()
   If TRIM$(Text3.TEXT) = "" Then
      Check1.Value = 1
      Label4.Caption = "Todos - Sin Clasificación"
   End If
End Sub

Private Sub Text3_LostFocus()
    If Text3.TEXT <> "" Then
      If Val(Text3.TEXT) < 256 Then
        VENDE% = Val(Text3.TEXT)
        VENDX$ = ECOARCH$("TATIMAT", Chr$(VENDE%))
        If VENDX <> "" Then
            Label4.Caption = VENDX$
            Check1.Value = 0
            Exit Sub
        End If
      End If
      Text3.TEXT = ""
      Label4.Caption = "Todos - Sin Clasificación"
      Check1.Value = 1
    End If
End Sub


Private Sub Text1_Change()
    CI1% = CODINT%(Text1.TEXT)
    If CI1% > 0 Then
         Label1.Caption = DESCRIT$(CI1%)
       Else
         Label1.Caption = ""
    End If
End Sub

Private Sub Text3_Change()
    If Val(Text3.TEXT) < 256 Then
        VENDE% = Val(Text3.TEXT)
        VENDX$ = ECOARCH$("TATIMAT", Chr$(VENDE%))
        If VENDX <> "" Then
            Label4.Caption = VENDX$
            Check1.Value = 0
            Exit Sub
          Else
            Label4.Caption = "Todos - Sin Clasificación"
            Check1.Value = 1
        End If
    End If
End Sub

Private Sub Text8_Change()
End Sub

Sub STODISP()
    '
    HOII% = HOY(HOS$)
    IAT& = 0            ' INDICE ARCHIVO TRANSFERENCIA
    '
    VECOPEN$ = PUNOCPEN$
    '
    Call CARDEPOS
    '
    TMT% = Val(DISPOSTO.Text3.TEXT)
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TIMAT$ = "" Then
        TMT% = 0
    End If
    '
    CI1% = CODINT%(DISPOSTO.Text1.TEXT)
    cod1$ = CODEXT$(CI1%)
    If TRIM$(cod1$) = "" Then Exit Sub
    CI2% = CODINT%(DISPOSTO.Text2.TEXT)
    cod2$ = CODEXT$(CI2%)
    If TRIM$(cod2$) = "" Then Exit Sub
    '
    NIVEFIL% = 1
    If DISPOSTO.Option2.Value = True Then NIVEFIL% = 2
    If DISPOSTO.Option1.Value = True Then NIVEFIL% = 3
    If DISPOSTO.Option4.Value = True Then NIVEFIL% = 4
    '
    Screen.MousePointer = 11
    Dim EXIS(16384, 6)
    Dim PUNOF$(16384)
    Dim PROTER%(16384)      ' MARCA PRODUCTO TERMINADO
    Call CAROEMI((-32767), 0, 0) ' PARA INICIALIZAR
    '
    If DISPOSTO.Check2.Value = 1 Then
       '
       UREMAS& = ULTREG&("MASTER")
       USER% = USNBR% Mod 100
       If USER% < 1 Then USER% = 1
       '
       While ULTREG&("SALSTOCK") < UREMAS&
         X$ = String$(39, 0) + Chr$(1)
         Call REGAPP("SALSTOCK", X$)
       Wend
       '
       While ULTREG&("SALDEPOS") < UREMAS&
         Y$ = REGLEIDO$("SALSTOCK", 1 + ULTREG&("SALDEPOS"))
         X$ = String$(256, 0)
         Call REPLA(X$, Y$, 1, 40)
         Call REPLA(X$, MKI$(HOII%), 253, 2)
         Call REPLA(X$, MKI$(HORANUM(Time$)), 255, 2)
         Call REGAPP("SALDEPOS", X$)
       Wend
       '
       Call CIERRARCH("SALSTOCK")   ' para liberar bloqueos
       Call CIERRARCH("SALDEPOS")   ' para liberar bloqueos
       '
       For U& = 1 To UREMAS&
         Call TRACE(20, U&, UREMAS&)
         X$ = String$(256, 0)
         Call REPLA(X$, MKI$(HOII%), 253, 2)
         Call REPLA(X$, MKI$(HORANUM(Time$)), 255, 2)
         Call GRAREG("!SALDEPOS", X$, U&)
       Next U&
       '
       FIN& = ULTREG&("MOVISTO")
       '
       For I& = 1 To FIN&
         Call TRACE(20, I&, FIN&)
         MVS$ = REGLEIDO$("MOVISTO", I&)
         '
         CI% = CVI(Mid$(MVS$, 3, 2))
         DEP% = Asc(Mid$(MVS$, 83, 1))
         SUCX% = IDEPOS%(DEP%)
         CANTI = CVD(Mid$(MVS$, 88, 8)) - CVD(Mid$(MVS$, 96, 8))
         '
         If CI% > 0 Then
           If CI% <= NUMAS% Then
             RGG& = CI% + 1: PX% = 4 * SUCX% - 3
             If PX% > 0 Then
               If PX% <= 245 Then
                 STRI$ = REGLEIDO$("!SALDEPOS", RGG&)
                 VALO# = CDbl(CANTI)
                 Call RECUPD(STRI$, PX%, 4, VALO#)
                 Call GRAREG("!SALDEPOS", STRI$, RGG&)
               End If
             End If
           End If
         End If
       Next I&
       '
       For I& = 1 To UREMAS&
         Call TRACE(20, I&, UREMAS&)
         X$ = REGLEIDO$("!SALDEPOS", I&)
         Call GRAREG("SALSTOCK", Left$(X$, 40), I&)
         Call GRAREG("SALDEPOS", X$, I&)
       Next I&
       '
       Call CIERRARCH("SALSTOCK")
       Call CIERRARCH("SALDEPOS")
       '
    End If
    '
    FIN& = NUMAS%
    For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       SST$ = REGLEIDO$("SALSTOCK", I& + 1) + Mid$(REGLEIDO$("SALDEPOS", I& + 1), 41)
       SSTT = 0
       For KI% = 1 To NUSU%
          K% = 4 * KI% - 3
          If DISTO%(KI%) >= 0 Then
            SSTT = SSTT + CVS(Mid$(SST$, K%, 4))
          End If
       Next KI%
       EXIS(I&, 1) = SSTT
       CIIX% = I&
       Call CARTRAST(CIIX%, "ST", "Stock al " + HOS$ + " - " + Time$, "", SSTT, HOII%)
       EXIS(I&, 4) = STOMIN(CIIX%)
    Next I&
    '
    FIN& = ULTREG&("OCOMPRA")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("OCOMPRA", U&)
       STAT% = Asc(Mid$(X$, 62, 1))
       If STAT% <= 0 Then
          CATOT# = CVS(Mid$(X$, 13, 4))
          CAPRA# = CVS(Mid$(X$, 58, 4))
          CAPEN = CATOT# - CAPRA#
          If CAPEN > 0 Then
             CIIX% = CVI(Mid$(X$, 11, 2))
             NC% = CVI(Mid$(X$, 7, 2))
             COMPROB$ = "OC." + TRIM$(Str$(CVS(Left$(X$, 4))))
             FENTRESOL% = CVI(Mid$(X$, 9, 2))
             EXIS(CIIX%, 2) = EXIS(CIIX%, 2) + CAPEN
             Call CARTRAST(CIIX%, "OC", RASOCLI$((-1) * NC%), COMPROB$, CAPEN, FENTRESOL%)
          End If
       End If
    Next U&
    '
    FIN& = ULTREG&("OCOMDETA")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("OCOMDETA", U&)
      STAT% = Asc(Mid$(X$, 124, 1))
      If STAT% < 8 Then
        NUORSE& = CVS(Mid$(X$, 117, 4))
        If NUORSE& < 1 Then        ' no es orden de servicio
          COEFCAN = CVS(Mid$(X$, 63, 4))
          If COEFCAN < 0.005 Then COEFCAN = 1
          CATOT# = CVS(Mid$(X$, 55, 4)) * COEFCAN
          CAPRA# = CVS(Mid$(X$, 99, 4)) * COEFCAN
          CAPEN = CATOT# - CAPRA#
          If CAPEN > 0 Then
             CIIX% = CVI(Mid$(X$, 9, 2))
             NC% = CVI(Mid$(X$, 7, 2))
             COMPROB$ = "OC." + TRIM$(Str$(CVS(Left$(X$, 4))))
             FENTRESOL% = CVI(Mid$(X$, 95, 2))
             EXIS(CIIX%, 2) = EXIS(CIIX%, 2) + CAPEN
             Call CARTRAST(CIIX%, "OC", RASOCLI$((-1) * NC%), COMPROB$, CAPEN, FENTRESOL%)
          End If
        End If
      End If
    Next U&
    '
    FIN& = ULTREG&("BOLREPEN")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("BOLREPEN", U&)
       CAPEN = CVD(Mid$(X$, 89, 8))
       If CAPEN > 0 Then
          CIIX% = CVI(Mid$(X$, 83, 2))
          NC% = CVI(Mid$(X$, 105, 2))
          COMPROB$ = "BR." + TRIM$(Str$(CVS(Left$(X$, 4))))
          FENTRESOL% = CVI(Mid$(X$, 81, 2))
          EXIS(CIIX%, 2) = EXIS(CIIX%, 2) + CAPEN
          Call CARTRAST(CIIX%, "BR", "RT." + TRIM$(Mid$(X$, 119, 16)) + " - " + RASOCLI$((-1) * NC%), COMPROB$, CAPEN, FENTRESOL%)
       End If
    Next U&
    '
    FIN& = ULTREG&("ORDEFABR")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("ORDEFABR", U&)
       STAT% = Asc(Mid$(X$, 108, 1))
       If STAT% > 0 Then
          If STAT% < 3 Then
             CATOT# = CVD(Mid$(X$, 89, 8))
             CAPRA# = CVD(Mid$(X$, 119, 8))
             CAPEN = CATOT# - CAPRA#
             If CAPEN > 0 Then
                CIIX% = CVI(Mid$(X$, 83, 2))
                REFE$ = Mid$(X$, 5, 64)
                COMPROB$ = "OF." + TRIM$(Str$(CVS(Left$(X$, 4))))
                FENTRESOL% = CVI(Mid$(X$, 105, 2))
                EXIS(CIIX%, 3) = EXIS(CIIX%, 3) + CAPEN
                Call CARTRAST(CIIX%, "OF", REFE$, COMPROB$, CAPEN, FENTRESOL%)
                PUNOF$(CIIX%) = PUNOF$(CIIX%) + MKS$(U&)
             End If
          End If
       End If
    Next U&
    '
    APEDPE$ = "PEDDETA": MODOMAI% = 0
    If CONTROL$("ADMIPED", "FORMACAR") = "ITEMIZADO" Then
       MODOMAI% = 1: APEDPE$ = "PEDCLIEN"
    End If
    FIN& = ULTREG&(APEDPE$)
    '
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$(APEDPE$, U&)
       COMPROB$ = "PC." + TRIM$(Str$(CVS(Left$(X$, 4))))
       NC% = CVI(Mid$(X$, 7, 2))
       If UCase$(Mid$(X$, 51, 1)) = "R" Then GoTo 19 ' SALTEA REPARACIONES
       '
       CAPEN = 0
       If MODOMAI% <> 1 Then
            CIIX% = CVI(Mid$(X$, 33, 2))
            CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
            FENTRESOL% = CVI(Mid$(X$, 25, 2))
            If Asc(Mid$(X$, 27, 1)) > 1 Then CAPEN = 0 ' Cumplido o Anulado
          Else
            CIIX% = CVI(Mid$(X$, 640 + 3, 2))
            Status% = Asc(Mid$(X$, 300, 1))
            If Status% <= 1 Then
               CACOPED = CVS(Mid$(X$, 513, 4))
               CAITPED = CVS(Mid$(X$, 645, 4))
               CACOENT = CVS(Mid$(X$, 512 + 77, 4))
               CAPEN = (CACOPED - CACOENT) * CAITPED
               FENTRESOL% = CVI(Mid$(X$, 512 + 63, 2))
            End If
       End If
       '
       If CAPEN > 0 Then
          If Abs(EXIS(CIIX%, 0)) <> 1 Then
             EXIS(CIIX%, 0) = DESCOMPO%(CIIX%)
             If EXIS(CIIX%, 0) <> 1 Then EXIS(CIIX%, 0) = (-1)
          End If
          '
          EXIS(CIIX%, 5) = EXIS(CIIX%, 5) + CAPEN
          Call CARTRAST(CIIX%, "PC", RASOCLI$(NC%), COMPROB$, (-1) * CAPEN, FENTRESOL%)
          '
          If EXIS(CIIX%, 0) = 1 Then ' items descompuestos
               EXIS(CIIX%, 3) = EXIS(CIIX%, 3) + CAPEN
               Call LEEEXPLO(CODEXT$(CIIX%), 1)
               For KKI% = 1 To CAIT%
                  EXIS(CIJ%(KKI%), 6) = EXIS(CIJ%(KKI%), 6) + CAPEN * USOI(KKI%)
               Next KKI%
          End If
       End If
19  Next U&
    '
    FIN& = ULTREG&("RESERVA")
    For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       XX$ = REGLEIDO$("RESERVA", I&)
       CIIX% = CVI(Left$(XX$, 2))
       TORE# = CVD(Mid$(XX$, 19, 8))
       CACO# = CVD(Mid$(XX$, 27, 8))
       '
       If UCase$(UNIMED$(CIIX%)) = "GK" Then
          TORE# = TORE# / 1000
          CACO# = CACO# / 1000
       End If
       '
       SIGNO% = Sgn(CVS(Mid$(XX$, 75, 4)))
       SARE = TORE# - CACO#
       If SARE * SIGNO% < 0 Then
          SARE = 0
       End If
       EXIS(CIIX%, 6) = EXIS(CIIX%, 6) + SARE
       If SARE >= 0.05 Then
          COMPROB$ = Mid$(XX$, 3, 12)
          REFE$ = "Reserva Fabric." + COMPROB$
          FENTRESOL% = CVI(Mid$(XX$, 17, 2))
          Call CARTRAST(CIIX%, "RF", REFE$, COMPROB$, (-1) * Abs(SARE), FENTRESOL%)
       End If
    Next I&
    '
    FIN& = ULTREG&("REPUREPA")
    For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       XX$ = REGLEIDO$("REPUREPA", I&)
       CIIX% = CVI(Mid$(XX$, 7, 2))
       TORE# = CVS(Mid$(XX$, 21, 4))
       CACO# = CVS(Mid$(XX$, 25, 4))
       '
       If UCase$(UNIMED$(CIIX%)) = "GK" Then
          TORE# = TORE# / 1000
          CACO# = CACO# / 1000
       End If
       '
       SARE = TORE# - CACO#
       If SARE < 0.05 Then GoTo 29
       '
       NOREPARA& = CVS(Left$(XX$, 4))
       REREP& = REGBUS&("PRESREPA", 1, MKS$(NOREPARA&))
       If REREP& > 0 Then
         YY$ = REGLEIDO$("PRESREPA", REREP&)
         If CVI(Mid$(YY$, 123, 2)) > 0 Then GoTo 29 ' O/Rep Cerrada
         If CVS(Mid$(YY$, 125, 4)) < 1 Then GoTo 29 ' falta O/Venta
       End If
       '
       EXIS(CIIX%, 6) = EXIS(CIIX%, 6) + SARE
       If SARE >= 0.05 Then
          COMPROB$ = Mid$(XX$, 3, 12)
          REFE$ = "Reserva Fabric." + COMPROB$
          FENTRESOL% = CVI(Mid$(XX$, 17, 2))
          Call CARTRAST(CIIX%, "RF", REFE$, COMPROB$, (-1) * Abs(SARE), FENTRESOL%)
       End If
29  Next I&
    '
    JJJ& = 0
    FIN& = ULTREG&("MACONSIG")
    For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       XX$ = REGLEIDO$("MACONSIG", I&)
       CIIX% = CVI(Mid$(XX$, 33, 2))
       TORE# = CVD(Mid$(XX$, 55, 8))
       CACO# = CVD(Mid$(XX$, 63, 8))
       '
       If CACO# < 0.005 Then
          NUOCO& = Val(Mid$(XX$, 24, 6))
          For TYU& = 1 To Len(VECOPEN$) Step 4
            If Abs(CVS(Mid$(VECOPEN$, TYU&, 4)) - NUOCO&) < 0.1 Then GoTo 42
          Next TYU&
          GoTo 49  ' FILTRA PARA O/C's PENDIENTES
       End If
       '
42     JJJ& = JJJ& + 1
       Call GRAREG("MACONSIG", XX$, JJJ&)
       '
       If UCase$(UNIMED$(CIIX%)) = "GK" Then
          TORE# = TORE# / 1000
          CACO# = CACO# / 1000
       End If
       '
       SARE = TORE# - CACO#
       If SARE < 0 Then
          SARE = 0
       End If
       If SARE >= 0.05 Then
          COMPROB$ = Mid$(XX$, 21, 12)
          NUOCO& = Val(Mid$(XX$, 24, 6))
          For TYU& = 1 To Len(VECOPEN$) Step 4
            If Abs(CVS(Mid$(VECOPEN$, TYU&, 4)) - NUOCO&) < 0.1 Then GoTo 45
          Next TYU&
          GoTo 49  ' FILTRA PARA O/C's PENDIENTES
          '
45        EXIS(CIIX%, 6) = EXIS(CIIX%, 6) + SARE
          REFE$ = "Consig." + RASOCLI$((-1) * CVI(Mid$(XX$, 3, 2)))
          FENTRESOL% = CVI(Mid$(XX$, 1, 2))
          Call CARTRAST(CIIX%, "RF", REFE$, COMPROB$, (-1) * Abs(SARE), FENTRESOL%)
       End If
49  Next I&
    Call SETULTREG("MACONSIG", JJJ&)
    Call CIERRARCH("MACONSIG")
    '
    FIN& = ULTREG&("STOREMOT")
    For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       XX$ = REGLEIDO$("STOREMOT", I&)
       CIIX% = CVI(Left$(XX$, 2))
       If CIIX% > 0 Then
         If CIIX% <= NUMAS% Then
           TIMO$ = UCase$(Mid$(XX$, 15, 2))
           TORE# = CVD(Mid$(XX$, 67, 8))
           CACO# = CVD(Mid$(XX$, 75, 8))
           SIGNO% = 1: COLUST% = 1
           SARE = TORE# - CACO#
           If Abs(SARE) >= 0.05 Then
             If TIMO$ = "OC" Then COLUST% = 2
             If TIMO$ = "BR" Then COLUST% = 2
             If TIMO$ = "OF" Then COLUST% = 3
             If TIMO$ = "PC" Then COLUST% = 5: SIGNO% = (-1)
             If TIMO$ = "RF" Then COLUST% = 6: SIGNO% = (-1)
             EXIS(CIIX%, COLUST%) = EXIS(CIIX%, COLUST%) + SARE * SIGNO%
           End If
         End If
       End If
    Next I&
    Call CIERRARCH("STOREMOT")
    '
    If NIVEFIL% = 4 Then GoTo 100 ' situacion productos terminados
    '
    J& = 0
    NVII& = NUINV%
    For I& = 1 To NVII&
        Call TRACE(20, I&, NVII&)
        II$ = REGLEIDO$("INVERT", I& + 1)
        CIIX% = CVI(Mid$(II$, 17, 2))
        If Mid$(II$, 16, 1) <> "`" Then
          If CIIX% > 0 Then
            If CIIX% <= NUMAS% Then
              CODE$ = CODEXT$(CIIX%)
              If CODE$ >= cod1$ Then
               If CODE$ <= cod2$ Then
                If TMT% = 0 Or TIMATE%(CIIX%) = TMT% Then
                  X$ = REGBLAN$("STODISP")
                  Call REPLA(X$, MKI$(CIIX%), 1, 2)
                  Call REPLA(X$, Chr$(TIMATE%(CIIX%)), 3, 1)
                  SUMA = 0: HAYMO% = 0
                  For K% = 1 To 6
                    SIG% = 1: If K% >= 4 Then SIG% = (-1)
                    SUMA = SUMA + SIG% * EXIS(CIIX%, K%)
                    If Abs(EXIS(CIIX%, K%)) >= 0.05 Then HAYMO% = 1
                    Call REPLA(X$, MKS$(EXIS(CIIX%, K%)), 4 * K% + 1, 4)
                  Next K%
                  If SUMA > 0.05 Then
                      Call REPLA(X$, MKS$(SUMA), 29, 4)
                    ElseIf SUMA < (-0.05) Then
                      Call REPLA(X$, MKS$((-1) * SUMA), 33, 4)
                      Call CAROEMI(CIIX%, Abs(SUMA), HOII%)
                  End If
                  '
                  If (NIVEFIL% = 1) Or (NIVEFIL% = 2 And HAYMO% = 1) Or (NIVEFIL% = 3 And SUMA <= (-0.05)) Then
                    Call REPLA(X$, MKS$(LOTESTAN(CIIX%)), 37, 4)
                    J& = J& + 1
                    Call GRAREG("STODISP", X$, J&)
                  End If
                End If
               End If
              End If
            End If
          End If
        End If
    Next I&
    '
    Call SETULTREG("STODISP", J&)
    Call CIERRARCH("STODISP")
    '
    Call SETULTREG("ATRASTOK", IAT&)
    XX$ = REGLEIDO$("ATRASTOK", 1)
    Call REPLA(XX$, HOS$, 17, 8)
    Call REPLA(XX$, Time$, 25, 8)
    Call GRAREG("ATRASTOK", XX$, 1)
    Call CIERRARCH("ATRASTOK")
    '
    Call CIERRARCH("*.*")
    '
    Call CARUTRA
    AAI% = 1
    If TRIM$(IDENSEDE$) <> "" Then
       If TRIM$(DESTITRA$) <> "" Then
          AAI% = COPYFILE%(LUDAT$ + "ATRASTOK.DAT", DESTITRA$)
       End If
    End If
    Screen.MousePointer = 1
    If AAI% <> 1 Then Exit Sub
    '
    If DISPOSTO.Option11.Value = True Then
         Call CONECRUN("*STODISC", "")
       ElseIf DISPOSTO.Option12.Value = True Then
         Call CONECRUN("*STODISL", "")
    End If
    Exit Sub
    '
100 FIN& = ULTREG&("ESTRUNUE")
    FINDOS& = ULTREG&("ESTRUDOS")
    For REG& = 1 To FIN&
       Call TRACE(20, REG&, FIN&)
       X$ = REGLEIDO$("ESTRUNUE", REG&)
       CIIX% = CVI(Left$(X$, 2))
       '
       If CIIX% > 0 Then
          If CIIX% <= NUMAS% Then
             If PROTER%(CIIX%) = 0 Then PROTER%(CIIX%) = 1
             '
             DPS$ = Mid$(X$, 3, 120)
             If REG& <= FINDOS& Then
                DPS$ = DPS$ + REGLEIDO$("ESTRUDOS", REG&)
             End If
             '
             For K% = 1 To Len(DPS$) Step 8
                DPSX$ = Mid$(DPS$, K%, 8)
                CIKK% = CVI(Left$(DPSX$, 2))
                If CIKK% > 0 Then
                  If CIKK% <= 32767 Then
                    If Abs(CVS(Mid$(DPSX$, 3))) >= 0.000005 Then
                      If PROTER%(CIIX%) = 0 Then PROTER%(CIIX%) = 1
                      PROTER%(CIKK%) = (-1)
                    End If
                  End If
                End If
             Next K%
          End If
       End If
       '
    Next REG&
    '
    J& = 0
    NVII& = NUINV%
    For I& = 1 To NVII&
        Call TRACE(20, I&, NVII&)
        II$ = REGLEIDO$("INVERT", I& + 1)
        CIIX% = CVI(Mid$(II$, 17, 2))
        If Mid$(II$, 16, 1) <> "`" Then
          If CIIX% > 0 Then
            If CIIX% <= NUMAS% Then
              If PROTER%(CIIX%) = 1 Then
                X0$ = REGBLAN$("SIPROTER")
                Call REPLA(X0$, MKI$(CIIX%), 1, 2)
                Call REPLA(X0$, MKS$(EXIS(CIIX%, 1)), 3, 4)
                If Len(PUNOF$(CIIX%)) < 4 Then
                     J& = J& + 1
                     Call GRAREG("SIPROTER", X0$, J&)
                   Else
                     For KK% = 1 To Len(PUNOF$(CIIX%)) Step 4
                       U& = CVS(Mid$(PUNOF$(CIIX%), KK%, 4))
                       X$ = REGLEIDO$("ORDEFABR", U&)
                       CATOF = CVD(Mid$(X$, 89, 8))
                       CAPEN = CATOF - CVD(Mid$(X$, 119, 8))
                       If CAPEN < 0 Then CAPEN = 0
                       COMPROB$ = "OF." + TRIM$(Str$(CVS(Left$(X$, 4))))
                       X1$ = X0$
                       Call REPLA(X1$, COMPROB$, 7, 12)
                       Call REPLA(X1$, MKS$(CATOF), 19, 4)
                       Call REPLA(X1$, MKS$(CAPEN), 23, 4)
                       J& = J& + 1
                       Call GRAREG("SIPROTER", X1$, J&)
                     Next KK%
                End If
              End If
            End If
          End If
        End If
    Next I&
    '
    Call SETULTREG("SIPROTER", J&)
    Call CIERRARCH("SIPROTER")
    '
    Call CONECRUN("*SIPROTE", "Situacion de Stocks Terminados")
    '
End Sub
'
Sub CARTRAST(CIIY%, COMO$, REMO$, NUCO$, CATRA, FEPRO%)
   '
   If IAT& < 1 Then
      If RECLEN%("ATRASTOK") < 256 Then
         Call COMUNI("Error en Controlador de Archivo 'ATRASTOK.RFS'.\  \Consulte a su Soporte de Sistemas.")
         Call FINAL("")
      End If
      HOII% = HOY(HOS$)
      Call CARUTRA
      XX$ = Space$(128) + String$(128, 0)
      Call REPLA(XX$, IDENSEDE$, 1, 8)
      Call REPLA(XX$, "HEADER", 9, 8)
      Call REPLA(XX$, HOS$, 17, 8)
      Call REPLA(XX$, Time$, 25, 8)
      Call REPLA(XX$, USERNAME$, 33, 32)
      Call REPLA(XX$, NOMTER$, 65, 32)
      Call REPLA(XX$, Space$(32), 97, 32)
      Call GRAREG("ATRASTOK", XX$, 1)
      IAT& = 1
   End If
   '
   If Abs(CATRA) >= 0.05 Then
      '
      XX$ = REGBLAN$("ATRASTOK")
      Call REPLA(XX$, IDENSEDE$, 1, 8)
      Call REPLA(XX$, "DISPOREM", 9, 8)
      Call REPLA(XX$, CODEXT$(CIIY%), 17, 16)
      Call REPLA(XX$, COMO$, 33, 2)
      Call REPLA(XX$, REMO$, 35, 50)
      Call REPLA(XX$, NUCO$, 85, 12)
      If CATRA > 0 Then
           Call REPLA(XX$, MKD$(CDbl(CATRA)), 97, 8)
         Else
           Call REPLA(XX$, MKD$(CDbl(Abs(CATRA))), 105, 8)
      End If
      Call REPLA(XX$, MKI$(FEPRO%), 113, 2)
      '
      IAT& = IAT& + 1
      Call GRAREG("ATRASTOK", XX$, IAT&)
      '
   End If
   '
End Sub
'
Sub CAROEMI(CIII%, CANOR, FEOR%)
   '
   Static CTX%, JCOM&, JFAB&
   Static ACON%(32767)
   '
   If CTX% = 0 Then
     If EXISTE%("PROMRP02.EXE") > 0 Then
       CTX% = (-1)
     End If
   End If
   '
   If CTX% = (-1) Then  ' lo hace por MRP
     Exit Sub
   End If
   '
   If CTX% = 0 Or CIII% = (-32767) Then
     FIN& = ULTREG&("ESTRUNUE")
     For REG& = 1 To FIN&
       Call TRACE(20, REG&, FIN&)
       X$ = REGLEIDO$("ESTRUNUE", REG&)
       CIIQ% = CVI(Left$(X$, 2))
       If CIIQ% > 0 Then
         If CIIQ% <= 32767 Then
           If ACON%(CIIQ%) = 0 Then
             DPSX$ = Mid$(X$, 3, 8)
             CIKK% = CVI(Left$(DPSX$, 2))
             If CIKK% > 0 Then
               If CIKK% <= 32767 Then
                 If Abs(CVS(Mid$(DPSX$, 3))) >= 0.000005 Then
                   ACON%(CIIQ%) = 1
                 End If
               End If
             End If
           End If
         End If
       End If
     Next REG&
     CTX% = 1
     JCON& = 0: Call SETULTREG("OCOMEMI", JCON&)
     JFAB& = 0: Call SETULTREG("ORFAEMI", JFAB&)
   End If
   '
   If CIII% < 1 Then Exit Sub
   '
   FF1$ = REGBLAN$("ORFAEMI")
   Call REPLA(FF1$, MKI$(CIII%), 1, 2)
   Call REPLA(FF1$, DESCRIT$(CIII%), 3, 26)
   Call REPLA(FF1$, CSTRING$(MKS$(CANOR), 4, 10, 1, 2), 29, 10)
   FEX = FEOR%
   Call REPLA(FF1$, FECHATEX$(FEX), 41, 8)
   Call REPLA(FF1$, MKD$(CDbl(CANOR)), 49, 8)
   Call REPLA(FF1$, MKI$(FEOR%), 57, 2)
   Call REPLA(FF1$, MKI$(CIII%), 59, 2)
   '
   If ACON(CIII%) = 1 Then
         JFAB& = JFAB& + 1
         Call GRAREG("ORFAEMI", FF1$, JFAB&)
      Else
         JCOM& = JCOM& + 1
         Call GRAREG("OCOMEMI", FF1$, JCOM&)
   End If
   '
End Sub


