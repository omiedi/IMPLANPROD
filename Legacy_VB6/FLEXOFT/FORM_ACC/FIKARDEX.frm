VERSION 5.00
Begin VB.Form FIKARDEX 
   BackColor       =   &H00FEFEE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ficha de Movimientos de Material (Ficha Kardex)"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture6 
      Height          =   1670
      Left            =   10100
      ScaleHeight     =   1605
      ScaleWidth      =   1545
      TabIndex        =   39
      Top             =   0
      Visible         =   0   'False
      Width           =   1600
      Begin VB.Image Image3 
         Height          =   1750
         Left            =   0
         Picture         =   "FIKARDEX.frx":0000
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   1650
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FEFEE0&
      Caption         =   "Opción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   4830
      TabIndex        =   35
      Top             =   630
      Visible         =   0   'False
      Width           =   950
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FEFEE0&
         Caption         =   "Lote"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   37
         Top             =   630
         Width           =   680
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FEFEE0&
         Caption         =   "D-Sec"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   36
         Top             =   420
         Value           =   -1  'True
         Width           =   680
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3950
      Index           =   4
      Left            =   7455
      TabIndex        =   34
      Top             =   1890
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3950
      Index           =   5
      Left            =   7900
      TabIndex        =   33
      Top             =   1890
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3950
      Index           =   6
      Left            =   9070
      TabIndex        =   32
      Top             =   1890
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3950
      Index           =   7
      Left            =   10120
      TabIndex        =   31
      Top             =   1890
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3950
      Index           =   8
      Left            =   11360
      TabIndex        =   30
      Top             =   1890
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3950
      Index           =   3
      Left            =   4500
      TabIndex        =   29
      Top             =   1890
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3950
      Index           =   2
      Left            =   3045
      TabIndex        =   28
      Top             =   1890
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3950
      Index           =   1
      Left            =   1525
      TabIndex        =   27
      Top             =   1890
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3950
      Index           =   0
      Left            =   1080
      TabIndex        =   26
      Top             =   1890
      Width           =   15
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FEFEE0&
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   105
      TabIndex        =   21
      Top             =   630
      Width           =   5640
      Begin VB.Label Label6 
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
         Height          =   495
         Left            =   210
         TabIndex        =   25
         Top             =   360
         Width           =   5265
      End
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   3
      Left            =   3045
      Picture         =   "FIKARDEX.frx":2398
      TabIndex        =   19
      Top             =   210
      Width           =   175
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   2
      Left            =   6615
      Picture         =   "FIKARDEX.frx":26A2
      TabIndex        =   18
      Top             =   210
      Width           =   175
   End
   Begin VB.TextBox Text4 
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
      Left            =   6090
      TabIndex        =   17
      Text            =   " "
      Top             =   210
      Width           =   540
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FEFEE0&
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
      Height          =   1020
      Left            =   5880
      TabIndex        =   10
      Top             =   630
      Width           =   4120
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FEFEE0&
         Caption         =   "Incluir Antiguos"
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
         Left            =   2310
         TabIndex        =   24
         Top             =   280
         Width           =   1695
      End
      Begin VB.CommandButton BOTSEL 
         Caption         =   "..."
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
         Index           =   1
         Left            =   3780
         Picture         =   "FIKARDEX.frx":29AC
         TabIndex        =   14
         Top             =   600
         Width           =   175
      End
      Begin VB.CommandButton BOTSEL 
         Caption         =   "..."
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
         Index           =   0
         Left            =   1785
         Picture         =   "FIKARDEX.frx":2CB6
         TabIndex        =   13
         Top             =   600
         Width           =   175
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   12
         Text            =   " "
         Top             =   600
         Width           =   960
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2835
         TabIndex        =   11
         Text            =   " "
         Top             =   600
         Width           =   960
      End
      Begin VB.Label Label10 
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
         Left            =   -105
         TabIndex        =   16
         Top             =   630
         Width           =   855
      End
      Begin VB.Label Label11 
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
         Left            =   1995
         TabIndex        =   15
         Top             =   630
         Width           =   750
      End
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
      Height          =   285
      Left            =   1050
      TabIndex        =   4
      Top             =   210
      Width           =   2010
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      Left            =   105
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   2205
      Width           =   11505
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00800000&
      FillColor       =   &H00400000&
      ForeColor       =   &H00400000&
      Height          =   1670
      Left            =   10100
      ScaleHeight     =   1605
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   22
         Top             =   840
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   880
         Picture         =   "FIKARDEX.frx":2FC0
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Terminar - Salir de Consulta de Stocks"
         Top             =   210
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   880
         Picture         =   "FIKARDEX.frx":310A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprime Ficha de Movimientos de Material (Ficha Kardex)"
         Top             =   860
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "FIKARDEX.frx":3774
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   1065
         Width           =   1515
      End
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"FIKARDEX.frx":5696
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      TabIndex        =   38
      Top             =   1920
      Visible         =   0   'False
      Width           =   11505
   End
   Begin VB.Line Line2 
      X1              =   11610
      X2              =   11610
      Y1              =   1920
      Y2              =   5880
   End
   Begin VB.Line Line1 
      X1              =   105
      X2              =   11600
      Y1              =   5860
      Y2              =   5860
   End
   Begin VB.Label Label5 
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
      Height          =   285
      Left            =   4515
      TabIndex        =   20
      Top             =   210
      Width           =   750
   End
   Begin VB.Label Label4 
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
      Height          =   285
      Left            =   6930
      TabIndex        =   9
      Top             =   210
      Width           =   3060
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dep:"
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
      Left            =   5040
      TabIndex        =   8
      Top             =   260
      Width           =   960
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "U-M:"
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
      Left            =   3570
      TabIndex        =   7
      Top             =   255
      Width           =   855
   End
   Begin VB.Label Label3 
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
      Left            =   0
      TabIndex        =   6
      Top             =   250
      Width           =   960
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"FIKARDEX.frx":5728
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      TabIndex        =   5
      Top             =   1920
      Width           =   11500
   End
   Begin VB.Menu menu 
      Caption         =   "Menu"
      Begin VB.Menu Imprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "FIKARDEX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BOTSEL_Click(Index As Integer)
    If Index < 2 Then
         Call SELECHO("SELFECHA")
         If VALACT1$("SELFECHA") <> "" Then
           If Index = 0 Then
                Text11.TEXT = VALACT1$("SELFECHA")
              Else
                Text12.TEXT = VALACT1$("SELFECHA")
           End If
         End If
      ElseIf Index = 2 Then
        Call SELECHO("TADEPOSI")
        If VALACT1$("TADEPOSI") <> "" Then
          Text4 = TRIM$(VALACT1$("TADEPOSI"))
        End If
      ElseIf Index = 3 Then
        Call SELECHO("MASTER")
        If VALACT1$("MASTER") <> "" Then
          Text1 = VALACT1$("MASTER")
        End If
    End If
End Sub

Private Sub Check1_Click()
   List1.Clear
   Call GENKARDEX
End Sub

Private Sub Command1_Click()
   Hide
End Sub

Private Sub Command8_Click()
   '
   Command8.Enabled = False
   ZBLA$ = REGBLAN$("COKARDEX")
   JJ& = 0
   FIN& = List1.ListCount
   '\\\OT-06-0390-MH-13/09/06///
   TODO% = 0
   For KU& = 1 To FIN&
     If List1.Selected(KU& - 1) = True Then GoTo 10
   Next KU&
   TODO% = 1
   GoTo 20
   '
10 OPXX% = COMALTER%("Seleccione Tipo de Listado\\Texto Seleccionado\Todo\Cancelar")
   TODO% = OPXX% - 1
   If TODO% > 1 Then GoTo 99
   '
20 For J1% = 1 To List1.ListCount
     '
     J2& = J1% + 1
     Call TRACE(24, J2&, FIN&)
     '
     If TODO% = 1 Or List1.Selected(J1% - 1) Then
        '
        x$ = List1.List(J1% - 1)
        FEC1% = FECHANUM(Left$(x$, 8))
        CM1$ = Mid$(x$, 11, 2)
        DOPRI1$ = Mid$(x$, 15, 12)
        DOSEC1$ = Mid$(x$, 29, 12)
        REFE1$ = Mid$(x$, 43, 26)
        dep1% = XVALO(Mid$(x$, 70, 3))
        ENTRA1# = XVALO(Mid$(x$, 75, 10))
        SALE1# = XVALO(Mid$(x$, 85, 10))
        SALD1# = XVALO(Mid$(x$, 95, 12))
        '
        YY$ = ZBLA$
        Call REPLA(YY$, MKI$(FEC1%), 1, 2)
        CI1% = CODINT%(Text1.TEXT)
        Call REPLA(YY$, MKI$(CI1%), 3, 2)
        Call REPLA(YY$, CM1$, 21, 2)
        Call REPLA(YY$, DOPRI1$, 23, 10)
        Call REPLA(YY$, DOSEC1$, 33, 10)
        Call REPLA(YY$, REFE1$, 43, 30)
        Call REPLA(YY$, Chr$(dep1%), 83, 1)
        Call REPLA(YY$, MKD$(ENTRA1#), 88, 8)
        Call REPLA(YY$, MKD$(SALE1#), 96, 8)
        Call REPLA(YY$, MKD$(SALD1#), 104, 8)
        JJ& = JJ& + 1
        Call GRAREG("COKARDEX", YY$, JJ&)
        '
     End If
     '
   Next J1%
   Call SETULTREG("COKARDEX", JJ&)
  
   '\\\OT-06-0390-MH-FIN///
   '
   Call HEADERS("COKARDEX", "")
   Call HEADERS("COKARDEX", "Artículo: " + Text1.TEXT + " - " + Label6.Caption)
   Call HEADERS("COKARDEX", "Período: " + Text11.TEXT + " - " + Text12.TEXT)
   Call FINAL("*LICARDEX")
   '
99 Command8.Enabled = True
   '
 End Sub

Private Sub Form_Load()
 If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture6.Visible = True
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2(0).BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Option1.BackColor = &HFCD7B6
      Me.Option2.BackColor = &HFCD7B6
      Me.Check1.BackColor = &HFCD7B6
      MARCOBARRA.Top = 1450
      MARCOBARRA.Left = 500
      Set MARCOBARRA.Container = Me.Picture6

 End If

   If EXISTE%("TRAZAB04.EXE") > 0 Or INTERPRE% = 1 Then
      Frame1.Width = 4630
      Label6.Width = 4280
      Frame3.Visible = True
      Label15.Visible = False
      Label16.Visible = True
   End If
   
   Text11 = "01/01/80"
   HOII% = HOY(HOS$)
   Text12 = HOS$
   Call Text4_Change
End Sub

Private Sub imprimir_Click()
   Command8_Click
End Sub

Private Sub List1_DblClick()
   TTXX$ = List1.TEXT
   CANTING = XVALO(Mid$(TTXX$, 75, 10))
   CANTSAL = XVALO(Mid$(TTXX$, 85, 10))
   '
   If CANTING > 0 Then
       IDLOT$ = TRIM$(Mid$(TTXX$, 129, 16))
       If IDLOT$ <> "" Then
         On Error Resume Next
         CONSUMLOT.Label4 = Text1
         If Err Then Exit Sub
         On Error GoTo 0
         CONSUMLOT.Label1 = IDLOT$
         CONSUMLOT.Show 1
       End If
     ElseIf CANTSAL > 0 Then
       On Error Resume Next
       LOCONSUM04.Label4 = TRIM$(Mid$(TTXX$, 15, 14))
       If Err Then Exit Sub
       On Error GoTo 0
       LOCONSUM04.Label5 = Mid$(TTXX$, 1, 8)
       LOCONSUM04.Label6 = TRIM$(Mid$(TTXX$, 43, 26))
       LOCONSUM04.Show 1
   End If
   '
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Option2.Value = True Then   ' solo si muestra por lotes
      If Button = 2 Then
         TTXX$ = List1.TEXT
         CANTING = XVALO(Mid$(TTXX$, 75, 10))
         CANTSAL = XVALO(Mid$(TTXX$, 85, 10))
         TIMOI$ = TRIM$(Mid$(TTXX$, 11, 2))
         '
         CIXI% = CODINT%(Text1)
         If CIXI% < 1 Then Exit Sub
         '
         If CANTSAL > 0 Or TIMOI$ = "AJ" Or TIMOI$ = "TR" Then
            Call ARMALILOT(CIXI%)
            Call FRESHECHO("!LILOTSE")
            Screen.MousePointer = 1
            If ULTREG("!LILOTSE") < 1 Then
              Call MENSERR(24, "No Hay Lotes Disponibles\para el Ajuste Propuesto.")
              Exit Sub
            End If
            '
            Call SELECHO("!LILOTSE")
            VDEVU$ = VALACT1$("!LILOTSE")
            VDEV2$ = VALACT2$("!LILOTSE")
            If TRIM$(VDEVU$) = "" Then Exit Sub
            '
            CAMAXI = XVALO(Mid$(VDEV2$, 21, 12))
            If CANTSAL > CAMAXI Then
               Call MENSERR(24, "Cantidad del Movimiento Supera\Saldo Fisico del Lote")
               Exit Sub
            End If
            '
            CMOX$ = TRIM$(Mid$(TTXX$, 11, 2))
            DOPRIX$ = TRIM$(Mid$(TTXX$, 15, 12))
            FEMOX$ = TRIM$(Left$(TTXX$, 8))
            FEMOI = Int(CDbl(CVDAT(FECHANUM(FEMOX$))))
            LOTEA$ = TRIM$(Mid$(TTXX$, 29, 16))
            TTZZ$ = "Aprueba Aplicar el Lote '" + VDEVU$ + "'\al Comprobante '" + CMOX$ + "' - '" + DOPRIX$ + "'"
            TTZZ$ = TTZZ$ + "\de Fecha '" + FEMOX$ + "' ?\\No, Cancelar\Si, Aplicar"
            If COMALTER%(TTZZ$) = 2 Then
               SQLX$ = "Cod_Inte = " & CIXI%
               SQLX$ = SQLX$ + " AND CodiMovi = '" & CMOX$ & "'"
               SQLX$ = SQLX$ + " AND DoPriCor = '" & DOPRIX$ & "'"
               SQLX$ = SQLX$ + " AND INT(Cdbl(FechMov)) = " & FEMOI
               If CANTSAL >= 0.05 Then SQLX$ = SQLX$ + " AND CantSalid = " & TRIM$(FORMATNUM(CANTSAL, "F12.4"))
               SQLX$ = SQLX$ + " AND IdenLote = '" & LOTEA$ & "' "
               '
               Screen.MousePointer = 11
               Movisto.FindFirst SQLX$
               If Movisto.NoMatch = False Then
                    Movisto.Edit
                    Movisto!IdenLote = VDEVU$
                    Movisto.Update
                    Call ACSALOTE(CIXI%, VDEVU$, "NOMESS")
                    TPI = List1.TopIndex
                    LSI = List1.ListIndex
                    Call GENKARDEX
                    On Error Resume Next
                      List1.TopIndex = TPI
                      List1.Selected(LSI) = True
                    On Error GoTo 0
                 Else
                    Call MENSERR(24, "Actualización Imposible")
               End If
               Screen.MousePointer = 1
            End If
         End If
      End If
   End If
End Sub

Private Sub Option1_Click()
   TPI = List1.TopIndex
   LSI = List1.ListIndex
   Call GENKARDEX
   On Error Resume Next
     List1.TopIndex = TPI
     List1.Selected(LSI) = True
   On Error GoTo 0
End Sub

Private Sub Option2_Click()
   TPI = List1.TopIndex
   LSI = List1.ListIndex
   Call GENKARDEX
   On Error Resume Next
     List1.TopIndex = TPI
     List1.Selected(LSI) = True
   On Error GoTo 0
End Sub

Private Sub salir_Click()
    Command1_Click
End Sub

Private Sub Text1_Change()
   '
   CI1% = CODINT%(Text1.TEXT)
   Label6.Caption = DESCRIT0$(CI1%)
   Label5.Caption = UNIMED$(CI1%)
   Call FRESHALL
   '
   Call GENKARDEX
End Sub

Private Sub Text1_DblClick()
   Call BOTSEL_Click(3)
End Sub

Private Sub Text11_Change()
   List1.Clear
   If FECHANUM(Text11) <> 0 Then
     Call GENKARDEX
   End If
End Sub

Private Sub Text11_DblClick()
   Call BOTSEL_Click(0)
End Sub

Private Sub Text12_Change()
   List1.Clear
   If FECHANUM(Text12) <> 0 Then
     Call GENKARDEX
   End If
End Sub

Private Sub Text12_DblClick()
   Call BOTSEL_Click(0)
End Sub

Private Sub Text3_Change()

End Sub

Private Sub Text4_Change()
   List1.Clear
   If Val(Text4) < 256 Then
     If Val(Text4) < 1 Or ECOARCH$("TADEPOSI", Chr$(Val(Text4))) = "" Then
         Label4 = "Todos los Depósitos"
       Else
         Label4 = ECOARCH$("TADEPOSI", Chr$(Val(Text4)))
     End If
     Call GENKARDEX
   End If
End Sub

Private Sub Text4_DblClick()
   Call BOTSEL_Click(2)
End Sub

Sub GENKARDEX()
   '
   Static DEPXX%, PRIDEP%
   If PRIDEP% < 1 Then
     Call CARDEPOS
     PRIDEP% = COSU%(1)
   End If
   '
   LU$ = LUDAT$
   If EXISTE%(LU$ + "COKARDEX.RFS") < 1 Then
     Call COPYSTRU("MOVISTO", "COKARDEX")
   End If
   REBLA$ = REGBLAN$("COKARDEX")
   '
   If Option1.Value = True Then
        Label15.Visible = True
        Label16.Visible = False
        Frame2(3).Left = 4500
     Else
        Label15.Visible = False
        Label16.Visible = True
        Frame2(3).Left = 4940
   End If
   List1.Clear
   CODL$ = Text1
5  CIXX% = CODINT%(CODL$)
   If CIXX% < 1 Then Exit Sub
   '
   DEPXX% = Val(Text4)
   If DEPXX% < 0 Then DEPXX% = 0
   '
   DESFE% = FECHANUM(Text11)
   DESDE = CDbl(CVDAT(DESFE%))
   HASFE% = FECHANUM(Text12)
   HASTA = CDbl(CVDAT(HASFE%))
   If DESDE > HASTA Then
     Call MENSERR(24, "Rango de Fechas no Válido")
     Exit Sub
   End If
   '
   INCLUVIE% = 0
   If Check1.Value = 1 Then INCLUVIE% = 1
   '
6  CIDD% = CIXX%
   JJ& = 0
   SARRAS# = 0
   INCLUPOS% = 0
   If CIDD% > 0 Then
     'Call BLOQARCH("COKARDEX")
     Call ABRESTOCKS
     '
     If INCLUVIE% = 1 Then
        Call ABRESTOCKVIE
        SQLX$ = " SELECT * FROM MOVISTO "
        SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIDD% & " "
        'SQLX$ = SQLX$ + "AND Cdbl(FechMov) >= " & DESDE & " "
        SQLX$ = SQLX$ + "AND INT(Cdbl(FechMov)) <= " & HASTA & " "
        If DEPXX% > 0 Then
          SQLX$ = SQLX$ + "AND DepoMovi = " & DEPXX% & " "
        End If
        SQLX$ = SQLX$ + "AND CodiMovi <> 'KU' "
        SQLX$ = SQLX$ + "ORDER BY FechMov ASC, FeReMovi ASC; "
        '
        Set MoviTemp = StockVie.OpenRecordset(SQLX$, dbOpenSnapshot)
        With MoviTemp
          On Error Resume Next
          .MoveFirst
          If Err < 1 Then
            Do While Not .EOF
              If CVINT(!FECHMOV) < DESFE% Then
                   SARRAS# = SARRAS# + !CANTINGRE - !CANTSALID
                 Else
                   YY$ = REBLA$
                   Call REPLA(YY$, MKI$(CVINT(!FECHMOV)), 1, 2)
                   Call REPLA(YY$, MKI$(!COD_INTE), 3, 2)
                   Call REPLA(YY$, !IdenLote, 5, 16)
                   Call REPLA(YY$, !CodiMovi, 21, 2)
                   Call REPLA(YY$, !DoPriCor, 23, 10)
                     If Option1.Value = True Then
                        Call REPLA(YY$, !DoSecCor, 33, 10)
                        Call REPLA(YY$, !ReferCor, 43, 30)
                     Else
                        Call REPLA(YY$, !IdenLote, 33, 16)
                        Call REPLA(YY$, !ReferCor, 49, 24)
                     End If
                   Call REPLA(YY$, MKD$(!VALOUNIT), 73, 8)
                     MONEX$ = "$": If !MoneVal = 2 Then MONEX$ = "U$"
                   Call REPLA(YY$, MONEX$, 81, 2)
                   Call REPLA(YY$, Chr$(!DEPOMOVI), 83, 1)
                   Call REPLA(YY$, MKI$(!Nume_Clie), 84, 2)
                   Call REPLA(YY$, MKI$(!Nume_Prov), 86, 2)
                   Call REPLA(YY$, MKD$(!CANTINGRE), 88, 8)
                   Call REPLA(YY$, MKD$(!CANTSALID), 96, 8)
                   Call REPLA(YY$, MKD$(0), 104, 8)
                   Call REPLA(YY$, Mid$(!DoPriCor, 11, 4), 112, 4)
                   Call REPLA(YY$, Mid$(!DoSecCor, 11, 4), 116, 4)
                   Call REPLA(YY$, CVINT(!FeReMovi), 122, 2)
                   Call REPLA(YY$, Chr$(!NumUsuar), 126, 1)
                   JJ& = JJ& + 1
                   'Call GRAREG("!COKARDEX", YY$, JJ&)
                   Call GRAREG("#COKARDEX", YY$, JJ&)
              End If
            .MoveNext
            Loop
          End If
        End With
        '
     End If
     '
     SQLX$ = " SELECT * FROM MOVISTO "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIDD% & " "
     'SQLX$ = SQLX$ + "AND Cdbl(FechMov) >= " & DESDE & " "
     'SQLX$ = SQLX$ + "AND INT(Cdbl(FechMov)) <= " & HASTA & " "
     If DEPXX% > 0 Then
       If DEPXX% = PRIDEP% Then
           SQLX$ = SQLX$ + "AND (DepoMovi = " & DEPXX% & " OR DepoMovi = 0 OR ISNULL(DepoMovi) = True OR ISEMPTY(DepoMovi) = True) "
         Else
           SQLX$ = SQLX$ + "AND DepoMovi = " & DEPXX% & " "
       End If
     End If
     If INCLUVIE% = 1 Then
       SQLX$ = SQLX$ + "AND CodiMovi <> 'KU' "
     End If
     SQLX$ = SQLX$ + "ORDER BY FechMov ASC, FeReMovi ASC; "
     '
     Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     With MoviTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           If CVINT(!FECHMOV) > HASFE% Then
             If INCLUPOS% = 0 Then
               OPXX% = COMALTER%("Se Han Detectado Movimientos POSTERIORES\a la Fecha Límite Seleccionada.\\No Considerarlos\Incluirlos")
               If OPXX% < 2 Then INCLUPOS% = (-1)
               If OPXX% > 1 Then INCLUPOS% = 1
             End If
           End If
           '
           If CVINT(!FECHMOV) < DESFE% Then
               SARRAS# = SARRAS# + !CANTINGRE - !CANTSALID
             ElseIf CVINT(!FECHMOV) <= HASFE% Or INCLUPOS% = 1 Then
               YY$ = REBLA$
               Call REPLA(YY$, MKI$(CVINT(!FECHMOV)), 1, 2)
               Call REPLA(YY$, MKI$(!COD_INTE), 3, 2)
               Call REPLA(YY$, !IdenLote, 5, 16)
               Call REPLA(YY$, !CodiMovi, 21, 2)
               Call REPLA(YY$, !DoPriCor, 23, 10)
                 If Option1.Value = True Then
                    Call REPLA(YY$, !DoSecCor, 33, 10)
                    Call REPLA(YY$, !ReferCor, 43, 30)
                 Else
                    Call REPLA(YY$, !IdenLote, 33, 16)
                    Call REPLA(YY$, !ReferCor, 49, 24)
                 End If
               Call REPLA(YY$, MKD$(!VALOUNIT), 73, 8)
                 MONEX$ = "$": If !MoneVal = 2 Then MONEX$ = "U$"
               Call REPLA(YY$, MONEX$, 81, 2)
               Call REPLA(YY$, Chr$(!DEPOMOVI), 83, 1)
               Call REPLA(YY$, MKI$(!Nume_Clie), 84, 2)
               Call REPLA(YY$, MKI$(!Nume_Prov), 86, 2)
               Call REPLA(YY$, MKD$(!CANTINGRE), 88, 8)
               Call REPLA(YY$, MKD$(!CANTSALID), 96, 8)
               Call REPLA(YY$, MKD$(0), 104, 8)
               Call REPLA(YY$, Mid$(!DoPriCor, 11, 4), 112, 4)
               Call REPLA(YY$, Mid$(!DoSecCor, 11, 4), 116, 4)
               Call REPLA(YY$, CVINT(!FeReMovi), 122, 2)
               Call REPLA(YY$, Chr$(!NumUsuar), 126, 1)
               JJ& = JJ& + 1
               'Call GRAREG("!COKARDEX", YY$, JJ&)
               Call GRAREG("#COKARDEX", YY$, JJ&)
           End If
         .MoveNext
         Loop
       End If
     End With
     Set MoviTemp = Nothing
     'Movisto.Close
     '
     On Error GoTo 0
     'Call SETULTREG("!COKARDEX", JJ&)
     Call SETULTREG("#COKARDEX", JJ&)
     Screen.MousePointer = 11
     '
     If INCLUVIE% = 1 Then
       Call BLOQARCH("COKARDEX")
       'Call TRACENTRAL("COKARDEX", "")
       Call FILEDUMP("COKARDEX")
       Call FILESORT("COKARDEX", "", 1, 1, "ASC")
       'Call TRALOCAL("COKARDEX", "")
       For KKUL& = 1 To JJ&
         CKX$ = REGLEIDO$("COKARDEX", KKUL&)
         Call GRAREG("#COKARDEX", CKX$, KKUL&)
       Next KKUL&
       Call CIERRARCH("COKARDEX")
     End If
     '
     'FIN& = ULTREG&("!COKARDEX")
     FIN& = ULTREG&("#COKARDEX")
     CASALDO# = 0
     If Abs(SARRAS#) >= 0.05 Then
       TTXX$ = Space$(256)
         FEX = DIANTE(DESFE%)
       Call REPLA(TTXX$, FECHATEX$(FEX), 1, 8)
         REFEM$ = "Saldo Anterior al " + FECHATEX$(DESFE%)
       Call REPLA(TTXX$, REFEM$, 43, 36)
         CASALDO# = SARRAS#
       Call REPLA(TTXX$, FORMATNUM$(CASALDO#, "F12.1"), 95, 12)
       List1.AddItem TTXX$
     End If
     '
     For KU& = 1 To FIN&
       'YY$ = REGLEIDO$("!COKARDEX", KU&)
       YY$ = REGLEIDO$("#COKARDEX", KU&)
       TTXX$ = Space$(256)
         FEX = CVI(Left$(YY$, 2))
       Call REPLA(TTXX$, FECHATEX$(FEX), 1, 8)
         CMOVI$ = Mid$(YY$, 21, 2)
       Call REPLA(TTXX$, CMOVI$, 11, 2)
         DOPRI$ = Mid$(YY$, 23, 10) + Mid$(YY$, 112, 4)
       Call REPLA(TTXX$, DOPRI$, 15, 12)
         DOSEC$ = Mid$(YY$, 33, 10) + Mid$(YY$, 116, 4)
       Call REPLA(TTXX$, DOSEC$, 29, 12)
         REFEM$ = Mid$(YY$, 43, 30)
       Call REPLA(TTXX$, REFEM$, 43, 26)
         If Option2.Value = True Then
            Call REPLA(TTXX$, Space$(40), 29, 40)
            DOSEC$ = Mid$(YY$, 33, 16)
            REFEM$ = Mid$(YY$, 49, 24)
            Call REPLA(TTXX$, DOSEC$, 29, 16)
            Call REPLA(TTXX$, REFEM$, 47, 22)
         End If
         DEPOM$ = FORMATNUM$(Asc(Mid$(YY$, 83, 1)), "I3")
       Call REPLA(TTXX$, DEPOM$, 70, 3)
         CAINGRE# = CVD(Mid$(YY$, 88, 8))
       Call REPLA(TTXX$, FORMATNUM$(CAINGRE#, "F10.1"), 75, 10)
         CASALI# = CVD(Mid$(YY$, 96, 8))
       Call REPLA(TTXX$, FORMATNUM$(CASALI#, "F10.1"), 85, 10)
         CASALDO# = CASALDO# + CAINGRE# - CASALI#
       Call REPLA(TTXX$, FORMATNUM$(CASALDO#, "F12.1"), 95, 12)
         IDLOT$ = Mid$(YY$, 5, 16)
       Call REPLA(TTXX$, IDLOT$, 129, 16)
       List1.AddItem TTXX$
     Next KU&
     Call FREEMEM("COKARDEX")
     Screen.MousePointer = 1
     '
 Exit Sub
     Call SETULTREG("COKARDEX", JJ&)
     Call CIERRARCH("COKARDEX")
     Screen.MousePointer = 1
     Call CONECRUN("*COKARDEX/" + CODD$, "Ficha Kardex Indirecta") ' + "/WAIT")
     Screen.MousePointer = 1
    Else
     Call MENSERR(24, "Codigo: " + TRIM$(CODD$) + "\sin Movimientos Según Condición SQL.")
   End If
   '
   'If TRIM$(CODL$) = "" Then GoTo 5
   Screen.MousePointer = 1
   GoTo 5
   '
End Sub

