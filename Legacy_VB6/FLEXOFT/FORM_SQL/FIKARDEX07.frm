VERSION 5.00
Begin VB.Form FIKARDEX07 
   BackColor       =   &H00FEFEE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ficha de Movimientos de Material (Ficha Kardex)"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "FIKARDEX07.frx":0000
      Left            =   7605
      List            =   "FIKARDEX07.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   35
      Text            =   "Combo1"
      Top             =   195
      Width           =   1935
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
      Height          =   1110
      Left            =   105
      TabIndex        =   21
      Top             =   630
      Width           =   4695
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
         Height          =   615
         Left            =   150
         TabIndex        =   25
         Top             =   315
         Width           =   4365
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
      Left            =   3570
      Picture         =   "FIKARDEX07.frx":0004
      TabIndex        =   19
      Top             =   215
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
      Left            =   5940
      Picture         =   "FIKARDEX07.frx":030E
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
      Height          =   300
      Left            =   5415
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
      Height          =   780
      Left            =   5400
      TabIndex        =   10
      Top             =   960
      Width           =   4500
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
         Left            =   2760
         TabIndex        =   24
         Top             =   0
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
         Left            =   4215
         Picture         =   "FIKARDEX07.frx":0618
         TabIndex        =   14
         Top             =   360
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
         Left            =   2115
         Picture         =   "FIKARDEX07.frx":0922
         TabIndex        =   13
         Top             =   360
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
         Left            =   1065
         TabIndex        =   12
         Text            =   " "
         Top             =   360
         Width           =   1065
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
         Left            =   3165
         TabIndex        =   11
         Text            =   " "
         Top             =   360
         Width           =   1065
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
         Left            =   120
         TabIndex        =   16
         Top             =   390
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
         Left            =   2325
         TabIndex        =   15
         Top             =   390
         Width           =   750
      End
   End
   Begin VB.TextBox Text1 
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
      Left            =   1050
      TabIndex        =   4
      Top             =   225
      Width           =   2535
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
      Height          =   1755
      Left            =   10100
      ScaleHeight     =   1695
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
         Picture         =   "FIKARDEX07.frx":0C2C
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
         Picture         =   "FIKARDEX07.frx":0D76
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprime Ficha de Movimientos de Material (Ficha Kardex)"
         Top             =   860
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "FIKARDEX07.frx":13E0
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   1065
         Width           =   1515
      End
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   9550
      Picture         =   "FIKARDEX07.frx":3302
      Stretch         =   -1  'True
      ToolTipText     =   "Detalle Movimientos por Lote Activo"
      Top             =   140
      Width           =   360
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Partidas:"
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
      Left            =   6510
      TabIndex        =   36
      Top             =   240
      Width           =   1065
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
      Height          =   315
      Left            =   4395
      TabIndex        =   20
      Top             =   210
      Width           =   405
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
      Left            =   5400
      TabIndex        =   9
      Top             =   555
      Width           =   4500
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
      Left            =   4755
      TabIndex        =   8
      Top             =   255
      Width           =   615
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
      Left            =   3870
      TabIndex        =   7
      Top             =   255
      Width           =   510
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
      Caption         =   $"FIKARDEX07.frx":360C
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
End
Attribute VB_Name = "FIKARDEX07"
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

Private Sub Combo1_Click()
  'SeleccionEnList
  For I = 0 To List1.ListCount - 1
     LOTMOV$ = Mid$(List1.List(I), 129, 16)
     If LOTMOV$ = Combo1.TEXT Then
          List1.Selected(I) = True
        Else
          List1.Selected(I) = False
     End If
   Next I
   
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
        DEP1% = XVALO(Mid$(x$, 70, 3))
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
        Call REPLA(YY$, Chr$(DEP1%), 83, 1)
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
   If Val(Text4) > 0 Or ECOARCH$("TADEPOSI", Chr$(Val(Text4))) <> "" Then
      Call HEADERS("COKARDEX", "Deposito: " & TRIM$(Text4) & " - " & ECOARCH$("TADEPOSI", Chr$(Val(Text4))))
   End If
   Call HEADERS("COKARDEX", "Artículo: " + Text1.TEXT + " - " + Label6.Caption)
   Call HEADERS("COKARDEX", "Período: " + Text11.TEXT + " - " + Text12.TEXT)
   Call FINAL("*LICARDEX")
   '
99 Command8.Enabled = True
   '
 End Sub

Private Sub Form_Load()
   Text11 = "01/01/80"
   HOII% = HOY(HOS$)
   Text12 = HOS$
   
   Call APLICACIONSKINS(Me)

   Call Text4_Change
   '\\\OT-08-0990-OD-30/10/08///
   LOMA% = XVALO(CONTROL$("MASTER", "LOMAXCOD"))
   If LOMA% > 15 Then Text1.MaxLength = LOMA%
   '\\\OT-08-0990-OD-FIN///

End Sub

Private Sub Image1_Click()
   IDLOT$ = TRIM$(Combo1.TEXT)
   If IDLOT$ <> "" Then
     On Error Resume Next
     CONSUMLOT07.Label4 = Text1
     If Err Then Exit Sub
     On Error GoTo 0
     CONSUMLOT07.Label1 = IDLOT$
     CONSUMLOT07.Show 1
   End If
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
         CONSUMLOT07.Label4 = Text1
         If Err Then Exit Sub
         On Error GoTo 0
         CONSUMLOT07.Label1 = IDLOT$
         CONSUMLOT07.Show 1
       End If
     ElseIf CANTSAL > 0 Then
       On Error Resume Next
       LOCONSUM07.Label4 = TRIM$(Mid$(TTXX$, 15, 14))
       If Err Then Exit Sub
       On Error GoTo 0
       LOCONSUM07.Label5 = Mid$(TTXX$, 1, 8)
       LOCONSUM07.Label6 = TRIM$(Mid$(TTXX$, 43, 26))
       LOCONSUM07.Show 1
   End If
   '
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   '
   If Button = 2 Then
      '
      LINEA$ = List1.List(List1.ListIndex)
      FECHA1$ = Mid$(LINEA$, 1, 8)
      FF% = FECHANUM(FECHA1$)
      FECHA2$ = FETEXCOR1$(FF%)
      LOTE$ = Mid(LINEA$, 15, 12)
      '
      CONDI$ = "DOPRICOR ='" & TRIM$(LOTE$) & "' AND FECHMOV='" & FECHA2$ & "'"
      '
      CODUSU% = XVALO(VALOBADA("MOVISTO", "NUMUSUAR", CONDI$))
      USU$ = USERNAM$(CODUSU%)
      FECHAREAL$ = Mid(SAFETEXT$(VALOBADA("MOVISTO", "FEREMOVI", CONDI$)), 1, 10)
      HORA$ = Mid(SAFETEXT$(VALOBADA("MOVISTO", "FEREMOVI", CONDI$)), 12, 8)
      '
      Call REPLA(VDEF$, USU$, 1, 20)
      Call REPLA(VDEF$, FECHAREAL$, 21, 10)
      Call REPLA(VDEF$, HORA$, 31, 8)
      DATOS$ = OBJPLA("DATOSMOV", VDEF$)
   End If
   '
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
   Combo1.Clear
   VECOMBO$ = Space$(1024)
   PINI& = 1
   '
   LU$ = LUDAT$
   If EXISTE%(LU$ + "COKARDEX.RFS") < 1 Then
     Call COPYSTRU("MOVISTO", "COKARDEX")
   End If
   REBLA$ = REGBLAN$("COKARDEX")
   '
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
     Call ABRESTOCKS
'     Call ABRESTOCKVIE
     
     If INCLUVIE% = 1 Then
     '
        SQLX1$ = " SELECT * FROM MOVIVIE WITH(NOLOCK) "
        SQLX1$ = SQLX1$ + "WHERE Cod_Inte = " & CIDD% & " "
        If DEPXX% > 0 Then
          If DEPXX% = PRIDEP% Then
              SQLX1$ = SQLX1$ + "AND (DepoMovi = 1 OR DepoMovi = 0 OR DepoMovi IS NULL)"
            Else
              SQLX1$ = SQLX1$ + "AND DepoMovi = " & DEPXX% & " "
          End If
        End If
        SQLX1$ = SQLX1$ + "AND CodiMovi <> 'KU' "
        SQLX1$ = SQLX1$ + "ORDER BY FechMov , FeReMovi ASC "
        '
        Dim MOVIVIETMP As ADODB.Recordset
        Set MOVIVIETMP = FLEXCONN.Execute(FILTSQL$(SQLX1$))
        With MOVIVIETMP
            Do While Not .EOF
              If CVINT(!fechMov) > HASFE% Then
                If INCLUPOS% = 0 Then
                  OPXX% = COMALTER%("Se Han Detectado Movimientos POSTERIORES\a la Fecha Límite Seleccionada.\\No Considerarlos\Incluirlos")
                  If OPXX% < 2 Then INCLUPOS% = (-1)
                  If OPXX% > 1 Then INCLUPOS% = 1
                End If
              End If
              '
              If CVINT(!fechMov) < DESFE% Then
                  SARRAS# = SARRAS# + XVALO(!CantIngre) - XVALO(!Cantsalid)
                ElseIf CVINT(!fechMov) <= HASFE% Or INCLUPOS% = 1 Then
                  YY$ = REBLA$
                  Call REPLA(YY$, MKI$(CVINT(!fechMov)), 1, 2)
                  Call REPLA(YY$, MKI$(XVALO(!COD_INTE)), 3, 2)
                  Call REPLA(YY$, SAFETEXT$(!idenlote), 5, 16)
                  Call REPLA(YY$, SAFETEXT$(!CodiMovi), 21, 2)
                  Call REPLA(YY$, SAFETEXT$(!DoPriCor), 23, 10)
                  Call REPLA(YY$, SAFETEXT$(!DoSecCor), 33, 10)
                  Call REPLA(YY$, SAFETEXT$(!ReferCor), 43, 30)
                  Call REPLA(YY$, MKD$(XVALO(!ValoUnit)), 73, 8)
                    MONEX$ = "$": If XVALO(!MoneVal) = 2 Then MONEX$ = "U$"
                  Call REPLA(YY$, MONEX$, 81, 2)
                  Call REPLA(YY$, Chr$(XVALO(!DepoMovi)), 83, 1)
                  Call REPLA(YY$, MKI$(XVALO(!Nume_Clie)), 84, 2)
                  Call REPLA(YY$, MKI$(XVALO(!Nume_Prov)), 86, 2)
                  Call REPLA(YY$, MKD$(XVALO(!CantIngre)), 88, 8)
                  Call REPLA(YY$, MKD$(XVALO(!Cantsalid)), 96, 8)
                  Call REPLA(YY$, MKD$(0), 104, 8)
                  Call REPLA(YY$, Mid$(SAFETEXT$(!DoPriCor), 11, 4), 112, 4)
                  Call REPLA(YY$, Mid$(SAFETEXT$(!DoSecCor), 11, 4), 116, 4)
                  Call REPLA(YY$, CVINT(!FeReMovi), 122, 2)
                  Call REPLA(YY$, Chr$(XVALO(!NumUsuar)), 126, 1)
                  JJ& = JJ& + 1
                  Call GRAREG("#COKARDEX", YY$, JJ&)
              End If
            .MoveNext
            Loop
        End With
        MOVIVIETMP.Close
        Set MOVIVIETMP = Nothing
     End If
     
     SQLX$ = " SELECT * FROM MOVISTO WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIDD% & " AND (VERIFICADO <> -1 OR VERIFICADO IS NULL) "
     If DEPXX% > 0 Then
       If DEPXX% = PRIDEP% Then
           SQLX$ = SQLX$ + "AND (DepoMovi = 1 OR DepoMovi = 0 OR DepoMovi IS NULL)"
         Else
           SQLX$ = SQLX$ + "AND DepoMovi = " & DEPXX% & " "
       End If
     End If
     If INCLUVIE% = 1 Then
       SQLX$ = SQLX$ + "AND CodiMovi <> 'KU' "
     End If
     SQLX$ = SQLX$ + "ORDER BY FechMov , FeReMovi ASC "

     '
     Dim MoviTemp As ADODB.Recordset
     Set MoviTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
     With MoviTemp
         Do While Not .EOF
           If CVINT(!fechMov) > HASFE% Then
             If INCLUPOS% = 0 Then
               OPXX% = COMALTER%("Se Han Detectado Movimientos POSTERIORES\a la Fecha Límite Seleccionada.\\No Considerarlos\Incluirlos")
               If OPXX% < 2 Then INCLUPOS% = (-1)
               If OPXX% > 1 Then INCLUPOS% = 1
             End If
           End If
           '
           If CVINT(!fechMov) < DESFE% Then
               SARRAS# = SARRAS# + XVALO(!CantIngre) - XVALO(!Cantsalid)
             ElseIf CVINT(!fechMov) <= HASFE% Or INCLUPOS% = 1 Then
               YY$ = REBLA$
               Call REPLA(YY$, MKI$(CVINT(!fechMov)), 1, 2)
               Call REPLA(YY$, MKI$(XVALO(!COD_INTE)), 3, 2)
               Call REPLA(YY$, SAFETEXT$(!idenlote), 5, 16)
               Call REPLA(YY$, SAFETEXT$(!CodiMovi), 21, 2)
               Call REPLA(YY$, SAFETEXT$(!DoPriCor), 23, 10)
               Call REPLA(YY$, SAFETEXT$(!DoSecCor), 33, 10)
               Call REPLA(YY$, SAFETEXT$(!ReferCor), 43, 30)
               Call REPLA(YY$, MKD$(XVALO(!ValoUnit)), 73, 8)
                 MONEX$ = "$": If XVALO(!MoneVal) = 2 Then MONEX$ = "U$"
               Call REPLA(YY$, MONEX$, 81, 2)
               Call REPLA(YY$, Chr$(XVALO(!DepoMovi)), 83, 1)
               Call REPLA(YY$, MKI$(XVALO(!Nume_Clie)), 84, 2)
               Call REPLA(YY$, MKI$(XVALO(!Nume_Prov)), 86, 2)
               Call REPLA(YY$, MKD$(XVALO(!CantIngre)), 88, 8)
               Call REPLA(YY$, MKD$(XVALO(!Cantsalid)), 96, 8)
               Call REPLA(YY$, MKD$(0), 104, 8)
               Call REPLA(YY$, Mid$(SAFETEXT$(!DoPriCor), 11, 4), 112, 4)
               Call REPLA(YY$, Mid$(SAFETEXT$(!DoSecCor), 11, 4), 116, 4)
               Call REPLA(YY$, CVINT(!FeReMovi), 122, 2)
               Call REPLA(YY$, Chr$(XVALO(!NumUsuar)), 126, 1)
               JJ& = JJ& + 1
               'Call GRAREG("!COKARDEX", YY$, JJ&)
               Call GRAREG("#COKARDEX", YY$, JJ&)
           End If
         .MoveNext
         Loop
     End With
     MoviTemp.Close
     Set MoviTemp = Nothing
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
       'Call REPLA(TTXX$, FORMATNUM$(CASALDO#, "F12.1"), 95, 12)
       Call REPLA(TTXX$, FORMATNUM$(CASALDO#, "F12." & CNTDC$(CIXX%) & ""), 95, 12)
       List1.AddItem TTXX$
     End If
     '
     For KU& = 1 To FIN&
       Call TRACE(20, KU&, FIN&)
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
         DEPOM$ = FORMATNUM$(Asc(Mid$(YY$, 83, 1)), "I3")
       Call REPLA(TTXX$, DEPOM$, 70, 3)
         CAINGRE# = CVD(Mid$(YY$, 88, 8))
       'Call REPLA(TTXX$, FORMATNUM$(CAINGRE#, "F10.1"), 75, 10)
       Call REPLA(TTXX$, FORMATNUM$(CAINGRE#, "F10." & CNTDC$(CIXX%) & ""), 75, 10)
         
         CASALI# = CVD(Mid$(YY$, 96, 8))
       'Call REPLA(TTXX$, FORMATNUM$(CASALI#, "F10.1"), 85, 10)
       Call REPLA(TTXX$, FORMATNUM$(CASALI#, "F10." & CNTDC$(CIXX%) & ""), 85, 10)
         CASALDO# = CASALDO# + CAINGRE# - CASALI#
       'Call REPLA(TTXX$, FORMATNUM$(CASALDO#, "F12.1"), 95, 12)
       Call REPLA(TTXX$, FORMATNUM$(CASALDO#, "F12." & CNTDC$(CIXX%) & ""), 95, 12)
         IDLOT$ = Mid$(YY$, 5, 16)
       Call REPLA(TTXX$, IDLOT$, 129, 16)
       '
       If TRIM$(IDLOT$) <> "" Then
         If InStr(VECOMBO$, IDLOT$) < 1 Then
           If PINI& > Len(VECOMBO$) Then VECOMBO$ = VECOMBO$ + Space$(1024)
           Mid$(VECOMBO$, PINI&, 16) = IDLOT$
           PINI& = PINI& + 16
         End If
       End If
       '
 '      Call CompletaCombo(IDLOT$)
       List1.AddItem TTXX$
     Next KU&
     Call FREEMEM("COKARDEX")
     '
     For KU& = 1 To PINI& - 1 Step 16
        Combo1.AddItem Mid$(VECOMBO$, KU&, 16)
     Next KU&
     '
     ' ESTO ES PARA POSICIONARSE AL FINAL DE LA LISTA
     TPXI& = List1.ListCount - 19
     If TPXI& < 0 Then TPXI& = 0
     List1.TopIndex = TPXI&
     '
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
   Screen.MousePointer = 1
   GoTo 5
   '
End Sub

Sub CompletaCombo(IDLOT$)
    If TRIM$(IDLOT$) <> "" Then
       Dim EXISTE As Boolean
       EXISTE = False
       '
       For I = 0 To Combo1.ListCount - 1
          If Combo1.List(I) = IDLOT$ Then
             EXISTE = True
          End If
       Next I
       '
       If EXISTE = False Then
          Combo1.AddItem IDLOT$
       End If
    End If
    '
End Sub
