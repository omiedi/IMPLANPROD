VERSION 5.00
Begin VB.Form AJUSTOBAS04 
   BackColor       =   &H00EEEEB0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ajustes de Stocks"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5295
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   6840
      Left            =   4410
      ScaleHeight     =   6780
      ScaleWidth      =   1110
      TabIndex        =   24
      Top             =   -105
      Width           =   1170
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   28
         Top             =   6000
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   105
            TabIndex        =   29
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
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
         Picture         =   "AJUSTOBAS04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   315
         Width           =   650
      End
      Begin VB.CommandButton Command29 
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
         Picture         =   "AJUSTOBAS04.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1000
         Width           =   650
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Kardx"
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
         Picture         =   "AJUSTOBAS04.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1800
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "AJUSTOBAS04.frx":075E
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   6250
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00EEEEB0&
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1850
      Left            =   210
      TabIndex        =   18
      Top             =   2835
      Width           =   4110
      Begin VB.Frame Frame7 
         BackColor       =   &H00EEEEB0&
         Caption         =   "Motivo - Referencia"
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
         Left            =   140
         TabIndex        =   21
         Top             =   945
         Width           =   3820
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
            Height          =   285
            Left            =   210
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   315
            Width           =   3480
         End
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00E0E0E0&
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
         Left            =   340
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   580
         Width           =   3480
      End
      Begin VB.TextBox Text10 
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
         Left            =   2730
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   580
         Width           =   1065
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   22
         Top             =   315
         Width           =   1065
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2730
      TabIndex        =   1
      Top             =   210
      Width           =   175
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   630
      Width           =   4110
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3675
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   210
      Width           =   645
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
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   210
      Width           =   1695
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   210
      TabIndex        =   2
      Top             =   1470
      Width           =   4110
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EEEEB0&
      Caption         =   "Ingreso de Ajustes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1850
      Left            =   210
      TabIndex        =   12
      Top             =   4725
      Width           =   4110
      Begin VB.TextBox Text7 
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
         Left            =   1785
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   735
         Width           =   1170
      End
      Begin VB.TextBox Text8 
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
         Left            =   1785
         TabIndex        =   4
         Top             =   1050
         Width           =   1170
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1785
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   420
         Width           =   1170
      End
      Begin VB.TextBox Text9 
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
         Left            =   1785
         TabIndex        =   5
         Top             =   1365
         Width           =   1170
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
         Height          =   600
         Left            =   3255
         Picture         =   "AJUSTOBAS04.frx":2680
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   420
         Width           =   645
      End
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
         Left            =   3255
         Picture         =   "AJUSTOBAS04.frx":298A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Entrada:"
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
         Left            =   105
         TabIndex        =   17
         Top             =   735
         Width           =   1590
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Salida:"
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
         Left            =   105
         TabIndex        =   16
         Top             =   1050
         Width           =   1590
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "StockTeórico:"
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
         Left            =   105
         TabIndex        =   15
         Top             =   420
         Width           =   1590
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Ajustado:"
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
         Left            =   105
         TabIndex        =   14
         Top             =   1365
         Width           =   1590
      End
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
      Left            =   2730
      TabIndex        =   9
      Top             =   255
      Width           =   855
   End
   Begin VB.Label Label1 
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
      Height          =   225
      Left            =   105
      TabIndex        =   8
      Top             =   255
      Width           =   855
   End
End
Attribute VB_Name = "AJUSTOBAS04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MODENTRAJU%
Private Sub BOTSEL_Click(Index As Integer)
    Call SELECHO("SELFECHA")
    If VALACT1$("SELFECHA") <> "" Then
        If Index = 0 Then
            Text11.TEXT = VALACT1$("SELFECHA")
          Else
            Text12.TEXT = VALACT1$("SELFECHA")
        End If
    End If
End Sub

Private Sub Command1_Click()
    PREPA% = 0
    Hide
End Sub

Private Sub Command13_Click()
    Command13.Enabled = False
    Call CONECRUN("ESTASTOK", "Estadísticas de Stocks")
    Command13.Enabled = True
End Sub

Private Sub Command14_Click()
    Command14.Enabled = False
    If App.EXEName = "INVENT02" Then
        PREPA% = 0
        Hide
      Else
        ACONEC$ = "INVENT02"
        If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "INVENTAR"
        Call CONECRUN(ACONEC$, "Inventarios Físicos y Valorizados")
    End If
    Command14.Enabled = True
End Sub

Private Sub COMMAND15_CLICK()
    Command15.Enabled = False
    FIKARDEX.Text1 = Text1
    FIKARDEX.Show 1
    Command15.Enabled = True
End Sub

Private Sub Command2_Click()
    COD$ = ""
    CI% = 0
    Text1.TEXT = ""
    Text2.TEXT = ""
    Text3.TEXT = ""
    Text6.TEXT = ""
    Text7.TEXT = ""
    Text8.TEXT = ""
    Text9.TEXT = ""
    Text1.SetFocus
End Sub


Private Sub Command29_Click()
    Call HELPONLINE("AJUSTOCK")
End Sub

Private Sub Command3_Click()
    Call AJUSTO3
    Call AJUSTO2
    Text6.TEXT = Text9.TEXT
    Text7.TEXT = ""
    Text8.TEXT = ""
    Text1.SetFocus
End Sub

Private Sub Command4_Click()
    '
    Call SELECHO("MASTER")
    COD$ = VALACT1$("MASTER")
    Text1.TEXT = COD$
    CI% = CODINT%(COD$)
    Text3.TEXT = DESCRIT$(CI%)
    Text2.TEXT = UNIMED$(CI%)
    '
End Sub
'

Private Sub Command7_Click()
    '
    DESNUM& = 1: If Val(Text14.TEXT) > 0 Then DESNUM& = Val(Text14.TEXT)
    HASNUM& = 999999
    If TRIM$(Text13.TEXT) <> "" Then
        If Val(Text13.TEXT) >= DESNUM& Then HASNUM& = Val(Text13.TEXT)
    End If
    DESFE% = 0: If FECHANUM(TRIM$(Text11.TEXT)) > 0 Then DESFE% = FECHANUM(TRIM$(Text11.TEXT))
    HASFE% = 32767
    If TRIM$(Text12.TEXT) <> "" Then
        If FECHANUM(TRIM$(Text12.TEXT)) >= DESFE% Then HASFE% = FECHANUM(TRIM$(Text12.TEXT))
    End If
    '
    Screen.MousePointer = 11
    Call BLANARCH("LISTRAJ")
    FIN& = ULTREG&("AJUSTOCK")
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        X$ = REGLEIDO$("AJUSTOCK", I&)
        FF% = CVI(Left$(X$, 2))
        NRO& = CVS(Mid$(X$, 3, 4))
        If FF% >= DESFE% Then
          If FF% <= HASFE% Then
            NRO& = CVS(Mid$(X$, 3, 4))
            If NRO& >= DESNUM& Then
              If NRO& <= HASNUM& Then
                REFE$ = Mid$(X$, 34, 47)
                CI1% = CVI(Mid$(X$, 7, 2))
                DP1% = Asc(Mid$(X$, 33, 1))
                CAN# = CVD(Mid$(X$, 25, 8))
                Y$ = REGBLAN$("LISTRAJ")
                Call REPLA(Y$, MKI$(FF%), 1, 2)
                Call REPLA(Y$, REFE$, 3, 48)
                Call REPLA(Y$, MKS$(NRO&), 51, 4)
                Call REPLA(Y$, MKI$(CI1%), 55, 2)
                Call REPLA(Y$, UNIMED$(CI1%), 57, 4)
                If CAN# > 0 Then
                    Call REPLA(Y$, MKD$(CAN#), 61, 8)
                    Call REPLA(Y$, MKD$(0), 69, 8)
                  Else
                    Call REPLA(Y$, MKD$(0), 61, 8)
                    Call REPLA(Y$, MKD$(Abs(CAN#)), 69, 8)
                End If
                Call REPLA(Y$, Chr$(DP1%), 80, 1)
                Call REGAPP("LISTRAJ", Y$)
              End If
            End If
          End If
        End If
    Next I&
    Call CIERRARCH("LISTRAJ")
    Screen.MousePointer = 1
    '
    OPCI% = 1
    If Option4.Value = True Then OPCI% = 2
    If OPCI% = 1 Then
        Call OPNOIN("")
        'Call CONECRUN("*COTRAST", "")
      ElseIf OPCI% = 2 Then
        Call CONECRUN("*LIAJUST", "")
    End If
    '

End Sub

Private Sub Form_Load()
    Call TRANSLABELS(Name)
    HOII% = HOY(HOS$)
    Text10.TEXT = HOS$
    Call AJUSTO1
End Sub
'

Private Sub List1_Click()
    '
    ILISTA& = List1.ListIndex
    Text4.TEXT = Left$(List1.TEXT, 20)
    Text6.TEXT = TRIM$(Mid$(List1.TEXT, 21))
    '
    Text7.TEXT = ""
    Text8.TEXT = ""
    Text9.TEXT = TRIM$(Mid$(List1.TEXT, 21))
    If Val(Text9.TEXT) < 0 Then
        Text7.TEXT = TRIM$(CSTRING$(MKD$((-1) * Val(Text9.TEXT)), 4, 10, 1, 2))
        Text9.TEXT = "0.0"
    End If
    Text9.SetFocus
End Sub

Private Sub Text1_Change()
    '
    Text2.TEXT = ""
    Text3.TEXT = ""
    If CODINT%(Text1.TEXT) > 0 Then
        COD$ = TRIM$(Text1.TEXT)
        CI% = CODINT%(COD$)
        Text2.TEXT = UNIMED$(CI%)
        Text3.TEXT = DESCRIT$(CI%)
        Call AJUSTO2               ' recalcular existencias a fecha
    End If
    '
End Sub
'

Private Sub Text1_DblClick()
    '
    Call SELECHO("MASTER")
    COD$ = VALACT1$("MASTER")
    Text1.TEXT = COD$
    CI% = CODINT%(COD$)
    Text3.TEXT = DESCRIT$(CI%)
    Text2.TEXT = UNIMED$(CI%)
    '
End Sub
'

Private Sub Text1_LostFocus()
    COD$ = TRIM$(Text1.TEXT)
    CI% = CODINT%(COD$)
    Text3.TEXT = DESCRIT$(CI%)
    Text2.TEXT = UNIMED$(CI%)
End Sub

Private Sub Text10_Change()
    If Visible = True Then
        If FECHANUM(Text10.TEXT) > 0 Then
            Call AJUSTO2
        End If
    End If
End Sub

Private Sub Text11_DblClick()
    Call SELECHO("SELFECHA")
    If VALACT1$("SELFECHA") <> "" Then
        Text11.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text12_DblClick()
    Call SELECHO("SELFECHA")
    If VALACT1$("SELFECHA") <> "" Then
        Text12.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text2_GotFocus()
    Text7.SetFocus

End Sub

Private Sub Text3_GotFocus()
    Text7.SetFocus

End Sub

Private Sub Text4_GotFocus()
    Text7.SetFocus

End Sub

Private Sub Text6_GotFocus()
    COD$ = Text1.TEXT
    If CODINT%(COD$) < 1 Then
        Text1.SetFocus
        Exit Sub
    End If
    ILISTA& = List1.ListIndex
    If ILISTA& < 0 Then
        List1.SetFocus
        Exit Sub
    End If
    Text9.SetFocus
End Sub
'

Private Sub Text7_Change()
    If MODENTRAJU% = 7 Then
        Text8.TEXT = ""
        CAJUS# = Val(Text6.TEXT) + Val(Text7.TEXT) - Val(Text8.TEXT)
        Text9.TEXT = Str$(CAJUS#)
    End If
End Sub
'

Private Sub Text8_Change()
    If MODENTRAJU% = 8 Then
        Text7.TEXT = ""
        CAJUS# = Val(Text6.TEXT) + Val(Text7.TEXT) - Val(Text8.TEXT)
        Text9.TEXT = Str$(CAJUS#)
    End If
End Sub
'

Private Sub Text9_Change()
    If MODENTRAJU% = 9 Then
       DIFER# = Val(Text9.TEXT) - Val(Text6.TEXT)
       Text7.TEXT = ""
       Text8.TEXT = ""
       If DIFER# > 0 Then
           Text7.TEXT = Str$(DIFER#)
         ElseIf DIFER# < 0 Then
           Text8.TEXT = Str$(Abs(DIFER#))
       End If
    End If
End Sub
'
Private Sub Text7_GotFocus()
    COD$ = Text1.TEXT
    If CODINT%(COD$) < 1 Then
        Text1.SetFocus
        Exit Sub
    End If
    ILISTA& = List1.ListIndex
    If ILISTA& < 0 Then
        List1.SetFocus
        Exit Sub
    End If
    Text7.SelStart = 0
    Text7.SelLength = Len(Text7.TEXT)
    MODENTRAJU% = 7
End Sub

Private Sub Text8_GotFocus()
    COD$ = Text1.TEXT
    If CODINT%(COD$) < 1 Then
        Text1.SetFocus
        Exit Sub
    End If
    ILISTA& = List1.ListIndex
    If ILISTA& < 0 Then
        List1.SetFocus
        Exit Sub
    End If
    Text8.SelStart = 0
    Text8.SelLength = Len(Text8.TEXT)
    MODENTRAJU% = 8
End Sub
'

Private Sub Text9_GotFocus()
    COD$ = Text1.TEXT
    If CODINT%(COD$) < 1 Then
        Text1.SetFocus
        Exit Sub
    End If
    ILISTA& = List1.ListIndex
    If ILISTA& < 0 Then
        List1.SetFocus
        Exit Sub
    End If
    Text9.SelStart = 0
    Text9.SelLength = Len(Text9.TEXT)
    MODENTRAJU% = 9
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command3.SetFocus
    End If
End Sub
'
Sub AJUSTO1()
    '
    HOOI% = HOY(HOS$)
    '
    List1.Clear
    Call CARDEPOS
    '
111 If NUSU% < 1 Then
        Call MENSERR(24, "Falta Definir\Depósitos y Sucursales.")
        Hide
        Call FRESHALL
        PREPA% = 0
        Exit Sub
    End If
    '
115 Text10.TEXT = HOS$
    Text5.TEXT = "Diferencia de Inventario"
    '
End Sub
'

Sub AJUSTO2()
    '
    FF% = FECHANUM(Text10.TEXT)
    If FF% < 1 Then
        Text10.TEXT = HOS$
        Text10.SetFocus
        Exit Sub
    End If
    '
    CI% = CODINT%(Text1)
    If CI% > 0 Then
      '
      Screen.MousePointer = 11
      Call VESARRAS(CI%)
      For KI% = 1 To 64
        CATEO#(KI%) = 0
      Next KI%
      '
      Call ABRESTOCKS
      SQLX$ = " SELECT * FROM MOVISTO "
      SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
      SQLX$ = SQLX$ + "AND INT(CDBL(FechMov))<= " & CDbl(CVDAT(FF%)) & " "
      '
      Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
      With MoviTemp
        On Error Resume Next
        .MoveFirst
        If Err < 1 Then
          Do While Not .EOF
            CANTI# = !CantIngre - !Cantsalid
            CANTIS = !CantIngre - !Cantsalid
            DEP% = !DepoMovi
            LOT$ = !IdenLote
            '
            COLUX% = 1
            For KI% = 1 To NUSU%
              If DEP% = COSU%(KI%) Then COLUX% = KI%
            Next KI%
            '
            CATEO#(COLUX%) = CATEO#(COLUX%) + CANTIS
          .MoveNext
          Loop
        End If
      End With
    End If
    '
    List1.Clear
    For U& = 1 To NUSU%
       CAN = CATEO#(U&)
       List1.AddItem SUCU$(U&) + CSTRING$(MKS$(CAN), 4, 10, 1, 2)
    Next U&
    '
    List1.ListIndex = ILISTA&
    List1.Refresh
    Text4.TEXT = Left$(List1.TEXT, 20)
    Text6.TEXT = TRIM$(Mid$(List1.TEXT, 21))
    Text7.TEXT = ""
    Text8.TEXT = ""
    Text9.TEXT = TRIM$(Mid$(List1.TEXT, 21))
    If XVALO(Text9.TEXT) < 0 Then
        Text7.TEXT = TRIM$(CSTRING$(MKD$((-1) * XVALO(Text9.TEXT)), 4, 10, 1, 2))
        Text9.TEXT = "0.0"
    End If
    Text9.SetFocus
    '
    Screen.MousePointer = 1
    '
End Sub
'

Sub AJUSTO3()
    '
    Static YAFORI%, FORIPRE$   ', UREMOVI&
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORAJUST"))
    YAFORI% = 1
    If FORIPRE$ <> "" Then
      LU$ = LUDAT$
      If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
        RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
        Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Vale de Ajuste de Stocks:\  \'" + RMCC$ + "'.")
        YAFORI% = (-1)
      End If
    End If
    '
    HOII% = HOY(HOS$)
    '
    Call ABRESTOCKS
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + "WHERE CodiMovi = 'AJ'"
    '\\\OT-05-0303-WAR-19/05/05///
    SQLX$ = SQLX$ + "ORDER BY VAL(MID$(DoPriCor,4)) DESC"
    '\\\OT-05-0303-WAR-FIN///
    UNUM& = 0
    Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    On Error Resume Next
    MoviTemp.MoveFirst
    If Err < 1 Then
      UNUM& = XVALO(Mid$(MoviTemp!DoPriCor, 4))
    End If
    On Error GoTo 0
    NRO& = 1 + UNUM&
    MoviTemp.Close
    '
    CAVIE = XVALO(Text6.TEXT)
    CANUE = XVALO(Text9.TEXT)
    '
    FF% = FECHANUM(Text10.TEXT)
    If FF% < 1 Then
      Call MENSERR(24, "Fecha Incorrecta")
      Text10.SetFocus
      Exit Sub
    End If
    '
    CIII% = CODINT%(Text1.TEXT)
    If CIII% < 1 Then
      Call MENSERR(24, "Falta Código de Material")
      Text1.SetFocus
      Exit Sub
    End If
    '
    LTX$ = LOTE$
    If Left$(UCase$(LOTE$), 5) = "TODOS" Or TRIM$(LOTE$) = "" Then
      LTX$ = Space$(16)
    End If
    ILISTA& = List1.ListIndex
    DEPO% = COSU%(ILISTA& + 1)
    FECHA% = FF%
    CANTI = CANUE - CAVIE
    '
    If Abs(CANTI) >= 0.05 Then
        '
        CMOV$ = TRIM$(CONTROL$("", "CODAJUST"))
        If CMOV$ = "" Then CMOV$ = "AJ"
        DOSECO$ = "AJ." + Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Left$(HOS$, 2)
        DOPRICO$ = "AS." + TRIM$(Str$(NRO&))
        NUORIT% = 0
        VUNI# = 0
        MONEII% = 0
        NC% = 0
        REFE$ = Text5.TEXT
        If TRIM$(REFE$) = "" Then
            REFE$ = "Ajuste de Inventario"
        End If
        '
        Call MOVISTOK(FF%, CIII%, LTX$, CMOV$, DOPRILA$, DOPRICO$, DOSELA$, DOSECO$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPO%, CANTI)
        '
237     If TRIM$(FORIPRE$) <> "" Then
          If YAFORI% > 0 Then
             LU$ = LUDAT$
             CODFOR$ = TRIM$(UCase$(FORIPRE$))
             '
             YAC = XVALO(ATRIFORM$(CODFOR$, "DETDESMM"))
             MARIZW = XVALO(ATRIFORM$(CODFOR$, "MARIZQMM"))
             MARSUW = XVALO(ATRIFORM$(CODFOR$, "MARSUPMM"))
             '
             INTERLI = XVALO(ATRIFORM$(CODFOR$, "INTERLIN"))
             If INTERLI < 1 Then
                Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
                INTERLI = 1
             End If
             '
             CACOPI% = XVALO(ATRIFORM$(CODFOR$, "CAN-COP"))
             If CACOPI% < 1 Or CACOPI% > 6 Then
                Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
                CACOPI% = 1
             End If
             '
             For COP% = 1 To CACOPI%
               '
               Call CARFORWIN
               '
               FEX = FECHA%
               Call PRINTFORWIN(FECHATEX$(FEX), "FECH-EMI", 0, 0)
               Call PRINTFORWIN(Time$, "HORA-EMI", 0, 0)
               Call PRINTFORWIN(USERNAME$, "USUARIO", 0, 0)
               Call PRINTFORWIN(Str$(NRO&), "NUME-COM", 0, 0)
               Call PRINTFORWIN(TRIM$(CODEXT$(CIII%)), "COD-MAT", 0, 0)
               Call PRINTFORWIN(DESCRIT$(CIII%), "DES-MAT", 0, 0)
               Call PRINTFORWIN(UNIMED$(CIII%), "UNIMED", 0, 0)
               Call PRINTFORWIN(ECOARCH$("ECODEP", Chr$(DEPO%)), "DEPOSITO", 0, 0)
               Call PRINTFORWIN(LTX$, "NRO-LOTE", 0, 0)
               CANS$ = CSTRING$(MKS$(CANTI), 3, 9, 1, 2)
               Call PRINTFORWIN(CANS$, "CANTIDAD", 0, 0)
               Call PRINTFORWIN(REFE$, "REF-MAT1", 0, 0)
               '
               'Printer.EndDoc
               Call SETSPOOL("ENDDOC", "")
               '
             Next COP%
             Call SAVESPOOL("Ajuste de Stocks" + Str$(NRO&), OKX%)
          End If
        End If
        '
        Call COMUNI("Ajuste Número" + Str$(NRO&) + "\Completo y Verificado")
        '
299 End If
    '
    Call CIERRARCH("*.*")
    '
End Sub
'



