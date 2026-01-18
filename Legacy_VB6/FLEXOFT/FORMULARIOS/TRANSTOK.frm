VERSION 5.00
Begin VB.Form TRANSTOK 
   BackColor       =   &H00E0E8FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transferencias de Stocks"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5610
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
   ScaleHeight     =   5670
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
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
      Height          =   800
      Left            =   2730
      Picture         =   "TRANSTOK.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Control de Transferencias por Rango de Fechas"
      Top             =   210
      Width           =   800
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4500
      ScaleHeight     =   75
      ScaleWidth      =   900
      TabIndex        =   26
      Top             =   1680
      Width           =   960
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Direct"
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
      Left            =   1890
      Picture         =   "TRANSTOK.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   210
      Width           =   800
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
      Left            =   210
      Picture         =   "TRANSTOK.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   210
      Width           =   800
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Kardex"
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
      Left            =   1050
      Picture         =   "TRANSTOK.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   210
      Width           =   800
   End
   Begin VB.CommandButton Command6 
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
      Left            =   3780
      Picture         =   "TRANSTOK.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   210
      Width           =   800
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E8FF&
      Caption         =   "Operaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   210
      TabIndex        =   17
      Top             =   1260
      Width           =   4110
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   18
         Top             =   420
         Width           =   2535
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
            Picture         =   "TRANSTOK.frx":106A
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Informar Trans- ferencias de Stock"
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
            TabIndex        =   19
            Top             =   105
            Width           =   1800
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E8FF&
      Caption         =   "Consultas - Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2805
      Left            =   210
      TabIndex        =   15
      Top             =   2730
      Width           =   5220
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E8FF&
         Caption         =   "Rango de Fechas"
         Height          =   1335
         Left            =   2580
         TabIndex        =   23
         Top             =   420
         Width           =   2430
         Begin VB.CommandButton BOTSEL1 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1995
            Picture         =   "TRANSTOK.frx":1374
            TabIndex        =   6
            Top             =   840
            Width           =   175
         End
         Begin VB.CommandButton BOTSEL 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1995
            Picture         =   "TRANSTOK.frx":167E
            TabIndex        =   4
            Top             =   420
            Width           =   175
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   945
            TabIndex        =   3
            Text            =   " "
            Top             =   420
            Width           =   1065
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   945
            TabIndex        =   5
            Text            =   " "
            Top             =   840
            Width           =   1065
         End
         Begin VB.Label Label2 
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
            Left            =   0
            TabIndex        =   25
            Top             =   420
            Width           =   855
         End
         Begin VB.Label Label1 
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
            Left            =   105
            TabIndex        =   24
            Top             =   840
            Width           =   750
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E8FF&
         Caption         =   "Rango Numérico"
         Height          =   1335
         Left            =   210
         TabIndex        =   20
         Top             =   420
         Width           =   2220
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   945
            TabIndex        =   2
            Top             =   840
            Width           =   1065
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   945
            TabIndex        =   1
            Top             =   420
            Width           =   1065
         End
         Begin VB.Label Label8 
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
            Left            =   105
            TabIndex        =   22
            Top             =   840
            Width           =   750
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
            Left            =   0
            TabIndex        =   21
            Top             =   420
            Width           =   855
         End
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
         Left            =   4095
         Picture         =   "TRANSTOK.frx":1988
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1995
         Width           =   855
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E8FF&
         Caption         =   "Salida"
         Height          =   705
         Left            =   210
         TabIndex        =   16
         Top             =   1890
         Width           =   3690
         Begin VB.OptionButton Option5 
            BackColor       =   &H00E0E8FF&
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
            TabIndex        =   7
            Top             =   315
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00E0E8FF&
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
            Left            =   2100
            TabIndex        =   8
            Top             =   315
            Width           =   1275
         End
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
      Height          =   800
      Left            =   4620
      Picture         =   "TRANSTOK.frx":1C92
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   210
      Width           =   800
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3990
      Picture         =   "TRANSTOK.frx":1DDC
      Top             =   2005
      Width           =   2010
   End
End
Attribute VB_Name = "TRANSTOK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RUSANJUA$

Private Sub BOTSEL_Click(Index As Integer)
    Call SELECHO("SELFECHA")
    Text2.TEXT = VALACT1$("SELFECHA")
End Sub

Private Sub BOTSEL1_Click(Index As Integer)
    Call SELECHO("SELFECHA")
    Text1.TEXT = VALACT1$("SELFECHA")
End Sub

Private Sub Command1_Click()
    Hide
    Call FRESHALL
End Sub

Private Sub Command2_Click()
    DESNUM& = 1: If Val(Text7.TEXT) > 0 Then DESNUM& = Val(Text7.TEXT)
    HASNUM& = 999999
    If TRIM$(Text8.TEXT) <> "" Then
        If Val(Text8.TEXT) >= DESNUM& Then HASNUM& = Val(Text8.TEXT)
    End If
    DESFE% = 0: If FECHANUM(TRIM$(Text2.TEXT)) > 0 Then DESFE% = FECHANUM(TRIM$(Text2.TEXT))
    HASFE% = 32767
    If TRIM$(Text1.TEXT) <> "" Then
        If FECHANUM(TRIM$(Text1.TEXT)) >= DESFE% Then HASFE% = FECHANUM(TRIM$(Text1.TEXT))
    End If
    '
    Screen.MousePointer = 11
    Call BLANARCH("LISTRAJ")
    FIN& = ULTREG&("MOVISTO")
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("MOVISTO", I&)
      If UCase$(Mid$(X$, 21, 2)) = "TR" Then
        FF% = CVI(Left$(X$, 2))
        If FF% >= DESFE% Then
          If FF% <= HASFE% Then
            CATRA# = CVD(Mid$(X$, 96, 8))
            If CATRA# >= 0.05 Then
              NRO& = Val(Mid$(X$, 26, 7))
              If NRO& >= DESNUM& Then
                If NRO& <= HASNUM& Then
                  REFE$ = Mid$(X$, 5, 16)
                  CI1% = CVI(Mid$(X$, 3, 2))
                  DP1% = Asc(Mid$(X$, 83, 1))
                  Z$ = REGLEIDO$("MOVISTO", I& + 1)
                  CATRB# = CVD(Mid$(Z$, 88, 8))
                  If Abs(CATRB# - CATRA#) < 0.005 Then
                    CI2% = CVI(Mid$(Z$, 3, 2))
                    DP2% = Asc(Mid$(Z$, 83, 1))
                    CAN# = CATRA#
                    '
                    Y$ = REGBLAN$("LISTRAJ")
                    Call REPLA(Y$, MKI$(FF%), 1, 2)
                    Call REPLA(Y$, REFE$, 3, 48)
                    Call REPLA(Y$, MKS$(NRO&), 51, 4)
                    Call REPLA(Y$, MKI$(CI1%), 55, 2)
                    Call REPLA(Y$, UNIMED$(CI1%), 57, 4)
                    Call REPLA(Y$, MKD$(0), 61, 8)
                    Call REPLA(Y$, MKD$(CAN#), 69, 8)
                    Call REPLA(Y$, Chr$(DP1%), 80, 1)
                    Call REGAPP("LISTRAJ", Y$)
                    Call REPLA(Y$, MKI$(CI2%), 55, 2)
                    Call REPLA(Y$, UNIMED$(CI2%), 57, 4)
                    Call REPLA(Y$, MKD$(CAN#), 61, 8)
                    Call REPLA(Y$, MKD$(0), 69, 8)
                    Call REPLA(Y$, Chr$(DP2%), 80, 1)
                    Call REGAPP("LISTRAJ", Y$)
                  End If
                End If
              End If
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
        Call CONECRUN("*LITRAST", "")
    End If
    '
End Sub

Private Sub Command26_Click()
   Call CONECRUN("FLSETUP/OPTRAJUST", "Configuración de Funcionamiento")
   Call FINAL("")
   ' OPTRAJUST.Show 1
End Sub

Private Sub Command27_Click()
    Call ACCDIR("TRANSFE")
End Sub

Private Sub Command3_Click()
    Call TRASTOCK
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call GENKARDEX("")
    'Call CONECRUN("*KARDEX", "")
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Call OPNOIN("")
End Sub

Private Sub Command6_Click()
    Call HELPONLINE("TRANSFE")
End Sub

Private Sub Command7_Click()
  '
  Command7.Enabled = False
  '\\\OT-06-0092-WAR-FIN///
  Dim DOPRIM$(), DOSEC$(), CANTI()
  '
  Call DESHASFECHA(DES%, HAS%, PERIO$)
  If TRIM$(PERIO$) = "" Then GoTo 99
  Screen.MousePointer = 11
  NRO1 = 0
  '
  FIN& = ULTREG&("MOVISTO")
  PRIMOTRA& = 0
  For J& = 1 To FIN&
    Call TRACE(24, J&, FIN&)
    X$ = REGLEIDO$("MOVISTO", J&)
    F1% = CVI(Mid$(X$, 1, 2))
    CM1$ = Mid$(X$, 21, 2)
    If F1% >= DES% Then
      If F1% <= HAS% Then
        If CM1$ = "TR" Then
          NRO1 = NRO1 + 1
          If PRIMOTRA& = 0 Then PRIMOTRA& = J&
        End If
      End If
    End If
  Next J&
  '
  ReDim DOPRIM$(NRO1)
  ReDim DOSEC$(NRO1)
  ReDim CANTI(NRO1)
  '
  Screen.MousePointer = 1
  FIN& = ULTREG&("MOVISTO")
  CONTI1& = 0
  For J& = PRIMOTRA& To FIN&
    Call TRACE(24, J&, FIN&)
    X$ = REGLEIDO$("MOVISTO", J&)
    F1% = CVI(Mid$(X$, 1, 2))
    CM1$ = Mid$(X$, 21, 2)
    If F1% >= DES% Then
      If F1% <= HAS% Then
        If CM1$ = "TR" Then
          DOP1$ = Mid$(X$, 23, 10)
          DOS1$ = Mid$(X$, 33, 10)
          QE1 = CVD(Mid$(X$, 88, 8))
          QS1 = CVD(Mid$(X$, 96, 8))
          CAMO1 = QE1 - QS1
          YAGREGO% = 0
          For U& = 0 To NRO1
            For B& = U& To NRO1
              If DOPRIM$(B&) <> "" Then
                If DOPRIM$(B&) = DOP1$ Then
                  If DOSEC$(B&) = DOS1$ Then
                    CANTI(B&) = CANTI(B&) + CAMO1
                    GoTo 20
                  End If
                End If
              End If
            Next B&
            '
            If YAGREGO% = 0 Then
              DOPRIM$(CONTI1&) = DOP1$
              DOSEC$(CONTI1&) = DOS1$
              CANTI(CONTI1&) = CAMO1
              CONTI1& = CONTI1& + 1
              YAGREGO% = 1
              GoTo 20
            End If
          Next U&
        End If
      End If
    End If
20 Next J&
  '
  HUBODIF% = 0
  FIN& = CONTI1&
  For U& = 0 To CONTI1&
    Call TRACE(24, U&, FIN&)
    If CANTI(U&) <> 0 Then
      MENS1$ = "Transferencia no Balanceada: " + DOPRIM$(U&) + "\ Cantidad = " + Str(Abs(CANTI(U&))) + "."
      Call MENSERR(24, MENS1$)
      HUBODIF% = 1
    End If
  Next U&
  '
  If HUBODIF% = 0 Then
    Call COMUNI("No se Encontraron Transferencias no Balanceadas.")
  End If
  '
99 Command7.Enabled = True
  Screen.MousePointer = 1
  '\\\OT-06-0092-WAR-15/03/06///
End Sub

Private Sub Form_Load()
   '
   RUSANJUA$ = TRIM$(CONTROL$("", "INTREMOT"))
   If RUSANJUA$ <> "" Then
     If Mid$(RUSANJUA$, 2, 1) <> ":" Then
       RUSANJUA$ = Left$(LUDAT$, 2) + RUSANJUA$
     End If
   End If
   If RUSANJUA$ <> "" Then
     'Image2.ToolTipText = "Doble Click para Compensacion de Stocks"
   End If
   '
   Call TRANSLABELS(Name)
   '
   Option5.Value = True
End Sub

Private Sub Image2_DblClick()
  If RUSANJUA$ <> "" Then
    Screen.MousePointer = 11
    Dim CIMOVI%(32767), CAMOVI(32767), DEPOMO%(32767)
    '
    Dim BUFLE As String * 128
    Dim BUFMAS As String * 128
    N1% = FILEOPEN%(RUSANJUA$ + "\MOVISTO.DAT", 0, 128)
    N2% = FILEOPEN%(RUSANJUA$ + "\MASTER.DAT", 0, 128)
    UMASJ% = LOF(N2%) / 128 - 1
    '
    NMS% = NUMAS%
    HOII% = HOY(HOS$)
    CAMOI% = 0
    JJ& = LOF(N1%) / 128
    For U& = 1 To JJ&
      Get #N1%, U&, BUFLE$
      If CVI(Mid$(BUFLE$, 127, 2)) = 0 Then
        CISJ% = CVI(Mid$(BUFLE$, 3, 2))
        If CISJ% > 0 Then
          If CISJ% <= UMASJ% Then
            Get #N2%, CISJ% + 1, BUFMAS$
            CODSJ$ = TRIM$(Mid$(BUFMAS$, 77, 15))
            CIXI% = CODINT%(CODSJ$)
            If CISJ% <> CIXI% Then
              XX$ = BUFLE$
              Call REPLA(XX$, MKI$(CIXI%), 3, 2)
              BUFLE$ = XX$
              Put #N1%, U&, BUFLE$
            End If
            '
            CAMO = CVD(Mid$(BUFLE$, 88, 8)) - CVD(Mid$(BUFLE$, 96, 8))
            DEMO% = Asc(Mid$(BUFLE$, 83, 1))
            For KKI% = 1 To CAMOI%
              If CIMOVI%(KKI%) = CIXI% Then
                If DEPOMO%(KKI%) = DEMO% Then
                  CAMOVI(KKI%) = CAMOVI(KKI%) + CAMO
                  GoTo 19
                End If
              End If
            Next KKI%
            CAMOI% = CAMOI% + 1
            If CAMOI% > 10000 Then
              Call MENSERR(24, "Error Fatal - Consulte")
              Close: End
              Exit Sub
            End If
            CIMOVI%(CAMOI%) = CIXI%
            DEPOMO%(CAMOI%) = DEMO%
            CAMOVI(CAMOI%) = CAMO
            '
          End If
        End If
      End If
19  Next U&
    '
    For KKI% = 1 To CAMOI%
      If Abs(CAMOVI(KKI%)) >= 0.05 Then
        CIXI% = CIMOVI%(KKI%)
        Call MOVISTO(HOII%, CIXI%, "", "SJ", "TR-" + TRIM$(Str$(HOII%)), "TR-" + TRIM$(Str$(HOII%)), "Trans.SJ/" + CODEXT$(CIXI%), 0, "", 0, DEPOMO%(KKI%), CAMOVI(KKI%))
        URE& = ULTREG&("MOVISTO")
        XX$ = REGLEIDO$("MOVISTO", URE&)
        TTXX$ = Mid$(XX$, 96, 8) + Mid$(XX$, 88, 8)
        Call REPLA(XX$, TTXX$, 88, 16)
        Call REPLA(XX$, MKI$(HOII%), 127, 2)
        BUFLE$ = XX$
        JJ& = JJ& + 1
        Put #N1%, JJ&, BUFLE$
      End If
    Next KKI%
    Call CIERRARCH("MOVISTO")
    '
    For U& = 1 To JJ&
      Get #N1%, U&, BUFLE$
      If CVI(Mid$(BUFLE$, 127, 2)) = 0 Then
        XX$ = BUFLE$
        Call REPLA(XX$, MKI$(HOII%), 127, 2)
        BUFLE$ = XX$
        Put #N1%, U&, BUFLE$
      End If
    Next U&
    '
    Close #N1%
    Close #N2%
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    Call TRASAN("MASTER")
    Call TRASAN("INVERT")
    Call TRASAN("ECOMAST")
    Call TRASAN("CIACTI")
    Call TRASAN("ESTRUNUE")
    Call TRASAN("ESTRUDOS")
    Call TRASAN("IMPLOSI")
    Call TRASAN("LIPRE001")
    Call TRASAN("LIPRE002")
    Call TRASAN("LIPRE003")
    Call TRASAN("LIPRE004")
    Call TRASAN("SECOPER")
    '
    Screen.MousePointer = 1
    MsgBox "Proceso Completo"
  End If
End Sub

Sub TRASAN(ANAME$)
    If TRIM$(ANAME$) <> "" Then
       U1% = COPYFILE%(LUDAT + ANAME$ + ".DAT", RUSANJUA$ + "\" + ANAME$ + ".DAT")
    End If
End Sub
'
Private Sub Text1_DblClick()
    Call SELECHO("SELFECHA")
    Text1.TEXT = VALACT1$("SELFECHA")
End Sub

Private Sub Text2_DblClick()
    Call SELECHO("SELFECHA")
    Text2.TEXT = VALACT1$("SELFECHA")
End Sub
