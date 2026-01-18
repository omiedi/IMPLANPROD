VERSION 5.00
Begin VB.Form RESTACO 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estadísticas de Compras"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5640
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4620
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   16
      Top             =   2625
      Width           =   855
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
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
      Height          =   855
      Left            =   210
      TabIndex        =   10
      Top             =   3150
      Width           =   4215
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
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
         Index           =   8
         Left            =   2415
         TabIndex        =   12
         Top             =   420
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
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
         Index           =   4
         Left            =   525
         TabIndex        =   11
         Top             =   420
         Width           =   1380
      End
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   4620
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   2415
      Visible         =   0   'False
      Width           =   855
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
      Height          =   855
      Left            =   4620
      Picture         =   "RESTACO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   1260
      Width           =   855
   End
   Begin VB.CommandButton Command2 
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
      Left            =   4620
      Picture         =   "RESTACO.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Terminar - Salir de esta Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ranking de Compras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   4215
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1365
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   18
         Top             =   1995
         Width           =   2430
         Begin VB.CommandButton Command20 
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
            Picture         =   "RESTACO.frx":0454
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Consulta de Ultima Factura por Proveedor"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Ultima Factura por Proveedor"
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
            Width           =   1485
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1050
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   1
         Top             =   1470
         Width           =   2430
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
            Picture         =   "RESTACO.frx":075E
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Ranking de Compras por Rubro por Rango de Fechas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Ranking de Com- pras por Rubro"
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
            TabIndex        =   3
            Top             =   60
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   4
         Top             =   945
         Width           =   2430
         Begin VB.CommandButton Command6 
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
            Picture         =   "RESTACO.frx":0A68
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Ranking de Compras por Proveedor por Rango de Fechas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Ranking Compras   por Proveedor"
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
            TabIndex        =   6
            Top             =   50
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   13
         Top             =   420
         Width           =   2430
         Begin VB.CommandButton Command1 
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
            Picture         =   "RESTACO.frx":0D72
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Resumen de Compras por Proveedor por Rango de Fechas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Resumen de  Com- pras por Proveedor"
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
            Left            =   680
            TabIndex        =   15
            Top             =   45
            Width           =   1905
         End
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4050
      Picture         =   "RESTACO.frx":107C
      Top             =   3485
      Width           =   2010
   End
End
Attribute VB_Name = "RESTACO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Command1.Enabled = False
    Call RESUCOMPRO
    Command1.Enabled = True
End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command20_Click()
  '\\\OT-06-0223-WAR-02/06/06///
Dim UFAC$(32767)
Dim FEUF%(32767)
Dim IUFAC#(32767)
  Command20.Enabled = False
  FIN& = ULTREG&("CUECORR")
  For IU& = 1 To FIN&
    Call TRACE(24, IU&, FIN&)
    X$ = REGLEIDO$("CUECORR", IU&)
    NP% = CVI(Mid$(X$, 5, 2))
    COMO1$ = Left$(Mid$(X$, 7, 18), 2)
    If NP% > 0 Then
      If COMO1$ <> "OP" And COMO1$ <> "FF" And COMO1$ <> "AS" Then
        F1% = CVI(Mid$(X$, 1, 2))
        If F1% >= FEUF%(NP%) Then
          '
          IMPO12# = CVD(Mid$(X$, 123, 8)) - CVD(Mid$(X$, 115, 8))
          If Abs(IMPO12#) > 0.005 Then
            COMO2$ = Mid$(X$, 7, 18)
            '
            UFAC$(NP%) = COMO2$
            FEUF%(NP%) = F1%
            IUFAC#(NP%) = IMPO12#
          End If
          '
        End If
      End If
    End If
  Next IU&
  '
  YBLA$ = REGBLAN$("UFACPRO")
  JJ& = 0
  For I% = 1 To 32767
    If I% = 32767 Then Exit For
    '
    If TRIM$(UFAC$(I%)) <> "" Then
      Y$ = YBLA$
      Call REPLA(Y$, MKI$(I%), 1, 2)
      Call REPLA(Y$, MKI$(FEUF%(I%)), 3, 2)
      Call REPLA(Y$, UFAC$(I%), 5, 18)
      Call REPLA(Y$, MKD$(IUFAC#(I%)), 23, 8)
      JJ& = JJ& + 1
      Call GRAREG("UFACPRO", Y$, JJ&)
    End If
    '
  Next I%
  '
  Call SETULTREG("UFACPRO", JJ&)
  Call FILESORT("UFACPRO", "UFACPRO", 1, 1, "ASC")
  Call FINAL("*UFACPRO")
  '
  Command20.Enabled = True
  '\\\OT-06-0223-WAR-FIN///
End Sub

Private Sub Command29_Click()
    Command29.Enabled = False
    Call HELPONLINE("RESTACOM")
    Command29.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call RANKPROVE
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    Call RANKGRUCOM
    Command7.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub
'
Sub RANKPROVE()
    '
    Dim TOTVENTA#(32767)
    '
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    J& = 0
    '
    CUEADU$ = AJUSTI$(CONTROL$("SUIVACO", "CUEADU"), 12)
    CUEADU1% = 0
    For U& = 1 To ULTREG&("ECUECON")
       X$ = REGLEIDO$("ECUECON", U&)
       If Left$(X$, 12) = CUEADU$ Then CUEADU1% = CVI(Mid$(X$, 41, 2)): DECADU$ = Mid$(X$, 13, 28)
    Next U&
    '
    FIN& = ULTREG&("CUECORR")
    TOTAGEN# = 0
    For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("CUECORR", U&)
        FF% = CVI(Left$(X$, 2))
        If FF% >= DES% Then
           If FF% <= HAS% Then
              NC% = CVI(Mid$(X$, 5, 2))
              If NC% > 0 Then
                 If NC% <= 32767 Then
                    SIG% = 1
                    TIX$ = UCase$(Mid$(X$, 7, 2))
                    CA$ = UCase$(Mid$(X$, 10, 1))
                    If TIX$ = "NC" Or TIX$ = "DP" Or TIX$ = "DC" Then SIG% = (-1)
                    '
                    If TIX$ <> "VC" Then
                     If TIX$ <> "FC" Then
                      If TIX$ <> "TF" Then
                       If TIX$ <> "RH" Then
                        If TIX$ <> "ND" Then
                         If TIX$ <> "NC" Then
                          If TIX$ <> "DP" Then
                           If TIX$ <> "DC" Then
                            GoTo 19
                           End If
                          End If
                         End If
                        End If
                       End If
                      End If
                     End If
                    End If
                    '
                    NETO# = (Int(100 * (Abs(CVD(Mid$(X$, 51, 8))) + Abs(CVD(Mid$(X$, 59, 8))) + Abs(CVD(Mid$(X$, 67, 8)))) + 0.5)) / 100
                    If CA$ = "C" Then NETO# = NETO# + Abs(CVD(Mid$(X$, 75, 8)))
                    '
                    Y6$ = Mid$(X$, 199, 80): IVADU# = 0
                    For UI% = 1 To 80 Step 10
                       If CVI(Mid$(Y6$, UI%, 2)) = CUEADU1% Then
                          IVADU# = IVADU# + Abs(CVD(Mid$(Y6$, UI% + 2, 8)))
                       End If
                    Next UI%
                    If IVADU# > NETO# Then IVADU# = NETO#
                    NETO# = NETO# - IVADU#
                    '
                    TOTVENTA#(NC%) = TOTVENTA#(NC%) + NETO# * SIG%
                    TOTAGEN# = TOTAGEN# + NETO# * SIG%
                 End If
              End If
           End If
        End If
19  Next U&
    '
    For NCI% = 1 To 32766
       If Abs(TOTVENTA#(NCI%)) > 0.005 Then
          RELIB$ = "A" + CSTRING$(MKD$(TOTAGEN# - TOTVENTA#(NCI%)), 4, 12, 2, 2)
          PORPAR = 100 * TOTVENTA#(NCI%) / TOTAGEN#
          RELIB$ = RELIB$ + CSTRING$(MKI$(NCI%), 1, 6, 0, 2)
          RELIB$ = RELIB$ + CSTRING$(MKS$(PORPAR), 4, 8, 2, 2)
          RELIB$ = RELIB$ + CSTRING$(MKD$(TOTVENTA#(NCI%)), 4, 12, 2, 2)
          RESTACO.List1.AddItem RELIB$
       End If
    Next NCI%
    '
    J& = 0
    FIN& = RESTACO.List1.ListCount
    '
    For U& = 1 To FIN&
        RESTACO.List1.ListIndex = U& - 1
        NC% = Val(Mid$(RESTACO.List1.Text, 14, 6))
        TONEVEN# = Val(Mid$(RESTACO.List1.Text, 28, 12))
        PORPAR = Val(Mid$(RESTACO.List1.Text, 20, 8))
        PORACU = PORACU + PORPAR
        '
        Z$ = REGBLAN$("RANKPRO")
        J& = J& + 1: JI% = J&
        Call REPLA(Z$, MKI$(JI%), 1, 2)
        Call REPLA(Z$, MKI$(NC%), 3, 2)
        Call REPLA(Z$, MKD$(TONEVEN#), 5, 8)
        Call REPLA(Z$, MKS$(PORPAR), 13, 4)
        Call REPLA(Z$, MKS$(PORACU), 17, 4)
        Call GRAREG("RANKPRO", Z$, J&)
        '
    Next U&
    '
    Call SETULTREG("RANKPRO", J&)
    Call CIERRARCH("RANKPRO")
    Call HEADERS("RANKPRO", "")
    Call HEADERS("RANKPRO", "Periodo: " + PERIO$)
    Screen.MousePointer = 1
    RESTACO.List1.Clear
    '
    If RESTACO.Option1(4).Value = True Then
         Call FINAL("*RANKPRO")
       Else
         Call FINAL("*LRANPRO")
    End If
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    '
End Sub
'
Sub RANKGRUCOM()

End Sub

Sub RESUCOMPRO()
    '
    Dim PUNCOMPRA$(32767)
    '
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        Exit Sub
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("CUECORR")
    '
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      SS0$ = REGLEIDO$("CUECORR", U&)
      FEMI% = CVI(Left$(SS0$, 2))
      If FEMI% >= DES% Then
        If FEMI% <= HAS% Then
          NCI% = CVI(Mid$(SS0$, 5, 2))
          TI$ = Mid$(SS0$, 7, 2)
          If TI$ = "VC" Or TI$ = "FC" Or TI$ = "TF" Or TI$ = "RH" Or TI$ = "ND" Or TI$ = "NC" Or TI$ = "DP" Then
            PUNCOMPRA$(NCI%) = PUNCOMPRA$(NCI%) + MKS$(U&)
          End If
        End If
      End If
    Next U&
    '
    JJ& = 0
    FIN& = ULTREG&("PROVED1")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("PROVED1", U&)
      NCI% = CVI(Left$(X$, 2))
      If NCI% > 0 Then
        If NCI% <= 32767 Then
          For KKI% = 1 To Len(PUNCOMPRA$(NCI%)) Step 4
            REGCUEC& = CVS(Mid$(PUNCOMPRA$(NCI%), KKI%, 4))
            SS0$ = REGLEIDO$("CUECORR", REGCUEC&)
            NUFA$ = Mid$(SS0$, 7, 18)
            FEMI% = CVI(Left$(SS0$, 2))
            SH1$ = Mid$(SS0$, 123, 8)
            SD1$ = Mid$(SS0$, 115, 8)
            IMPO# = CVD(SH1$) - CVD(SD1$)
            '
            XXX$ = REGBLAN$("RECOPRO")
            Call REPLA(XXX$, MKI$(NCI%), 1, 2)
            Call REPLA(XXX$, MKI$(FEMI%), 3, 2)
            Call REPLA(XXX$, NUFA$, 5, 18)
            Call REPLA(XXX$, MKD$(IMPO#), 23, 8)
            JJ& = JJ& + 1
            Call GRAREG("RECOPRO", XXX$, JJ&)
          Next KKI%
        End If
      End If
    Next U&
    '
    Call SETULTREG("RECOPRO", JJ&)
    Call CIERRARCH("RECOPRO")
    Call HEADERS("RECOPRO", "")
    Call HEADERS("RECOPRO", "Periodo: " + PERIO$)
    '
    Call FINAL("*SOCOPRO")
    Screen.MousePointer = 1
    '
End Sub
