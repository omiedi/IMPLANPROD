VERSION 5.00
Begin VB.Form ANACOCHE 
   BackColor       =   &H00EEEEC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Análisis de Cobranza"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5265
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
   ScaleHeight     =   2820
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4220
      ScaleHeight     =   75
      ScaleWidth      =   810
      TabIndex        =   6
      Top             =   1890
      Width           =   875
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00EEEEC0&
      Caption         =   "Período"
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
      Left            =   160
      TabIndex        =   3
      Top             =   210
      Width           =   3900
      Begin VB.TextBox MESAN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1365
         TabIndex        =   9
         Text            =   " "
         Top             =   840
         Width           =   2220
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
         Height          =   360
         Left            =   3570
         TabIndex        =   8
         Top             =   840
         Width           =   175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mes y año:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   -945
         TabIndex        =   10
         Top             =   945
         Width           =   2220
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
      Left            =   160
      TabIndex        =   2
      Top             =   1785
      Width           =   3900
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
         Left            =   2100
         TabIndex        =   5
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
         Left            =   735
         TabIndex        =   4
         Top             =   420
         Width           =   1170
      End
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
      Left            =   4220
      Picture         =   "ANACOCHE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Emite Reporte Solicitado"
      Top             =   1155
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
      Left            =   4220
      Picture         =   "ANACOCHE.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Abandona la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   4050
      Picture         =   "ANACOCHE.frx":0454
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   4725
   End
End
Attribute VB_Name = "ANACOCHE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command10_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   Call ANACOCHEQ
   Command12.Enabled = True
End Sub

Private Sub Command4_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
        Exit Sub
    End If
    MESAN.Text = MESASEL$
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID
    Call TRANSLABELS(Name)
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    FEX = HOY(HO$)
    Text3 = FECHATEX$(FEX)
    Text2 = "01" + Mid$(Text3, 3)
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Sub ANACOCHEQ()
    '
    Dim INIPER%(12), FINPER%(12)
    MMXX0$ = TRIM$(MESAN.Text)
    mmxx1$ = MESANO$(MMXX0$, DES%, HAS%)
    If TRIM$(mmxx1$) = "" Then
        Call MENSERR(24, "Ingrese el Mes a Listar")
        Exit Sub
    End If
    '
    If Option11.Value = True Then
        Call MENSERR(24, "Imposible por Pantalla.\Exporte a Archivo de Texto.")
        Exit Sub
    End If
    '
    REAVA& = 0
    Screen.MousePointer = 11
    '
    DD1% = DES%
    FEX = DD1%
    SMX0% = SEMANU%(FEX, DELX, ALX)
    INIPER%(0) = DELX: FINPER%(0) = ALX
    For KKI% = 1 To 5
      FEX = DIPOST(ALX)
      SMX1% = SEMANU%(FEX, DELX, ALX)
      INIPER%(KKI%) = DELX: FINPER%(KKI%) = ALX
    Next KKI%
    '
    FEX = INIPER%(3)
    FETE$ = FECHATEX$(FEX)
    MACT% = 1 + Val(Mid$(FETE$, 4, 2))
    If MACT% > 12 Then MACT% = 1
    '
    FEX = FINPER%(4)
    INIPER%(5) = DIPOST(FEX)
    While Val(Mid$(FECHATEX$(FEX), 4, 2)) <> MACT%
      FEX = DIPOST(FEX)
    Wend
10  FEY = DIPOST(FEX)
    If Val(Mid$(FECHATEX$(FEY), 4, 2)) = MACT% Then
       FEX = FEY
       GoTo 10
    End If
    FINPER%(5) = FEX
    '
    MACT% = MACT + 1
    If MACT% > 12 Then MACT% = 1
    '
    FEX = FINPER%(5)
    INIPER%(6) = DIPOST(FEX)
    While Val(Mid$(FECHATEX$(FEX), 4, 2)) <> MACT%
      FEX = DIPOST(FEX)
    Wend
20  FEY = DIPOST(FEX)
    If Val(Mid$(FECHATEX$(FEY), 4, 2)) = MACT% Then
       FEX = FEY
       GoTo 20
    End If
    FINPER%(6) = FEX
    '
    MACT% = MACT + 1
    If MACT% > 12 Then MACT% = 1
    '
    FEX = FINPER%(6)
    INIPER%(7) = DIPOST(FEX)
    While Val(Mid$(FECHATEX$(FEX), 4, 2)) <> MACT%
      FEX = DIPOST(FEX)
    Wend
30  FEY = DIPOST(FEX)
    If Val(Mid$(FECHATEX$(FEY), 4, 2)) = MACT% Then
       FEX = FEY
       GoTo 30
    End If
    FINPER%(7) = 32767
    '
    JJ& = 0
    REBLA$ = REGBLAN$("ANACHERE")
    For U& = 1 To ULTREG&("CHECAR")
      X$ = REGLEIDO$("CHECAR", U&)
      FECOBRA% = CVI(Mid$(X$, 45, 2))
      If FECOBRA% >= DES% Then
        If FECOBRA% <= HAS% Then
          NURECI = CVS(Mid$(X$, 47, 4))
          NCLIE% = CVI(Mid$(X$, 3, 2))
          BANCHE$ = TRIM$(Mid$(X$, 17, 16))
          If BANCHE$ = "" Then BANCHE$ = TRIM$(ECOARCH$("MECOBRA", Left$(X$, 2)))
          NUCHEBA$ = BANCHE$
          If NUCHEBA$ <> "" And TRIM$(Mid$(X$, 5, 12)) <> "" Then NUCHEBA$ = NUCHEBA$ + " - "
          NUCHEBA$ = NUCHEBA$ + TRIM$(Mid$(X$, 5, 12))
          IMPOCHE# = CVD(Mid$(X$, 33, 8))
          FACRE% = CVI(Mid$(X$, 41, 2))
          If FACRE% < INIPER%(0) Then FACRE% = INIPER%(0)
          '
          Y$ = REBLA$
          Call REPLA(Y$, MKS$(NURECI), 1, 4)
          Call REPLA(Y$, MKI$(FECOBRA%), 5, 2)
          Call REPLA(Y$, MKI$(NCLIE%), 7, 2)
          Call REPLA(Y$, NUCHEBA$, 9, 32)
          Call REPLA(Y$, MKI$(FACRE%), 41, 2)
          For KKI% = 0 To 7
            If FACRE% <= FINPER%(KKI%) Then
              Call REPLA(Y$, MKD$(IMPOCHE#), 43 + 8 * KKI%, 8)
              GoTo 40
            End If
          Next KKI%
40        Call REPLA(Y$, MKD$(IMPOCHE#), 107, 8)
          JJ& = JJ& + 1
          Call GRAREG("ANACHERE", Y$, JJ&)
        End If
      End If
    Next U&
    '
    Call SETULTREG("ANACHERE", JJ&)
99  Call CIERRARCH("*.*")
    Call HEADERS("ANACHERE", "")
    Call HEADERS("ANACHERE", "Corresponde a: " + mmxx1$)
    Call FINAL("*LIANACHE")
    Screen.MousePointer = 1
   '
End Sub

Private Sub List1_Click()

End Sub

Private Sub MESAN_DblClick()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
        Exit Sub
    End If
    MESAN.Text = MESASEL$
End Sub
