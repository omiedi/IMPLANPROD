VERSION 5.00
Begin VB.Form ITECOM04 
   BackColor       =   &H00FFF0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de Items a Comprar"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
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
      Left            =   420
      MaxLength       =   5
      TabIndex        =   15
      Text            =   " "
      Top             =   555
      Width           =   960
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
      Height          =   285
      Left            =   1365
      TabIndex        =   14
      Top             =   555
      Width           =   195
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
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
      Left            =   1830
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   " "
      Top             =   210
      Width           =   4485
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
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
      Left            =   1830
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   " "
      Top             =   540
      Width           =   4485
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
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
      Left            =   1830
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   " "
      Top             =   870
      Width           =   750
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
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
      Left            =   2670
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   " "
      Top             =   870
      Width           =   3645
   End
   Begin VB.CommandButton Command9 
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
      Left            =   9240
      TabIndex        =   9
      Top             =   840
      Width           =   195
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00E0E0E0&
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
      Left            =   8295
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   " "
      Top             =   840
      Width           =   960
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  'Right Justify
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
      Left            =   8295
      TabIndex        =   7
      Text            =   " "
      Top             =   210
      Width           =   1120
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
      Height          =   3840
      Left            =   105
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   1365
      Width           =   9570
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00800000&
      Height          =   7575
      Left            =   9765
      ScaleHeight     =   7515
      ScaleWidth      =   1005
      TabIndex        =   0
      Top             =   0
      Width           =   1065
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
         Height          =   650
         Left            =   105
         Picture         =   "ITECOM04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Aprueba la Selección de Items a Comprar"
         Top             =   945
         Width           =   650
      End
      Begin VB.CommandButton BOTEXIT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "ITECOM04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Cancela la Pre-Selección"
         Top             =   210
         Width           =   650
      End
      Begin VB.CommandButton BOTREC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "ITECOM04.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Aprueba la Selección de Items Cumplidos"
         Top             =   6555
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   1
         Top             =   6090
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   2
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -240
         Picture         =   "ITECOM04.frx":091E
         Stretch         =   -1  'True
         Top             =   4725
         Width           =   1410
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   210
      TabIndex        =   19
      Top             =   315
      Width           =   960
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   6765
      TabIndex        =   18
      Top             =   900
      Width           =   1380
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Número O/Compra:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   6195
      TabIndex        =   17
      Top             =   285
      Width           =   2010
   End
   Begin VB.Label Label8 
      BackColor       =   &H00EEEEEE&
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
      Left            =   420
      TabIndex        =   16
      Top             =   870
      Width           =   1170
   End
End
Attribute VB_Name = "ITECOM04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTEXIT_Click()
    '
    Call CIERRARCH("*.*")
    Unload ITECOM04
    '
End Sub

Private Sub Command1_Click()
   '
   Command1.Enabled = False
   '
   NPRO% = Val(Text1)
   If REGICLI&((-1) * NPRO%) < 1 Then
     Call MENSERR(24, "Seleccione Proveedor")
     Call Command2_Click
     GoTo 99
   End If
   '
   HOII% = HOY(HOS$)
   JJ& = 0
   REBLA$ = REGBLAN$("OCOMDETA")
   For KKI% = 1 To List1.ListCount
     If List1.Selected(KKI% - 1) = True Then
       '
       TTX1$ = List1.List(KKI% - 1)
       CODE$ = Left$(TTX1$, 16)
       CANTU = Val(Mid$(TTX1$, 65, 12))
       FENTREIT% = FECHANUM(Mid$(TTX1$, 79, 8))
       If FENTREIT% < HOII% Then
         FENTREIT% = HOII%
       End If
       '
       ZZ$ = REBLA$
         CI0% = CODINT%(CODE$)
       Call REPLA(ZZ$, MKI$(CI0%), 9, 2)
       Call REPLA(ZZ$, UNIMED$(CI0%), 11, 4)
       Call REPLA(ZZ$, MKS$(CANTU), 51, 4)
       Call REPLA(ZZ$, UNIMED$(CI0%), 59, 4)
       Call REPLA(ZZ$, MKS$(1), 63, 4)
       Call REPLA(ZZ$, MKI$(FENTREIT%), 95, 2)
       '
       UC$ = FORCCOD04.ULTICOM$(NPRO%, CI0%)
       MONULCO% = Asc(Mid$(UC$, 122, 1)): If MONULCO% < 1 Then MONULCO% = 1
       PREUNID# = CVD(Mid$(UC$, 67, 8)) * TICAMONE(MONULCO%)
       PORDESCUEI = CVS(Mid$(UC$, 75, 4))
       If PREUNID# < 0.00005 Or MODOULPRE% = 2 Then
         PREUNID# = COSUNI(CI0%) * TICAMONE(MONECOSTO%(CI0%)) * COEFI
         PORDESCUEI = 0
       End If
       PREUNID# = PREUNID# / TICAMONE(MONEMI%)
       Call REPLA(ZZ$, MKD$(PREUNID#), 67, 8)
       Call REPLA(ZZ$, MKS$(PORDESCUEI), 75, 4)
       '
       Call REPLA(ZZ$, MKI$(KKI% - 1), 125, 2)
       '
       JJ& = JJ& + 1
       Call GRAREG("!OCOMDETA", ZZ$, JJ&)
       '
     End If
   Next KKI%
   Call SETULTREG("!OCOMDETA", JJ&)
   '
   If JJ& < 1 Then
       Call MENSERR(2, "No se han Marcado Items a Comprar.")
     Else
       FORCCOD04.Text1 = Text1
       Call FORCCOD04.CALTOCOM
       On Error GoTo 0
       FORCCOD04.Show 1
       Call BLANFORMU
       Call BOTEXIT_Click
   End If
99 Command1.Enabled = True
   '
End Sub

Private Sub Command2_Click()
   '
   Call SELECHO("ACREDOR")
   NCLIE% = Val(TRIM$(VALACT1$("ACREDOR")))
   If NCLIE% > 0 Then
     Text1 = TRIM$(Str$(NCLIE%))
   End If
   '
End Sub

Private Sub Command9_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text9.TEXT = VALACT1$("SELFECHA")
   End If
End Sub

Private Sub Form_Load()
    '
    Call BLANFORMU
    '
End Sub

Private Sub Text1_Change()
    '
    NCLIE% = Val(Text1)
    Call BLANFORMU
    If REGICLI&((-1) * NCLIE%) > 0 Then
      Call CARDACLI
    End If
    '
End Sub

Private Sub Text1_DblClick()
   Call Command2_Click
End Sub

Sub CARDACLI()
   '
   Screen.MousePointer = 11
   '
   NC% = (-1) * Val(Text1.TEXT)
   If NC% < 0 Then
     If STATCLI%(NC%) > 0 Then
       Label8 = CODICLI$(NC%)
       Text2.TEXT = RASOCLI$(NC%)
       Text3.TEXT = DOMICLI$(NC%)
       Text4.TEXT = CPOSCLI$(NC%)
       Text5.TEXT = LOCACLI$(NC%)
     End If
   End If
   '
   Screen.MousePointer = 1
   '
End Sub

Sub BLANFORMU()
   '
   Label8 = ""
   Text2.TEXT = ""
   Text3.TEXT = ""
   Text4.TEXT = ""
   Text5.TEXT = ""
   '
   HOII% = HOY(HOS$)
   Text9.TEXT = HOS$
   Text11.TEXT = TRIM$(Str$(PRONUOCO&))
   '
End Sub

Function PRONUOCO&()
   '
   PRONU& = 1
   Call ABRECOMPRAS
   '
   SQLX$ = "SELECT Nume_Ocom FROM OcomEnca "
   SQLX$ = SQLX$ + "ORDER BY Nume_Ocom DESC;"
   Set OEncaTemp = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
   On Error Resume Next
   With OEncaTemp
     .MoveFirst
     If Err < 1 Then
       PRONU& = 1 + !Nume_Ocom
     End If
   End With
   On Error GoTo 0
   '
   PRONUOCO& = PRONU&
   '
End Function

Private Sub Text11_GotFocus()
   On Error Resume Next 'x err '5' llamada a procedimiento no valida.
   Text1.SetFocus
   Screen.MousePointer = 1
   On Error GoTo 0
End Sub

Private Sub Text9_GotFocus()
   Text1.SetFocus
End Sub
