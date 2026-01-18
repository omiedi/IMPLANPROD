VERSION 5.00
Begin VB.Form INFONORM 
   BackColor       =   &H00D0D0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes Normalizados"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture16 
      Height          =   5900
      Left            =   10605
      ScaleHeight     =   5835
      ScaleWidth      =   1035
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   1100
      Begin VB.Image Image3 
         Height          =   5900
         Left            =   0
         Picture         =   "INFONORM.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   3900
      Left            =   10080
      TabIndex        =   30
      Top             =   1260
      Width           =   15
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      Height          =   3900
      Left            =   8400
      TabIndex        =   29
      Top             =   1260
      Width           =   15
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   3900
      Left            =   6720
      TabIndex        =   28
      Top             =   1260
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3900
      Left            =   2205
      TabIndex        =   27
      Top             =   1260
      Width           =   15
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3900
      Left            =   680
      TabIndex        =   26
      Top             =   1260
      Width           =   15
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00800000&
      Height          =   8600
      Left            =   10605
      ScaleHeight     =   8535
      ScaleWidth      =   1215
      TabIndex        =   19
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton Command17 
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
         Height          =   675
         Left            =   120
         Picture         =   "INFONORM.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir Informe Normalizado"
         Top             =   4305
         Width           =   615
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Tbls"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         Picture         =   "INFONORM.frx":2A02
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Tabla Indice de Informes Normalizados"
         Top             =   2820
         Width           =   615
      End
      Begin VB.CommandButton Command18 
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
         Height          =   675
         Left            =   120
         Picture         =   "INFONORM.frx":2B4C
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton Command19 
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
         Height          =   675
         Left            =   120
         Picture         =   "INFONORM.frx":2E56
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Plan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         Picture         =   "INFONORM.frx":2FA0
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Plan General de Cuentas"
         Top             =   2100
         Width           =   615
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   555
         TabIndex        =   20
         Top             =   3990
         Width           =   615
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -315
         Picture         =   "INFONORM.frx":30EA
         Top             =   5145
         Width           =   1515
      End
   End
   Begin VB.CommandButton Command21 
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
      Height          =   315
      Left            =   10215
      TabIndex        =   16
      Top             =   735
      Width           =   175
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7350
      TabIndex        =   15
      Text            =   "Combo1"
      Top             =   735
      Width           =   1275
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
      Height          =   315
      Left            =   4725
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   735
      Width           =   1800
   End
   Begin VB.TextBox Text11 
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
      Left            =   4725
      TabIndex        =   9
      Top             =   255
      Width           =   435
   End
   Begin VB.CommandButton Command31 
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
      Left            =   5145
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   255
      Width           =   175
   End
   Begin VB.PictureBox Picture1 
      Height          =   460
      Left            =   210
      ScaleHeight     =   405
      ScaleWidth      =   10095
      TabIndex        =   3
      Top             =   5120
      Width           =   10155
      Begin VB.CommandButton Command22 
         Caption         =   "Ver Mayor"
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
         Left            =   460
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Muestra el Mayor de la Cuenta Activa"
         Top             =   70
         Width           =   1500
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6580
         TabIndex        =   5
         Top             =   60
         Width           =   1500
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8260
         TabIndex        =   4
         Top             =   60
         Width           =   1500
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
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
      Height          =   3405
      Left            =   210
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   1730
      Width           =   10170
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00E0E0E0&
      Height          =   450
      Left            =   210
      ScaleHeight     =   390
      ScaleWidth      =   10095
      TabIndex        =   0
      Top             =   1365
      Width           =   10155
      Begin VB.Label Label27 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cuenta              Detalle / Referencia                                                       Debe                 Haber"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   525
         TabIndex        =   1
         Top             =   45
         Width           =   11355
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   9105
      TabIndex        =   18
      Top             =   735
      Width           =   1110
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Per:"
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
      Left            =   6720
      TabIndex        =   17
      Top             =   795
      Width           =   2325
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Forma:"
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
      Left            =   6300
      TabIndex        =   14
      Top             =   795
      Width           =   960
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo:"
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
      Left            =   3675
      TabIndex        =   12
      Top             =   795
      Width           =   960
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Informe:"
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
      Left            =   3675
      TabIndex        =   11
      Top             =   315
      Width           =   960
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5385
      TabIndex        =   10
      Top             =   255
      Width           =   4965
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Normalizados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   525
      TabIndex        =   7
      Top             =   630
      Width           =   2535
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Informes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   315
      TabIndex        =   6
      Top             =   105
      Width           =   2535
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu tablaindicenormalizados 
         Caption         =   "Tabla Indice de Informes Normalizados "
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu imprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu ss 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu accesodirecto 
      Caption         =   "Accesos Directos"
      Begin VB.Menu ACCDIRECTO 
         Caption         =   "Accesos Directo"
      End
   End
End
Attribute VB_Name = "INFONORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RECAR%

Private Sub ACCDIRECTO_Click()
    Call Command10_Click
End Sub

'
Private Sub Combo1_Click()
   NUINF% = Val(Text11)
   If NUINF% < 1 Or TRIM$(Label4) = "" Then Exit Sub
   '
   XX$ = FILTERGETRECORD$("ENCAINF", 1, MKI$(NUINF%))
   Call REPLA(XX$, Chr$(Val(Combo1.TEXT)), 48, 1)
   Call FILTERPUTRECORD("ENCAINF", 1, MKI$(NUINF%), XX$)
   '
   Call CARGAINF
End Sub

Private Sub Combo2_Click()
   NUINF% = Val(Text11)
   If NUINF% < 1 Or TRIM$(Label4) = "" Then Exit Sub
   '
   XX$ = FILTERGETRECORD$("ENCAINF", 1, MKI$(NUINF%))
   Call REPLA(XX$, Chr$(Val(Combo2.TEXT)), 49, 1)
   Call FILTERPUTRECORD("ENCAINF", 1, MKI$(NUINF%), XX$)
   '
   Call CARGAINF
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   FPLANCUEN.Show 1
   Command10.Enabled = True
End Sub

Private Sub Command16_Click()
   Command16.Enabled = False
   Call TRALOCAL("ENCAINF", "")
   VTB% = VENTABU%("ENCAINF")
   If VTB% >= 0 Then
     For KU& = 1 To ULTREG&("!ENCAINF")
       XX$ = REGLEIDO$("!ENCAINF", KU&)
       KI% = KU&
       Call REPLA(XX$, MKI$(KI%), 1, 2)
       Call GRAREG("!ENCAINF", XX$, KU&)
     Next KU&
     Call CIERRARCH("!ENCAINF")
     Call TRACENTRAL("ENCAINF", "")
   End If
   Command16.Enabled = True
End Sub

Private Sub Command17_Click()
   '
   '30/04/07 (OT-07-0175)
   Command17.Enabled = False
   '
   J& = 0
   If List1.ListCount > 0 Then
      For i& = 0 To List1.ListCount - 1
          YY$ = List1.List(i&)
          NORDEN$ = TRIM(Mid$(YY$, 1, 3))
          CODCTA$ = TRIM(Mid$(YY$, 5, 12))
          TOTDEB# = XVALO(Mid$(YY$, 56, 12))
          TOTHAB# = XVALO(Mid$(YY$, 70, 12))
          '
          XX$ = REGBLAN$("INFNORM")
          Call REPLA(XX$, NORDEN$, 1, 3)
          Call REPLA(XX$, CODCTA$, 4, 12)
          Call REPLA(XX$, MKD$(TOTDEB#), 16, 8)
          Call REPLA(XX$, MKD$(TOTHAB#), 24, 8)
          J& = J& + 1
          Call GRAREG("INFNORM", XX$, J&)
      Next i&
      '
      Call SETULTREG("INFNORM", J&)
      '
      Call HEADERS("INFNORM", "")
      Call HEADERS("INFNORM", "Informe: " + Label4.Caption + " - " + Label6.Caption)
      Call FINAL("*INFNORM")
   End If
   '
   Command17.Enabled = True
   '30/04/07 (FIN)
End Sub

Private Sub Command19_Click()
   RECAR% = 1
   Hide
End Sub

Private Sub Command21_Click()
   '
   If TRIM$(Label4) = "" Then
     Exit Sub
   End If
   '
10 Call SELECHO("SELFECHA")
   FESELEC$ = VALACT1$("SELFECHA")
   '
   If FESELEC$ <> "" Then
     FESEL% = FECHANUM(FESELEC$)
     If FESEL% < COMEJ% Or FESEL% > FINEJ% Then
       Call MENSERR(24, "Fecha no Válida")
       GoTo 10
     End If
     Label6 = FESELEC$
   End If
   '
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   If List1.ListIndex >= 0 Then
     COEXT$ = TRIM$(Mid$(List1.TEXT, 5, 12))
   End If
   CUEII% = CUEINT%(COEXT)
   If CUEII% > 0 Then
     Call INICONTA.LMAYORPAN(CUEII%)
   End If
99 Command22.Enabled = True
End Sub

Private Sub Command31_Click()
   Call SELECHO("ENCAINF")
   VDEVU$ = VALACT1$("ENCAINF")
   If VDEVU$ <> "" Then
     Text11 = VDEVU$
   End If
End Sub

Private Sub Form_Activate()
   If RECAR% > 0 Then
     OPENFORMS = DoEvents
     Call Text11_Change
     RECAR% = 0
   End If
End Sub

Private Sub Form_Load()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.Visible = True
      MARCOBARRA.Top = 5500
      MARCOBARRA.Left = 150
      MARCOBARRA.ZOrder 0

    End If

   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
   End If
   '
   Combo1.Clear
   Combo1.AddItem "1.- Saldos"
   Combo1.AddItem "2.- Res.Mens"
   Combo1.AddItem "3.- Res.Trim"
   Combo1.AddItem "4.- Res.Sem"
   Combo1.AddItem "5.- Mov Dia"
   '
   Combo2.Clear
   Combo2.AddItem "1.- Lineal"
   Combo2.AddItem "2.- Acumul"
   '
   HOII% = HOY(HOS$)
   Label6 = HOS$
End Sub

Private Sub imprimir_Click()
    Call Command17_Click
End Sub

Private Sub Label27_DblClick()
   Call List1_DblClick
End Sub

Private Sub Label4_DblClick()
   Call Command31_Click
End Sub

Private Sub Label6_Change()
   Call CARGAINF
End Sub

Private Sub Label6_DblClick()
   Call Command21_Click
End Sub

Private Sub List1_DblClick()
   NUINF% = Val(Text11)
   If NUINF% < 1 Or TRIM$(Label4) = "" Then Exit Sub
   '
   Call FILTERFILE("INFORMES", "!INFORMES", 1, MKI$(NUINF%))
   VTB% = VENTABU%("INFORMES")
   If VTB% >= 0 Then
     Call FILTERUPDATE("INFORMES", "", 1, MKI$(NUINF%))
     Call CARGAINF
   End If
End Sub

Private Sub salir_Click()
    Call Command19_Click
End Sub

Private Sub tablaindicenormalizados_Click()
   Call Command16_Click
End Sub

Private Sub Text11_Change()
   Label4 = ""
   If Val(Text11) > 0 Then
     If Val(Text11) < 32767 Then
       XX$ = FILTERGETRECORD$("ENCAINF", 1, MKI$(Val(Text11)))
       Label4 = TRIM$(Mid$(XX$, 3, 45))
       TIPOINF% = Asc(Mid$(XX$, 48, 1))
       FORMINF% = Asc(Mid$(XX$, 49, 1))
       '
       Combo1.TEXT = ""
       For KU% = 1 To Combo1.ListCount
         If Val(Left$(Combo1.List(KU% - 1), 1)) = TIPOINF% Then
           Combo1.TEXT = Combo1.List(KU% - 1)
           Exit For
         End If
       Next KU%
       '
       Combo2.TEXT = ""
       For KU% = 1 To Combo2.ListCount
         If Val(Left$(Combo2.List(KU% - 1), 1)) = FORMINF% Then
           Combo2.TEXT = Combo2.List(KU% - 1)
           Exit For
         End If
       Next KU%
       '
     End If
     Call CARGAINF
   End If
End Sub

Private Sub Text11_DblClick()
   Call Command31_Click
End Sub

Sub CARGAINF()
   '
   Dim CUEI%(1024), NVLI%(1024), SALFIN#(1024), SALINI#(1024), IMPORT#(1024)
   CONTA% = 0
   List1.Clear
   Label21 = "": Label22 = ""
   OPENFORMS = DoEvents
   NUINF% = Val(Text11)
   If NUINF% < 1 Or TRIM$(Label4) = "" Then Exit Sub
   '
   FESEL% = FECHANUM(Label6)
   If FESEL% < COMEJ% Or FESEL% > FINEJ% Then Exit Sub
   '
   RFF$ = RECFILT$("INFORMES", 1, MKI$(NUINF%))
   For KU& = 1 To Len(RFF$) Step 4
     REG& = CVS(Mid$(RFF$, KU&, 4))
     XX$ = REGLEIDO$("INFORMES", REG&)
     CUEII% = CVI(Mid$(XX$, 3, 2))
     NIVEI% = CVI(Mid$(XX$, 5, 2))
     If NIVEI% < 0 Then NIVEI% = 0
     If NIVEI% > 1 Then NIVEI% = 1
     CONTA% = CONTA% + 1
     CUEI%(CONTA%) = CUEII%
     NVLI%(CONTA%) = NIVEI%
     SALFIN#(CONTA%) = 0
     SALINI#(CONTA%) = 0
   Next KU&
   '
   For KI% = 1 To CONTA%
     CUEII% = CUEI%(KI%)
     SALFIN#(KI%) = SALCUENTA#(CUEII%, FESEL%)
   Next KI%
   '
   TI% = Val(Combo1.TEXT)
   If TI% = 5 Then ' movimientos diarios
       FEX = FESEL%
       FEINI% = DIANTE(FEX)
       For KI% = 1 To CONTA%
         CUEII% = CUEI%(KI%)
         SALINI#(KI%) = SALCUENTA#(CUEII%, FEINI%)
       Next KI%
     ElseIf TI% = 2 Then ' movimientos diarios
       FEX = FESEL%
       FET$ = "01" + Mid$(FECHATEX$(FEX), 3)
       FEINI% = DIANTE(FECHANUM(FET$))
       For KI% = 1 To CONTA%
         CUEII% = CUEI%(KI%)
         SALINI#(KI%) = SALCUENTA#(CUEII%, FEINI%)
       Next KI%
   End If
   '
   For KI% = 1 To CONTA%
     IMPORT#(KI%) = SALFIN#(KI%) - SALINI#(KI%)
   Next KI%
   '
   CONTB% = 0
   For KI% = 1 To CONTA%
     If Abs(IMPORT#(KI%)) >= 0.005 Then
       CONTB% = CONTB% + 1
       CUEI%(CONTB%) = CUEI%(KI%)
       NVLI%(CONTB%) = NVLI%(KI%)
       SALFIN#(CONTB%) = SALFIN#(KI%)
       SALINI#(CONTB%) = SALINI#(KI%)
       IMPORT#(CONTB%) = IMPORT#(KI%)
     End If
   Next KI%
   CONTA% = CONTB%
   '
   SUMACU# = 0: SUDEBE# = 0: SUHABE# = 0
   For KI% = 1 To CONTA%
     FO% = Val(Combo2.TEXT)
     NIVEI% = NVLI%(KI%)
     CUEII% = CUEI%(KI%)
     If FO% = 1 Then
         NVLX$ = FORMATNUM$(KI%, "I3")
       Else
         NVLX$ = " 0.": If NIVEI% = 1 Then NVLX$ = " .1"
     End If
     '
     TTXX$ = Space$(128)
     Call REPLA(TTXX$, NVLX$, 1, 3)
     Call REPLA(TTXX$, CODCUE$(CUEII%), 5, 16)
     Call REPLA(TTXX$, DENOCUE$(CUEII%), 18, 40)
     '
     SALDO1# = IMPORT#(KI%)
     SALDO2# = 0
     If FO% = 1 Then
       If SALDO1# < 0 Then
         SALDO2# = Abs(SALDO1#)
         SALDO1# = 0
       End If
     End If
     '
     If FO% = 2 Then
       SUMACU# = SUMACU# + SALDO1#
       If KI% = CONTA% Or NVLI%(KI% + 1) = 0 Then
         SALDO2# = SUMACU#
         SUMACU# = 0
       End If
     End If
     '
     SUDEBE# = SUDEBE# + SALDO1#
     SUHABE = SUHABE# + SALDO2#
     '
     SALDOF1$ = FORMATNUM$(SALDO1#, "F12.2")
     SALDOF2$ = FORMATNUM$(SALDO2#, "F12.2")
     Call REPLA(TTXX$, SALDOF1$, 56, 12)
     Call REPLA(TTXX$, SALDOF2$, 70, 12)
     '
     List1.AddItem TTXX$
   Next KI%
   '
   Label21 = FORMATNUM$(SUDEBE#, "F12.2")
   Label22 = FORMATNUM$(SUHABE#, "F12.2")
End Sub

Function SALCUENTA#(CUEII%, FESEL%)
   '
   Call ABRECABAN
   Call ABREPLANCUEN
   '
   ' CARGAR MOVIMIENTOS DE DIARIO
   SUCUI% = 0
   SQLX$ = "SELECT NuIdAsien, IDEBEPASE, IHABEPASE "
   SQLX$ = SQLX$ + " FROM DETAASIEN WHERE "
   'SQLX$ = SQLX$ + "(CDBL(FECHASIEN) >= " & CDbl(CVDAT(COMEJ%)) & " "
   SQLX$ = SQLX$ + "CDBL(FECHASIEN) <= " & CDbl(CVDAT(FESEL%)) & " "
   SQLX$ = SQLX$ + "AND CUECOINT = " & CUEII% & " "
   SQLX$ = SQLX$ + "AND SUCUASIen = " & SUCUI% & " "
   SQLX$ = SQLX$ + "AND STATPASE > 0; "
   Set MODIASEL = Contable.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   '
   SALCU# = 0
   With MODIASEL
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
'         TIASI% = 0
'         NUAS$ = !NuIdAsien
'         Encasien.FindFirst "NUIDASIEN = '" & NUAS$ & "'"
'         If Encasien.NoMatch = False Then
'           TIASI% = Encasien!TIPOASIEN
'         End If
'         If TIASI% <> 9 Then
          SALCU# = SALCU# + !IDEBEPASE - !IHABEPASE
'         End If
         .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   '
19 MODIASEL.Close
   
   ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "WHERE CueContaI = " & CUEII% & " "
   SQLX$ = SQLX$ + " AND (NuAsidiar = '' or isnull(NuAsiDiar) = true OR IsEmpty(NuAsiDiar) = true) "
   SQLX$ = SQLX$ + " AND CDBL(FechConta) <= " & CDbl(CVDAT(FESEL%)) & " "
   Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   With CABATEMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
         SALCU# = SALCU# + !IDEBECOMP - !IHABECOMP
         .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   '
   SALCUENTA# = SALCU#
   '
End Function
