VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form INFONOR07 
   BackColor       =   &H0095C5E1&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes Normalizados"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid GRID2 
      Height          =   765
      Left            =   960
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   1349
      _Version        =   393216
      Rows            =   6
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   8433840
      BackColorSel    =   -2147483643
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      Redraw          =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   25
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H002C3D46&
      ForeColor       =   &H80000008&
      Height          =   8600
      Left            =   10605
      ScaleHeight     =   8565
      ScaleWidth      =   1245
      TabIndex        =   18
      Top             =   0
      Width           =   1275
      Begin VB.PictureBox MARCOBARRA 
         Height          =   195
         Left            =   120
         ScaleHeight     =   135
         ScaleWidth      =   585
         TabIndex        =   34
         Top             =   5280
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   35
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "INFONOR07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir Informe Normalizado"
         Top             =   4305
         Width           =   650
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Tbls"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "INFONOR07.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Tabla Indice de Informes Normalizados"
         Top             =   2820
         Width           =   650
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "INFONOR07.frx":07B4
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   960
         Width           =   650
      End
      Begin VB.CommandButton Command19 
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
         Height          =   690
         Left            =   120
         Picture         =   "INFONOR07.frx":0ABE
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   650
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Plan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "INFONOR07.frx":0C08
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Plan General de Cuentas"
         Top             =   2100
         Width           =   650
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
      TabIndex        =   15
      Top             =   735
      Width           =   175
   End
   Begin VB.ComboBox Combo2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7350
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   735
      Width           =   1275
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4725
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   735
      Width           =   1800
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4725
      TabIndex        =   8
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   255
      Width           =   175
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   460
      Left            =   210
      ScaleHeight     =   435
      ScaleWidth      =   10125
      TabIndex        =   2
      Top             =   5120
      Width           =   10155
      Begin VB.CommandButton Command22 
         Caption         =   "Ver Mayor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   460
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Muestra el Mayor de la Cuenta Activa"
         Top             =   70
         Width           =   1500
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6580
         TabIndex        =   4
         Top             =   60
         Width           =   1500
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8260
         TabIndex        =   3
         Top             =   60
         Width           =   1500
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3810
      Left            =   210
      ScaleHeight     =   3780
      ScaleWidth      =   10125
      TabIndex        =   0
      Top             =   1365
      Width           =   10155
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   3900
         Left            =   440
         TabIndex        =   32
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   3900
         Left            =   1965
         TabIndex        =   31
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   3900
         Left            =   6480
         TabIndex        =   30
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   3900
         Left            =   8160
         TabIndex        =   29
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame5 
         Caption         =   "Frame5"
         Height          =   3900
         Left            =   9840
         TabIndex        =   28
         Top             =   0
         Width           =   15
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
         Left            =   -10
         MultiSelect     =   2  'Extended
         TabIndex        =   33
         Top             =   380
         Width           =   10155
      End
      Begin VB.Label Label27 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"INFONOR07.frx":0D52
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   10110
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "INFONOR07.frx":0E06
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   3120
      OleObjectBlob   =   "INFONOR07.frx":103A
      TabIndex        =   26
      Top             =   1080
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   9105
      TabIndex        =   17
      Top             =   735
      Width           =   1110
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Período"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8640
      TabIndex        =   16
      Top             =   795
      Width           =   405
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Forma"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6660
      TabIndex        =   13
      Top             =   795
      Width           =   600
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3675
      TabIndex        =   11
      Top             =   795
      Width           =   960
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Informe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3675
      TabIndex        =   10
      Top             =   315
      Width           =   960
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5385
      TabIndex        =   9
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
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   525
      TabIndex        =   6
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
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   315
      TabIndex        =   5
      Top             =   105
      Width           =   2535
   End
End
Attribute VB_Name = "INFONOR07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RECAR%
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
   FPLANCUEN07.Show 1
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
     On Error Resume Next
     Call INICONTA07.LMAYORPAN(CUEII%)
     If Err Then
        Call INICONTA12.LMAYORPAN(CUEII%)
     End If
     On Error GoTo 0
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
     Openforms = DoEvents
     Call Text11_Change
     RECAR% = 0
   End If
End Sub

Private Sub Form_Load()
'   EPAC$ = TRIM$(EMPREAC$)
'   If EPAC$ <> "" Then
'       Caption = Caption + "  -  " + EPAC$
'   End If
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
      EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
      Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
   End If
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me)
   
   Call REMPLAGRID(Me, GRID2, Picture5, Label27, List1)
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
   Openforms = DoEvents
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
   Call REMPLAGRID(Me, GRID2, Picture5, Label27, List1)
   Label21 = FORMATNUM$(SUDEBE#, "F12.2")
   Label22 = FORMATNUM$(SUHABE#, "F12.2")
End Sub

Function SALCUENTA#(CUEII%, FESEL%)
   '
   Call ABRECABAN
   Call ABREPLANCUEN
   '
               ' SUCUI% = UNINEGO%
               SQLX$ = "SELECT SUM(IDEBEPASE) AS SUDEBE, SUM(IHABEPASE) AS SUHABE"
               SQLX$ = SQLX$ + " FROM DETAASIEN WITH(NOLOCK) WHERE "
               SQLX$ = SQLX$ + "FECHASIEN <= '" & FETEXCOR1$(FESEL%, "HASTA") & "' "
               SQLX$ = SQLX$ + "AND CUECOINT = " & CUEII% & " "
               ' SQLX$ = SQLX$ + "AND SUCUASIen = " & SUCUI% & " "
               SQLX$ = SQLX$ + "AND STATPASE > 0; "
               Dim MODIASEL As ADODB.Recordset
               Set MODIASEL = New ADODB.Recordset
               Set MODIASEL = FLEXCONN.Execute(FILTSQL$(SQLX$))
               SALCU# = XVALO(MODIASEL!SUDEBE) - XVALO(MODIASEL!SUHABE)
               GoTo 19
   ' CARGAR MOVIMIENTOS DE DIARIO
   ' SUCUI% = UNINEGO%
   SQLX$ = "SELECT NuIdAsien, IDEBEPASE, IHABEPASE "
   SQLX$ = SQLX$ + " FROM DETAASIEN WHERE "
   SQLX$ = SQLX$ + "FECHASIEN <= '" & CVDAT(FESEL%) & "' "
   SQLX$ = SQLX$ + "AND CUECOINT = " & CUEII% & " "
   ' SQLX$ = SQLX$ + "AND SUCUASIen = " & SUCUI% & " "
   SQLX$ = SQLX$ + "AND STATPASE > 0; "
   '
   'Dim MODIASEL As ADODB.Recordset
   Set MODIASEL = New ADODB.Recordset
   MODIASEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

   
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
          SALCU# = SALCU# + !idebepase - !ihabepase
'         End If
         .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   '
19 MODIASEL.Close
   Set MODIASEL = Nothing
             '
             SQLX$ = "SELECT SUM(IDEBECOMP) AS SUDEBE, SUM(IHABECOMP) AS SUHABE  FROM CAJABANC WITH(NOLOCK) "
             SQLX$ = SQLX$ + "WHERE CueContaI = " & CUEII% & " "
             SQLX$ = SQLX$ + " AND (NuAsidiar = '' or NuAsiDiar is null) "
             SQLX$ = SQLX$ + " AND FechConta <= '" & FETEXCOR1$(FESEL%, "HASTA") & "' "
             Dim CABATEMP As ADODB.Recordset
             Set CABATEMP = New ADODB.Recordset
             Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
             SALCU# = SALCU# + XVALO(CABATEMP!SUDEBE) - XVALO(CABATEMP!SUHABE)
             CABATEMP.Close
             Set CABATEMP = Nothing
             GoTo 90
             '
   
   ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "WHERE CueContaI = " & CUEII% & " "
   SQLX$ = SQLX$ + " AND (NuAsidiar = '' or NuAsiDiar is null) "
   SQLX$ = SQLX$ + " AND FechConta <= '" & CVDAT(FESEL%) & "' "
   'Dim CABATEMP As ADODB.Recordset
   Set CABATEMP = New ADODB.Recordset
   CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   '
   With CABATEMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
         SALCU# = SALCU# + !idebecomp - !ihabecomp
         .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   Set CABATEMP = Nothing
   '
90 SALCUENTA# = SALCU#
   '
End Function
