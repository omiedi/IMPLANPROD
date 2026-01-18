VERSION 5.00
Begin VB.Form ANOTASEG09 
   BackColor       =   &H00E0E0DD&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anotador Seguimiento Comercial"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton BOTGUARDA 
      Caption         =   "Guarda Anotacion Cambiada"
      Height          =   255
      Left            =   6825
      TabIndex        =   34
      Top             =   6195
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000040&
      Height          =   1740
      Left            =   8715
      ScaleHeight     =   1680
      ScaleWidth      =   1155
      TabIndex        =   29
      Top             =   0
      Width           =   1215
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   460
         Left            =   120
         Picture         =   "ANOTASEG09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Ver Oportunidades Para este Prospecto"
         Top             =   660
         Width           =   600
      End
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
         Height          =   460
         Left            =   120
         Picture         =   "ANOTASEG09.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Preparar y Enviar un E-Mail"
         Top             =   1125
         Width           =   600
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   460
         Left            =   120
         Picture         =   "ANOTASEG09.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   210
         Width           =   600
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0DD&
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1620
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   8460
      Begin VB.CommandButton Command35 
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
         Left            =   1980
         TabIndex        =   33
         ToolTipText     =   "Selecciona Oportunidades para el Contacto Activo"
         Top             =   1185
         Width           =   175
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
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
         Left            =   1470
         TabIndex        =   27
         Top             =   735
         Width           =   1695
      End
      Begin VB.Label Label8 
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
         Left            =   2250
         TabIndex        =   26
         Top             =   1200
         Width           =   6060
      End
      Begin VB.Label Label7 
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
         Left            =   300
         TabIndex        =   25
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label3 
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
         Height          =   540
         Left            =   3255
         TabIndex        =   24
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contacto o Prospecto:"
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
         Left            =   -210
         TabIndex        =   23
         Top             =   495
         Width           =   2400
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Oportunidad:"
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
         TabIndex        =   22
         ToolTipText     =   "Doble Click para limpiar"
         Top             =   945
         Width           =   1200
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0DD&
      Caption         =   "Novedades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1815
      Left            =   105
      TabIndex        =   9
      Top             =   6350
      Width           =   10620
      Begin VB.CommandButton Command15 
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
         Height          =   280
         Left            =   8520
         TabIndex        =   10
         Top             =   1360
         Width           =   175
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1175
         Left            =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   600
         Width           =   6435
      End
      Begin VB.CommandButton Command5 
         Caption         =   "O.K"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   745
         Left            =   8800
         Picture         =   "ANOTASEG09.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1020
         Width           =   615
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0C0C0&
         Height          =   320
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   10515
         TabIndex        =   12
         Top             =   315
         Width           =   10580
         Begin VB.Line Line9 
            X1              =   -100
            X2              =   9400
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   1155
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Label Label5 
            Caption         =   "Label5"
            Height          =   15
            Left            =   0
            TabIndex        =   14
            Top             =   420
            Width           =   1380
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha     Detalle                                                   Usuario"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   105
            TabIndex        =   13
            Top             =   0
            Width           =   10200
         End
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   7545
         TabIndex        =   11
         Top             =   1350
         Width           =   1000
      End
      Begin VB.Line Line3 
         X1              =   7500
         X2              =   9500
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Próx.Contacto:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7560
         TabIndex        =   18
         Top             =   1080
         Width           =   1380
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   330
         Left            =   0
         TabIndex        =   17
         Top             =   630
         Width           =   1065
      End
      Begin VB.Line Line8 
         X1              =   0
         X2              =   11665
         Y1              =   1780
         Y2              =   1780
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7560
         TabIndex        =   20
         Top             =   720
         Width           =   4005
      End
      Begin VB.Line Line7 
         X1              =   10580
         X2              =   10580
         Y1              =   315
         Y2              =   1900
      End
      Begin VB.Line Line6 
         X1              =   7450
         X2              =   10350
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line5 
         X1              =   7450
         X2              =   7450
         Y1              =   600
         Y2              =   1750
      End
      Begin VB.Label Label32 
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
         Height          =   1170
         Left            =   7440
         TabIndex        =   19
         Top             =   600
         Width           =   2265
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Label35"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8380
         TabIndex        =   16
         Top             =   840
         Width           =   4000
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0DD&
      Caption         =   "Anotaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   4450
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Doble-Click para Ordemiento por Fecha"
      Top             =   1785
      Width           =   9300
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   4005
         Left            =   8980
         TabIndex        =   8
         Top             =   530
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   4005
         Left            =   7080
         TabIndex        =   6
         Top             =   530
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   4110
         Left            =   1000
         TabIndex        =   5
         Top             =   530
         Width           =   15
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3810
         IntegralHeight  =   0   'False
         ItemData        =   "ANOTASEG09.frx":08A8
         Left            =   0
         List            =   "ANOTASEG09.frx":08AA
         MultiSelect     =   2  'Extended
         TabIndex        =   4
         Top             =   600
         Width           =   9250
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   435
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   9195
         TabIndex        =   3
         Top             =   315
         Width           =   9260
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha   Detalle                                            Usuario/Termin"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   80
            TabIndex        =   7
            Top             =   30
            Width           =   10200
         End
      End
      Begin VB.Line Line1 
         X1              =   9250
         X2              =   9250
         Y1              =   320
         Y2              =   4430
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   10605
         Y1              =   4410
         Y2              =   4410
      End
   End
   Begin VB.Label Label31 
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
      Height          =   1170
      Left            =   240
      TabIndex        =   28
      Top             =   6960
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   -105
      TabIndex        =   1
      Top             =   2625
      Width           =   1695
   End
End
Attribute VB_Name = "ANOTASEG09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub BOTGUARDA_Click()
   '
   COCLI$ = TRIM$(Label26)
   CODOP$ = TRIM$(Label7)
   '
   ANOTAX$ = ""
   For KKI% = 1 To List1.ListCount
      ANOTAX$ = ANOTAX$ + List1.List(KKI% - 1) + CRLF
   Next KKI%
   '
   If COCLI$ <> "" Then
      CONDI$ = "CODI_CLI = '" & COCLI$ & "'"
      Call ACTUREGISTRO("BSEGCOM", "HIS_CONTAC", CONDI$, ANOTAX$, RAFE&, "NOMESS")
      If CODOP$ <> "" Then
         CONDI$ = "CODI_CLI = '" & COCLI$ & "' and CodOport = '" & CODOP$ & "'"
         Call ACTUREGISTRO("OPORTUNI", "HIS_OPORTU", CONDI$, ANOTAX$, RAFE&, "NOMESS")
      End If
      BOTGUARDA.Visible = False
   End If
   '
End Sub

Private Sub Command1_Click()
    CODICLIX$ = Label26
    If CODICLIX$ <> "" Then
       CODIOPO$ = Label7
       If CODIOPO$ = "" Then
          Call SELECHO("!INDOPORT/INDICE DE OPORTUNIDADES")
          CODIOPO$ = TRIM$(VALACT1$("!INDOPORT"))
       End If
       '
       If CODIOPO$ <> "" Then
          Label7 = CODIOPO$
          Call STDFORM("NUOPOR", CODIOPO$)
       End If
    End If
End Sub

Private Sub Command15_Click()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text12.TEXT = VDEV$
    End If
End Sub

Private Sub Command2_Click()
   BOTGUARDA.Visible = False
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command20_Click()
   '
   Command20.Enabled = False
   PREPAMAI09.COMMAND16.Visible = False
   PREPAMAI09.Label26 = Label26
   PREPAMAI09.Label10 = Label7
   PREPAMAI09.Show 1
   If PREPAMAI09.Label1 = "OK" Then
      PERSODES$ = PREPAMAI09.Combo1.TEXT
      For KKI% = 1 To FISECOM09.Text6.Count
         If TRIM$(FISECOM09.Text6(KKI% - 1)) <> "" Then
            If FISECOM09.Text6(KKI% - 1) = PREPAMAI09.Combo1.TEXT Then
               Call FISECOM09.Text6_DblClick(KKI% - 1)
               GoTo 80
            End If
         End If
      Next KKI%
      Call FISECOM09.Text5_DblClick
      '
80    PREPAMAI09.Label1 = ""
   End If
   Unload PREPAMAI09
   Command20.Enabled = True
   '
End Sub


Private Sub Command35_Click()
    Call SELECHO("!INDOPORT/INDICE DE OPORTUNIDADES")
    CODIOPO$ = TRIM$(VALACT1$("!INDOPORT"))
    If CODIOPO$ <> "" Then
       Label7 = CODIOPO$
    End If
End Sub

Private Sub Command5_Click()
   '\\\OT-05-0966-OD-27/10/08///
   Command5.Enabled = False
   HOII% = HOY(HOS$)
   '
   COCLI$ = TRIM$(Label26)
   CODOP$ = TRIM$(Label7)
   If COCLI$ <> "" Then
     FEC% = FECHANUM(Text12)
     If FEC% <= HOY(HO$) Then FEC% = HOY(HO$)
     FEPRO = FETEXCOR1$(FEC%)
     CONDI$ = "CODI_CLI = '" & COCLI$ & "'"
     Call ACTUREGISTRO("BSEGCOM", "FechProxCto", CONDI$, FEPRO, RAFE&, "NOMESS")
     '
     If TRIM$(Text11) <> "" Then
       Call FRATEXTO(Text11, 56)
       TTXX1$ = ""
       For KKI% = 1 To CARETEX%
         TTXX$ = Space$(76)
         Call REPLA(TTXX$, RETEX$(KKI%), 11, 56)
         If KKI% = 1 Then
           Call REPLA(TTXX$, Label34, 1, 8)
           Call REPLA(TTXX$, USERTER$, 69, 24)
         End If
         TTXX1$ = TTXX1$ + TTXX$ + Chr$(10) + Chr$(13)
       Next KKI%
       TTXX1$ = TTXX1$ + String$(88, "=") + Chr$(10) + Chr$(13)
       '
       CONDI$ = "CODI_CLI = '" & COCLI$ & "'"
       ANOTANTE$ = VALOBADA("BSEGCOM", "HIS_CONTAC", CONDI$)
       Call ACTUREGISTRO("BSEGCOM", "HIS_CONTAC", CONDI$, ANOTANTE$ + TTXX1$, RAFE&)
       '
       If CODOP$ <> "" Then
          CONDI$ = "CODI_CLI = '" & COCLI$ & "' and CodOport = '" & CODOP$ & "'"
          ANOTANTE$ = VALOBADA("OPORTUNI", "HIS_OPORTU", CONDI$)
          Call ACTUREGISTRO("OPORTUNI", "HIS_OPORTU", CONDI$, ANOTANTE$ + TTXX1$, RAFE&)
       End If
       '
       Text11 = ""
       Call CARGANOTA
       '
     End If
   End If
   '
   Command5.Enabled = True
   '
End Sub

Sub CARGANOTA()
   BOTGUARDA.Visible = False
   COCLI$ = TRIM$(Label26)
   CODOP$ = TRIM$(Label7)
   List1.Clear
   If COCLI$ <> "" Then
      If CODOP$ <> "" Then
           CONDI$ = "CODI_CLI = '" & COCLI$ & "' and CodOport = '" & CODOP$ & "'"
           ANOTANTE$ = VALOBADA("OPORTUNI", "HIS_OPORTU", CONDI$)
         Else
15         CONDI$ = "CODI_CLI = '" & COCLI$ & "'"
           ANOTANTE$ = VALOBADA("BSEGCOM", "HIS_CONTAC", CONDI$)
      End If
      '
      XX$ = ANOTANTE$: LONGI& = Len(XX$)
      INICIO& = 1
      Do While INICIO& < LONGI&
         FIN& = InStr(INICIO&, XX$, Chr$(10))
         FINB& = InStr(INICIO&, XX$, CRLF$)
         If FINB& > 0 And FINB& < FIN& Then FIN& = FINB&
         If FIN& > 0 Then
            TTXX$ = Mid$(XX$, INICIO&, FIN& - INICIO&)
            List1.AddItem TTXX$
            INICIO& = FIN& + 2
          Else
            TTXX$ = Mid$(XX$, INICIO&)
            List1.AddItem TTXX$
            Exit Do
         End If
      Loop
      '
      TPI = List1.ListCount - 12
      If TPI < 0 Then TPI = 0
      List1.TopIndex = TPI
      '
   End If
   '
End Sub

Sub ORDENOTA()
   '
   Dim FEXI%(1024), BLOQ$(1024)
   COCLI$ = TRIM$(Label26)
   CODOP$ = TRIM$(Label7)
   List1.Clear: MODO% = 0
   If COCLI$ <> "" Then
      If CODOP$ <> "" Then
           CONDI$ = "CODI_CLI = '" & COCLI$ & "' and CodOport = '" & CODOP$ & "'"
           ANOTANTE$ = VALOBADA("OPORTUNI", "HIS_OPORTU", CONDI$)
           MODO% = 1
           If TRIM$(ANOTANTE$) = "" Then
              Label7 = ""
              GoTo 15
           End If
         Else
15         CONDI$ = "CODI_CLI = '" & COCLI$ & "'"
           ANOTANTE$ = VALOBADA("BSEGCOM", "HIS_CONTAC", CONDI$)
           MODO% = 2
      End If
      '
      XX$ = ANOTANTE$: LONGI& = Len(XX$)
      INICIO& = 1: IBLO% = 0
      Do While INICIO& < LONGI&
         FIN& = InStr(INICIO&, XX$, Chr$(10))
         FINB& = InStr(INICIO&, XX$, CRLF$)
         If FINB& > 0 And FINB& < FIN& Then FIN& = FINB&
         If FIN& > 0 Then
            TTXX$ = Mid$(XX$, INICIO&, FIN& - INICIO&)
            If Left$(TTXX$, 8) <> "        " Then
              FEXII% = FECHANUM(Left$(TTXX$, 8))
              If FEXII% > 0 Then
                 IBLO% = IBLO% + 1
                 FEXI%(IBLO%) = FEXII%
              End If
            End If
            BLOQ$(IBLO%) = BLOQ$(IBLO%) + TTXX$ + CRLF
            ' List1.AddItem TTXX$
            INICIO& = FIN& + 2
          Else
            TTXX$ = Mid$(XX$, INICIO&)
            If Left$(TTXX$, 8) <> "        " Then
              FEXII% = FECHANUM(Left$(TTXX$, 8))
              If FEXII% > 0 Then
                 IBLO% = IBLO% + 1
                 FEXI%(IBLO%) = FEXII%
              End If
            End If
            BLOQ$(IBLO%) = BLOQ$(IBLO%) + TTXX$
            Exit Do
         End If
      Loop
      '
      ' ORDENAR
      For i% = 1 To IBLO%
        For J% = i% + 1 To IBLO%
          If FEXI%(i%) > FEXI%(J%) Then
            FEXII% = FEXI%(i%)
            BLOQI$ = BLOQ$(i%)
            FEXI%(i%) = FEXI%(J%)
            BLOQ$(i%) = BLOQ$(J%)
            FEXI%(J%) = FEXII%
            BLOQ$(J%) = BLOQI$
          End If
        Next J%
      Next i%
      '
      ' VOLVER A ARMAR EL BLOQUE
      ANOTANTE$ = ""
      For i% = 1 To IBLO%
        ANOTANTE$ = ANOTANTE$ + BLOQ$(i%)
      Next i%
      '
      ' PONER EN EL LIST
      List1.Clear
      XX$ = ANOTANTE$: LONGI& = Len(XX$)
      INICIO& = 1
      Do While INICIO& < LONGI&
         FIN& = InStr(INICIO&, XX$, Chr$(10))
         FINB& = InStr(INICIO&, XX$, CRLF$)
         If FINB& > 0 And FINB& < FIN& Then FIN& = FINB&
         If FIN& > 0 Then
            TTXX$ = Mid$(XX$, INICIO&, FIN& - INICIO&)
            List1.AddItem TTXX$
            INICIO& = FIN& + 2
          Else
            TTXX$ = Mid$(XX$, INICIO&)
            List1.AddItem TTXX$
            Exit Do
         End If
      Loop
      '
      TPI = List1.ListCount - 12
      If TPI < 0 Then TPI = 0
      List1.TopIndex = TPI
      '
      If USNBR% Mod 100 = 1 Then
         BOTGUARDA.Visible = True
      End If
      '
   End If
   '
End Sub

Private Sub Form_Load()
   '
   Call CENTRAFORM(Me)
   Label9 = USERNAM$(USNBR%)
   HOII% = HOY(HOS$)
   Label34 = HOS$
   Text12 = FECHATEX$(DIPOST(HOII%))
   '
End Sub

Private Sub Frame1_DblClick()
   Call ORDENOTA
End Sub

Private Sub Label21_DblClick()
    Label7 = ""
End Sub

Private Sub Label26_Change()
    Label7 = ""
    Label3 = VALOBADA("BSEGCOM", "Raso_Cli", "CODI_CLI = '" & Label26 & "'")
    CONDI$ = "CodOport <>'' AND CODI_CLI = '" & Label26 & "'"
    Call CARGALISEL("!INDOPORT", "OPORTUNI", "CodOport/A12;RefeOport/A48", CONDI$, "NOMESS")
    Call CARGANOTA
    CONDI2$ = "CODI_CLI = '" & Label26 & "'"
    Text12 = FECHATEX$(CVINT(VALOBADA("BSEGCOM", "FechProxCto", CONDI2$)))
End Sub

Private Sub Label34_DblClick()
   Call SELECHO("FECHA")
   If VALACT1$("FECHA") <> "" Then
      Label34 = VALACT1$("FECHA")
   End If
End Sub

Private Sub Label7_Change()
    Label8 = VALOBADA("OPORTUNI", "RefeOport", "CodOport = '" & Label7 & "'")
    Call CARGANOTA
End Sub

Private Sub Label7_DblClick()
    Call Command35_Click
End Sub

Private Sub List1_Click()
  On Error Resume Next
  List1.SetFocus
  On Error GoTo 0
End Sub

Private Sub List1_DblClick()
  '
  TTXX$ = List1.TEXT
  PPXX% = InStr(TTXX$, "(#")
  If PPXX% > 0 Then
     PPYY% = InStr(PPXX%, TTXX$, "#)")
     If PPYY% > PPXX% + 2 Then
        NUMAI$ = TRIM$(Mid$(TTXX$, PPXX% + 2, PPYY% - PPXX% - 2))
        SHOWMAIL09.Label13 = NUMAI$
        SHOWMAIL09.Show 1
        Exit Sub
     End If
  End If
  '
  PPXX% = InStr(TTXX$, "($")
  If PPXX% > 0 Then
     PPYY% = InStr(PPXX%, TTXX$, "$)")
     If PPYY% > PPXX% + 2 Then
        NUMAI$ = TRIM$(Mid$(TTXX$, PPXX% + 2, PPYY% - PPXX% - 2))
        SHINMAIL09.Label13 = NUMAI$
        SHINMAIL09.Show 1
        Exit Sub
     End If
  End If
  '
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
     ILISTX& = List1.ListIndex
     If ILISTX& >= 0 Then
       If DERACCE%("BORRANOT/NOMESS", "Borrar Registros de Anotaciones") >= 2 Then
         BOTGUARDA.Visible = True
         List1.RemoveItem ILISTX&
         If ILISTX& < List1.ListCount Then
             List1.ListIndex = ILISTX&
           Else
             List1.ListIndex = List1.ListCount - 1
         End If
         If List1.ListIndex >= 0 Then List1.Selected(List1.ListIndex) = True
       End If
     End If
   End If
End Sub

Private Sub Text12_DblClick()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text12.TEXT = VDEV$
    End If
End Sub
