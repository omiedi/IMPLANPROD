VERSION 5.00
Begin VB.Form ANOTACARPI 
   BackColor       =   &H00EEDDEE&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anotador Carpeta de Importacion /  Exportacion"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000040&
      Height          =   1815
      Left            =   8700
      ScaleHeight     =   1755
      ScaleWidth      =   1155
      TabIndex        =   26
      Top             =   120
      Width           =   1215
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   60
         ScaleHeight     =   75
         ScaleWidth      =   630
         TabIndex        =   31
         Top             =   960
         Width           =   690
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   32
            Top             =   120
            Width           =   12000
         End
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
         Height          =   500
         Left            =   120
         Picture         =   "ANOTACARPI.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Cerrar Carpeta de Expotacion"
         Top             =   90
         Width           =   600
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -320
         Picture         =   "ANOTACARPI.frx":014A
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   1200
         Width           =   1515
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00EEDDEE&
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
      Height          =   1830
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
         Left            =   3180
         TabIndex        =   28
         ToolTipText     =   "Selecciona Carpeta de Exportacion"
         Top             =   300
         Width           =   175
      End
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
         Height          =   285
         Left            =   7320
         TabIndex        =   30
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Alta:"
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
         Left            =   6060
         TabIndex        =   29
         ToolTipText     =   "Doble Click para limpiar"
         Top             =   390
         Width           =   1200
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
         Height          =   570
         Left            =   240
         TabIndex        =   24
         Top             =   840
         Width           =   8055
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
         Left            =   1500
         TabIndex        =   23
         Top             =   330
         Width           =   1695
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Carpeta:"
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
         Left            =   240
         TabIndex        =   22
         ToolTipText     =   "Doble Click para limpiar"
         Top             =   390
         Width           =   1200
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EEDDEE&
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
      Left            =   120
      TabIndex        =   9
      Top             =   6600
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
         Left            =   8820
         Picture         =   "ANOTACARPI.frx":206C
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   975
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
            Caption         =   "Fecha   Detalle                                              Usuario"
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
         Height          =   1140
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
      BackColor       =   &H00EEDDEE&
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
      Top             =   2040
      Width           =   9300
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   3885
         Left            =   9000
         TabIndex        =   8
         Top             =   540
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   3915
         Left            =   7140
         TabIndex        =   6
         Top             =   510
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   3900
         Left            =   1020
         TabIndex        =   5
         Top             =   525
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
         ItemData        =   "ANOTACARPI.frx":2376
         Left            =   -15
         List            =   "ANOTACARPI.frx":2378
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
      TabIndex        =   25
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
Attribute VB_Name = "ANOTACARPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub ANOTA_TRANSACCIONES(NORXX$, TXTX$)

    If NORXX$ <> "" Then
        ANOTANTE$ = VALOBADA("CARPEIMP", "His_Carpi", "CodCarpImp = '" & NORXX$ & "'")
        ANOTA$ = ANOTANTE$ + TXTX$ + Chr$(10) + Chr$(13)
        Call ACTUREGISTRO("CARPEIMP", "His_Carpi", "CodCarpImp = '" & NORXX$ & "'", ANOTA$, RAFE&)
    End If

End Sub


Private Sub Command1_Click()
     Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      FECH$ = VALACT1$("SELFECHA")
      Label6 = FECH$
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
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command20_Click()
   '
Exit Sub

   Call FISECOM07.Text5_DblClick
   Exit Sub

   EMPRE$ = TRIM$(Left$(FISECOM.Text4, 59))
   CALNUM$ = TRIM$(FISECOM.Text20)
   CPOST$ = TRIM$(FISECOM.Text21)
   If CPOST$ <> "" Then CPOST$ = CPOST$ + " "
   LOCAX$ = CPOST$ + TRIM$(FISECOM.Text22)
   PROPA$ = TRIM$(FISECOM.Text19)
   If PROPA$ <> "" Then PROPA$ = PROPA$ + " - "
   PROPA$ = TRIM$(PROPA$ + TRIM$(FISECOM.Text23))
   '
   CADEST% = 0
   If TRIM$(FISECOM.Text6(0)) <> "" Then
      CADEST% = CADEST% + 1
      DESTIX$ = TRIM$(FISECOM.Text6(0))
   End If
   '
   TTXX$ = ""
   If EMPRE$ <> "" Then TTXX$ = TTXX$ + EMPRE$ + Chr$(13) + Chr$(10)
   If CALNUM$ <> "" Then TTXX$ = TTXX$ + CALNUM$ + Chr$(13) + Chr$(10)
   If LOCAX$ <> "" Then TTXX$ = TTXX$ + LOCAX$ + Chr$(13) + Chr$(10)
   If PROPA$ <> "" Then TTXX$ = TTXX$ + PROPA$ + Chr$(13) + Chr$(10)
   '
   MMT4$ = TRIM$(DESTIX$)
   If MMT4$ <> "" Then MMT4$ = MMT4$ + " - "
   MAILMESS.Text4 = TRIM$(MMT4$ + EMPRE$)
   MAILMESS.Text5 = FISECOM.Text5
   MAILMESS.Text20 = "FLEXOFT - "
   '
   HOII% = HOY(HOS$)
   MAILMESS.Label5 = "Buenos Aires, " + FETEXTO$(HOII%) + ".-"
   While MAILMESS.Label5.Width < MAILMESS.Text2.Width - 1500
     MAILMESS.Label5 = " " + MAILMESS.Label5
   Wend
   '
   TTYY$ = MAILMESS.Label5 + Chr$(13) + Chr$(10)
   If CADEST% < 1 Then
       TTYY$ = TTYY$ + "Sres." + Chr$(13) + Chr$(10)
    Else
       TTYY$ = TTYY$ + "Sr./Sra." + Chr$(13) + Chr$(10)
       TTYY$ = TTYY$ + DESTIX$ + Chr$(13) + Chr$(10)
   End If
   TTYY$ = TTYY$ + TTXX$
   SST = 4 + Len(TTYY$)
   '
   For KKII% = 1 To 5: TTYY$ = TTYY$ + Chr$(13) + Chr$(10): Next KKII%
   TTYY$ = TTYY$ + Space$(75) + "Ing. Ernesto P. Bauer" + Chr$(13) + Chr$(10)
   TTYY$ = TTYY$ + "===================================================================" + Chr$(13) + Chr$(10)
   TTYY$ = TTYY$ + "FLEXOFT Soluciones Informáticas Integrales" + Chr$(13) + Chr$(10)
   TTYY$ = TTYY$ + "Bauness 2170 - 5to. 'D' - 1431 Buenos Aires - Argentina" + Chr$(13) + Chr$(10)
   TTYY$ = TTYY$ + "Tel (00-54-11) 4524-1284 / 4541-1515 - e-mail: info@flexoft.com.ar" + Chr$(13) + Chr$(10)
   TTYY$ = TTYY$ + "===================================================================" + Chr$(13) + Chr$(10)
   '
   MAILMESS.List1.Clear
   For KU& = 1 To List1.ListCount
     MAILMESS.List1.AddItem List1.List(KU& - 1)
   Next KU&
   TPU& = MAILMESS.List1.ListCount - 7
   If TPU& < 0 Then TPU& = 0
   MAILMESS.List1.TopIndex = TPU&
   '
   MAILMESS.Text2 = TTYY$ ' = TTYY$ + Text1
   MAILMESS.Text2.SelStart = SST
   MAILMESS.Text2.SelLength = 0
   '
   ATENTI% = 1
   MAILMESS.Label26 = Label26
   MAILMESS.Show 1
   '
End Sub

Private Sub Command35_Click()

'    Call SELECHO("!INDOPORT/INDICE DE OPORTUNIDADES")
'    CODIOPO$ = TRIM$(VALACT1$("!INDOPORT"))
    '
    On Error Resume Next
    MODO% = 1
    A = MENUEXP09.Command1.Caption
    If Err > 0 Then MODO% = 0
    On Error GoTo 0
    CONDI$ = "CodcarpImp <>'' AND STATUS < 9 AND MODO = " & MODO%
    CONDI$ = CONDI$ + "ORDER BY CodCarpImp"
    Call CARGALISEL("!INCARPIM", "CARPEIMP", "CodcarpImp/A12;FechAlta/d8;RefeCarpImp/A48", CONDI$, "NOMESS")
    Call SELECHO("!INCARPIM/INDICE DE CARPETAS DE IMPORTACION\EXPORTACION")
    '
    CODIOPO$ = TRIM$(VALACT1$("!INCARPIM"))
    DESCRI$ = VALACT2$("!INCARPIM")
        If CODIOPO$ <> "" Then
       Label7 = CODIOPO$
       Label8 = DESCRI$
    End If
End Sub

Private Sub Command5_Click()
   '\\\OT-05-0966-OD-27/10/08///
   Command5.Enabled = False
'   HOII% = HOY(HOS$)
'   '
'   COCLI$ = TRIM$(Label26)
'   CODOP$ = TRIM$(Label7)
'   If COCLI$ <> "" Then
'     FEC% = FECHANUM(Text12)
'     If FEC% <= HOY(HO$) Then FEC% = HOY(HO$)
'     FEPRO = FETEXCOR1$(FEC%)
'     CONDI$ = "CODI_CLI = '" & COCLI$ & "'"
'     Call ACTUREGISTRO("BSEGCOM", "FechProxCto", CONDI$, FEPRO, RAFE&, "NOMESS")
'     '
'     If TRIM$(Text11) <> "" Then
'       Call FRATEXTO(Text11, 56)
'       TTXX1$ = ""
'       For KKI% = 1 To CARETEX%
'         TTXX$ = Space$(76)
'         Call REPLA(TTXX$, RETEX$(KKI%), 11, 56)
'         If KKI% = 1 Then
'           Call REPLA(TTXX$, HOS$, 1, 8)
'           Call REPLA(TTXX$, USERTER$, 69, 24)
'         End If
'         TTXX1$ = TTXX1$ + TTXX$ + Chr$(10) + Chr$(13)
'       Next KKI%
'       TTXX1$ = TTXX1$ + String$(88, "=") + Chr$(10) + Chr$(13)
'       '
'       CONDI$ = "CODI_CLI = '" & COCLI$ & "'"
'       ANOTANTE$ = VALOBADA("BSEGCOM", "HIS_CONTAC", CONDI$)
'       Call ACTUREGISTRO("BSEGCOM", "HIS_CONTAC", CONDI$, ANOTANTE$ + TTXX1$, RAFE&)
'       '
'       If CODOP$ <> "" Then
'          CONDI$ = "CODI_CLI = '" & COCLI$ & "' and CodOport = '" & CODOP$ & "'"
'          ANOTANTE$ = VALOBADA("OPORTUNI", "HIS_OPORTU", CONDI$)
'          Call ACTUREGISTRO("OPORTUNI", "HIS_OPORTU", CONDI$, ANOTANTE$ + TTXX1$, RAFE&)
'       End If
'       '
'       Text11 = ""
'       Call CARGANOTA
'       '
'     End If
'   End If
'
     CODTRAN$ = Label7
     If TRIM$(Text11) <> "" Then
        Call FRATEXTO(Text11, 56)
        For KKI% = 1 To CARETEX%
          TTXX$ = Space$(76)
          Call REPLA(TTXX$, RETEX$(KKI%), 11, 56)
          If KKI% = 1 Then
            Call REPLA(TTXX$, HOS$, 1, 8)
            Call REPLA(TTXX$, USERNAME$, 69, 24)
          End If
          Call REPLA(CODTRAN$, ANOTACARPI.Label7, 1, 16)
          Call ANOTA_TRANSACCIONES(CODTRAN$, TTXX$)
        Next KKI%
        Call ANOTA_TRANSACCIONES(CODTRAN$, String$(86, "="))
        Text11 = ""
        '
        List1.Clear
        SQLX$ = "SELECT * FROM CARPEIMP"
        SQLX$ = SQLX$ + " WHERE CodCarpImp = '" & TRIM$(CODTRAN$) & "'"
        '
        Dim ANOTA As New ADODB.Recordset
        Set ANOTA = New ADODB.Recordset
25      On Error Resume Next
        ANOTA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
        End If
        On Error GoTo 0
        '
        With ANOTA
          On Error Resume Next
          .MoveFirst
          If Err < 1 Then
             On Error GoTo 0
             XX$ = SAFETEXT$(!His_Carpi): LONGI& = Len(XX$)
             INICIO& = 1
             Do While INICIO& < LONGI&
                 FIN& = InStr(INICIO&, XX$, Chr$(10))
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
          End If
          On Error GoTo 0
        End With
        ANOTA.Close
        Set ANOTA = Nothing
     End If
   '
   Text11 = ""
   Command5.Enabled = True
   '
End Sub




Sub CARGANOTA()
'   COCLI$ = TRIM$(Label26)
'   CODOP$ = TRIM$(Label7)
'   List1.Clear
'   If COCLI$ <> "" Then
'      If CODOP$ <> "" Then
'           CONDI$ = "CODI_CLI = '" & COCLI$ & "' and CodOport = '" & CODOP$ & "'"
'           ANOTANTE$ = VALOBADA("OPORTUNI", "HIS_OPORTU", CONDI$)
'         Else
'           CONDI$ = "CODI_CLI = '" & COCLI$ & "'"
'           ANOTANTE$ = VALOBADA("BSEGCOM", "HIS_CONTAC", CONDI$)
'      End If
'      '
'      XX$ = ANOTANTE$: LONGI& = Len(XX$)
'      INICIO& = 1
'      Do While INICIO& < LONGI&
'         FIN& = InStr(INICIO&, XX$, Chr$(10))
'         If FIN& > 0 Then
'            TTXX$ = Mid$(XX$, INICIO&, FIN& - INICIO&)
'            List1.AddItem TTXX$
'            INICIO& = FIN& + 2
'          Else
'            TTXX$ = Mid$(XX$, INICIO&)
'            List1.AddItem TTXX$
'            Exit Do
'         End If
'      Loop
'      '
'      TPI = List1.ListCount - 12
'      If TPI < 0 Then TPI = 0
'      List1.TopIndex = TPI
'      '
'   End If
   '
       List1.Clear
    CODTRAN$ = Label7
    If CODTRAN$ <> "" Then
        XX$ = VALOBADA("CARPEIMP", "His_Carpi", "CodCarpImp = '" & CODTRAN$ & "'")
        LONGI& = Len(XX$): INICIO& = 1
        '
        Do While INICIO& < LONGI&
           FIN& = InStr(INICIO&, XX$, Chr$(10))
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
    End If
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
'    Label8 = VALOBADA("OPORTUNI", "RefeOport", "CodOport = '" & Label7 & "'")
    
    FECHA% = CVINT(VALOBADA("CARPEIMP", "FechAlta", "CodCarpImp = '" & Label7 & "'", "NOMESS"))
    Label6 = FECHATEX$(FECHA%)
    Call CARGANOTA
End Sub

Private Sub Label7_DblClick()
    Call Command35_Click
End Sub

Private Sub List1_DblClick()
  '
  Exit Sub
  NPX& = XVALO(FISECOM07.Label26)
  Call TRALOFIL("ANOTASEG", 1, MKS$(NPX&))
  VTB% = VENTABU%("ANOTASEG")
  If VTB% >= 0 Then
    Call TRACENFIL("ANOTASEG", 1, MKS$(NPX&))
  End If
  '
  List1.Clear
  RFF$ = RECFILT$("ANOTASEG", 1, MKS$(NPX&))
  '
  For U& = 1 To Len(RFF$) Step 4
    RENOTA& = CVS(Mid$(RFF$, U&, 4))
    XX$ = REGLEIDO$("ANOTASEG", RENOTA&)
    For KU% = 5 To 92
      If Asc(Mid$(XX$, KU%)) < 32 Then
        Mid$(XX$, KU%, 1) = " "
      End If
    Next KU%
    List1.AddItem RTrim$(Mid$(XX$, 5, 92))
  Next U&
  '
  TPI = List1.ListCount - 12
  If TPI < 0 Then TPI = 0
  List1.TopIndex = TPI
  '
End Sub

Private Sub Text12_DblClick()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text12.TEXT = VDEV$
    End If
End Sub
