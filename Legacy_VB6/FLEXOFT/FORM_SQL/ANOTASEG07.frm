VERSION 5.00
Begin VB.Form ANOTASEG07 
   BackColor       =   &H00EEDDEE&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anotador Seguimiento Comercial"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000012&
      Height          =   675
      Left            =   9120
      ScaleHeight     =   615
      ScaleWidth      =   1395
      TabIndex        =   37
      Top             =   550
      Width           =   1455
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
         Height          =   600
         Left            =   750
         Picture         =   "ANOTASEG07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   0
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
         Height          =   600
         Left            =   50
         Picture         =   "ANOTASEG07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Preparar y Enviar un mail"
         Top             =   0
         Width           =   600
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
      Height          =   1590
      Left            =   45
      TabIndex        =   22
      Top             =   160
      Width           =   10620
      Begin VB.Frame Frame9 
         BackColor       =   &H00EEDDEE&
         Caption         =   "Nro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   730
         Left            =   8030
         TabIndex        =   23
         Top             =   315
         Width           =   960
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
            Left            =   105
            TabIndex        =   24
            Top             =   315
            Width           =   750
         End
      End
      Begin VB.Label Label14 
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
         Left            =   9660
         TabIndex        =   36
         Top             =   1155
         Width           =   750
      End
      Begin VB.Label Label13 
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
         Left            =   4830
         TabIndex        =   35
         Top             =   1155
         Width           =   3585
      End
      Begin VB.Label Label11 
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
         Left            =   1155
         TabIndex        =   34
         Top             =   1155
         Width           =   2115
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
         Left            =   4620
         TabIndex        =   33
         Top             =   735
         Width           =   3270
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
         Left            =   1155
         TabIndex        =   32
         Top             =   735
         Width           =   3270
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
         Height          =   285
         Left            =   1155
         TabIndex        =   31
         Top             =   420
         Width           =   6735
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa:"
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
         TabIndex        =   30
         Top             =   500
         Width           =   960
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contactos:"
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
         TabIndex        =   29
         Top             =   800
         Width           =   960
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Especialidad:"
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
         Left            =   2940
         TabIndex        =   28
         Top             =   1230
         Width           =   1800
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Profesión:"
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
         TabIndex        =   27
         Top             =   1230
         Width           =   960
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Empleados:"
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
         Left            =   8505
         TabIndex        =   26
         Top             =   1230
         Width           =   1065
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Observac:"
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
         TabIndex        =   25
         Top             =   1680
         Width           =   960
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
      Height          =   1950
      Left            =   45
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
         Left            =   9455
         TabIndex        =   10
         Top             =   1470
         Width           =   175
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1175
         Left            =   1155
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   735
         Width           =   7155
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
         Left            =   9810
         Picture         =   "ANOTASEG07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1140
         Width           =   750
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0C0C0&
         Height          =   425
         Left            =   0
         ScaleHeight     =   360
         ScaleWidth      =   10515
         TabIndex        =   12
         Top             =   315
         Width           =   10580
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
            Top             =   75
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
         Left            =   8500
         TabIndex        =   11
         Top             =   1470
         Width           =   1000
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
         Left            =   8400
         TabIndex        =   18
         Top             =   1200
         Width           =   1380
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " "
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
         Left            =   80
         TabIndex        =   17
         Top             =   800
         Width           =   1065
      End
      Begin VB.Line Line8 
         X1              =   0
         X2              =   11665
         Y1              =   1900
         Y2              =   1900
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
         Left            =   8400
         TabIndex        =   21
         Top             =   840
         Width           =   4000
      End
      Begin VB.Line Line7 
         X1              =   10580
         X2              =   10580
         Y1              =   315
         Y2              =   1900
      End
      Begin VB.Line Line6 
         X1              =   8295
         X2              =   10605
         Y1              =   1130
         Y2              =   1130
      End
      Begin VB.Line Line5 
         X1              =   8295
         X2              =   8295
         Y1              =   735
         Y2              =   1920
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
         Height          =   1165
         Left            =   0
         TabIndex        =   20
         Top             =   735
         Width           =   1170
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
         Left            =   8295
         TabIndex        =   19
         Top             =   735
         Width           =   2265
      End
      Begin VB.Line Line3 
         X1              =   8295
         X2              =   10605
         Y1              =   1130
         Y2              =   1130
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
      Left            =   45
      TabIndex        =   2
      Top             =   1830
      Width           =   10620
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   3795
         Left            =   10290
         TabIndex        =   8
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   4005
         Left            =   8085
         TabIndex        =   6
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   4110
         Left            =   1155
         TabIndex        =   5
         Top             =   630
         Width           =   15
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3660
         Left            =   0
         MultiSelect     =   2  'Extended
         TabIndex        =   4
         Top             =   735
         Width           =   10580
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   435
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   10515
         TabIndex        =   3
         Top             =   315
         Width           =   10580
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha     Detalle                                                   Usuario/Terminal"
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
            Top             =   105
            Width           =   10200
         End
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   10605
         Y1              =   4410
         Y2              =   4410
      End
      Begin VB.Line Line1 
         X1              =   10580
         X2              =   10580
         Y1              =   320
         Y2              =   6225
      End
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
Attribute VB_Name = "ANOTASEG07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


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

Private Sub Command5_Click()
   '\\\OT-05-0966-OD-27/10/08///
   Command5.Enabled = False
   HOII% = HOY(HOS$)
   '
   COCLI$ = TRIM$(Label26)
   If COCLI$ <> "" Then
     If TRIM$(Text11) <> "" Then
       Call FRATEXTO(Text11, 56)
       For KKI% = 1 To CARETEX%
         TTXX$ = Space$(76)
         Call REPLA(TTXX$, RETEX$(KKI%), 11, 56)
         If KKI% = 1 Then
           Call REPLA(TTXX$, HOS$, 1, 8)
           Call REPLA(TTXX$, USERTER$, 69, 24)
         End If
         Call REGAPP("ANOTASEG", Left$(MKS$(NPX&) + TTXX$, 96))
          TTXX1$ = TTXX1$ + TTXX$ + Chr$(10) + Chr$(13)
       Next KKI%
          'TTXX1$ = TTXX1$ + String$(88, "=")
       Call REGAPP("ANOTASEG", MKS$(NPX&) + String$(88, "="))
       Call CIERRARCH("ANOTASEG")
       Text11 = ""
       SQLX$ = "SELECT HIS_CONTAC FROM BSEGCOM"
       SQLX$ = SQLX$ + " WHERE CODI_CLI = '" & COCLI$ & "'"
       Dim ANOTMP As New ADODB.Recordset
       Set ANOTMP = New ADODB.Recordset
       ANOTMP.OPEN FILTSQL$ (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       With ANOTMP
          If Not (.BOF And .EOF) Then
             ANOTA$ = SAFETEXT$(!HIS_CONTAC)
             ANOTA$ = ANOTA$ + TTXX1$
             ANOTA$ = ANOTA$ + String$(88, "=") + Chr$(10) + Chr$(13)
             !HIS_CONTAC = ANOTA$
             .Update
           End If
       End With
       ANOTMP.Close
       Set ANOTMP = Nothing
       
      List1.Clear
      
      ANOTACION$ = VALOBADA("BSEGCOM", "HIS_CONTAC", "CODI_CLI = '" & COCLI$ & "'")
      XX$ = ANOTACION$: LONGI& = Len(XX$)
        INICIO& = 1
        Do While INICIO& < LONGI&
            FIN& = InStr(INICIO&, XX$, Chr$(10))
            If FIN& > 0 Then
               TTXX$ = Mid$(XX$, INICIO&, FIN& - INICIO&)
               ANOTASEG07.List1.AddItem TTXX$
               INICIO& = FIN& + 2
            Else
               TTXX$ = Mid$(XX$, INICIO&)
               ANOTASEG07.List1.AddItem TTXX$
               Exit Do
            End If
        Loop

     '
     TPI = ANOTASEG07.List1.ListCount - 12
     If TPI < 0 Then TPI = 0
     ANOTASEG07.List1.TopIndex = TPI

       '
       '
     End If
   End If
   '
   Command5.Enabled = True
   '
   '\\\OT-05-0966-OD-FIN///

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

Private Sub Label34_DblClick()
   Call SELECHO("FECHA")
   If VALACT1$("FECHA") <> "" Then
      Label34 = VALACT1$("FECHA")
   End If
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
